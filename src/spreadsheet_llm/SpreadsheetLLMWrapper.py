import json
import logging
import pandas as pd
import openpyxl
from typing import Optional, List
from pydantic import BaseModel, Field
from langchain_core.language_models.chat_models import BaseChatModel
from langchain_core.messages import HumanMessage
from spreadsheet_llm.IndexColumnConverter import IndexColumnConverter
from spreadsheet_llm.SheetCompressor import SheetCompressor

logger = logging.getLogger(__name__)


class CellRangeItem(BaseModel):
    """Simple schema for cell range with optional title."""

    title: Optional[str] = Field(
        None, description="Optional title or label for this range"
    )
    range: str = Field(description="Cell range address (e.g., 'A1:B5', 'C3')")


class CellRangeList(BaseModel):
    """List of cell ranges with reasoning (Chain of Thought)."""

    reasoning: str = Field(
        description="Step-by-step reasoning process for identifying the cell ranges. "
        "Explain what patterns you noticed, how you identified meaningful ranges, "
        "and why you chose these specific ranges."
    )
    items: List[CellRangeItem] = Field(
        description="List of cell ranges found in the spreadsheet"
    )


RECOGNIZE_PROMPT = """Instruction:
Given an inverted index mapping cell content to their locations in a spreadsheet.

Format:
   Each line consists of the cell content and the cell addresses where it appears, separated by '|'.
   Format: content|cell1,cell2,cell3,...
   Example:
   Eagles|B12,B39,B44,B48,B52,B54
   Purple City|D12,J12,F16,D18,E41,E49
   ${Integer}|G14:G15,H15,G16:I17

   Data formats are prefixed with '${}', like '${Integer}' or '${yyyy/mm/dd}'.
   Cells are separated by commas.

Your task is to:
1. First, provide step-by-step reasoning about what patterns you observe in the data
2. Explain how you identify meaningful cell ranges
3. Then, return the identified cell ranges with optional descriptive titles

Use Chain of Thought reasoning to ensure accurate identification.
"""


class SpreadsheetLLMWrapper:

    def __init__(self, format_aware: bool = False, llm: Optional[BaseChatModel] = None):
        """
        Initialize SpreadsheetLLMWrapper.

        Args:
            format_aware: If True, enables format-aware aggregation in dict output.
                         Groups cells by both value AND data type (e.g., "100 (Integer)").
                         If False (default), groups cells only by value.
            llm: LangChain ChatModel instance (e.g., ChatOpenAI, ChatAnthropic, etc.).
                 If None, a default ChatOpenAI model will be created.
        """
        self.format_aware = format_aware
        self.llm = llm

    def read_spreadsheet(self, file):

        wb = openpyxl.load_workbook(file)
        return wb

    # Takes a file, compresses it
    def compress_spreadsheet(self, wb):
        sheet_compressor = SheetCompressor(format_aware=self.format_aware)
        # header=None: treat all rows as data, don't auto-generate column names
        sheet = pd.read_excel(wb, engine="openpyxl", header=None)
        # sheet = sheet.apply(
        #     lambda x: x.str.replace("\n", "<br>") if x.dtype == "object" else x
        # )

        # Reset index and column names to integers
        sheet = sheet.reset_index(drop=True)
        sheet.columns = list(range(len(sheet.columns)))

        logger.info(f"Original sheet shape: {sheet.shape} (rows x cols)")
        logger.debug(f"First 5 rows:\n{sheet.head()}")

        # Structural-anchor-based Extraction
        sheet = sheet_compressor.anchor(sheet)
        logger.info(f"After anchor, sheet shape: {sheet.shape} (rows x cols)")

        # Encoding
        markdown = sheet_compressor.encode(
            wb, sheet
        )  # Paper encodes first then anchors; I chose to do this in reverse
        logger.info(f"Encoded markdown shape: {markdown.shape}")
        logger.debug(f"Markdown columns: {markdown.columns.tolist()}")
        logger.debug(f"First 10 markdown entries:\n{markdown.head(10)}")

        # Data-Format Aggregation
        markdown["Category"] = markdown["Value"].apply(
            lambda x: sheet_compressor.get_category(x)
        )
        category_dict = sheet_compressor.inverted_category(markdown)
        logger.info(f"Categories found: {set(category_dict.values())}")
        logger.info(f"Number of unique values: {len(category_dict)}")

        try:
            areas = sheet_compressor.identical_cell_aggregation(sheet, category_dict)
            logger.info(f"Number of areas identified: {len(areas)}")
            logger.debug("First 10 areas:")
            for i, area in enumerate(areas[:10]):
                logger.debug(f"  Area {i}: {area}")
        except RecursionError:
            logger.error("RecursionError in identical_cell_aggregation")
            return

        # Inverted-index Translation
        compress_dict = sheet_compressor.inverted_index(markdown)
        logger.info(f"Compress dict entries: {len(compress_dict)}")

        return areas, compress_dict, sheet_compressor

    def serialize_areas(self, areas, sheet_compressor):
        """Serialize areas to string representation.

        Args:
            areas: List of area tuples
            sheet_compressor: SheetCompressor instance with row/column mappings

        Returns:
            Serialized string representation of areas
        """
        string = ""
        converter = IndexColumnConverter()
        for i in areas:
            # Map compressed indices back to original indices
            original_row_start = sheet_compressor.row_mapping[i[0][0]]
            original_col_start = sheet_compressor.column_mapping[i[0][1]]
            original_row_end = sheet_compressor.row_mapping[i[1][0]]
            original_col_end = sheet_compressor.column_mapping[i[1][1]]

            string += (
                "("
                + i[2]
                + "|"
                + converter.parse_colindex(original_col_start + 1)
                + str(original_row_start + 1)
                + ":"
                + converter.parse_colindex(original_col_end + 1)
                + str(original_row_end + 1)
                + "), "
            )
        return string

    def write_areas(self, file, areas, sheet_compressor):
        """Write serialized areas to file."""
        string = self.serialize_areas(areas, sheet_compressor)
        with open(file, "w+", encoding="utf-8") as f:
            f.writelines(string)

    def serialize_dict(self, dict):
        """Serialize dictionary to string representation.

        Args:
            dict: Dictionary to serialize

        Returns:
            Serialized string representation of dictionary
        """
        string = ""
        for key, value in dict.items():
            # Skip empty keys
            if not key or str(key).strip() == "":
                continue
            # value is now list[str], join with comma
            value_str = ",".join(value) if isinstance(value, list) else str(value)
            string += str(key) + "|" + value_str + "\n"
        return string

    def write_dict(self, file, dict):
        """Write serialized dictionary to file."""
        string = self.serialize_dict(dict)
        with open(file, "w+", encoding="utf-8") as f:
            f.writelines(string)

    def serialize_mapping(self, sheet_compressor):
        """Serialize coordinate mapping to JSON string.

        Args:
            sheet_compressor: SheetCompressor instance with row/column mappings

        Returns:
            JSON string representation of coordinate mapping
        """
        mapping = sheet_compressor.get_coordinate_mapping()
        return json.dumps(mapping, indent=2, ensure_ascii=False)

    def write_mapping(self, file, sheet_compressor):
        """Write coordinate mapping from compressed to original coordinates as JSON"""
        mapping_str = self.serialize_mapping(sheet_compressor)
        with open(file, "w+", encoding="utf-8") as f:
            f.write(mapping_str)

    def convert_compressed_to_original(
        self, compressed_coord: str, sheet_compressor
    ) -> str:
        """
        Convert compressed coordinate(s) to original coordinate(s).

        Args:
            compressed_coord: Compressed coordinate string, can be:
                - Single cell: "A1", "B5"
                - Range: "A1:B5", "C3:D10"
                - Multiple ranges: "A1,B2:B5,C3"
            sheet_compressor: SheetCompressor instance with row/column mappings

        Returns:
            Original coordinate string in the same format as input

        Examples:
            "A1" -> "A1" (if A1 maps to A1)
            "B5" -> "C10" (if compressed B5 maps to original C10)
            "A1:B3" -> "A1:C5" (range conversion)
            "A1,B2:B5" -> "A1,C3:C8" (multiple ranges)
        """
        converter = IndexColumnConverter()

        def convert_single_cell(cell: str) -> str:
            """Convert a single cell coordinate"""
            # Parse cell coordinate (e.g., "A1" -> col=0, row=0)
            match = converter.parse_cell(cell)
            if not match:
                logger.warning(f"Invalid cell format: {cell}")
                return cell

            col_str, row_str = match
            compressed_col = (
                converter.parse_cellindex(col_str) - 1
            )  # Convert to 0-based
            compressed_row = int(row_str) - 1  # Convert to 0-based

            # Map to original indices
            original_row = sheet_compressor.row_mapping.get(
                compressed_row, compressed_row
            )
            original_col = sheet_compressor.column_mapping.get(
                compressed_col, compressed_col
            )

            # Convert back to cell notation
            original_col_str = converter.parse_colindex(original_col + 1)
            original_row_str = str(original_row + 1)

            return f"{original_col_str}{original_row_str}"

        def convert_range(range_str: str) -> str:
            """Convert a cell range (e.g., 'A1:B5')"""
            if ":" in range_str:
                start, end = range_str.split(":")
                return f"{convert_single_cell(start)}:{convert_single_cell(end)}"
            else:
                return convert_single_cell(range_str)

        # Handle multiple ranges separated by commas
        parts = compressed_coord.split(",")
        converted_parts = [convert_range(part.strip()) for part in parts]

        return ",".join(converted_parts)

    def recognize(self, compress_dict, user_prompt: str | None = None) -> CellRangeList:
        """
        Use LLM to extract structured cell ranges from spreadsheet content with Chain of Thought reasoning.

        Returns compressed coordinates. Use recognize_original() to get original coordinates.

        Returns a CellRangeList Pydantic model with:
        - reasoning: Step-by-step Chain of Thought explanation
        - items: List of CellRangeItem objects (each with 'title' and 'range')

        Args:
            compress_dict: Inverted index dictionary from compression
            user_prompt: Optional user instruction. If None, extracts all meaningful ranges.

        Returns:
            CellRangeList Pydantic model with compressed coordinates

        Example:
            >>> from langchain_openai import ChatOpenAI
            >>> llm = ChatOpenAI(model="gpt-4o-mini")
            >>> wrapper = SpreadsheetLLMWrapper(llm=llm)
            >>> wb = wrapper.read_spreadsheet("data.xlsx")
            >>> areas, compress_dict, compressor = wrapper.compress_spreadsheet(wb)
            >>> result = wrapper.recognize(compress_dict)
            >>> print(result.reasoning)  # See the LLM's thought process
            >>> for item in result.items:
            ...     print(f"{item.title or 'N/A'}: {item.range}")
        """
        # Check if LLM is initialized
        if self.llm is None:
            raise ValueError(
                "LLM is not initialized. Please provide an LLM instance when creating "
                "SpreadsheetLLMWrapper, e.g., SpreadsheetLLMWrapper(llm=ChatOpenAI(model='gpt-4o-mini'))"
            )

        # Serialize the compressed data
        dict_str = self.serialize_dict(compress_dict)

        # Build the user message with instruction and input
        user_message = RECOGNIZE_PROMPT

        if user_prompt:
            user_message += f"\nAdditional instructions: {user_prompt}\n"

        user_message += "\nINPUT:\n"
        user_message += dict_str

        logger.info(
            f"Sending structured recognition request to LLM: {self.llm.__class__.__name__}"
        )
        logger.info("Output schema: CellRangeList")
        logger.debug(f"User message length: {len(user_message)} chars")

        try:
            # Create a structured output version of the LLM
            structured_llm = self.llm.with_structured_output(CellRangeList)

            # Call LLM using LangChain (only user message, no system message)
            messages = [
                HumanMessage(content=user_message),
            ]

            response = structured_llm.invoke(messages)
            if isinstance(response, dict):
                response = CellRangeList(**response)
            elif not isinstance(response, CellRangeList):
                response = CellRangeList(**response.model_dump())

            logger.info(
                f"Structured recognition completed. Found {len(response.items)} ranges."
            )
            logger.debug(f"Reasoning: {response.reasoning}")
            return response

        except Exception as e:
            logger.error(f"LLM structured output error: {str(e)}")
            raise

    def recognize_original(
        self, compress_dict, sheet_compressor, user_prompt: str | None = None
    ) -> CellRangeList:
        """
        Use LLM to extract structured cell ranges and convert to original spreadsheet coordinates.

        This method wraps recognize() and automatically converts compressed coordinates
        to original spreadsheet coordinates using the sheet_compressor mappings.

        Args:
            compress_dict: Inverted index dictionary from compression
            sheet_compressor: SheetCompressor instance with row/column mappings
            user_prompt: Optional user instruction. If None, extracts all meaningful ranges.

        Returns:
            CellRangeList Pydantic model with original coordinates

        Example:
            >>> from langchain_openai import ChatOpenAI
            >>> llm = ChatOpenAI(model="gpt-4o-mini")
            >>> wrapper = SpreadsheetLLMWrapper(llm=llm)
            >>> wb = wrapper.read_spreadsheet("data.xlsx")
            >>> areas, compress_dict, compressor = wrapper.compress_spreadsheet(wb)
            >>> result = wrapper.recognize_original(compress_dict, compressor)
            >>> print(result.reasoning)
            >>> for item in result.items:
            ...     print(f"{item.title or 'N/A'}: {item.range}")  # Original coordinates
        """
        # Call recognize to get compressed coordinates
        compressed_result = self.recognize(compress_dict, user_prompt)

        # Convert all ranges to original coordinates
        logger.info("Converting compressed coordinates to original coordinates")
        converted_items = []
        for item in compressed_result.items:
            original_range = self.convert_compressed_to_original(
                item.range, sheet_compressor
            )
            converted_items.append(
                CellRangeItem(title=item.title, range=original_range)
            )

        logger.debug(f"Converted {len(converted_items)} ranges to original coordinates")

        # Return new CellRangeList with original coordinates
        return CellRangeList(
            reasoning=compressed_result.reasoning, items=converted_items
        )
