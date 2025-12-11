"""
BANTMS Category Map Sheet Generator
====================================

Generates the Category_Map reference sheet for:
- Auto-categorization of bank transactions
- Pattern matching rules
- Cash Flow section mappings
- Line item assignments

This sheet drives the automatic categorization engine
for bank statement transactions.
"""

from datetime import date
from typing import List, Dict, Optional

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from openpyxl.styles import (
    Font, PatternFill, Border, Side, Alignment
)

from ..config.constants import CashFlowSection, CashFlowDirection
from ..config.categories import (
    CategoryDefinition,
    CFSection,
    CFDirection,
    get_all_categories,
    CASH_FLOW_STRUCTURE
)
from ..config.styles import (
    ColorPalette, Typography, NumberFormats,
    ColumnWidths, RowHeights
)


# =============================================================================
# SECTION 1: COLUMN DEFINITIONS
# =============================================================================

CATEGORY_MAP_COLUMNS: List[Dict] = [
    {
        "header": "Pattern_ID",
        "width": 12,
        "format": "@",
        "alignment": "center",
        "description": "Unique Pattern Identifier"
    },
    {
        "header": "Description_Pattern",
        "width": 30,
        "format": "@",
        "alignment": "left",
        "description": "Text pattern to search for (wildcards supported)"
    },
    {
        "header": "Category_Code",
        "width": 15,
        "format": "@",
        "alignment": "center",
        "description": "Unique category code"
    },
    {
        "header": "Category_Name",
        "width": 25,
        "format": "@",
        "alignment": "left",
        "description": "Human-readable category name"
    },
    {
        "header": "CF_Section",
        "width": 14,
        "format": "@",
        "alignment": "center",
        "description": "Cash Flow section (Operating/Investing/Financing)"
    },
    {
        "header": "CF_Line_Item",
        "width": 30,
        "format": "@",
        "alignment": "left",
        "description": "Cash Flow Statement line item"
    },
    {
        "header": "Sign_Direction",
        "width": 12,
        "format": "@",
        "alignment": "center",
        "description": "Inflow or Outflow"
    },
    {
        "header": "Priority",
        "width": 10,
        "format": "0",
        "alignment": "center",
        "description": "Pattern matching priority (lower = higher priority)"
    },
    {
        "header": "Is_Active",
        "width": 10,
        "format": "@",
        "alignment": "center",
        "description": "Whether pattern is active (TRUE/FALSE)"
    },
    {
        "header": "Notes",
        "width": 35,
        "format": "@",
        "alignment": "left",
        "description": "Additional notes or examples"
    },
]


# =============================================================================
# SECTION 2: STYLE DEFINITIONS
# =============================================================================

def create_section_header_style(section: str) -> Dict:
    """Create section header style based on CF section."""
    section_colors = {
        "Operating": ColorPalette.SECTION_OPERATING,
        "Investing": ColorPalette.SECTION_INVESTING,
        "Financing": ColorPalette.SECTION_FINANCING,
    }

    return {
        "font": Font(
            name=Typography.FONT_FAMILY,
            size=Typography.SIZE_BODY,
            bold=True,
            color=ColorPalette.HEADER_DARK_NAVY
        ),
        "fill": PatternFill(
            start_color=section_colors.get(section, ColorPalette.ROW_LIGHT_GRAY),
            end_color=section_colors.get(section, ColorPalette.ROW_LIGHT_GRAY),
            fill_type="solid"
        ),
        "border": Border(
            left=Side(style="thin", color=ColorPalette.BORDER_THIN_GRAY),
            right=Side(style="thin", color=ColorPalette.BORDER_THIN_GRAY),
            top=Side(style="medium", color=ColorPalette.HEADER_MEDIUM_BLUE),
            bottom=Side(style="thin", color=ColorPalette.BORDER_THIN_GRAY)
        )
    }


# =============================================================================
# SECTION 3: CATEGORY MAP GENERATOR
# =============================================================================

class CategoryMapSheetGenerator:
    """
    Generator for Category_Map reference sheet.

    Creates a formatted Excel sheet with:
    - Pattern matching rules for auto-categorization
    - Cash Flow section assignments
    - Line item mappings
    - Priority ordering
    """

    SHEET_NAME = "Category_Map"
    TABLE_NAME = "tbl_Category_Map"
    TITLE_ROW = 1
    HEADER_ROW = 3
    DATA_START_ROW = 4

    def __init__(self, workbook: Workbook):
        """
        Initialize the Category Map sheet generator.

        Args:
            workbook: The openpyxl Workbook to add the sheet to.
        """
        self.workbook = workbook
        self.sheet: Optional[Worksheet] = None
        self.categories: List[CategoryDefinition] = []

    def generate(self) -> Worksheet:
        """
        Generate the Category_Map sheet.

        Returns:
            The generated Worksheet.
        """
        self.categories = get_all_categories()

        # Create or get sheet
        if self.SHEET_NAME in self.workbook.sheetnames:
            self.sheet = self.workbook[self.SHEET_NAME]
        else:
            self.sheet = self.workbook.create_sheet(self.SHEET_NAME)

        # Build sheet components
        self._write_title()
        self._write_headers()
        self._write_data()
        self._apply_formatting()
        self._create_table()
        self._set_column_widths()
        self._create_named_ranges()
        self._add_data_validation()
        self._add_instructions()

        return self.sheet

    def _write_title(self):
        """Write sheet title and description."""
        # Title
        self.sheet.cell(
            row=self.TITLE_ROW,
            column=1,
            value="CATEGORY MAP - Transaction Auto-Categorization Rules"
        )
        title_cell = self.sheet.cell(row=self.TITLE_ROW, column=1)
        title_cell.font = Font(
            name=Typography.FONT_FAMILY,
            size=Typography.SIZE_TITLE,
            bold=True,
            color=ColorPalette.HEADER_DARK_NAVY
        )

        # Merge title cells
        self.sheet.merge_cells(
            start_row=self.TITLE_ROW,
            start_column=1,
            end_row=self.TITLE_ROW,
            end_column=len(CATEGORY_MAP_COLUMNS)
        )

        # Subtitle
        self.sheet.cell(
            row=self.TITLE_ROW + 1,
            column=1,
            value="Pattern matching rules for auto-categorizing bank statement transactions"
        )
        subtitle_cell = self.sheet.cell(row=self.TITLE_ROW + 1, column=1)
        subtitle_cell.font = Font(
            name=Typography.FONT_FAMILY,
            size=Typography.SIZE_SMALL,
            italic=True,
            color=ColorPalette.TEXT_DARK_GRAY
        )

    def _write_headers(self):
        """Write column headers."""
        header_style = {
            "font": Font(
                name=Typography.FONT_FAMILY,
                size=Typography.SIZE_HEADER,
                bold=True,
                color=ColorPalette.TEXT_WHITE
            ),
            "fill": PatternFill(
                start_color=ColorPalette.HEADER_DARK_NAVY,
                end_color=ColorPalette.HEADER_DARK_NAVY,
                fill_type="solid"
            ),
            "alignment": Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            ),
            "border": Border(
                bottom=Side(style="medium", color=ColorPalette.BORDER_MEDIUM_BLACK)
            )
        }

        for col_idx, col_def in enumerate(CATEGORY_MAP_COLUMNS, start=1):
            cell = self.sheet.cell(
                row=self.HEADER_ROW,
                column=col_idx,
                value=col_def["header"]
            )

            cell.font = header_style["font"]
            cell.fill = header_style["fill"]
            cell.alignment = header_style["alignment"]
            cell.border = header_style["border"]

    def _write_data(self):
        """Write category pattern data rows."""
        row_idx = self.DATA_START_ROW
        pattern_id = 1

        # Iterate through categories and their patterns
        for category in self.categories:
            # Skip if no patterns (like UNCATEGORIZED)
            if not category.description_patterns:
                continue

            # Get section-based styling
            section_name = category.section.value
            is_first_of_section = self._is_first_of_section(category)

            # Write each pattern as a separate row
            for pattern in category.description_patterns:
                self._write_pattern_row(
                    row=row_idx,
                    pattern_id=pattern_id,
                    pattern=pattern,
                    category=category,
                    is_section_start=is_first_of_section and pattern == category.description_patterns[0]
                )
                row_idx += 1
                pattern_id += 1

        # Add UNCATEGORIZED as the last row
        self._write_uncategorized_row(row_idx, pattern_id)

    def _is_first_of_section(self, category: CategoryDefinition) -> bool:
        """Check if this is the first category of its section."""
        for cat in self.categories:
            if cat.section == category.section:
                return cat == category
        return False

    def _write_pattern_row(
        self,
        row: int,
        pattern_id: int,
        pattern: str,
        category: CategoryDefinition,
        is_section_start: bool
    ):
        """Write a single pattern row."""
        section_name = category.section.value

        # Determine row style
        if is_section_start:
            style = create_section_header_style(section_name)
        else:
            is_alternate = (row - self.DATA_START_ROW) % 2 == 1
            fill_color = ColorPalette.ROW_LIGHT_GRAY if is_alternate else ColorPalette.ROW_WHITE
            style = {
                "font": Font(
                    name=Typography.FONT_FAMILY,
                    size=Typography.SIZE_BODY,
                    color=ColorPalette.TEXT_BLACK
                ),
                "fill": PatternFill(
                    start_color=fill_color,
                    end_color=fill_color,
                    fill_type="solid"
                ),
                "border": Border(
                    left=Side(style="thin", color=ColorPalette.BORDER_THIN_GRAY),
                    right=Side(style="thin", color=ColorPalette.BORDER_THIN_GRAY),
                    top=Side(style="thin", color=ColorPalette.BORDER_THIN_GRAY),
                    bottom=Side(style="thin", color=ColorPalette.BORDER_THIN_GRAY)
                )
            }

        # Write cells
        data = [
            (f"PAT-{pattern_id:04d}", "center"),    # Pattern_ID
            (pattern, "left"),                       # Description_Pattern
            (category.code, "center"),               # Category_Code
            (category.name, "left"),                 # Category_Name
            (section_name, "center"),                # CF_Section
            (category.line_item, "left"),            # CF_Line_Item
            (category.direction.value, "center"),    # Sign_Direction
            (category.priority, "center"),           # Priority
            ("TRUE", "center"),                      # Is_Active
            ("", "left"),                            # Notes
        ]

        for col_idx, (value, h_align) in enumerate(data, start=1):
            cell = self.sheet.cell(row=row, column=col_idx, value=value)
            cell.font = style["font"]
            cell.fill = style["fill"]
            cell.border = style["border"]
            cell.alignment = Alignment(horizontal=h_align, vertical="center")

    def _write_uncategorized_row(self, row: int, pattern_id: int):
        """Write the UNCATEGORIZED catch-all row."""
        style = {
            "font": Font(
                name=Typography.FONT_FAMILY,
                size=Typography.SIZE_BODY,
                bold=True,
                color=ColorPalette.STATUS_TEXT_PENDING
            ),
            "fill": PatternFill(
                start_color=ColorPalette.STATUS_PENDING_AMBER,
                end_color=ColorPalette.STATUS_PENDING_AMBER,
                fill_type="solid"
            ),
            "border": Border(
                left=Side(style="thin", color=ColorPalette.BORDER_THIN_GRAY),
                right=Side(style="thin", color=ColorPalette.BORDER_THIN_GRAY),
                top=Side(style="medium", color=ColorPalette.STATUS_TEXT_PENDING),
                bottom=Side(style="medium", color=ColorPalette.STATUS_TEXT_PENDING)
            )
        }

        data = [
            (f"PAT-{pattern_id:04d}", "center"),
            ("*", "left"),  # Catch-all pattern
            ("UNCATEGORIZED", "center"),
            ("Uncategorized", "left"),
            ("Operating", "center"),
            ("Other Operating Items", "left"),
            ("Outflow", "center"),
            (999, "center"),
            ("TRUE", "center"),
            ("Catch-all for unmatched transactions - REVIEW REQUIRED", "left"),
        ]

        for col_idx, (value, h_align) in enumerate(data, start=1):
            cell = self.sheet.cell(row=row, column=col_idx, value=value)
            cell.font = style["font"]
            cell.fill = style["fill"]
            cell.border = style["border"]
            cell.alignment = Alignment(horizontal=h_align, vertical="center")

        self.last_data_row = row

    def _apply_formatting(self):
        """Apply number formats to data columns."""
        # Priority column formatting
        priority_col = 8
        for row in range(self.DATA_START_ROW, self.last_data_row + 1):
            cell = self.sheet.cell(row=row, column=priority_col)
            cell.number_format = "0"

    def _create_table(self):
        """Create Excel table for category map data."""
        start_col = "A"
        end_col = get_column_letter(len(CATEGORY_MAP_COLUMNS))

        table_range = f"{start_col}{self.HEADER_ROW}:{end_col}{self.last_data_row}"

        table = Table(
            displayName=self.TABLE_NAME,
            ref=table_range
        )

        table_style = TableStyleInfo(
            name="TableStyleMedium9",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
        )
        table.tableStyleInfo = table_style

        self.sheet.add_table(table)

    def _set_column_widths(self):
        """Set column widths based on definitions."""
        for col_idx, col_def in enumerate(CATEGORY_MAP_COLUMNS, start=1):
            col_letter = get_column_letter(col_idx)
            self.sheet.column_dimensions[col_letter].width = col_def["width"]

        # Set row heights
        self.sheet.row_dimensions[self.TITLE_ROW].height = RowHeights.TITLE
        self.sheet.row_dimensions[self.HEADER_ROW].height = RowHeights.HEADER

    def _create_named_ranges(self):
        """Create named ranges for formula references."""
        # Named range for category codes
        cat_code_range = f"'{self.SHEET_NAME}'!$C${self.DATA_START_ROW}:$C${self.last_data_row}"
        self.workbook.defined_names.add(
            self.workbook.defined_names.new(name="Category_Codes", value=cat_code_range)
        )

        # Named range for pattern lookup
        pattern_range = f"'{self.SHEET_NAME}'!$B${self.DATA_START_ROW}:$C${self.last_data_row}"
        self.workbook.defined_names.add(
            self.workbook.defined_names.new(name="Pattern_Lookup", value=pattern_range)
        )

        # Named range for CF line items
        line_item_range = f"'{self.SHEET_NAME}'!$C${self.DATA_START_ROW}:$F${self.last_data_row}"
        self.workbook.defined_names.add(
            self.workbook.defined_names.new(name="CF_Line_Item_Lookup", value=line_item_range)
        )

    def _add_data_validation(self):
        """Add data validation dropdowns."""
        # CF_Section dropdown
        section_validation = DataValidation(
            type="list",
            formula1='"Operating,Investing,Financing"',
            allow_blank=False
        )
        section_col = get_column_letter(5)
        section_validation.add(f"{section_col}{self.DATA_START_ROW}:{section_col}1000")
        self.sheet.add_data_validation(section_validation)

        # Sign_Direction dropdown
        direction_validation = DataValidation(
            type="list",
            formula1='"Inflow,Outflow"',
            allow_blank=False
        )
        direction_col = get_column_letter(7)
        direction_validation.add(f"{direction_col}{self.DATA_START_ROW}:{direction_col}1000")
        self.sheet.add_data_validation(direction_validation)

        # Is_Active dropdown
        active_validation = DataValidation(
            type="list",
            formula1='"TRUE,FALSE"',
            allow_blank=False
        )
        active_col = get_column_letter(9)
        active_validation.add(f"{active_col}{self.DATA_START_ROW}:{active_col}1000")
        self.sheet.add_data_validation(active_validation)

    def _add_instructions(self):
        """Add usage instructions at the bottom."""
        instructions_row = self.last_data_row + 3

        instructions = [
            "",
            "AUTO-CATEGORIZATION FORMULA (use in Bank_Statement_Actual):",
            '=IFERROR(INDEX(tbl_Category_Map[Category_Code],AGGREGATE(15,6,ROW(tbl_Category_Map[Description_Pattern])-MIN(ROW(tbl_Category_Map[Description_Pattern]))+1/(ISNUMBER(SEARCH(tbl_Category_Map[Description_Pattern],[@Description]))*(tbl_Category_Map[Is_Active]="TRUE")),1)),"UNCATEGORIZED")',
            "",
            "PATTERN SYNTAX:",
            "• Use * as wildcard (e.g., *ARAMCO* matches 'PAYMENT TO ARAMCO CO')",
            "• Patterns are case-insensitive",
            "• Lower priority number = higher matching priority",
            "• Set Is_Active to FALSE to disable a pattern without deleting",
            "",
            "ADDING NEW PATTERNS:",
            "• Add new rows to the table with unique Pattern_ID",
            "• Ensure Category_Code matches existing category or create new one",
            "• Test pattern matching before activating",
        ]

        for idx, text in enumerate(instructions):
            cell = self.sheet.cell(row=instructions_row + idx, column=1, value=text)
            if text.startswith("AUTO-CATEGORIZATION") or text.startswith("PATTERN SYNTAX") or text.startswith("ADDING"):
                cell.font = Font(
                    name=Typography.FONT_FAMILY,
                    size=Typography.SIZE_BODY,
                    bold=True,
                    color=ColorPalette.HEADER_DARK_NAVY
                )
            elif text.startswith("="):
                cell.font = Font(
                    name="Consolas",
                    size=Typography.SIZE_SMALL,
                    color=ColorPalette.VALUE_POSITIVE_GREEN
                )
            else:
                cell.font = Font(
                    name=Typography.FONT_FAMILY,
                    size=Typography.SIZE_SMALL,
                    color=ColorPalette.TEXT_DARK_GRAY
                )


# =============================================================================
# SECTION 4: UTILITY FUNCTIONS
# =============================================================================

def generate_category_map_sheet(workbook: Workbook) -> Worksheet:
    """
    Convenience function to generate Category_Map sheet.

    Args:
        workbook: The openpyxl Workbook.

    Returns:
        The generated Worksheet.
    """
    generator = CategoryMapSheetGenerator(workbook)
    return generator.generate()


def get_categorization_formula() -> str:
    """
    Get the auto-categorization formula for use in Bank_Statement_Actual.

    Returns:
        Excel formula string for pattern-based categorization.
    """
    return (
        '=IFERROR('
        'INDEX(tbl_Category_Map[Category_Code],'
        'AGGREGATE(15,6,'
        'ROW(tbl_Category_Map[Description_Pattern])'
        '-MIN(ROW(tbl_Category_Map[Description_Pattern]))+1/'
        '(ISNUMBER(SEARCH(tbl_Category_Map[Description_Pattern],[@Description]))'
        '*(tbl_Category_Map[Is_Active]="TRUE")),'
        '1)),'
        '"UNCATEGORIZED")'
    )


def get_line_item_lookup_formula() -> str:
    """
    Get the line item lookup formula based on category code.

    Returns:
        Excel formula string for looking up CF_Line_Item.
    """
    return (
        '=IFERROR('
        'INDEX(tbl_Category_Map[CF_Line_Item],'
        'MATCH([@Category_Code],tbl_Category_Map[Category_Code],0)),'
        '"Other")'
    )
