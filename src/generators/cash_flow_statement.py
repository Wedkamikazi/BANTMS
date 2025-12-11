"""
BANTMS Cash Flow Statement Sheet Generator
===========================================

Generates the Cash_Flow_Statement summary sheet with:
- Operating Activities section
- Investing Activities section
- Financing Activities section
- Actual vs Budget/Forecast comparison
- Variance analysis
- Period-based aggregation

All amounts in SAR (base currency).
Pulls data from Bank_Statement_Actual via SUMPRODUCT formulas.
"""

from datetime import date, datetime
from typing import List, Dict, Optional

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter
from openpyxl.styles import (
    Font, PatternFill, Border, Side, Alignment
)

from ..config.categories import CASH_FLOW_STRUCTURE, CFSection
from ..config.styles import (
    ColorPalette, Typography, NumberFormats,
    ColumnWidths, RowHeights
)


# =============================================================================
# SECTION 1: COLUMN DEFINITIONS
# =============================================================================

CF_STATEMENT_COLUMNS: List[Dict] = [
    {
        "header": "Line Item",
        "width": 40,
        "alignment": "left"
    },
    {
        "header": "Category_Code",
        "width": 0,  # Hidden column
        "alignment": "center"
    },
    {
        "header": "Actual (SAR)",
        "width": 18,
        "alignment": "right"
    },
    {
        "header": "Budget (SAR)",
        "width": 18,
        "alignment": "right"
    },
    {
        "header": "Variance",
        "width": 16,
        "alignment": "right"
    },
    {
        "header": "Variance %",
        "width": 12,
        "alignment": "right"
    },
]


# =============================================================================
# SECTION 2: LINE ITEM DEFINITIONS
# =============================================================================

# Operating Activities line items
OPERATING_LINE_ITEMS: List[Dict] = [
    {
        "name": "Collections from Customers",
        "code": "AR-COLLECT",
        "type": "inflow",
        "actual_formula": '=IFERROR(SUMPRODUCT((tbl_Bank_Statement[CF_Line_Item]="Collections from Customers")*(tbl_Bank_Statement[Credit])),0)',
        "budget_formula": '=IFERROR(SUMIF(tbl_AR_Master[Status],"Expected",tbl_AR_Master[Weighted_Amount_SAR]),0)'
    },
    {
        "name": "Payments to Suppliers",
        "code": "AP-SUPPLIER",
        "type": "outflow",
        "actual_formula": '=IFERROR(-SUMPRODUCT((tbl_Bank_Statement[CF_Line_Item]="Payments to Suppliers")*(tbl_Bank_Statement[Debit])),0)',
        "budget_formula": '=IFERROR(-SUMIFS(tbl_AP_Master[Amount_SAR],tbl_AP_Master[Status],"Pending",tbl_AP_Master[Payment_Type],"Supplier"),0)'
    },
    {
        "name": "Government Payments",
        "code": "GOV-*",
        "type": "outflow",
        "actual_formula": '=IFERROR(-SUMPRODUCT((tbl_Bank_Statement[CF_Line_Item]="Government Payments")*(tbl_Bank_Statement[Debit])),0)',
        "budget_formula": '=IFERROR(-SUMIFS(tbl_AP_Master[Amount_SAR],tbl_AP_Master[Status],"Pending",tbl_AP_Master[Payment_Type],"Government"),0)'
    },
    {
        "name": "Payroll Disbursements",
        "code": "HR-*",
        "type": "outflow",
        "actual_formula": '=IFERROR(-SUMPRODUCT((tbl_Bank_Statement[CF_Line_Item]="Payroll Disbursements")*(tbl_Bank_Statement[Debit])),0)',
        "budget_formula": '=0'  # Could link to Payroll_Schedule
    },
    {
        "name": "Utilities & Rent",
        "code": "UTIL-*",
        "type": "outflow",
        "actual_formula": '=IFERROR(-SUMPRODUCT((tbl_Bank_Statement[CF_Line_Item]="Utilities & Rent")*(tbl_Bank_Statement[Debit])),0)',
        "budget_formula": '=IFERROR(-SUMIFS(tbl_AP_Master[Amount_SAR],tbl_AP_Master[Status],"Pending",tbl_AP_Master[Payment_Type],"Utilities"),0)'
    },
    {
        "name": "Other Operating Expenses",
        "code": "OPEX-*",
        "type": "outflow",
        "actual_formula": '=IFERROR(-SUMPRODUCT((tbl_Bank_Statement[CF_Line_Item]="Other Operating Expenses")*(tbl_Bank_Statement[Debit])),0)',
        "budget_formula": '=0'
    },
]

# Investing Activities line items
INVESTING_LINE_ITEMS: List[Dict] = [
    {
        "name": "Time Deposit Placements",
        "code": "TD-PLACE",
        "type": "outflow",
        "actual_formula": '=IFERROR(-SUMPRODUCT((tbl_Bank_Statement[CF_Line_Item]="Time Deposit Placements")*(tbl_Bank_Statement[Debit])),0)',
        "budget_formula": '=0'
    },
    {
        "name": "Time Deposit Maturities",
        "code": "TD-MAT",
        "type": "inflow",
        "actual_formula": '=IFERROR(SUMPRODUCT((tbl_Bank_Statement[CF_Line_Item]="Time Deposit Maturities")*(tbl_Bank_Statement[Credit])),0)',
        "budget_formula": '=IFERROR(SUMIFS(tbl_TD_Register[Maturity_Amount_SAR],tbl_TD_Register[Status],"Active",tbl_TD_Register[Days_To_Maturity],"<=30",tbl_TD_Register[Days_To_Maturity],">=0"),0)'
    },
    {
        "name": "Interest Received on TDs",
        "code": "TD-INT",
        "type": "inflow",
        "actual_formula": '=IFERROR(SUMPRODUCT((tbl_Bank_Statement[CF_Line_Item]="Interest Received on TDs")*(tbl_Bank_Statement[Credit])),0)',
        "budget_formula": '=IFERROR(SUMIF(tbl_TD_Register[Status],"Active",tbl_TD_Register[Profit_SAR]),0)'
    },
    {
        "name": "Purchase of Fixed Assets",
        "code": "CAPEX-PUR",
        "type": "outflow",
        "actual_formula": '=IFERROR(-SUMPRODUCT((tbl_Bank_Statement[CF_Line_Item]="Purchase of Fixed Assets")*(tbl_Bank_Statement[Debit])),0)',
        "budget_formula": '=0'
    },
]

# Financing Activities line items
FINANCING_LINE_ITEMS: List[Dict] = [
    {
        "name": "Loan Drawdowns",
        "code": "LOAN-DRAW",
        "type": "inflow",
        "actual_formula": '=IFERROR(SUMPRODUCT((tbl_Bank_Statement[CF_Line_Item]="Loan Drawdowns")*(tbl_Bank_Statement[Credit])),0)',
        "budget_formula": '=0'
    },
    {
        "name": "Loan Repayments - Principal",
        "code": "LOAN-PRIN",
        "type": "outflow",
        "actual_formula": '=IFERROR(-SUMPRODUCT((tbl_Bank_Statement[CF_Line_Item]="Loan Repayments - Principal")*(tbl_Bank_Statement[Debit])),0)',
        "budget_formula": '=0'
    },
    {
        "name": "Loan Repayments - Interest",
        "code": "LOAN-INT",
        "type": "outflow",
        "actual_formula": '=IFERROR(-SUMPRODUCT((tbl_Bank_Statement[CF_Line_Item]="Loan Repayments - Interest")*(tbl_Bank_Statement[Debit])),0)',
        "budget_formula": '=0'
    },
]


# =============================================================================
# SECTION 3: SHEET GENERATOR
# =============================================================================

class CashFlowStatementGenerator:
    """
    Generator for Cash Flow Statement summary sheet.

    Creates a formatted statement with:
    - Operating/Investing/Financing sections
    - Actual vs Budget comparison
    - Variance analysis
    - Subtotals and net change
    """

    SHEET_NAME = "Cash_Flow_Statement"
    TITLE_ROW = 1
    PERIOD_ROW = 3
    HEADER_ROW = 5
    DATA_START_ROW = 7

    def __init__(self, workbook: Workbook):
        """Initialize the Cash Flow Statement generator."""
        self.workbook = workbook
        self.sheet: Optional[Worksheet] = None
        self.current_row = self.DATA_START_ROW

    def generate(self) -> Worksheet:
        """Generate the Cash Flow Statement sheet."""
        if self.SHEET_NAME in self.workbook.sheetnames:
            self.sheet = self.workbook[self.SHEET_NAME]
        else:
            self.sheet = self.workbook.create_sheet(self.SHEET_NAME)

        # Build sheet
        self._write_title()
        self._write_period_selector()
        self._write_headers()

        # Write sections
        operating_subtotal_row = self._write_operating_section()
        investing_subtotal_row = self._write_investing_section()
        financing_subtotal_row = self._write_financing_section()

        # Write summary
        self._write_summary(
            operating_subtotal_row,
            investing_subtotal_row,
            financing_subtotal_row
        )

        # Formatting
        self._set_column_widths()
        self._hide_code_column()
        self._freeze_panes()

        return self.sheet

    def _write_title(self):
        """Write statement title."""
        self.sheet.cell(row=self.TITLE_ROW, column=1, value="CASH FLOW STATEMENT")
        title_cell = self.sheet.cell(row=self.TITLE_ROW, column=1)
        title_cell.font = Font(
            name=Typography.FONT_FAMILY,
            size=18,
            bold=True,
            color=ColorPalette.HEADER_DARK_NAVY
        )
        self.sheet.merge_cells("A1:F1")

        # Subtitle
        self.sheet.cell(
            row=self.TITLE_ROW + 1, column=1,
            value="All amounts in SAR (Saudi Riyal)"
        )
        self.sheet.cell(row=self.TITLE_ROW + 1, column=1).font = Font(
            name=Typography.FONT_FAMILY,
            size=11,
            italic=True,
            color=ColorPalette.TEXT_DARK_GRAY
        )

    def _write_period_selector(self):
        """Write period selection reference."""
        self.sheet.cell(row=self.PERIOD_ROW, column=1, value="Period:")
        self.sheet.cell(row=self.PERIOD_ROW, column=1).font = Font(bold=True)

        self.sheet.cell(
            row=self.PERIOD_ROW, column=2,
            value='=IFERROR(Selected_Period,TEXT(TODAY(),"MMM-YYYY"))'
        )
        self.sheet.cell(row=self.PERIOD_ROW, column=2).font = Font(bold=True)

    def _write_headers(self):
        """Write column headers."""
        header_style = {
            "font": Font(
                name=Typography.FONT_FAMILY,
                size=12,
                bold=True,
                color=ColorPalette.TEXT_WHITE
            ),
            "fill": PatternFill(
                start_color=ColorPalette.HEADER_DARK_NAVY,
                end_color=ColorPalette.HEADER_DARK_NAVY,
                fill_type="solid"
            ),
            "border": Border(
                bottom=Side(style="medium", color=ColorPalette.BORDER_MEDIUM_BLACK)
            )
        }

        for col_idx, col_def in enumerate(CF_STATEMENT_COLUMNS, start=1):
            cell = self.sheet.cell(
                row=self.HEADER_ROW,
                column=col_idx,
                value=col_def["header"]
            )
            cell.font = header_style["font"]
            cell.fill = header_style["fill"]
            cell.border = header_style["border"]
            cell.alignment = Alignment(
                horizontal=col_def["alignment"],
                vertical="center"
            )

    def _write_section_header(self, title: str, section_color: str):
        """Write a section header row."""
        cell = self.sheet.cell(row=self.current_row, column=1, value=title)
        cell.font = Font(
            name=Typography.FONT_FAMILY,
            size=12,
            bold=True,
            color=ColorPalette.TEXT_WHITE
        )
        cell.fill = PatternFill(
            start_color=ColorPalette.HEADER_MEDIUM_BLUE,
            end_color=ColorPalette.HEADER_MEDIUM_BLUE,
            fill_type="solid"
        )

        # Merge section header
        self.sheet.merge_cells(
            start_row=self.current_row,
            start_column=1,
            end_row=self.current_row,
            end_column=6
        )

        self.current_row += 1

    def _write_line_item(self, item: Dict, is_alternate: bool = False):
        """Write a single line item row."""
        fill_color = ColorPalette.ROW_LIGHT_GRAY if is_alternate else ColorPalette.ROW_WHITE

        style = {
            "font": Font(
                name=Typography.FONT_FAMILY,
                size=11,
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

        row = self.current_row

        # Column A: Line Item Name (indented)
        cell_a = self.sheet.cell(row=row, column=1, value=f"  {item['name']}")
        cell_a.font = style["font"]
        cell_a.fill = style["fill"]
        cell_a.border = style["border"]

        # Column B: Category Code (hidden)
        cell_b = self.sheet.cell(row=row, column=2, value=item['code'])
        cell_b.font = style["font"]
        cell_b.fill = style["fill"]
        cell_b.border = style["border"]

        # Column C: Actual
        cell_c = self.sheet.cell(row=row, column=3, value=item['actual_formula'])
        cell_c.font = style["font"]
        cell_c.fill = style["fill"]
        cell_c.border = style["border"]
        cell_c.number_format = "#,##0.00"
        cell_c.alignment = Alignment(horizontal="right")

        # Column D: Budget
        cell_d = self.sheet.cell(row=row, column=4, value=item['budget_formula'])
        cell_d.font = style["font"]
        cell_d.fill = style["fill"]
        cell_d.border = style["border"]
        cell_d.number_format = "#,##0.00"
        cell_d.alignment = Alignment(horizontal="right")

        # Column E: Variance (Actual - Budget)
        cell_e = self.sheet.cell(row=row, column=5, value=f"=C{row}-D{row}")
        cell_e.font = style["font"]
        cell_e.fill = style["fill"]
        cell_e.border = style["border"]
        cell_e.number_format = "#,##0.00"
        cell_e.alignment = Alignment(horizontal="right")

        # Column F: Variance %
        cell_f = self.sheet.cell(
            row=row, column=6,
            value=f"=IFERROR(E{row}/ABS(D{row}),0)"
        )
        cell_f.font = style["font"]
        cell_f.fill = style["fill"]
        cell_f.border = style["border"]
        cell_f.number_format = "0.0%"
        cell_f.alignment = Alignment(horizontal="right")

        self.current_row += 1

    def _write_subtotal(self, title: str, start_row: int, end_row: int) -> int:
        """Write a subtotal row and return its row number."""
        row = self.current_row

        style = {
            "font": Font(
                name=Typography.FONT_FAMILY,
                size=11,
                bold=True,
                color=ColorPalette.HEADER_DARK_NAVY
            ),
            "fill": PatternFill(
                start_color=ColorPalette.HEADER_LIGHT_BLUE,
                end_color=ColorPalette.HEADER_LIGHT_BLUE,
                fill_type="solid"
            ),
            "border": Border(
                top=Side(style="medium", color=ColorPalette.BORDER_MEDIUM_BLACK),
                bottom=Side(style="double", color=ColorPalette.BORDER_MEDIUM_BLACK)
            )
        }

        # Column A: Subtotal title
        cell_a = self.sheet.cell(row=row, column=1, value=title)
        cell_a.font = style["font"]
        cell_a.fill = style["fill"]
        cell_a.border = style["border"]

        # Column B: Empty
        cell_b = self.sheet.cell(row=row, column=2, value="")
        cell_b.fill = style["fill"]
        cell_b.border = style["border"]

        # Column C: Sum of Actual
        cell_c = self.sheet.cell(row=row, column=3, value=f"=SUM(C{start_row}:C{end_row})")
        cell_c.font = style["font"]
        cell_c.fill = style["fill"]
        cell_c.border = style["border"]
        cell_c.number_format = "#,##0.00"
        cell_c.alignment = Alignment(horizontal="right")

        # Column D: Sum of Budget
        cell_d = self.sheet.cell(row=row, column=4, value=f"=SUM(D{start_row}:D{end_row})")
        cell_d.font = style["font"]
        cell_d.fill = style["fill"]
        cell_d.border = style["border"]
        cell_d.number_format = "#,##0.00"
        cell_d.alignment = Alignment(horizontal="right")

        # Column E: Variance
        cell_e = self.sheet.cell(row=row, column=5, value=f"=C{row}-D{row}")
        cell_e.font = style["font"]
        cell_e.fill = style["fill"]
        cell_e.border = style["border"]
        cell_e.number_format = "#,##0.00"
        cell_e.alignment = Alignment(horizontal="right")

        # Column F: Variance %
        cell_f = self.sheet.cell(row=row, column=6, value=f"=IFERROR(E{row}/ABS(D{row}),0)")
        cell_f.font = style["font"]
        cell_f.fill = style["fill"]
        cell_f.border = style["border"]
        cell_f.number_format = "0.0%"
        cell_f.alignment = Alignment(horizontal="right")

        self.current_row += 2  # Add blank row after subtotal
        return row

    def _write_operating_section(self) -> int:
        """Write Operating Activities section."""
        self._write_section_header(
            "CASH FLOWS FROM OPERATING ACTIVITIES",
            ColorPalette.SECTION_OPERATING
        )

        start_row = self.current_row
        for idx, item in enumerate(OPERATING_LINE_ITEMS):
            self._write_line_item(item, is_alternate=idx % 2 == 1)
        end_row = self.current_row - 1

        return self._write_subtotal(
            "NET CASH FROM OPERATING ACTIVITIES",
            start_row,
            end_row
        )

    def _write_investing_section(self) -> int:
        """Write Investing Activities section."""
        self._write_section_header(
            "CASH FLOWS FROM INVESTING ACTIVITIES",
            ColorPalette.SECTION_INVESTING
        )

        start_row = self.current_row
        for idx, item in enumerate(INVESTING_LINE_ITEMS):
            self._write_line_item(item, is_alternate=idx % 2 == 1)
        end_row = self.current_row - 1

        return self._write_subtotal(
            "NET CASH FROM INVESTING ACTIVITIES",
            start_row,
            end_row
        )

    def _write_financing_section(self) -> int:
        """Write Financing Activities section."""
        self._write_section_header(
            "CASH FLOWS FROM FINANCING ACTIVITIES",
            ColorPalette.SECTION_FINANCING
        )

        start_row = self.current_row
        for idx, item in enumerate(FINANCING_LINE_ITEMS):
            self._write_line_item(item, is_alternate=idx % 2 == 1)
        end_row = self.current_row - 1

        return self._write_subtotal(
            "NET CASH FROM FINANCING ACTIVITIES",
            start_row,
            end_row
        )

    def _write_summary(
        self,
        operating_row: int,
        investing_row: int,
        financing_row: int
    ):
        """Write the summary section with net change."""
        row = self.current_row

        # Grand total style
        style = {
            "font": Font(
                name=Typography.FONT_FAMILY,
                size=12,
                bold=True,
                color=ColorPalette.TEXT_WHITE
            ),
            "fill": PatternFill(
                start_color=ColorPalette.HEADER_DARK_NAVY,
                end_color=ColorPalette.HEADER_DARK_NAVY,
                fill_type="solid"
            ),
            "border": Border(
                top=Side(style="thick", color=ColorPalette.BORDER_MEDIUM_BLACK),
                bottom=Side(style="thick", color=ColorPalette.BORDER_MEDIUM_BLACK)
            )
        }

        # NET INCREASE/(DECREASE) IN CASH
        cell_a = self.sheet.cell(row=row, column=1, value="NET INCREASE/(DECREASE) IN CASH")
        cell_a.font = style["font"]
        cell_a.fill = style["fill"]
        cell_a.border = style["border"]

        cell_b = self.sheet.cell(row=row, column=2, value="")
        cell_b.fill = style["fill"]
        cell_b.border = style["border"]

        # Sum of all three sections
        cell_c = self.sheet.cell(
            row=row, column=3,
            value=f"=C{operating_row}+C{investing_row}+C{financing_row}"
        )
        cell_c.font = style["font"]
        cell_c.fill = style["fill"]
        cell_c.border = style["border"]
        cell_c.number_format = "#,##0.00"
        cell_c.alignment = Alignment(horizontal="right")

        cell_d = self.sheet.cell(
            row=row, column=4,
            value=f"=D{operating_row}+D{investing_row}+D{financing_row}"
        )
        cell_d.font = style["font"]
        cell_d.fill = style["fill"]
        cell_d.border = style["border"]
        cell_d.number_format = "#,##0.00"
        cell_d.alignment = Alignment(horizontal="right")

        cell_e = self.sheet.cell(row=row, column=5, value=f"=C{row}-D{row}")
        cell_e.font = style["font"]
        cell_e.fill = style["fill"]
        cell_e.border = style["border"]
        cell_e.number_format = "#,##0.00"
        cell_e.alignment = Alignment(horizontal="right")

        cell_f = self.sheet.cell(row=row, column=6, value=f"=IFERROR(E{row}/ABS(D{row}),0)")
        cell_f.font = style["font"]
        cell_f.fill = style["fill"]
        cell_f.border = style["border"]
        cell_f.number_format = "0.0%"
        cell_f.alignment = Alignment(horizontal="right")

    def _set_column_widths(self):
        """Set column widths."""
        for col_idx, col_def in enumerate(CF_STATEMENT_COLUMNS, start=1):
            col_letter = get_column_letter(col_idx)
            self.sheet.column_dimensions[col_letter].width = col_def["width"]

    def _hide_code_column(self):
        """Hide the Category_Code column."""
        self.sheet.column_dimensions["B"].hidden = True

    def _freeze_panes(self):
        """Freeze header row."""
        self.sheet.freeze_panes = f"A{self.DATA_START_ROW}"


# =============================================================================
# SECTION 4: UTILITY FUNCTIONS
# =============================================================================

def generate_cash_flow_statement_sheet(workbook: Workbook) -> Worksheet:
    """
    Convenience function to generate Cash Flow Statement sheet.

    Args:
        workbook: The openpyxl Workbook.

    Returns:
        The generated Worksheet.
    """
    generator = CashFlowStatementGenerator(workbook)
    return generator.generate()
