"""
BANTMS TD Register Sheet Generator
===================================

Generates the TD_Register (Time Deposits) sheet with:
- Fixed deposit portfolio tracking
- Maturity date calculations
- Interest/profit calculations
- Currency conversion to SAR
- Status tracking (Active/Action_Required/Closed)
"""

from datetime import date, datetime, timedelta
from decimal import Decimal
from typing import List, Dict, Optional
import random

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from openpyxl.styles import (
    Font, PatternFill, Border, Side, Alignment
)

from ..config.constants import (
    CurrencyFramework,
    SaudiBanks,
    TDStatus
)
from ..config.styles import (
    ColorPalette, Typography, NumberFormats,
    ColumnWidths, RowHeights
)


# =============================================================================
# SECTION 1: COLUMN DEFINITIONS
# =============================================================================

TD_REGISTER_COLUMNS: List[Dict] = [
    {
        "header": "TD_ID",
        "width": 12,
        "format": "@",
        "alignment": "center",
        "description": "Unique time deposit identifier",
        "formula": None
    },
    {
        "header": "Bank",
        "width": 12,
        "format": "@",
        "alignment": "center",
        "description": "Bank code",
        "formula": None
    },
    {
        "header": "Bank_Name",
        "width": 22,
        "format": "@",
        "alignment": "left",
        "description": "Bank full name",
        "formula": None
    },
    {
        "header": "Account_No",
        "width": 18,
        "format": "@",
        "alignment": "left",
        "description": "TD account number",
        "formula": None
    },
    {
        "header": "CCY",
        "width": 8,
        "format": "@",
        "alignment": "center",
        "description": "Currency code",
        "formula": None
    },
    {
        "header": "Start_Date",
        "width": 12,
        "format": "DD-MMM-YYYY",
        "alignment": "center",
        "description": "Deposit start date",
        "formula": None
    },
    {
        "header": "Maturity_Date",
        "width": 12,
        "format": "DD-MMM-YYYY",
        "alignment": "center",
        "description": "Maturity date",
        "formula": None
    },
    {
        "header": "Days",
        "width": 8,
        "format": "0",
        "alignment": "center",
        "description": "Tenor in days",
        "formula": '=IFERROR([@Maturity_Date]-[@Start_Date],0)'
    },
    {
        "header": "Principal_Original",
        "width": 16,
        "format": "#,##0.00",
        "alignment": "right",
        "description": "Principal in original currency",
        "formula": None
    },
    {
        "header": "Rate_Percent",
        "width": 10,
        "format": "0.00%",
        "alignment": "center",
        "description": "Annual interest rate",
        "formula": None
    },
    {
        "header": "Profit_Original",
        "width": 14,
        "format": "#,##0.00",
        "alignment": "right",
        "description": "Profit in original currency",
        "formula": '=IFERROR([@Principal_Original]*[@Rate_Percent]*[@Days]/365,0)'
    },
    {
        "header": "Maturity_Amount_Original",
        "width": 18,
        "format": "#,##0.00",
        "alignment": "right",
        "description": "Total maturity amount in original currency",
        "formula": '=IFERROR([@Principal_Original]+[@Profit_Original],0)'
    },
    {
        "header": "FX_Rate",
        "width": 10,
        "format": "#,##0.0000",
        "alignment": "right",
        "description": "Exchange rate to SAR",
        "formula": '=IFERROR(IF([@CCY]="SAR",1,INDEX(FX_Rate_Lookup,MATCH([@CCY],CCY_Codes,0),3)),1)'
    },
    {
        "header": "Principal_SAR",
        "width": 16,
        "format": "#,##0.00",
        "alignment": "right",
        "description": "Principal in SAR",
        "formula": '=IFERROR([@Principal_Original]*[@FX_Rate],0)'
    },
    {
        "header": "Profit_SAR",
        "width": 14,
        "format": "#,##0.00",
        "alignment": "right",
        "description": "Profit in SAR",
        "formula": '=IFERROR([@Profit_Original]*[@FX_Rate],0)'
    },
    {
        "header": "Maturity_Amount_SAR",
        "width": 18,
        "format": "#,##0.00",
        "alignment": "right",
        "description": "Total maturity amount in SAR",
        "formula": '=IFERROR([@Maturity_Amount_Original]*[@FX_Rate],0)'
    },
    {
        "header": "Status",
        "width": 14,
        "format": "@",
        "alignment": "center",
        "description": "Active/Action_Required/Closed",
        "formula": '=IF([@Maturity_Date]>TODAY(),"Active",IF([@Maturity_Date]=TODAY(),"Action_Required","Closed"))'
    },
    {
        "header": "Days_To_Maturity",
        "width": 14,
        "format": "0",
        "alignment": "center",
        "description": "Days until maturity",
        "formula": '=IFERROR([@Maturity_Date]-TODAY(),0)'
    },
    {
        "header": "Week_ID",
        "width": 10,
        "format": "00",
        "alignment": "center",
        "description": "Saudi week number for maturity",
        "formula": '=IFERROR(INT(([@Maturity_Date]-DATE(YEAR([@Maturity_Date]),1,1)-MOD(WEEKDAY(DATE(YEAR([@Maturity_Date]),1,1))+1,7)+1)/7)+1,"")'
    },
    {
        "header": "Rollover_Intent",
        "width": 12,
        "format": "@",
        "alignment": "center",
        "description": "Rollover/Withdraw/Pending",
        "formula": None
    },
    {
        "header": "Purpose",
        "width": 20,
        "format": "@",
        "alignment": "left",
        "description": "Purpose/allocation",
        "formula": None
    },
    {
        "header": "Notes",
        "width": 25,
        "format": "@",
        "alignment": "left",
        "description": "Additional notes",
        "formula": None
    },
]


# =============================================================================
# SECTION 2: SAMPLE DATA GENERATOR
# =============================================================================

def generate_sample_td_data() -> List[Dict]:
    """
    Generate sample TD data for demonstration.

    Returns:
        List of TD dictionaries with realistic data.
    """
    # Banks for TDs
    banks = [
        ("SNB", "Saudi National Bank"),
        ("RAJHI", "Al Rajhi Bank"),
        ("RIYAD", "Riyad Bank"),
        ("SABB", "Saudi British Bank"),
        ("ANB", "Arab National Bank"),
        ("BSF", "Banque Saudi Fransi"),
    ]

    tds = []
    td_id = 1

    # Generate 15-20 time deposits
    num_tds = random.randint(15, 20)

    for i in range(num_tds):
        bank_code, bank_name = random.choice(banks)

        # Tenor options: 30, 60, 90, 180, 365 days
        tenor = random.choice([30, 60, 90, 180, 365])

        # Start date: within last 6 months
        start_date = date.today() - timedelta(days=random.randint(0, 180))
        maturity_date = start_date + timedelta(days=tenor)

        # Principal amount: 1M to 10M SAR
        principal = random.randint(1000000, 10000000)

        # Currency: mostly SAR, some USD
        ccy = "SAR" if random.random() < 0.8 else "USD"

        # Interest rate based on tenor and currency
        if ccy == "SAR":
            base_rate = 5.0 + tenor / 365 * 1.5  # 5-6.5% for SAR
        else:
            base_rate = 4.5 + tenor / 365 * 1.0  # 4.5-5.5% for USD

        rate = base_rate + random.uniform(-0.5, 0.5)
        rate = round(rate, 2) / 100  # Convert to decimal

        # Rollover intent
        if maturity_date > date.today():
            rollover = random.choice(["Rollover", "Withdraw", "Pending"])
        else:
            rollover = "Completed"

        # Purpose
        purposes = [
            "Working Capital Reserve",
            "Project Fund",
            "Tax Provision",
            "Dividend Reserve",
            "Contingency Fund",
            "Growth Capital",
        ]

        tds.append({
            "TD_ID": f"TD-{td_id:04d}",
            "Bank": bank_code,
            "Bank_Name": bank_name,
            "Account_No": f"TD{bank_code}{td_id:06d}",
            "CCY": ccy,
            "Start_Date": start_date,
            "Maturity_Date": maturity_date,
            "Principal_Original": float(principal),
            "Rate_Percent": rate,
            "Rollover_Intent": rollover,
            "Purpose": random.choice(purposes),
            "Notes": "",
        })
        td_id += 1

    # Sort by maturity date
    tds.sort(key=lambda x: x["Maturity_Date"])

    return tds


# =============================================================================
# SECTION 3: SHEET GENERATOR
# =============================================================================

class TDRegisterSheetGenerator:
    """
    Generator for TD_Register (Time Deposits) sheet.

    Creates a formatted Excel sheet with:
    - TD portfolio tracking
    - Interest/profit calculations
    - Currency conversion to SAR
    - Maturity alerts
    """

    SHEET_NAME = "TD_Register"
    TABLE_NAME = "tbl_TD_Register"
    TITLE_ROW = 1
    HEADER_ROW = 3
    DATA_START_ROW = 4

    def __init__(self, workbook: Workbook):
        """
        Initialize the TD Register sheet generator.

        Args:
            workbook: The openpyxl Workbook to add the sheet to.
        """
        self.workbook = workbook
        self.sheet: Optional[Worksheet] = None
        self.tds: List[Dict] = []
        self.last_data_row = self.DATA_START_ROW

    def generate(
        self,
        tds: Optional[List[Dict]] = None,
        include_sample_data: bool = True
    ) -> Worksheet:
        """
        Generate the TD_Register sheet.

        Args:
            tds: Optional list of TD dictionaries.
            include_sample_data: Whether to include sample data (default True).

        Returns:
            The generated Worksheet.
        """
        if tds:
            self.tds = tds
        elif include_sample_data:
            self.tds = generate_sample_td_data()
        else:
            self.tds = []

        # Create or get sheet
        if self.SHEET_NAME in self.workbook.sheetnames:
            self.sheet = self.workbook[self.SHEET_NAME]
        else:
            self.sheet = self.workbook.create_sheet(self.SHEET_NAME)

        # Build sheet components
        self._write_title()
        self._write_headers()
        self._write_data()
        self._create_table()
        self._set_column_widths()
        self._add_data_validation()
        self._freeze_panes()
        self._add_summary_section()
        self._add_instructions()

        return self.sheet

    def _write_title(self):
        """Write sheet title and description."""
        self.sheet.cell(
            row=self.TITLE_ROW,
            column=1,
            value="TD REGISTER - Time Deposit Portfolio"
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
            end_column=len(TD_REGISTER_COLUMNS)
        )

        # Subtitle
        self.sheet.cell(
            row=self.TITLE_ROW + 1,
            column=1,
            value="Fixed Deposit Tracking | Interest Calculations | All amounts in SAR"
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

        for col_idx, col_def in enumerate(TD_REGISTER_COLUMNS, start=1):
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
        """Write TD data rows with formulas."""
        for row_idx, td in enumerate(self.tds):
            excel_row = self.DATA_START_ROW + row_idx
            is_alternate = row_idx % 2 == 1

            self._write_td_row(excel_row, td, is_alternate)

        self.last_data_row = self.DATA_START_ROW + len(self.tds) - 1
        if self.last_data_row < self.DATA_START_ROW:
            self.last_data_row = self.DATA_START_ROW

    def _write_td_row(self, row: int, td: Dict, is_alternate: bool):
        """Write a single TD row."""
        fill_color = ColorPalette.ROW_LIGHT_GRAY if is_alternate else ColorPalette.ROW_WHITE

        # Status-based styling (calculate status)
        maturity_date = td.get("Maturity_Date")
        today = date.today()

        if maturity_date:
            if maturity_date > today:
                status = "Active"
                status_color = fill_color
            elif maturity_date == today:
                status = "Action_Required"
                status_color = ColorPalette.STATUS_URGENT_PURPLE
            else:
                status = "Closed"
                status_color = ColorPalette.STATUS_PAID_GREEN
        else:
            status_color = fill_color

        # Highlight TDs maturing within 7 days
        if maturity_date and (maturity_date - today).days <= 7 and maturity_date > today:
            status_color = ColorPalette.STATUS_PENDING_AMBER

        style = {
            "font": Font(
                name=Typography.FONT_FAMILY,
                size=Typography.SIZE_BODY,
                color=ColorPalette.TEXT_BLACK
            ),
            "fill": PatternFill(
                start_color=status_color,
                end_color=status_color,
                fill_type="solid"
            ),
            "border": Border(
                left=Side(style="thin", color=ColorPalette.BORDER_THIN_GRAY),
                right=Side(style="thin", color=ColorPalette.BORDER_THIN_GRAY),
                top=Side(style="thin", color=ColorPalette.BORDER_THIN_GRAY),
                bottom=Side(style="thin", color=ColorPalette.BORDER_THIN_GRAY)
            )
        }

        for col_idx, col_def in enumerate(TD_REGISTER_COLUMNS, start=1):
            header = col_def["header"]

            # Determine cell value
            if col_def["formula"]:
                value = col_def["formula"]
            elif header in td:
                value = td[header]
            else:
                value = None

            cell = self.sheet.cell(row=row, column=col_idx, value=value)
            cell.font = style["font"]
            cell.fill = style["fill"]
            cell.border = style["border"]

            h_align = col_def["alignment"]
            cell.alignment = Alignment(horizontal=h_align, vertical="center")
            cell.number_format = col_def["format"]

    def _create_table(self):
        """Create Excel table for TD data."""
        start_col = "A"
        end_col = get_column_letter(len(TD_REGISTER_COLUMNS))

        table_end_row = max(self.last_data_row, self.DATA_START_ROW)
        table_range = f"{start_col}{self.HEADER_ROW}:{end_col}{table_end_row}"

        table = Table(
            displayName=self.TABLE_NAME,
            ref=table_range
        )

        table_style = TableStyleInfo(
            name="TableStyleMedium5",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
        )
        table.tableStyleInfo = table_style

        self.sheet.add_table(table)

    def _set_column_widths(self):
        """Set column widths based on definitions."""
        for col_idx, col_def in enumerate(TD_REGISTER_COLUMNS, start=1):
            col_letter = get_column_letter(col_idx)
            self.sheet.column_dimensions[col_letter].width = col_def["width"]

        self.sheet.row_dimensions[self.TITLE_ROW].height = RowHeights.TITLE
        self.sheet.row_dimensions[self.HEADER_ROW].height = RowHeights.HEADER

    def _add_data_validation(self):
        """Add data validation for input columns."""
        # Bank dropdown
        bank_codes = ",".join(SaudiBanks.BANK_CODES)
        bank_validation = DataValidation(
            type="list",
            formula1=f'"{bank_codes}"',
            allow_blank=False
        )
        bank_col = get_column_letter(2)
        bank_validation.add(f"{bank_col}{self.DATA_START_ROW}:{bank_col}5000")
        self.sheet.add_data_validation(bank_validation)

        # Currency dropdown
        ccy_codes = ",".join(CurrencyFramework.CURRENCY_CODES)
        ccy_validation = DataValidation(
            type="list",
            formula1=f'"{ccy_codes}"',
            allow_blank=False
        )
        ccy_col = get_column_letter(5)
        ccy_validation.add(f"{ccy_col}{self.DATA_START_ROW}:{ccy_col}5000")
        self.sheet.add_data_validation(ccy_validation)

        # Rollover Intent dropdown
        rollover_validation = DataValidation(
            type="list",
            formula1='"Rollover,Withdraw,Pending,Completed"',
            allow_blank=False
        )
        rollover_col = get_column_letter(20)
        rollover_validation.add(f"{rollover_col}{self.DATA_START_ROW}:{rollover_col}5000")
        self.sheet.add_data_validation(rollover_validation)

    def _freeze_panes(self):
        """Freeze header rows and key columns."""
        self.sheet.freeze_panes = f"D{self.DATA_START_ROW}"

    def _add_summary_section(self):
        """Add portfolio summary section."""
        summary_row = self.last_data_row + 3

        # Summary header
        self.sheet.cell(row=summary_row, column=1, value="PORTFOLIO SUMMARY")
        self.sheet.cell(row=summary_row, column=1).font = Font(
            name=Typography.FONT_FAMILY,
            size=Typography.SIZE_HEADER,
            bold=True,
            color=ColorPalette.HEADER_DARK_NAVY
        )

        # Summary metrics
        metrics = [
            ("Total Principal (SAR):", f'=SUMIF(tbl_TD_Register[Status],"Active",tbl_TD_Register[Principal_SAR])'),
            ("Total Expected Profit (SAR):", f'=SUMIF(tbl_TD_Register[Status],"Active",tbl_TD_Register[Profit_SAR])'),
            ("Total at Maturity (SAR):", f'=SUMIF(tbl_TD_Register[Status],"Active",tbl_TD_Register[Maturity_Amount_SAR])'),
            ("Active TDs Count:", f'=COUNTIF(tbl_TD_Register[Status],"Active")'),
            ("Maturing This Week:", f'=COUNTIFS(tbl_TD_Register[Status],"Active",tbl_TD_Register[Days_To_Maturity],"<=7")'),
            ("Action Required:", f'=COUNTIF(tbl_TD_Register[Status],"Action_Required")'),
        ]

        for idx, (label, formula) in enumerate(metrics):
            label_row = summary_row + 1 + idx
            self.sheet.cell(row=label_row, column=1, value=label)
            self.sheet.cell(row=label_row, column=1).font = Font(
                name=Typography.FONT_FAMILY,
                size=Typography.SIZE_BODY,
                bold=True
            )

            self.sheet.cell(row=label_row, column=2, value=formula)
            self.sheet.cell(row=label_row, column=2).number_format = "#,##0.00"
            self.sheet.cell(row=label_row, column=2).font = Font(
                name=Typography.FONT_FAMILY,
                size=Typography.SIZE_BODY,
                color=ColorPalette.VALUE_POSITIVE_GREEN
            )

    def _add_instructions(self):
        """Add usage instructions at the bottom."""
        instructions_row = self.last_data_row + 12

        instructions = [
            "",
            "STATUS LOGIC:",
            "• Active: Maturity date > Today",
            "• Action_Required: Maturity date = Today",
            "• Closed: Maturity date < Today",
            "",
            "PROFIT CALCULATION:",
            "• Profit = Principal × Rate × Days / 365",
            "• All amounts converted to SAR",
            "",
            "ROLLOVER INTENT:",
            "• Rollover: Renew TD at maturity",
            "• Withdraw: Cash out at maturity",
            "• Pending: Decision not yet made",
        ]

        for idx, text in enumerate(instructions):
            cell = self.sheet.cell(row=instructions_row + idx, column=1, value=text)
            if text.startswith("STATUS") or text.startswith("PROFIT") or text.startswith("ROLLOVER"):
                cell.font = Font(
                    name=Typography.FONT_FAMILY,
                    size=Typography.SIZE_BODY,
                    bold=True,
                    color=ColorPalette.HEADER_DARK_NAVY
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

def generate_td_register_sheet(
    workbook: Workbook,
    include_sample_data: bool = True
) -> Worksheet:
    """
    Convenience function to generate TD_Register sheet.

    Args:
        workbook: The openpyxl Workbook.
        include_sample_data: Whether to include sample data.

    Returns:
        The generated Worksheet.
    """
    generator = TDRegisterSheetGenerator(workbook)
    return generator.generate(include_sample_data=include_sample_data)


def get_profit_formula() -> str:
    """
    Get the profit calculation formula.

    Returns:
        Excel formula string for profit calculation.
    """
    return '=IFERROR([@Principal_Original]*[@Rate_Percent]*[@Days]/365,0)'


def get_status_formula() -> str:
    """
    Get the TD status formula.

    Returns:
        Excel formula string for status determination.
    """
    return '=IF([@Maturity_Date]>TODAY(),"Active",IF([@Maturity_Date]=TODAY(),"Action_Required","Closed"))'
