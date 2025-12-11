"""
BANTMS Bank Statement Actual Sheet Generator
=============================================

Generates the Bank_Statement_Actual sheet for:
- Importing actual bank transactions
- Auto-categorization via Category_Map patterns
- Currency conversion to SAR (base currency)
- Saudi Week ID calculation (Sunday start)
- Period aggregation fields (Month, Quarter)

This is the primary source of actual cash flow data.
"""

from datetime import date, datetime
from decimal import Decimal
from typing import List, Dict, Optional, Tuple

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from openpyxl.styles import (
    Font, PatternFill, Border, Side, Alignment, Protection
)

from ..config.constants import (
    SaudiCalendarConfig,
    CurrencyFramework,
    SaudiBanks
)
from ..config.styles import (
    ColorPalette, Typography, NumberFormats,
    ColumnWidths, RowHeights
)


# =============================================================================
# SECTION 1: COLUMN DEFINITIONS
# =============================================================================

BANK_STATEMENT_COLUMNS: List[Dict] = [
    {
        "header": "Transaction_ID",
        "width": 14,
        "format": "@",
        "alignment": "center",
        "description": "Auto-generated transaction identifier",
        "formula": None,
        "protected": True
    },
    {
        "header": "Date",
        "width": 12,
        "format": "DD-MMM-YYYY",
        "alignment": "center",
        "description": "Transaction date",
        "formula": None,
        "protected": False
    },
    {
        "header": "Value_Date",
        "width": 12,
        "format": "DD-MMM-YYYY",
        "alignment": "center",
        "description": "Value date for interest calculation",
        "formula": None,
        "protected": False
    },
    {
        "header": "Reference_No",
        "width": 18,
        "format": "@",
        "alignment": "left",
        "description": "Bank reference number",
        "formula": None,
        "protected": False
    },
    {
        "header": "Description",
        "width": 40,
        "format": "@",
        "alignment": "left",
        "description": "Transaction description (raw from bank)",
        "formula": None,
        "protected": False
    },
    {
        "header": "Debit",
        "width": 15,
        "format": "#,##0.00",
        "alignment": "right",
        "description": "Debit amount (outflow)",
        "formula": None,
        "protected": False
    },
    {
        "header": "Credit",
        "width": 15,
        "format": "#,##0.00",
        "alignment": "right",
        "description": "Credit amount (inflow)",
        "formula": None,
        "protected": False
    },
    {
        "header": "Running_Balance",
        "width": 16,
        "format": "#,##0.00",
        "alignment": "right",
        "description": "Running balance per bank statement",
        "formula": None,
        "protected": False
    },
    {
        "header": "Bank_Code",
        "width": 10,
        "format": "@",
        "alignment": "center",
        "description": "Bank identifier code",
        "formula": None,
        "protected": False
    },
    {
        "header": "Account_No",
        "width": 18,
        "format": "@",
        "alignment": "left",
        "description": "Bank account number",
        "formula": None,
        "protected": False
    },
    {
        "header": "Original_CCY",
        "width": 10,
        "format": "@",
        "alignment": "center",
        "description": "Original transaction currency",
        "formula": None,
        "protected": False
    },
    {
        "header": "Original_Amount",
        "width": 15,
        "format": "#,##0.00",
        "alignment": "right",
        "description": "Amount in original currency",
        "formula": None,
        "protected": False
    },
    {
        "header": "FX_Rate_Applied",
        "width": 12,
        "format": "#,##0.0000",
        "alignment": "right",
        "description": "Exchange rate applied (to SAR)",
        "formula": '=IFERROR(IF([@Original_CCY]="SAR",1,INDEX(FX_Rate_Lookup,MATCH([@Original_CCY],CCY_Codes,0),3)),1)',
        "protected": True
    },
    {
        "header": "SAR_Equivalent",
        "width": 16,
        "format": "#,##0.00",
        "alignment": "right",
        "description": "Amount converted to SAR",
        "formula": '=IFERROR(IF([@Original_CCY]="SAR",[@Original_Amount],[@Original_Amount]*[@FX_Rate_Applied]),0)',
        "protected": True
    },
    {
        "header": "Category_Code",
        "width": 14,
        "format": "@",
        "alignment": "center",
        "description": "Auto-categorized from Category_Map",
        "formula": '=IFERROR(INDEX(tbl_Category_Map[Category_Code],AGGREGATE(15,6,ROW(tbl_Category_Map[Description_Pattern])-MIN(ROW(tbl_Category_Map[Description_Pattern]))+1/(ISNUMBER(SEARCH(tbl_Category_Map[Description_Pattern],[@Description]))*(tbl_Category_Map[Is_Active]="TRUE")),1)),"UNCATEGORIZED")',
        "protected": True
    },
    {
        "header": "CF_Line_Item",
        "width": 25,
        "format": "@",
        "alignment": "left",
        "description": "Cash Flow line item",
        "formula": '=IFERROR(INDEX(tbl_Category_Map[CF_Line_Item],MATCH([@Category_Code],tbl_Category_Map[Category_Code],0)),"Other")',
        "protected": True
    },
    {
        "header": "CF_Section",
        "width": 12,
        "format": "@",
        "alignment": "center",
        "description": "Cash Flow section",
        "formula": '=IFERROR(INDEX(tbl_Category_Map[CF_Section],MATCH([@Category_Code],tbl_Category_Map[Category_Code],0)),"Operating")',
        "protected": True
    },
    {
        "header": "Week_ID",
        "width": 10,
        "format": "00",
        "alignment": "center",
        "description": "Saudi week number (Sunday start)",
        "formula": '=IFERROR(INT(([@Date]-DATE(YEAR([@Date]),1,1)-MOD(WEEKDAY(DATE(YEAR([@Date]),1,1))+1,7)+1)/7)+1,"")',
        "protected": True
    },
    {
        "header": "Month_Period",
        "width": 10,
        "format": "@",
        "alignment": "center",
        "description": "Year-Month period",
        "formula": '=IFERROR(TEXT([@Date],"YYYY-MM"),"")',
        "protected": True
    },
    {
        "header": "Quarter",
        "width": 10,
        "format": "@",
        "alignment": "center",
        "description": "Year-Quarter period",
        "formula": '=IFERROR(YEAR([@Date])&"-Q"&ROUNDUP(MONTH([@Date])/3,0),"")',
        "protected": True
    },
    {
        "header": "Day_Type",
        "width": 8,
        "format": "@",
        "alignment": "center",
        "description": "W=Working, WE=Weekend, WED=Payment Day",
        "formula": '=IFERROR(IF(WEEKDAY([@Date],2)=3,"WED",IF(OR(WEEKDAY([@Date],2)=5,WEEKDAY([@Date],2)=6),"WE","W")),"")',
        "protected": True
    },
]


# =============================================================================
# SECTION 2: SAMPLE DATA GENERATOR
# =============================================================================

def generate_sample_transactions() -> List[Dict]:
    """
    Generate sample bank transactions for demonstration.

    Returns:
        List of transaction dictionaries with realistic Saudi business data.
    """
    from datetime import timedelta
    import random

    base_date = date.today() - timedelta(days=90)
    transactions = []
    running_balance = Decimal("5000000.00")  # Starting balance 5M SAR
    tx_id = 1

    # Sample transaction templates
    templates = [
        # AR Collections
        {"desc": "ARAMCO PAYMENT REF:AR2024001", "type": "credit", "min": 500000, "max": 2000000, "ccy": "SAR"},
        {"desc": "SABIC COLLECTION INV-2024-0123", "type": "credit", "min": 300000, "max": 800000, "ccy": "SAR"},
        {"desc": "CUSTOMER PAYMENT - AL RAJHI TRANSFER", "type": "credit", "min": 50000, "max": 200000, "ccy": "SAR"},
        {"desc": "USD RECEIPT FROM EXPORT CUSTOMER", "type": "credit", "min": 50000, "max": 150000, "ccy": "USD"},

        # AP Payments
        {"desc": "SUPPLIER PAYMENT - VENDOR ID: V2001", "type": "debit", "min": 100000, "max": 500000, "ccy": "SAR"},
        {"desc": "AP PAYMENT TO AL-MARAI CO", "type": "debit", "min": 50000, "max": 150000, "ccy": "SAR"},
        {"desc": "STRATEGIC SUPPLIER - JARIR TRADING", "type": "debit", "min": 20000, "max": 80000, "ccy": "SAR"},

        # Government
        {"desc": "ZATCA VAT PAYMENT Q4-2024", "type": "debit", "min": 200000, "max": 500000, "ccy": "SAR"},
        {"desc": "GOSI CONTRIBUTION DEC-2024", "type": "debit", "min": 150000, "max": 300000, "ccy": "SAR"},

        # Payroll
        {"desc": "WPS SALARY TRANSFER DEC-2024", "type": "debit", "min": 800000, "max": 1500000, "ccy": "SAR"},
        {"desc": "EMPLOYEE BONUS PAYMENT", "type": "debit", "min": 50000, "max": 200000, "ccy": "SAR"},

        # Utilities
        {"desc": "SEC ELECTRICITY BILL REF:E2024120001", "type": "debit", "min": 30000, "max": 80000, "ccy": "SAR"},
        {"desc": "STC TELECOM INVOICE", "type": "debit", "min": 5000, "max": 15000, "ccy": "SAR"},

        # TD Activity
        {"desc": "TD PLACEMENT 3M @ SNB", "type": "debit", "min": 1000000, "max": 5000000, "ccy": "SAR"},
        {"desc": "TD MATURITY REF:TD2024001", "type": "credit", "min": 1000000, "max": 5000000, "ccy": "SAR"},
        {"desc": "TD INTEREST RECEIVED SNB", "type": "credit", "min": 10000, "max": 50000, "ccy": "SAR"},

        # Loans
        {"desc": "LOAN PRINCIPAL REPAYMENT - SABB", "type": "debit", "min": 200000, "max": 500000, "ccy": "SAR"},
        {"desc": "LOAN INTEREST PAYMENT - SABB", "type": "debit", "min": 20000, "max": 80000, "ccy": "SAR"},

        # Bank Charges
        {"desc": "BANK SERVICE CHARGE", "type": "debit", "min": 500, "max": 2000, "ccy": "SAR"},
        {"desc": "SWIFT TRANSFER FEE", "type": "debit", "min": 100, "max": 500, "ccy": "SAR"},

        # EUR/GBP transactions
        {"desc": "EUR PAYMENT FROM EUROPEAN CLIENT", "type": "credit", "min": 20000, "max": 100000, "ccy": "EUR"},
        {"desc": "GBP RECEIPT - UK SUBSIDIARY", "type": "credit", "min": 30000, "max": 80000, "ccy": "GBP"},
    ]

    banks = list(SaudiBanks.BANKS.keys())[:5]  # Use first 5 banks
    account_numbers = {
        "SNB": "SA1234567890123456789012",
        "RAJHI": "SA2345678901234567890123",
        "RIYAD": "SA3456789012345678901234",
        "SABB": "SA4567890123456789012345",
        "ANB": "SA5678901234567890123456",
    }

    for day_offset in range(90):
        current_date = base_date + timedelta(days=day_offset)

        # Skip weekends (Friday=4, Saturday=5 in Python)
        if current_date.weekday() in (4, 5):
            continue

        # Generate 1-5 transactions per day
        num_transactions = random.randint(1, 5)

        for _ in range(num_transactions):
            template = random.choice(templates)
            bank = random.choice(banks)
            amount = Decimal(str(random.randint(template["min"], template["max"])))

            # Apply FX rate for non-SAR
            fx_rates = {
                "SAR": Decimal("1.0000"),
                "USD": Decimal("3.7500"),
                "EUR": Decimal("4.0500"),
                "GBP": Decimal("4.7000"),
                "AED": Decimal("1.0200"),
                "SGD": Decimal("2.7800"),
            }
            fx_rate = fx_rates.get(template["ccy"], Decimal("1.0000"))
            sar_amount = amount * fx_rate

            if template["type"] == "debit":
                debit = sar_amount
                credit = Decimal("0")
                running_balance -= sar_amount
            else:
                debit = Decimal("0")
                credit = sar_amount
                running_balance += sar_amount

            transactions.append({
                "Transaction_ID": f"TXN-{tx_id:06d}",
                "Date": current_date,
                "Value_Date": current_date,
                "Reference_No": f"REF{current_date.strftime('%Y%m%d')}{tx_id:04d}",
                "Description": template["desc"],
                "Debit": float(debit) if debit > 0 else None,
                "Credit": float(credit) if credit > 0 else None,
                "Running_Balance": float(running_balance),
                "Bank_Code": bank,
                "Account_No": account_numbers.get(bank, "SA0000000000000000000000"),
                "Original_CCY": template["ccy"],
                "Original_Amount": float(amount),
            })
            tx_id += 1

    return transactions


# =============================================================================
# SECTION 3: SHEET GENERATOR
# =============================================================================

class BankStatementSheetGenerator:
    """
    Generator for Bank_Statement_Actual sheet.

    Creates a formatted Excel sheet with:
    - Transaction data import structure
    - Auto-categorization formulas
    - Currency conversion to SAR
    - Saudi Week ID calculation
    - Period aggregation fields
    """

    SHEET_NAME = "Bank_Statement_Actual"
    TABLE_NAME = "tbl_Bank_Statement"
    TITLE_ROW = 1
    HEADER_ROW = 3
    DATA_START_ROW = 4

    def __init__(self, workbook: Workbook):
        """
        Initialize the Bank Statement sheet generator.

        Args:
            workbook: The openpyxl Workbook to add the sheet to.
        """
        self.workbook = workbook
        self.sheet: Optional[Worksheet] = None
        self.transactions: List[Dict] = []
        self.last_data_row = self.DATA_START_ROW

    def generate(
        self,
        transactions: Optional[List[Dict]] = None,
        include_sample_data: bool = True
    ) -> Worksheet:
        """
        Generate the Bank_Statement_Actual sheet.

        Args:
            transactions: Optional list of transaction dictionaries.
            include_sample_data: Whether to include sample data (default True).

        Returns:
            The generated Worksheet.
        """
        if transactions:
            self.transactions = transactions
        elif include_sample_data:
            self.transactions = generate_sample_transactions()
        else:
            self.transactions = []

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
        self._add_data_validation()
        self._freeze_panes()
        self._add_instructions()

        return self.sheet

    def _write_title(self):
        """Write sheet title and description."""
        self.sheet.cell(
            row=self.TITLE_ROW,
            column=1,
            value="BANK STATEMENT - Actual Transactions"
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
            end_column=len(BANK_STATEMENT_COLUMNS)
        )

        # Subtitle
        self.sheet.cell(
            row=self.TITLE_ROW + 1,
            column=1,
            value="Import bank transactions | Auto-categorization | All amounts in SAR"
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

        for col_idx, col_def in enumerate(BANK_STATEMENT_COLUMNS, start=1):
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
        """Write transaction data rows with formulas."""
        for row_idx, tx in enumerate(self.transactions):
            excel_row = self.DATA_START_ROW + row_idx
            is_alternate = row_idx % 2 == 1

            self._write_transaction_row(excel_row, tx, is_alternate)

        self.last_data_row = self.DATA_START_ROW + len(self.transactions) - 1
        if self.last_data_row < self.DATA_START_ROW:
            self.last_data_row = self.DATA_START_ROW

    def _write_transaction_row(self, row: int, tx: Dict, is_alternate: bool):
        """Write a single transaction row."""
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

        for col_idx, col_def in enumerate(BANK_STATEMENT_COLUMNS, start=1):
            header = col_def["header"]

            # Determine cell value
            if col_def["formula"]:
                # Use formula
                value = col_def["formula"]
            elif header in tx:
                value = tx[header]
            else:
                value = None

            cell = self.sheet.cell(row=row, column=col_idx, value=value)
            cell.font = style["font"]
            cell.fill = style["fill"]
            cell.border = style["border"]

            # Apply alignment
            h_align = col_def["alignment"]
            cell.alignment = Alignment(horizontal=h_align, vertical="center")

            # Apply number format
            cell.number_format = col_def["format"]

            # Color negative amounts in red
            if header in ("Debit", "SAR_Equivalent") and value and not col_def["formula"]:
                if isinstance(value, (int, float)) and value > 0:
                    cell.font = Font(
                        name=Typography.FONT_FAMILY,
                        size=Typography.SIZE_BODY,
                        color=ColorPalette.VALUE_NEGATIVE_RED
                    )

            # Color positive amounts in green
            if header == "Credit" and value and not col_def["formula"]:
                if isinstance(value, (int, float)) and value > 0:
                    cell.font = Font(
                        name=Typography.FONT_FAMILY,
                        size=Typography.SIZE_BODY,
                        color=ColorPalette.VALUE_POSITIVE_GREEN
                    )

    def _apply_formatting(self):
        """Apply conditional formatting and number formats."""
        # Number formats are applied in _write_transaction_row
        pass

    def _create_table(self):
        """Create Excel table for bank statement data."""
        start_col = "A"
        end_col = get_column_letter(len(BANK_STATEMENT_COLUMNS))

        # Ensure at least one data row for table
        table_end_row = max(self.last_data_row, self.DATA_START_ROW)
        table_range = f"{start_col}{self.HEADER_ROW}:{end_col}{table_end_row}"

        table = Table(
            displayName=self.TABLE_NAME,
            ref=table_range
        )

        table_style = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
        )
        table.tableStyleInfo = table_style

        self.sheet.add_table(table)

    def _set_column_widths(self):
        """Set column widths based on definitions."""
        for col_idx, col_def in enumerate(BANK_STATEMENT_COLUMNS, start=1):
            col_letter = get_column_letter(col_idx)
            self.sheet.column_dimensions[col_letter].width = col_def["width"]

        # Set row heights
        self.sheet.row_dimensions[self.TITLE_ROW].height = RowHeights.TITLE
        self.sheet.row_dimensions[self.HEADER_ROW].height = RowHeights.HEADER

    def _add_data_validation(self):
        """Add data validation for input columns."""
        # Bank_Code dropdown
        bank_codes = ",".join(SaudiBanks.BANK_CODES)
        bank_validation = DataValidation(
            type="list",
            formula1=f'"{bank_codes}"',
            allow_blank=True
        )
        bank_col = get_column_letter(9)  # Column I
        bank_validation.add(f"{bank_col}{self.DATA_START_ROW}:{bank_col}5000")
        self.sheet.add_data_validation(bank_validation)

        # Currency dropdown
        ccy_codes = ",".join(CurrencyFramework.CURRENCY_CODES)
        ccy_validation = DataValidation(
            type="list",
            formula1=f'"{ccy_codes}"',
            allow_blank=False
        )
        ccy_col = get_column_letter(11)  # Column K
        ccy_validation.add(f"{ccy_col}{self.DATA_START_ROW}:{ccy_col}5000")
        self.sheet.add_data_validation(ccy_validation)

        # Date validation
        date_validation = DataValidation(
            type="date",
            operator="greaterThan",
            formula1="2020-01-01",
            allow_blank=False
        )
        date_col = get_column_letter(2)  # Column B
        date_validation.add(f"{date_col}{self.DATA_START_ROW}:{date_col}5000")
        self.sheet.add_data_validation(date_validation)

    def _freeze_panes(self):
        """Freeze header rows and key columns."""
        self.sheet.freeze_panes = f"F{self.DATA_START_ROW}"  # Freeze columns A-E and header rows

    def _add_instructions(self):
        """Add usage instructions at the bottom."""
        instructions_row = self.last_data_row + 3

        instructions = [
            "",
            "DATA IMPORT INSTRUCTIONS:",
            "• Paste bank statement data starting from row 4",
            "• Required columns: Date, Description, Debit OR Credit, Original_CCY, Original_Amount",
            "• Formula columns (grayed out) will auto-calculate",
            "",
            "SAUDI CALENDAR:",
            "• Week starts Sunday, ends Saturday",
            "• Weekend: Friday & Saturday (WE)",
            "• Payment Day: Wednesday (WED)",
            "• Week_ID: 01-52 based on Saudi calendar",
            "",
            "CURRENCY CONVERSION:",
            "• Base currency: SAR (Saudi Riyal)",
            "• FX rates lookup from FX_Rates sheet",
            "• All SAR_Equivalent values for reporting",
            "",
            "AUTO-CATEGORIZATION:",
            "• Category_Code auto-populated from Category_Map patterns",
            "• Review UNCATEGORIZED transactions and update Category_Map",
        ]

        for idx, text in enumerate(instructions):
            cell = self.sheet.cell(row=instructions_row + idx, column=1, value=text)
            if text.startswith("DATA IMPORT") or text.startswith("SAUDI") or text.startswith("CURRENCY") or text.startswith("AUTO-"):
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

def generate_bank_statement_sheet(
    workbook: Workbook,
    include_sample_data: bool = True
) -> Worksheet:
    """
    Convenience function to generate Bank_Statement_Actual sheet.

    Args:
        workbook: The openpyxl Workbook.
        include_sample_data: Whether to include sample transactions.

    Returns:
        The generated Worksheet.
    """
    generator = BankStatementSheetGenerator(workbook)
    return generator.generate(include_sample_data=include_sample_data)


def get_week_id_formula() -> str:
    """
    Get the Saudi Week ID calculation formula.

    Saudi calendar: Week starts Sunday (day 7 in Excel WEEKDAY with type 2).

    Returns:
        Excel formula string for Week ID calculation.
    """
    return '=IFERROR(INT(([@Date]-DATE(YEAR([@Date]),1,1)-MOD(WEEKDAY(DATE(YEAR([@Date]),1,1))+1,7)+1)/7)+1,"")'


def get_day_type_formula() -> str:
    """
    Get the day type classification formula.

    W = Working day (Sunday-Thursday)
    WE = Weekend (Friday-Saturday)
    WED = Payment day (Wednesday)

    Returns:
        Excel formula string for day type classification.
    """
    return '=IFERROR(IF(WEEKDAY([@Date],2)=3,"WED",IF(OR(WEEKDAY([@Date],2)=5,WEEKDAY([@Date],2)=6),"WE","W")),"")'
