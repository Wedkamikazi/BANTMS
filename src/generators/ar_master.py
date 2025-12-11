"""
BANTMS AR Master Sheet Generator
=================================

Generates the AR_Master (Accounts Receivable) sheet with:
- Expected collections tracking
- Collection probability weighting
- Currency conversion to SAR
- Saudi Week ID calculation
- Aging analysis fields
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
    CollectionStatus,
    ConfirmationStatus
)
from ..config.styles import (
    ColorPalette, Typography, NumberFormats,
    ColumnWidths, RowHeights
)


# =============================================================================
# SECTION 1: COLUMN DEFINITIONS
# =============================================================================

AR_MASTER_COLUMNS: List[Dict] = [
    {
        "header": "Collection_ID",
        "width": 14,
        "format": "@",
        "alignment": "center",
        "description": "Unique collection identifier",
        "formula": None
    },
    {
        "header": "Customer_ID",
        "width": 12,
        "format": "@",
        "alignment": "center",
        "description": "Customer reference code",
        "formula": None
    },
    {
        "header": "Customer_Name",
        "width": 30,
        "format": "@",
        "alignment": "left",
        "description": "Customer company name",
        "formula": None
    },
    {
        "header": "Customer_Tier",
        "width": 12,
        "format": "0",
        "alignment": "center",
        "description": "1=Strategic, 2=Regular, 3=Small",
        "formula": None
    },
    {
        "header": "Invoice_No",
        "width": 16,
        "format": "@",
        "alignment": "left",
        "description": "Invoice number",
        "formula": None
    },
    {
        "header": "Invoice_Date",
        "width": 12,
        "format": "DD-MMM-YYYY",
        "alignment": "center",
        "description": "Invoice date",
        "formula": None
    },
    {
        "header": "Expected_Date",
        "width": 14,
        "format": "DD-MMM-YYYY",
        "alignment": "center",
        "description": "Expected collection date",
        "formula": None
    },
    {
        "header": "Amount_Original_CCY",
        "width": 16,
        "format": "#,##0.00",
        "alignment": "right",
        "description": "Amount in original currency",
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
        "header": "FX_Rate",
        "width": 10,
        "format": "#,##0.0000",
        "alignment": "right",
        "description": "Exchange rate to SAR",
        "formula": '=IFERROR(IF([@CCY]="SAR",1,INDEX(FX_Rate_Lookup,MATCH([@CCY],CCY_Codes,0),3)),1)'
    },
    {
        "header": "Amount_SAR",
        "width": 16,
        "format": "#,##0.00",
        "alignment": "right",
        "description": "Amount in SAR",
        "formula": '=IFERROR([@Amount_Original_CCY]*[@FX_Rate],0)'
    },
    {
        "header": "Collection_Probability",
        "width": 14,
        "format": "0%",
        "alignment": "center",
        "description": "Probability of collection (0-100%)",
        "formula": None
    },
    {
        "header": "Weighted_Amount_SAR",
        "width": 16,
        "format": "#,##0.00",
        "alignment": "right",
        "description": "Probability-weighted amount",
        "formula": '=IFERROR([@Amount_SAR]*[@Collection_Probability],0)'
    },
    {
        "header": "Status",
        "width": 12,
        "format": "@",
        "alignment": "center",
        "description": "Expected/Received/Delayed/Disputed",
        "formula": None
    },
    {
        "header": "Actual_Receipt_Date",
        "width": 14,
        "format": "DD-MMM-YYYY",
        "alignment": "center",
        "description": "Actual receipt date (when received)",
        "formula": None
    },
    {
        "header": "Week_ID",
        "width": 10,
        "format": "00",
        "alignment": "center",
        "description": "Saudi week number for Expected_Date",
        "formula": '=IFERROR(INT(([@Expected_Date]-DATE(YEAR([@Expected_Date]),1,1)-MOD(WEEKDAY(DATE(YEAR([@Expected_Date]),1,1))+1,7)+1)/7)+1,"")'
    },
    {
        "header": "Days_Until_Due",
        "width": 12,
        "format": "0",
        "alignment": "center",
        "description": "Days until expected date",
        "formula": '=IFERROR([@Expected_Date]-TODAY(),"")'
    },
    {
        "header": "Aging_Bucket",
        "width": 12,
        "format": "@",
        "alignment": "center",
        "description": "Aging category",
        "formula": '=IFERROR(IF([@Days_Until_Due]>30,"Future",IF([@Days_Until_Due]>0,"Due Soon",IF([@Days_Until_Due]>-30,"0-30 Days",IF([@Days_Until_Due]>-60,"31-60 Days",IF([@Days_Until_Due]>-90,"61-90 Days","90+ Days"))))),"")'
    },
    {
        "header": "Confirmation",
        "width": 12,
        "format": "@",
        "alignment": "center",
        "description": "Confirmed/Estimated/Tentative",
        "formula": None
    },
    {
        "header": "Payment_Method",
        "width": 14,
        "format": "@",
        "alignment": "center",
        "description": "Expected payment method",
        "formula": None
    },
    {
        "header": "Notes",
        "width": 30,
        "format": "@",
        "alignment": "left",
        "description": "Additional notes",
        "formula": None
    },
]


# =============================================================================
# SECTION 2: SAMPLE DATA GENERATOR
# =============================================================================

def generate_sample_ar_data() -> List[Dict]:
    """
    Generate sample AR data for demonstration.

    Returns:
        List of collection dictionaries with realistic Saudi business data.
    """
    # Sample customers by tier
    customers = {
        1: [  # Tier 1 - Strategic
            {"id": "C-001", "name": "Saudi Aramco", "prob": 0.98},
            {"id": "C-002", "name": "SABIC", "prob": 0.95},
            {"id": "C-003", "name": "Saudi Telecom Company", "prob": 0.95},
            {"id": "C-004", "name": "Ma'aden Mining", "prob": 0.92},
            {"id": "C-005", "name": "ACWA Power", "prob": 0.90},
        ],
        2: [  # Tier 2 - Regular
            {"id": "C-010", "name": "Al Faisaliah Group", "prob": 0.85},
            {"id": "C-011", "name": "Bin Laden Group", "prob": 0.82},
            {"id": "C-012", "name": "Abdul Latif Jameel", "prob": 0.85},
            {"id": "C-013", "name": "Al Muhaidib Group", "prob": 0.80},
            {"id": "C-014", "name": "Olayan Group", "prob": 0.88},
        ],
        3: [  # Tier 3 - Small
            {"id": "C-020", "name": "Local Contractor A", "prob": 0.70},
            {"id": "C-021", "name": "Trading Company B", "prob": 0.65},
            {"id": "C-022", "name": "Services Provider C", "prob": 0.72},
            {"id": "C-023", "name": "Retail Store D", "prob": 0.68},
            {"id": "C-024", "name": "Small Business E", "prob": 0.60},
        ],
    }

    collections = []
    base_date = date.today() - timedelta(days=60)
    collection_id = 1

    payment_methods = ["Bank Transfer", "Check", "Wire Transfer", "LC"]

    for tier, tier_customers in customers.items():
        for customer in tier_customers:
            # Generate 2-5 invoices per customer
            num_invoices = random.randint(2, 5)

            for i in range(num_invoices):
                invoice_date = base_date + timedelta(days=random.randint(0, 90))
                expected_date = invoice_date + timedelta(days=random.choice([30, 45, 60, 90]))

                # Amount varies by tier
                if tier == 1:
                    amount = random.randint(500000, 3000000)
                elif tier == 2:
                    amount = random.randint(100000, 800000)
                else:
                    amount = random.randint(20000, 150000)

                # Currency (mostly SAR, some USD for tier 1)
                if tier == 1 and random.random() < 0.4:
                    ccy = random.choice(["USD", "EUR"])
                else:
                    ccy = "SAR"

                # Probability based on customer and randomness
                base_prob = customer["prob"]
                prob = max(0.5, min(1.0, base_prob + random.uniform(-0.1, 0.05)))

                # Status distribution
                if expected_date < date.today():
                    if random.random() < 0.7:
                        status = "Received"
                        actual_date = expected_date + timedelta(days=random.randint(-5, 10))
                    else:
                        status = random.choice(["Delayed", "Disputed"])
                        actual_date = None
                else:
                    status = "Expected"
                    actual_date = None

                collections.append({
                    "Collection_ID": f"AR-{collection_id:06d}",
                    "Customer_ID": customer["id"],
                    "Customer_Name": customer["name"],
                    "Customer_Tier": tier,
                    "Invoice_No": f"SI-{invoice_date.strftime('%Y%m')}-{collection_id:04d}",
                    "Invoice_Date": invoice_date,
                    "Expected_Date": expected_date,
                    "Amount_Original_CCY": float(amount),
                    "CCY": ccy,
                    "Collection_Probability": prob,
                    "Status": status,
                    "Actual_Receipt_Date": actual_date,
                    "Confirmation": random.choice(["Confirmed", "Estimated", "Tentative"]),
                    "Payment_Method": random.choice(payment_methods),
                    "Notes": "",
                })
                collection_id += 1

    return collections


# =============================================================================
# SECTION 3: SHEET GENERATOR
# =============================================================================

class ARMasterSheetGenerator:
    """
    Generator for AR_Master (Accounts Receivable) sheet.

    Creates a formatted Excel sheet with:
    - Expected collection tracking
    - Probability-weighted amounts
    - Currency conversion to SAR
    - Aging analysis
    """

    SHEET_NAME = "AR_Master"
    TABLE_NAME = "tbl_AR_Master"
    TITLE_ROW = 1
    HEADER_ROW = 3
    DATA_START_ROW = 4

    def __init__(self, workbook: Workbook):
        """
        Initialize the AR Master sheet generator.

        Args:
            workbook: The openpyxl Workbook to add the sheet to.
        """
        self.workbook = workbook
        self.sheet: Optional[Worksheet] = None
        self.collections: List[Dict] = []
        self.last_data_row = self.DATA_START_ROW

    def generate(
        self,
        collections: Optional[List[Dict]] = None,
        include_sample_data: bool = True
    ) -> Worksheet:
        """
        Generate the AR_Master sheet.

        Args:
            collections: Optional list of collection dictionaries.
            include_sample_data: Whether to include sample data (default True).

        Returns:
            The generated Worksheet.
        """
        if collections:
            self.collections = collections
        elif include_sample_data:
            self.collections = generate_sample_ar_data()
        else:
            self.collections = []

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
        self._add_instructions()

        return self.sheet

    def _write_title(self):
        """Write sheet title and description."""
        self.sheet.cell(
            row=self.TITLE_ROW,
            column=1,
            value="AR MASTER - Accounts Receivable Management"
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
            end_column=len(AR_MASTER_COLUMNS)
        )

        # Subtitle
        self.sheet.cell(
            row=self.TITLE_ROW + 1,
            column=1,
            value="Expected Collections | Probability Weighting | All amounts in SAR"
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

        for col_idx, col_def in enumerate(AR_MASTER_COLUMNS, start=1):
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
        """Write collection data rows with formulas."""
        for row_idx, collection in enumerate(self.collections):
            excel_row = self.DATA_START_ROW + row_idx
            is_alternate = row_idx % 2 == 1

            self._write_collection_row(excel_row, collection, is_alternate)

        self.last_data_row = self.DATA_START_ROW + len(self.collections) - 1
        if self.last_data_row < self.DATA_START_ROW:
            self.last_data_row = self.DATA_START_ROW

    def _write_collection_row(self, row: int, collection: Dict, is_alternate: bool):
        """Write a single collection row."""
        fill_color = ColorPalette.ROW_LIGHT_GRAY if is_alternate else ColorPalette.ROW_WHITE

        # Status-based styling
        status = collection.get("Status", "Expected")
        status_colors = {
            "Received": ColorPalette.STATUS_PAID_GREEN,
            "Expected": fill_color,
            "Delayed": ColorPalette.STATUS_POSTPONED_1_ORANGE,
            "Disputed": ColorPalette.STATUS_POSTPONED_2_RED,
        }
        row_fill_color = status_colors.get(status, fill_color)

        style = {
            "font": Font(
                name=Typography.FONT_FAMILY,
                size=Typography.SIZE_BODY,
                color=ColorPalette.TEXT_BLACK
            ),
            "fill": PatternFill(
                start_color=row_fill_color,
                end_color=row_fill_color,
                fill_type="solid"
            ),
            "border": Border(
                left=Side(style="thin", color=ColorPalette.BORDER_THIN_GRAY),
                right=Side(style="thin", color=ColorPalette.BORDER_THIN_GRAY),
                top=Side(style="thin", color=ColorPalette.BORDER_THIN_GRAY),
                bottom=Side(style="thin", color=ColorPalette.BORDER_THIN_GRAY)
            )
        }

        for col_idx, col_def in enumerate(AR_MASTER_COLUMNS, start=1):
            header = col_def["header"]

            # Determine cell value
            if col_def["formula"]:
                value = col_def["formula"]
            elif header in collection:
                value = collection[header]
            else:
                value = None

            cell = self.sheet.cell(row=row, column=col_idx, value=value)
            cell.font = style["font"]
            cell.fill = style["fill"]
            cell.border = style["border"]

            h_align = col_def["alignment"]
            cell.alignment = Alignment(horizontal=h_align, vertical="center")
            cell.number_format = col_def["format"]

            # Color amounts green for positive cash flow
            if header in ("Amount_SAR", "Weighted_Amount_SAR") and not col_def["formula"]:
                if isinstance(value, (int, float)) and value > 0:
                    cell.font = Font(
                        name=Typography.FONT_FAMILY,
                        size=Typography.SIZE_BODY,
                        color=ColorPalette.VALUE_POSITIVE_GREEN
                    )

    def _create_table(self):
        """Create Excel table for AR data."""
        start_col = "A"
        end_col = get_column_letter(len(AR_MASTER_COLUMNS))

        table_end_row = max(self.last_data_row, self.DATA_START_ROW)
        table_range = f"{start_col}{self.HEADER_ROW}:{end_col}{table_end_row}"

        table = Table(
            displayName=self.TABLE_NAME,
            ref=table_range
        )

        table_style = TableStyleInfo(
            name="TableStyleMedium4",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
        )
        table.tableStyleInfo = table_style

        self.sheet.add_table(table)

    def _set_column_widths(self):
        """Set column widths based on definitions."""
        for col_idx, col_def in enumerate(AR_MASTER_COLUMNS, start=1):
            col_letter = get_column_letter(col_idx)
            self.sheet.column_dimensions[col_letter].width = col_def["width"]

        self.sheet.row_dimensions[self.TITLE_ROW].height = RowHeights.TITLE
        self.sheet.row_dimensions[self.HEADER_ROW].height = RowHeights.HEADER

    def _add_data_validation(self):
        """Add data validation for input columns."""
        # Customer Tier dropdown
        tier_validation = DataValidation(
            type="list",
            formula1='"1,2,3"',
            allow_blank=False
        )
        tier_col = get_column_letter(4)
        tier_validation.add(f"{tier_col}{self.DATA_START_ROW}:{tier_col}5000")
        self.sheet.add_data_validation(tier_validation)

        # Currency dropdown
        ccy_codes = ",".join(CurrencyFramework.CURRENCY_CODES)
        ccy_validation = DataValidation(
            type="list",
            formula1=f'"{ccy_codes}"',
            allow_blank=False
        )
        ccy_col = get_column_letter(9)
        ccy_validation.add(f"{ccy_col}{self.DATA_START_ROW}:{ccy_col}5000")
        self.sheet.add_data_validation(ccy_validation)

        # Probability validation (0-100%)
        prob_validation = DataValidation(
            type="decimal",
            operator="between",
            formula1="0",
            formula2="1",
            allow_blank=False
        )
        prob_col = get_column_letter(12)
        prob_validation.add(f"{prob_col}{self.DATA_START_ROW}:{prob_col}5000")
        self.sheet.add_data_validation(prob_validation)

        # Status dropdown
        status_options = ",".join([s.value for s in CollectionStatus])
        status_validation = DataValidation(
            type="list",
            formula1=f'"{status_options}"',
            allow_blank=False
        )
        status_col = get_column_letter(14)
        status_validation.add(f"{status_col}{self.DATA_START_ROW}:{status_col}5000")
        self.sheet.add_data_validation(status_validation)

        # Confirmation dropdown
        confirm_options = ",".join([c.value for c in ConfirmationStatus])
        confirm_validation = DataValidation(
            type="list",
            formula1=f'"{confirm_options}"',
            allow_blank=False
        )
        confirm_col = get_column_letter(19)
        confirm_validation.add(f"{confirm_col}{self.DATA_START_ROW}:{confirm_col}5000")
        self.sheet.add_data_validation(confirm_validation)

    def _freeze_panes(self):
        """Freeze header rows and key columns."""
        self.sheet.freeze_panes = f"D{self.DATA_START_ROW}"

    def _add_instructions(self):
        """Add usage instructions at the bottom."""
        instructions_row = self.last_data_row + 3

        instructions = [
            "",
            "COLLECTION PROBABILITY:",
            "• Enter probability as decimal (0.95 = 95%)",
            "• Weighted_Amount_SAR = Amount_SAR * Probability",
            "• Use for cash flow forecasting",
            "",
            "AGING BUCKETS:",
            "• Future: >30 days from today",
            "• Due Soon: 0-30 days from today",
            "• 0-30 Days: Overdue 0-30 days",
            "• 31-60 Days: Overdue 31-60 days",
            "• 61-90 Days: Overdue 61-90 days",
            "• 90+ Days: Overdue >90 days",
            "",
            "STATUS CODES:",
            "• Expected: Awaiting collection",
            "• Received: Payment received",
            "• Delayed: Past due date",
            "• Disputed: Under dispute",
        ]

        for idx, text in enumerate(instructions):
            cell = self.sheet.cell(row=instructions_row + idx, column=1, value=text)
            if text.startswith("COLLECTION") or text.startswith("AGING") or text.startswith("STATUS"):
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

def generate_ar_master_sheet(
    workbook: Workbook,
    include_sample_data: bool = True
) -> Worksheet:
    """
    Convenience function to generate AR_Master sheet.

    Args:
        workbook: The openpyxl Workbook.
        include_sample_data: Whether to include sample data.

    Returns:
        The generated Worksheet.
    """
    generator = ARMasterSheetGenerator(workbook)
    return generator.generate(include_sample_data=include_sample_data)


def get_weighted_amount_formula() -> str:
    """
    Get the weighted amount formula for probability-adjusted forecasting.

    Returns:
        Excel formula string for weighted amount calculation.
    """
    return '=IFERROR([@Amount_SAR]*[@Collection_Probability],0)'


def get_aging_bucket_formula() -> str:
    """
    Get the aging bucket classification formula.

    Returns:
        Excel formula string for aging bucket classification.
    """
    return (
        '=IFERROR('
        'IF([@Days_Until_Due]>30,"Future",'
        'IF([@Days_Until_Due]>0,"Due Soon",'
        'IF([@Days_Until_Due]>-30,"0-30 Days",'
        'IF([@Days_Until_Due]>-60,"31-60 Days",'
        'IF([@Days_Until_Due]>-90,"61-90 Days","90+ Days"))))),"")'
    )
