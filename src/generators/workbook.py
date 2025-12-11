"""
BANTMS Workbook Generator
==========================

Main entry point for generating the complete Cash Flow Management Excel workbook.

Creates all sheets in the correct order:
1. FX_Rates (reference data)
2. Category_Map (categorization rules)
3. Bank_Statement_Actual (actual transactions)
4. AP_Master (accounts payable)
5. AR_Master (accounts receivable)
6. TD_Register (time deposits)
7. Cash_Flow_Statement (summary reporting)

All amounts are converted to SAR (base currency).
Saudi calendar is used throughout (Sunday start, Fri/Sat weekend).
"""

from datetime import datetime
from pathlib import Path
from typing import Optional, List, Dict

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.worksheet import Worksheet

from .fx_rates import FXRatesSheetGenerator
from .category_map import CategoryMapSheetGenerator
from .bank_statement import BankStatementSheetGenerator
from .ap_master import APMasterSheetGenerator
from .ar_master import ARMasterSheetGenerator
from .td_register import TDRegisterSheetGenerator
from .cash_flow_statement import CashFlowStatementGenerator

from ..config.styles import ColorPalette, Typography


# =============================================================================
# SECTION 1: WORKBOOK CONFIGURATION
# =============================================================================

class WorkbookConfig:
    """Configuration for the Cash Flow workbook."""

    # Workbook metadata
    TITLE: str = "BANTMS Cash Flow Management System"
    AUTHOR: str = "Treasury Management Team"
    VERSION: str = "1.0.0"

    # Sheet order
    SHEET_ORDER: List[str] = [
        "Dashboard",
        "Cash_Flow_Statement",
        "Treasury_Calendar",
        "Bank_Statement_Actual",
        "AP_Master",
        "AR_Master",
        "TD_Register",
        "FX_Rates",
        "Category_Map",
    ]

    # Default file name
    DEFAULT_FILENAME: str = "BANTMS_CashFlow_{date}.xlsx"


# =============================================================================
# SECTION 2: DASHBOARD GENERATOR
# =============================================================================

class DashboardSheetGenerator:
    """
    Generator for the Dashboard sheet.

    Contains:
    - Period selector
    - Key metrics summary
    - Navigation links
    """

    SHEET_NAME = "Dashboard"

    def __init__(self, workbook: Workbook):
        self.workbook = workbook
        self.sheet: Optional[Worksheet] = None

    def generate(self) -> Worksheet:
        """Generate the Dashboard sheet."""
        if self.SHEET_NAME in self.workbook.sheetnames:
            self.sheet = self.workbook[self.SHEET_NAME]
        else:
            self.sheet = self.workbook.create_sheet(self.SHEET_NAME, 0)

        self._write_header()
        self._write_period_selector()
        self._write_key_metrics()
        self._write_navigation()
        self._set_column_widths()

        return self.sheet

    def _write_header(self):
        """Write dashboard header."""
        self.sheet.cell(row=1, column=1, value="BANTMS CASH FLOW MANAGEMENT")
        header_cell = self.sheet.cell(row=1, column=1)
        header_cell.font = Font(
            name=Typography.FONT_FAMILY,
            size=20,
            bold=True,
            color=ColorPalette.HEADER_DARK_NAVY
        )

        self.sheet.merge_cells("A1:H1")

        # Subtitle
        self.sheet.cell(
            row=2, column=1,
            value="Saudi Treasury Management | Base Currency: SAR | Week Starts Sunday"
        )
        self.sheet.cell(row=2, column=1).font = Font(
            name=Typography.FONT_FAMILY,
            size=12,
            italic=True,
            color=ColorPalette.TEXT_DARK_GRAY
        )

        # Generation timestamp
        self.sheet.cell(
            row=3, column=1,
            value=f"Generated: {datetime.now().strftime('%d-%b-%Y %H:%M')}"
        )
        self.sheet.cell(row=3, column=1).font = Font(
            name=Typography.FONT_FAMILY,
            size=10,
            color=ColorPalette.TEXT_DARK_GRAY
        )

    def _write_period_selector(self):
        """Write period selector section."""
        # Section header
        self.sheet.cell(row=5, column=1, value="PERIOD SELECTOR")
        self.sheet.cell(row=5, column=1).font = Font(
            name=Typography.FONT_FAMILY,
            size=14,
            bold=True,
            color=ColorPalette.TEXT_WHITE
        )
        self.sheet.cell(row=5, column=1).fill = PatternFill(
            start_color=ColorPalette.HEADER_DARK_NAVY,
            end_color=ColorPalette.HEADER_DARK_NAVY,
            fill_type="solid"
        )
        self.sheet.merge_cells("A5:B5")

        # Period type
        self.sheet.cell(row=6, column=1, value="Period Type:")
        self.sheet.cell(row=6, column=2, value="Monthly")
        self.sheet.cell(row=6, column=2).font = Font(bold=True)

        # Selected period
        self.sheet.cell(row=7, column=1, value="Selected Period:")
        self.sheet.cell(row=7, column=2, value=datetime.now().strftime("%b-%Y"))
        self.sheet.cell(row=7, column=2).font = Font(bold=True)

        # Create named cell for period selector
        self.workbook.defined_names.add(
            self.workbook.defined_names.new(
                name="Selected_Period",
                value=f"'{self.SHEET_NAME}'!$B$7"
            )
        )

    def _write_key_metrics(self):
        """Write key metrics summary section."""
        # Section header
        self.sheet.cell(row=9, column=1, value="KEY METRICS (SAR)")
        self.sheet.cell(row=9, column=1).font = Font(
            name=Typography.FONT_FAMILY,
            size=14,
            bold=True,
            color=ColorPalette.TEXT_WHITE
        )
        self.sheet.cell(row=9, column=1).fill = PatternFill(
            start_color=ColorPalette.HEADER_DARK_NAVY,
            end_color=ColorPalette.HEADER_DARK_NAVY,
            fill_type="solid"
        )
        self.sheet.merge_cells("A9:D9")

        # Metrics
        metrics = [
            ("Current Cash Balance:", '=IFERROR(SUMIF(tbl_Bank_Statement[CF_Section],"Operating",tbl_Bank_Statement[SAR_Equivalent]),0)', "right"),
            ("Pending AP (This Week):", '=IFERROR(SUMIFS(tbl_AP_Master[Amount_SAR],tbl_AP_Master[Status],"Pending",tbl_AP_Master[Week_ID],INT((TODAY()-DATE(YEAR(TODAY()),1,1))/7)+1),0)', "right"),
            ("Expected AR (This Week):", '=IFERROR(SUMIFS(tbl_AR_Master[Weighted_Amount_SAR],tbl_AR_Master[Status],"Expected",tbl_AR_Master[Week_ID],INT((TODAY()-DATE(YEAR(TODAY()),1,1))/7)+1),0)', "right"),
            ("Active TD Portfolio:", '=IFERROR(SUMIF(tbl_TD_Register[Status],"Active",tbl_TD_Register[Principal_SAR]),0)', "right"),
            ("TDs Maturing This Week:", '=IFERROR(COUNTIFS(tbl_TD_Register[Status],"Active",tbl_TD_Register[Days_To_Maturity],"<=7",tbl_TD_Register[Days_To_Maturity],">=0"),0)', "right"),
        ]

        for idx, (label, formula, align) in enumerate(metrics):
            row = 10 + idx
            self.sheet.cell(row=row, column=1, value=label)
            self.sheet.cell(row=row, column=1).font = Font(bold=True)

            self.sheet.cell(row=row, column=2, value=formula)
            self.sheet.cell(row=row, column=2).number_format = "#,##0.00"
            self.sheet.cell(row=row, column=2).font = Font(
                color=ColorPalette.VALUE_POSITIVE_GREEN
            )

    def _write_navigation(self):
        """Write navigation links section."""
        # Section header
        self.sheet.cell(row=17, column=1, value="QUICK NAVIGATION")
        self.sheet.cell(row=17, column=1).font = Font(
            name=Typography.FONT_FAMILY,
            size=14,
            bold=True,
            color=ColorPalette.TEXT_WHITE
        )
        self.sheet.cell(row=17, column=1).fill = PatternFill(
            start_color=ColorPalette.HEADER_DARK_NAVY,
            end_color=ColorPalette.HEADER_DARK_NAVY,
            fill_type="solid"
        )
        self.sheet.merge_cells("A17:B17")

        # Navigation links
        links = [
            ("Cash Flow Statement", "Cash_Flow_Statement"),
            ("Bank Statement", "Bank_Statement_Actual"),
            ("Accounts Payable", "AP_Master"),
            ("Accounts Receivable", "AR_Master"),
            ("Time Deposits", "TD_Register"),
            ("FX Rates", "FX_Rates"),
            ("Category Map", "Category_Map"),
        ]

        for idx, (label, sheet_name) in enumerate(links):
            row = 18 + idx
            self.sheet.cell(row=row, column=1, value=f"• {label}")
            # Note: Hyperlinks to sheets can be added here if needed

    def _set_column_widths(self):
        """Set column widths."""
        self.sheet.column_dimensions["A"].width = 25
        self.sheet.column_dimensions["B"].width = 20
        self.sheet.column_dimensions["C"].width = 15
        self.sheet.column_dimensions["D"].width = 15


# =============================================================================
# SECTION 3: MAIN WORKBOOK GENERATOR
# =============================================================================

class CashFlowWorkbookGenerator:
    """
    Main generator for the complete Cash Flow Management workbook.

    Creates all sheets and ensures proper linking and formulas.
    """

    def __init__(self):
        """Initialize the workbook generator."""
        self.workbook: Optional[Workbook] = None
        self.include_sample_data: bool = True

    def generate(
        self,
        include_sample_data: bool = True,
        output_path: Optional[str] = None
    ) -> str:
        """
        Generate the complete Cash Flow workbook.

        Args:
            include_sample_data: Whether to include sample/demo data.
            output_path: Optional output file path. Auto-generates if not provided.

        Returns:
            Path to the generated Excel file.
        """
        self.include_sample_data = include_sample_data

        # Create new workbook
        self.workbook = Workbook()

        # Remove default sheet
        if "Sheet" in self.workbook.sheetnames:
            del self.workbook["Sheet"]

        # Generate sheets in order
        self._generate_reference_sheets()
        self._generate_source_sheets()
        self._generate_summary_sheets()

        # Set document properties
        self._set_document_properties()

        # Reorder sheets
        self._reorder_sheets()

        # Generate output path
        if not output_path:
            output_path = self._generate_output_path()

        # Save workbook
        self.workbook.save(output_path)

        return output_path

    def _generate_reference_sheets(self):
        """Generate reference/lookup sheets."""
        print("Generating FX_Rates sheet...")
        fx_generator = FXRatesSheetGenerator(self.workbook)
        fx_generator.generate()

        print("Generating Category_Map sheet...")
        cat_generator = CategoryMapSheetGenerator(self.workbook)
        cat_generator.generate()

    def _generate_source_sheets(self):
        """Generate source data sheets."""
        print("Generating Bank_Statement_Actual sheet...")
        bank_generator = BankStatementSheetGenerator(self.workbook)
        bank_generator.generate(include_sample_data=self.include_sample_data)

        print("Generating AP_Master sheet...")
        ap_generator = APMasterSheetGenerator(self.workbook)
        ap_generator.generate(include_sample_data=self.include_sample_data)

        print("Generating AR_Master sheet...")
        ar_generator = ARMasterSheetGenerator(self.workbook)
        ar_generator.generate(include_sample_data=self.include_sample_data)

        print("Generating TD_Register sheet...")
        td_generator = TDRegisterSheetGenerator(self.workbook)
        td_generator.generate(include_sample_data=self.include_sample_data)

    def _generate_summary_sheets(self):
        """Generate summary/dashboard sheets."""
        print("Generating Dashboard sheet...")
        dash_generator = DashboardSheetGenerator(self.workbook)
        dash_generator.generate()

        print("Generating Cash_Flow_Statement sheet...")
        cf_generator = CashFlowStatementGenerator(self.workbook)
        cf_generator.generate()

    def _set_document_properties(self):
        """Set workbook document properties."""
        self.workbook.properties.title = WorkbookConfig.TITLE
        self.workbook.properties.creator = WorkbookConfig.AUTHOR
        self.workbook.properties.description = (
            "Enterprise Cash Flow Management System for Saudi Treasury Operations. "
            "Base currency: SAR. Week starts Sunday. Generated by BANTMS."
        )
        self.workbook.properties.version = WorkbookConfig.VERSION

    def _reorder_sheets(self):
        """Reorder sheets according to configuration."""
        # Get current sheet names
        current_sheets = self.workbook.sheetnames

        # Determine target order based on existing sheets
        target_order = []
        for sheet_name in WorkbookConfig.SHEET_ORDER:
            if sheet_name in current_sheets:
                target_order.append(sheet_name)

        # Add any remaining sheets not in the config
        for sheet_name in current_sheets:
            if sheet_name not in target_order:
                target_order.append(sheet_name)

        # Reorder
        for idx, sheet_name in enumerate(target_order):
            if sheet_name in self.workbook.sheetnames:
                current_idx = self.workbook.sheetnames.index(sheet_name)
                self.workbook.move_sheet(sheet_name, offset=idx - current_idx)

    def _generate_output_path(self) -> str:
        """Generate default output file path."""
        date_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = WorkbookConfig.DEFAULT_FILENAME.format(date=date_str)
        return filename


# =============================================================================
# SECTION 4: CONVENIENCE FUNCTIONS
# =============================================================================

def generate_cashflow_workbook(
    output_path: Optional[str] = None,
    include_sample_data: bool = True
) -> str:
    """
    Generate a complete Cash Flow Management workbook.

    Args:
        output_path: Optional path for the output file.
        include_sample_data: Whether to include sample/demo data.

    Returns:
        Path to the generated Excel file.
    """
    generator = CashFlowWorkbookGenerator()
    return generator.generate(
        include_sample_data=include_sample_data,
        output_path=output_path
    )


def main():
    """Main entry point for command-line usage."""
    import argparse

    parser = argparse.ArgumentParser(
        description="Generate BANTMS Cash Flow Management Excel Workbook"
    )
    parser.add_argument(
        "-o", "--output",
        help="Output file path",
        default=None
    )
    parser.add_argument(
        "--no-sample-data",
        action="store_true",
        help="Generate without sample data"
    )

    args = parser.parse_args()

    print("=" * 60)
    print("BANTMS Cash Flow Workbook Generator")
    print("=" * 60)
    print(f"Base Currency: SAR (Saudi Riyal)")
    print(f"Supported Currencies: USD, AED, EUR, GBP, SGD")
    print(f"Calendar: Saudi (Sunday start, Fri/Sat weekend)")
    print("=" * 60)

    output_path = generate_cashflow_workbook(
        output_path=args.output,
        include_sample_data=not args.no_sample_data
    )

    print("=" * 60)
    print(f"Workbook generated successfully: {output_path}")
    print("=" * 60)


if __name__ == "__main__":
    main()
