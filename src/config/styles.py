"""
BANTMS Styling Configuration Module
====================================

Contains all Excel styling specifications:
- Color Palette
- Typography Settings
- Number Formats
- Conditional Formatting Rules
- Border Styles
"""

from typing import Dict, NamedTuple, Optional
from dataclasses import dataclass


# =============================================================================
# SECTION 1: COLOR PALETTE
# =============================================================================

class ColorPalette:
    """
    Enterprise Color Palette for Excel Workbook

    Based on professional treasury/finance standards.
    """

    # Primary Header Colors
    HEADER_DARK_NAVY: str = "1E3A5F"      # Dark navy for main headers
    HEADER_MEDIUM_BLUE: str = "3D5A80"    # Medium blue for sub-headers
    HEADER_LIGHT_BLUE: str = "E3F2FD"     # Light blue for totals row

    # Text Colors
    TEXT_WHITE: str = "FFFFFF"
    TEXT_BLACK: str = "000000"
    TEXT_DARK_GRAY: str = "333333"

    # Value Colors (for amounts)
    VALUE_POSITIVE_GREEN: str = "2E7D32"   # Dark green for positive
    VALUE_NEGATIVE_RED: str = "C62828"     # Dark red for negative
    VALUE_NEUTRAL_BLACK: str = "000000"

    # Row Alternation
    ROW_LIGHT_GRAY: str = "F5F5F5"        # Light gray for alternating rows
    ROW_WHITE: str = "FFFFFF"

    # Border Colors
    BORDER_THIN_GRAY: str = "D0D0D0"      # Thin gray for grid
    BORDER_MEDIUM_BLACK: str = "000000"    # Medium black for sections

    # Status Colors (for payment/collection status)
    STATUS_PAID_GREEN: str = "C8E6C9"      # Light green background
    STATUS_PENDING_AMBER: str = "FFF3E0"   # Light amber background
    STATUS_POSTPONED_1_ORANGE: str = "FFE0B2"  # Light orange
    STATUS_POSTPONED_2_RED: str = "FFCDD2"     # Light red
    STATUS_URGENT_PURPLE: str = "E1BEE7"       # Light purple

    # Status Text Colors
    STATUS_TEXT_PAID: str = "1B5E20"
    STATUS_TEXT_PENDING: str = "E65100"
    STATUS_TEXT_POSTPONED: str = "BF360C"
    STATUS_TEXT_URGENT: str = "6A1B9A"

    # Section Colors for Cash Flow
    SECTION_OPERATING: str = "E8F5E9"      # Light green
    SECTION_INVESTING: str = "E3F2FD"      # Light blue
    SECTION_FINANCING: str = "FFF8E1"      # Light yellow

    # Day Type Colors (for Treasury Calendar)
    DAY_WORKING: str = "FFFFFF"
    DAY_WEEKEND: str = "ECEFF1"            # Light gray for weekends
    DAY_PAYMENT: str = "E8F5E9"            # Light green for Wednesday


# =============================================================================
# SECTION 2: TYPOGRAPHY SETTINGS
# =============================================================================

@dataclass
class FontStyle:
    """Font style specification."""
    name: str = "Calibri"
    size: int = 11
    bold: bool = False
    italic: bool = False
    color: str = "000000"


class Typography:
    """Typography configuration for the workbook."""

    # Font Family
    FONT_FAMILY: str = "Calibri"

    # Font Sizes
    SIZE_TITLE: int = 16
    SIZE_SECTION_HEADER: int = 14
    SIZE_HEADER: int = 12
    SIZE_BODY: int = 11
    SIZE_SMALL: int = 10
    SIZE_FOOTNOTE: int = 9

    # Pre-defined Font Styles
    TITLE: FontStyle = FontStyle(
        name="Calibri",
        size=16,
        bold=True,
        color="1E3A5F"
    )

    SECTION_HEADER: FontStyle = FontStyle(
        name="Calibri",
        size=14,
        bold=True,
        color="FFFFFF"
    )

    COLUMN_HEADER: FontStyle = FontStyle(
        name="Calibri",
        size=12,
        bold=True,
        color="FFFFFF"
    )

    SUB_HEADER: FontStyle = FontStyle(
        name="Calibri",
        size=11,
        bold=True,
        color="FFFFFF"
    )

    BODY: FontStyle = FontStyle(
        name="Calibri",
        size=11,
        bold=False,
        color="000000"
    )

    BODY_BOLD: FontStyle = FontStyle(
        name="Calibri",
        size=11,
        bold=True,
        color="000000"
    )

    TOTAL_ROW: FontStyle = FontStyle(
        name="Calibri",
        size=11,
        bold=True,
        color="1E3A5F"
    )

    POSITIVE_VALUE: FontStyle = FontStyle(
        name="Calibri",
        size=11,
        bold=False,
        color="2E7D32"
    )

    NEGATIVE_VALUE: FontStyle = FontStyle(
        name="Calibri",
        size=11,
        bold=False,
        color="C62828"
    )


# =============================================================================
# SECTION 3: NUMBER FORMATS
# =============================================================================

class NumberFormats:
    """Number format strings for Excel."""

    # Currency Formats (SAR as base)
    SAR_FULL: str = '_ "SAR" * #,##0.00_ ;_ "SAR" * -#,##0.00_ ;_ "SAR" * "-"??_ ;_ @_ '
    SAR_NO_SYMBOL: str = '_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * "-"??_ ;_ @_ '
    SAR_THOUSANDS: str = '_ * #,##0_ ;_ * -#,##0_ ;_ * "-"??_ ;_ @_ '
    SAR_MILLIONS: str = '_ * #,##0.0,,"M"_ ;_ * -#,##0.0,,"M"_ ;_ * "-"??_ ;_ @_ '

    # USD Format
    USD_FULL: str = '_ "$" * #,##0.00_ ;_ "$" * -#,##0.00_ ;_ "$" * "-"??_ ;_ @_ '

    # AED Format
    AED_FULL: str = '_ "AED" * #,##0.00_ ;_ "AED" * -#,##0.00_ ;_ "AED" * "-"??_ ;_ @_ '

    # EUR Format
    EUR_FULL: str = '_ "€" * #,##0.00_ ;_ "€" * -#,##0.00_ ;_ "€" * "-"??_ ;_ @_ '

    # GBP Format
    GBP_FULL: str = '_ "£" * #,##0.00_ ;_ "£" * -#,##0.00_ ;_ "£" * "-"??_ ;_ @_ '

    # SGD Format
    SGD_FULL: str = '_ "S$" * #,##0.00_ ;_ "S$" * -#,##0.00_ ;_ "S$" * "-"??_ ;_ @_ '

    # Generic Number Formats
    NUMBER_2DP: str = '#,##0.00'
    NUMBER_0DP: str = '#,##0'
    NUMBER_ACCOUNTING: str = '_(#,##0.00_);_((#,##0.00);_("-"??_);_(@_)'

    # Percentage Formats
    PERCENT_2DP: str = '0.00%'
    PERCENT_1DP: str = '0.0%'
    PERCENT_0DP: str = '0%'

    # Date Formats
    DATE_FULL: str = 'DD-MMM-YYYY'
    DATE_SHORT: str = 'DD/MM/YY'
    DATE_ISO: str = 'YYYY-MM-DD'
    DATE_MONTH_YEAR: str = 'MMM-YYYY'
    DATE_DAY_MONTH: str = 'DD-MMM'

    # Time Formats
    TIME_24H: str = 'HH:MM:SS'
    TIME_12H: str = 'h:mm AM/PM'

    # DateTime Format
    DATETIME_FULL: str = 'DD-MMM-YYYY HH:MM:SS'

    # Text Format
    TEXT: str = '@'

    @classmethod
    def get_currency_format(cls, currency_code: str) -> str:
        """Get the appropriate currency format for a currency code."""
        formats = {
            "SAR": cls.SAR_FULL,
            "USD": cls.USD_FULL,
            "AED": cls.AED_FULL,
            "EUR": cls.EUR_FULL,
            "GBP": cls.GBP_FULL,
            "SGD": cls.SGD_FULL,
        }
        return formats.get(currency_code, cls.SAR_NO_SYMBOL)


# =============================================================================
# SECTION 4: BORDER STYLES
# =============================================================================

class BorderStyle:
    """Border style constants."""
    NONE: str = None
    THIN: str = "thin"
    MEDIUM: str = "medium"
    THICK: str = "thick"
    DOUBLE: str = "double"
    HAIR: str = "hair"
    DASHED: str = "dashed"
    DOTTED: str = "dotted"


@dataclass
class BorderConfig:
    """Complete border configuration for a cell."""
    left: Optional[str] = None
    right: Optional[str] = None
    top: Optional[str] = None
    bottom: Optional[str] = None
    left_color: str = "D0D0D0"
    right_color: str = "D0D0D0"
    top_color: str = "D0D0D0"
    bottom_color: str = "D0D0D0"


class Borders:
    """Pre-defined border configurations."""

    # Grid border (thin gray)
    GRID: BorderConfig = BorderConfig(
        left=BorderStyle.THIN,
        right=BorderStyle.THIN,
        top=BorderStyle.THIN,
        bottom=BorderStyle.THIN,
        left_color=ColorPalette.BORDER_THIN_GRAY,
        right_color=ColorPalette.BORDER_THIN_GRAY,
        top_color=ColorPalette.BORDER_THIN_GRAY,
        bottom_color=ColorPalette.BORDER_THIN_GRAY
    )

    # Section border (medium black bottom)
    SECTION_BOTTOM: BorderConfig = BorderConfig(
        bottom=BorderStyle.MEDIUM,
        bottom_color=ColorPalette.BORDER_MEDIUM_BLACK
    )

    # Total row border (medium black top and bottom)
    TOTAL_ROW: BorderConfig = BorderConfig(
        top=BorderStyle.MEDIUM,
        bottom=BorderStyle.DOUBLE,
        top_color=ColorPalette.BORDER_MEDIUM_BLACK,
        bottom_color=ColorPalette.BORDER_MEDIUM_BLACK
    )

    # Header bottom border
    HEADER_BOTTOM: BorderConfig = BorderConfig(
        bottom=BorderStyle.MEDIUM,
        bottom_color=ColorPalette.BORDER_MEDIUM_BLACK
    )


# =============================================================================
# SECTION 5: ALIGNMENT SETTINGS
# =============================================================================

class Alignment:
    """Alignment constants."""

    # Horizontal
    H_LEFT: str = "left"
    H_CENTER: str = "center"
    H_RIGHT: str = "right"
    H_GENERAL: str = "general"

    # Vertical
    V_TOP: str = "top"
    V_CENTER: str = "center"
    V_BOTTOM: str = "bottom"


@dataclass
class AlignmentConfig:
    """Complete alignment configuration."""
    horizontal: str = "general"
    vertical: str = "center"
    wrap_text: bool = False
    shrink_to_fit: bool = False
    indent: int = 0


class Alignments:
    """Pre-defined alignment configurations."""

    # Standard alignments
    LEFT: AlignmentConfig = AlignmentConfig(horizontal="left", vertical="center")
    CENTER: AlignmentConfig = AlignmentConfig(horizontal="center", vertical="center")
    RIGHT: AlignmentConfig = AlignmentConfig(horizontal="right", vertical="center")

    # Header alignments
    HEADER_CENTER: AlignmentConfig = AlignmentConfig(
        horizontal="center",
        vertical="center",
        wrap_text=True
    )

    # Number alignments
    NUMBER: AlignmentConfig = AlignmentConfig(horizontal="right", vertical="center")

    # Text with wrap
    TEXT_WRAP: AlignmentConfig = AlignmentConfig(
        horizontal="left",
        vertical="top",
        wrap_text=True
    )


# =============================================================================
# SECTION 6: COLUMN WIDTH PRESETS
# =============================================================================

class ColumnWidths:
    """Standard column width presets."""

    # Identifier columns
    ID_SMALL: float = 8.0
    ID_MEDIUM: float = 12.0
    ID_LARGE: float = 15.0

    # Date columns
    DATE_SHORT: float = 10.0
    DATE_FULL: float = 12.0
    DATETIME: float = 18.0

    # Amount columns
    AMOUNT_SMALL: float = 12.0
    AMOUNT_MEDIUM: float = 15.0
    AMOUNT_LARGE: float = 18.0

    # Description columns
    DESC_SHORT: float = 20.0
    DESC_MEDIUM: float = 30.0
    DESC_LONG: float = 40.0

    # Name columns
    NAME_SHORT: float = 15.0
    NAME_MEDIUM: float = 25.0
    NAME_LONG: float = 35.0

    # Status columns
    STATUS: float = 12.0

    # Currency columns
    CURRENCY_CODE: float = 6.0

    # Percentage columns
    PERCENT: float = 10.0


# =============================================================================
# SECTION 7: ROW HEIGHT PRESETS
# =============================================================================

class RowHeights:
    """Standard row height presets."""

    TITLE: float = 30.0
    SECTION_HEADER: float = 25.0
    HEADER: float = 20.0
    BODY: float = 15.0
    SPACER: float = 8.0
    TOTAL: float = 20.0


# =============================================================================
# SECTION 8: CELL STYLE PRESETS
# =============================================================================

@dataclass
class CellStyle:
    """Complete cell style configuration."""
    font: FontStyle = None
    fill_color: Optional[str] = None
    border: Optional[BorderConfig] = None
    alignment: Optional[AlignmentConfig] = None
    number_format: Optional[str] = None


class CellStyles:
    """Pre-defined cell style configurations."""

    # Title style
    TITLE: CellStyle = CellStyle(
        font=Typography.TITLE,
        alignment=Alignments.LEFT
    )

    # Section header style
    SECTION_HEADER: CellStyle = CellStyle(
        font=Typography.SECTION_HEADER,
        fill_color=ColorPalette.HEADER_DARK_NAVY,
        alignment=Alignments.HEADER_CENTER
    )

    # Column header style
    COLUMN_HEADER: CellStyle = CellStyle(
        font=Typography.COLUMN_HEADER,
        fill_color=ColorPalette.HEADER_MEDIUM_BLUE,
        border=Borders.HEADER_BOTTOM,
        alignment=Alignments.HEADER_CENTER
    )

    # Body cell style
    BODY: CellStyle = CellStyle(
        font=Typography.BODY,
        border=Borders.GRID,
        alignment=Alignments.LEFT
    )

    # Number cell style
    NUMBER: CellStyle = CellStyle(
        font=Typography.BODY,
        border=Borders.GRID,
        alignment=Alignments.NUMBER,
        number_format=NumberFormats.NUMBER_2DP
    )

    # Currency cell style (SAR)
    CURRENCY_SAR: CellStyle = CellStyle(
        font=Typography.BODY,
        border=Borders.GRID,
        alignment=Alignments.NUMBER,
        number_format=NumberFormats.SAR_NO_SYMBOL
    )

    # Date cell style
    DATE: CellStyle = CellStyle(
        font=Typography.BODY,
        border=Borders.GRID,
        alignment=Alignments.CENTER,
        number_format=NumberFormats.DATE_FULL
    )

    # Percentage cell style
    PERCENT: CellStyle = CellStyle(
        font=Typography.BODY,
        border=Borders.GRID,
        alignment=Alignments.NUMBER,
        number_format=NumberFormats.PERCENT_2DP
    )

    # Total row style
    TOTAL: CellStyle = CellStyle(
        font=Typography.TOTAL_ROW,
        fill_color=ColorPalette.HEADER_LIGHT_BLUE,
        border=Borders.TOTAL_ROW,
        alignment=Alignments.NUMBER
    )

    # Positive value style
    POSITIVE: CellStyle = CellStyle(
        font=Typography.POSITIVE_VALUE,
        alignment=Alignments.NUMBER
    )

    # Negative value style
    NEGATIVE: CellStyle = CellStyle(
        font=Typography.NEGATIVE_VALUE,
        alignment=Alignments.NUMBER
    )
