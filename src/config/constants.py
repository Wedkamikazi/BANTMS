"""
BANTMS Constants Configuration Module
=====================================

Contains all system-wide constants for:
- Saudi Arabian Calendar System
- Currency Framework (SAR as base)
- Bank Codes
- Vendor Tier Classifications
- Payment Type Definitions
- Category Codes for Cash Flow Classification
"""

from enum import Enum
from typing import Dict, List, NamedTuple
from decimal import Decimal
from datetime import date


# =============================================================================
# SECTION 1: SAUDI ARABIAN CALENDAR SYSTEM
# =============================================================================

class SaudiCalendarConfig:
    """
    Saudi Arabian Calendar Configuration

    - Week starts: Sunday (day 0)
    - Week ends: Saturday (day 6)
    - Weekend days: Friday (day 5) and Saturday (day 6)
    - Business days: Sunday through Thursday
    - Payment day policy: Wednesday only (except Urgent/Payroll)
    """

    # Week configuration (Sunday = 0, Saturday = 6)
    WEEK_START_DAY: int = 6  # Sunday in Python's weekday() is 6
    WEEK_END_DAY: int = 5    # Saturday in Python's weekday() is 5

    # Weekend days (Friday = 4, Saturday = 5 in Python)
    WEEKEND_DAYS: tuple = (4, 5)  # Friday and Saturday

    # Business days (Sunday = 6, Monday = 0, Tuesday = 1, Wednesday = 2, Thursday = 3)
    BUSINESS_DAYS: tuple = (6, 0, 1, 2, 3)  # Sunday through Thursday

    # Standard payment day (Wednesday = 2 in Python's weekday())
    PAYMENT_DAY: int = 2  # Wednesday

    # Day names in order (Sunday start)
    DAY_NAMES: tuple = (
        "Sunday", "Monday", "Tuesday", "Wednesday",
        "Thursday", "Friday", "Saturday"
    )

    # Day abbreviations (Sunday start)
    DAY_ABBREV: tuple = ("Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat")

    # Day type indicators
    DAY_TYPE_WORKING: str = "W"
    DAY_TYPE_WEEKEND: str = "WE"
    DAY_TYPE_PAYMENT: str = "WED"  # Wednesday payment day

    @classmethod
    def is_weekend(cls, weekday: int) -> bool:
        """Check if a Python weekday (0=Monday) is a Saudi weekend day."""
        return weekday in cls.WEEKEND_DAYS

    @classmethod
    def is_business_day(cls, weekday: int) -> bool:
        """Check if a Python weekday is a Saudi business day."""
        return weekday in cls.BUSINESS_DAYS

    @classmethod
    def is_payment_day(cls, weekday: int) -> bool:
        """Check if a Python weekday is the standard payment day (Wednesday)."""
        return weekday == cls.PAYMENT_DAY

    @classmethod
    def get_day_type(cls, weekday: int) -> str:
        """Get the day type indicator for a Python weekday."""
        if cls.is_payment_day(weekday):
            return cls.DAY_TYPE_PAYMENT
        elif cls.is_weekend(weekday):
            return cls.DAY_TYPE_WEEKEND
        else:
            return cls.DAY_TYPE_WORKING


# =============================================================================
# SECTION 2: CURRENCY FRAMEWORK (SAR AS BASE)
# =============================================================================

class CurrencyInfo(NamedTuple):
    """Currency information structure."""
    code: str
    name: str
    symbol: str
    rate_to_sar: Decimal
    decimal_places: int


class CurrencyFramework:
    """
    Currency Framework Configuration

    - Base currency: SAR (Saudi Riyal)
    - Supported currencies: USD, AED, EUR, GBP, SGD
    - All summary figures display in SAR equivalent
    """

    # Base currency
    BASE_CURRENCY: str = "SAR"
    BASE_CURRENCY_NAME: str = "Saudi Riyal"
    BASE_CURRENCY_SYMBOL: str = "﷼"

    # Supported currencies with exchange rates to SAR
    # Note: Rates are indicative and should be updated from FX_Rates sheet
    CURRENCIES: Dict[str, CurrencyInfo] = {
        "SAR": CurrencyInfo(
            code="SAR",
            name="Saudi Riyal",
            symbol="﷼",
            rate_to_sar=Decimal("1.0000"),
            decimal_places=2
        ),
        "USD": CurrencyInfo(
            code="USD",
            name="US Dollar",
            symbol="$",
            rate_to_sar=Decimal("3.7500"),
            decimal_places=2
        ),
        "AED": CurrencyInfo(
            code="AED",
            name="UAE Dirham",
            symbol="د.إ",
            rate_to_sar=Decimal("1.0200"),
            decimal_places=2
        ),
        "EUR": CurrencyInfo(
            code="EUR",
            name="Euro",
            symbol="€",
            rate_to_sar=Decimal("4.0500"),
            decimal_places=2
        ),
        "GBP": CurrencyInfo(
            code="GBP",
            name="British Pound",
            symbol="£",
            rate_to_sar=Decimal("4.7000"),
            decimal_places=2
        ),
        "SGD": CurrencyInfo(
            code="SGD",
            name="Singapore Dollar",
            symbol="S$",
            rate_to_sar=Decimal("2.7800"),
            decimal_places=2
        ),
    }

    # Currency codes list (for dropdowns)
    CURRENCY_CODES: tuple = ("SAR", "USD", "AED", "EUR", "GBP", "SGD")

    @classmethod
    def get_rate_to_sar(cls, currency_code: str) -> Decimal:
        """Get exchange rate to SAR for a currency."""
        if currency_code in cls.CURRENCIES:
            return cls.CURRENCIES[currency_code].rate_to_sar
        raise ValueError(f"Unknown currency code: {currency_code}")

    @classmethod
    def convert_to_sar(cls, amount: Decimal, from_currency: str) -> Decimal:
        """Convert an amount to SAR."""
        rate = cls.get_rate_to_sar(from_currency)
        return amount * rate

    @classmethod
    def get_currency_symbol(cls, currency_code: str) -> str:
        """Get the symbol for a currency."""
        if currency_code in cls.CURRENCIES:
            return cls.CURRENCIES[currency_code].symbol
        return currency_code


# =============================================================================
# SECTION 3: SAUDI BANK CODES
# =============================================================================

class BankInfo(NamedTuple):
    """Bank information structure."""
    code: str
    name: str
    swift: str
    short_name: str


class SaudiBanks:
    """Saudi Arabian Bank Configuration."""

    BANKS: Dict[str, BankInfo] = {
        "SNB": BankInfo(
            code="SNB",
            name="Saudi National Bank",
            swift="NCBKSAJE",
            short_name="SNB"
        ),
        "RAJHI": BankInfo(
            code="RAJHI",
            name="Al Rajhi Bank",
            swift="RJHISARI",
            short_name="Al Rajhi"
        ),
        "RIYAD": BankInfo(
            code="RIYAD",
            name="Riyad Bank",
            swift="RIABORIS",
            short_name="Riyad"
        ),
        "SABB": BankInfo(
            code="SABB",
            name="Saudi British Bank",
            swift="SABBSARI",
            short_name="SABB"
        ),
        "ANB": BankInfo(
            code="ANB",
            name="Arab National Bank",
            swift="ARNBSARI",
            short_name="ANB"
        ),
        "BSF": BankInfo(
            code="BSF",
            name="Banque Saudi Fransi",
            swift="BSFRSARI",
            short_name="BSF"
        ),
        "ALINMA": BankInfo(
            code="ALINMA",
            name="Alinma Bank",
            swift="AABORSAR",
            short_name="Alinma"
        ),
        "ALBILAD": BankInfo(
            code="ALBILAD",
            name="Bank AlBilad",
            swift="BBILSARI",
            short_name="AlBilad"
        ),
        "ALJAZIRA": BankInfo(
            code="ALJAZIRA",
            name="Bank AlJazira",
            swift="BJAZSAJE",
            short_name="AlJazira"
        ),
        "GIB": BankInfo(
            code="GIB",
            name="Gulf International Bank",
            swift="GABORSAR",
            short_name="GIB"
        ),
    }

    # Bank codes list (for dropdowns)
    BANK_CODES: tuple = tuple(BANKS.keys())


# =============================================================================
# SECTION 4: VENDOR TIER CLASSIFICATION
# =============================================================================

class VendorTier(Enum):
    """
    Vendor Tier Classification

    Tier determines payment snap direction:
    - Tier 1 (Critical): Snap to Wednesday BEFORE due date
    - Tier 2 (Standard): Snap to Wednesday AFTER due date
    - Tier 3 (Flexible): Snap to Wednesday AFTER due date
    """
    TIER_1_CRITICAL = 1
    TIER_2_STANDARD = 2
    TIER_3_FLEXIBLE = 3


class VendorTierConfig:
    """Vendor Tier Configuration with snap direction."""

    TIERS: Dict[int, Dict] = {
        1: {
            "name": "Critical",
            "description": "Mission-critical suppliers, government entities",
            "snap_direction": "BEFORE",
            "max_postpone": 0,
            "examples": ["Saudi Aramco", "SABIC", "GOSI", "ZATCA", "SEC"]
        },
        2: {
            "name": "Standard",
            "description": "Regular business suppliers",
            "snap_direction": "AFTER",
            "max_postpone": 1,
            "examples": ["Al-Marai", "Jarir", "Extra Electronics"]
        },
        3: {
            "name": "Flexible",
            "description": "Flexible payment terms suppliers",
            "snap_direction": "AFTER",
            "max_postpone": 2,
            "examples": ["Office supplies", "Maintenance contractors"]
        },
    }

    @classmethod
    def get_snap_direction(cls, tier: int) -> str:
        """Get snap direction for a tier."""
        return cls.TIERS.get(tier, cls.TIERS[3])["snap_direction"]

    @classmethod
    def get_max_postpone(cls, tier: int) -> int:
        """Get maximum postponement count for a tier."""
        return cls.TIERS.get(tier, cls.TIERS[3])["max_postpone"]


# =============================================================================
# SECTION 5: PAYMENT TYPE DEFINITIONS
# =============================================================================

class PaymentType(Enum):
    """Payment Type Classification."""
    SUPPLIER = "Supplier"
    GOVERNMENT = "Government"
    STRATEGIC = "Strategic"
    UTILITIES = "Utilities"
    RENT = "Rent"
    PAYROLL = "Payroll"
    LOAN = "Loan"
    TAX = "Tax"
    INSURANCE = "Insurance"
    OTHER = "Other"


# =============================================================================
# SECTION 6: PAYMENT STATUS DEFINITIONS
# =============================================================================

class PaymentStatus(Enum):
    """Payment Status Classification."""
    PENDING = "Pending"
    PAID = "Paid"
    POSTPONED_1 = "Postponed-1"
    POSTPONED_2 = "Postponed-2"
    URGENT = "Urgent"
    CANCELLED = "Cancelled"
    ON_HOLD = "On Hold"


class CollectionStatus(Enum):
    """Collection Status Classification."""
    EXPECTED = "Expected"
    RECEIVED = "Received"
    DELAYED = "Delayed"
    DISPUTED = "Disputed"
    WRITTEN_OFF = "Written Off"


class TDStatus(Enum):
    """Time Deposit Status Classification."""
    ACTIVE = "Active"
    ACTION_REQUIRED = "Action_Required"
    CLOSED = "Closed"
    ROLLED_OVER = "Rolled_Over"


# =============================================================================
# SECTION 7: CONFIRMATION STATUS
# =============================================================================

class ConfirmationStatus(Enum):
    """Confirmation Status for forecasting reliability."""
    CONFIRMED = "Confirmed"
    ESTIMATED = "Estimated"
    TENTATIVE = "Tentative"


# =============================================================================
# SECTION 8: PERIOD SELECTOR OPTIONS
# =============================================================================

class PeriodType(Enum):
    """Period Type for filtering and aggregation."""
    DAILY = "Daily"
    WEEKLY = "Weekly"
    MONTHLY = "Monthly"
    QUARTERLY = "Quarterly"
    YEARLY = "Yearly"


class PeriodSelectorConfig:
    """Period Selector Configuration."""

    # Quarter definitions
    QUARTERS: Dict[str, tuple] = {
        "Q1": (1, 2, 3),
        "Q2": (4, 5, 6),
        "Q3": (7, 8, 9),
        "Q4": (10, 11, 12),
    }

    # Week ID format: 01-52
    WEEK_ID_FORMAT: str = "{:02d}"

    # Month format
    MONTH_FORMAT: str = "{:%b-%Y}"

    # Quarter format
    QUARTER_FORMAT: str = "Q{quarter} {year}"


# =============================================================================
# SECTION 9: CASH FLOW SECTION DEFINITIONS
# =============================================================================

class CashFlowSection(Enum):
    """Cash Flow Statement Section Classification."""
    OPERATING = "Operating"
    INVESTING = "Investing"
    FINANCING = "Financing"


class CashFlowDirection(Enum):
    """Cash Flow Direction."""
    INFLOW = "Inflow"
    OUTFLOW = "Outflow"
