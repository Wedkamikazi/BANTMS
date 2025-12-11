"""
BANTMS Category Configuration Module
=====================================

Contains all transaction categorization rules:
- Category Codes for Cash Flow Classification
- Description Patterns for Auto-Categorization
- Cash Flow Line Item Mappings
"""

from typing import Dict, List, NamedTuple
from enum import Enum
from dataclasses import dataclass


# =============================================================================
# SECTION 1: CASH FLOW SECTION ENUMERATION
# =============================================================================

class CFSection(Enum):
    """Cash Flow Statement Section."""
    OPERATING = "Operating"
    INVESTING = "Investing"
    FINANCING = "Financing"


class CFDirection(Enum):
    """Cash Flow Direction."""
    INFLOW = "Inflow"
    OUTFLOW = "Outflow"


# =============================================================================
# SECTION 2: CATEGORY DEFINITION
# =============================================================================

@dataclass
class CategoryDefinition:
    """
    Complete category definition for transaction classification.

    Attributes:
        code: Unique category code (e.g., "AR-COLLECT")
        name: Human-readable name
        section: Cash Flow section (Operating/Investing/Financing)
        line_item: Cash Flow Statement line item
        direction: Inflow or Outflow
        description_patterns: List of patterns to match in transaction descriptions
        priority: Pattern matching priority (lower = higher priority)
    """
    code: str
    name: str
    section: CFSection
    line_item: str
    direction: CFDirection
    description_patterns: List[str]
    priority: int = 100


# =============================================================================
# SECTION 3: OPERATING ACTIVITIES CATEGORIES
# =============================================================================

class OperatingCategories:
    """Operating Activities Category Definitions."""

    # Collections from Customers
    AR_COLLECT = CategoryDefinition(
        code="AR-COLLECT",
        name="Customer Collections",
        section=CFSection.OPERATING,
        line_item="Collections from Customers",
        direction=CFDirection.INFLOW,
        description_patterns=[
            "*ARAMCO*",
            "*SABIC*",
            "*CUSTOMER*",
            "*COLLECTION*",
            "*RECEIPT*",
            "*AR RECEIPT*",
            "*RECEIVABLE*",
            "*SALES RECEIPT*",
            "*PAYMENT RECEIVED*",
            "*CUST PMT*",
            "*INVOICE PAYMENT*",
        ],
        priority=10
    )

    AR_REFUND = CategoryDefinition(
        code="AR-REFUND",
        name="Customer Refunds",
        section=CFSection.OPERATING,
        line_item="Customer Refunds",
        direction=CFDirection.OUTFLOW,
        description_patterns=[
            "*CUSTOMER REFUND*",
            "*AR REFUND*",
            "*SALES REFUND*",
            "*REFUND TO CUSTOMER*",
        ],
        priority=15
    )

    # Payments to Suppliers
    AP_SUPPLIER = CategoryDefinition(
        code="AP-SUPPLIER",
        name="Supplier Payments",
        section=CFSection.OPERATING,
        line_item="Payments to Suppliers",
        direction=CFDirection.OUTFLOW,
        description_patterns=[
            "*SUPPLIER*",
            "*VENDOR*",
            "*AP PAYMENT*",
            "*TRADE PAYABLE*",
            "*PURCHASE*",
            "*PROCUREMENT*",
            "*MATERIAL*",
            "*INVENTORY*",
        ],
        priority=20
    )

    AP_STRATEGIC = CategoryDefinition(
        code="AP-STRATEGIC",
        name="Strategic Supplier Payments",
        section=CFSection.OPERATING,
        line_item="Strategic Supplier Payments",
        direction=CFDirection.OUTFLOW,
        description_patterns=[
            "*AL-MARAI*",
            "*ALMARAI*",
            "*PANDA*",
            "*JARIR*",
            "*EXTRA*",
        ],
        priority=18
    )

    # Government Payments
    GOV_TAX = CategoryDefinition(
        code="GOV-TAX",
        name="Tax Payments",
        section=CFSection.OPERATING,
        line_item="Government Payments",
        direction=CFDirection.OUTFLOW,
        description_patterns=[
            "*ZATCA*",
            "*GAZT*",
            "*TAX*",
            "*VAT*",
            "*ZAKAT*",
            "*WITHHOLDING*",
        ],
        priority=5
    )

    GOV_GOSI = CategoryDefinition(
        code="GOV-GOSI",
        name="GOSI Payments",
        section=CFSection.OPERATING,
        line_item="Government Payments",
        direction=CFDirection.OUTFLOW,
        description_patterns=[
            "*GOSI*",
            "*SOCIAL INSURANCE*",
            "*SOCIAL SECURITY*",
        ],
        priority=5
    )

    GOV_LICENSE = CategoryDefinition(
        code="GOV-LICENSE",
        name="Government Licenses & Fees",
        section=CFSection.OPERATING,
        line_item="Government Payments",
        direction=CFDirection.OUTFLOW,
        description_patterns=[
            "*LICENSE*",
            "*PERMIT*",
            "*GOVERNMENT FEE*",
            "*GOV FEE*",
            "*MINISTRY*",
            "*MUNICIPALITY*",
            "*BALADIYA*",
        ],
        priority=6
    )

    # Payroll
    HR_PAYROLL = CategoryDefinition(
        code="HR-PAYROLL",
        name="Payroll Disbursements",
        section=CFSection.OPERATING,
        line_item="Payroll Disbursements",
        direction=CFDirection.OUTFLOW,
        description_patterns=[
            "*SALARY*",
            "*SALARIES*",
            "*PAYROLL*",
            "*WAGES*",
            "*WPS*",
            "*EMPLOYEE*",
            "*STAFF*",
        ],
        priority=8
    )

    HR_EOS = CategoryDefinition(
        code="HR-EOS",
        name="End of Service Benefits",
        section=CFSection.OPERATING,
        line_item="Payroll Disbursements",
        direction=CFDirection.OUTFLOW,
        description_patterns=[
            "*END OF SERVICE*",
            "*EOS*",
            "*INDEMNITY*",
            "*GRATUITY*",
            "*TERMINATION*",
        ],
        priority=9
    )

    HR_BENEFITS = CategoryDefinition(
        code="HR-BENEFITS",
        name="Employee Benefits",
        section=CFSection.OPERATING,
        line_item="Payroll Disbursements",
        direction=CFDirection.OUTFLOW,
        description_patterns=[
            "*MEDICAL*",
            "*INSURANCE PREMIUM*",
            "*STAFF BENEFIT*",
            "*EMPLOYEE BENEFIT*",
            "*ALLOWANCE*",
            "*BONUS*",
        ],
        priority=11
    )

    # Utilities & Rent
    UTIL_ELECTRICITY = CategoryDefinition(
        code="UTIL-ELEC",
        name="Electricity",
        section=CFSection.OPERATING,
        line_item="Utilities & Rent",
        direction=CFDirection.OUTFLOW,
        description_patterns=[
            "*SEC*",
            "*SAUDI ELECTRICITY*",
            "*ELECTRICITY*",
            "*ELECTRIC*",
            "*POWER*",
        ],
        priority=12
    )

    UTIL_WATER = CategoryDefinition(
        code="UTIL-WATER",
        name="Water",
        section=CFSection.OPERATING,
        line_item="Utilities & Rent",
        direction=CFDirection.OUTFLOW,
        description_patterns=[
            "*WATER*",
            "*NWC*",
            "*NATIONAL WATER*",
        ],
        priority=12
    )

    UTIL_TELECOM = CategoryDefinition(
        code="UTIL-TELECOM",
        name="Telecommunications",
        section=CFSection.OPERATING,
        line_item="Utilities & Rent",
        direction=CFDirection.OUTFLOW,
        description_patterns=[
            "*STC*",
            "*MOBILY*",
            "*ZAIN*",
            "*TELECOM*",
            "*TELEPHONE*",
            "*INTERNET*",
        ],
        priority=12
    )

    RENT_OFFICE = CategoryDefinition(
        code="RENT-OFFICE",
        name="Office Rent",
        section=CFSection.OPERATING,
        line_item="Utilities & Rent",
        direction=CFDirection.OUTFLOW,
        description_patterns=[
            "*RENT*",
            "*LEASE*",
            "*RENTAL*",
            "*OFFICE SPACE*",
            "*PROPERTY*",
        ],
        priority=14
    )

    # Other Operating Expenses
    OPEX_PROFESSIONAL = CategoryDefinition(
        code="OPEX-PROF",
        name="Professional Services",
        section=CFSection.OPERATING,
        line_item="Other Operating Expenses",
        direction=CFDirection.OUTFLOW,
        description_patterns=[
            "*CONSULTANT*",
            "*LEGAL*",
            "*AUDIT*",
            "*ACCOUNTING*",
            "*ADVISORY*",
            "*PROFESSIONAL*",
        ],
        priority=30
    )

    OPEX_OFFICE = CategoryDefinition(
        code="OPEX-OFFICE",
        name="Office Expenses",
        section=CFSection.OPERATING,
        line_item="Other Operating Expenses",
        direction=CFDirection.OUTFLOW,
        description_patterns=[
            "*OFFICE SUPPLIES*",
            "*STATIONERY*",
            "*PRINTING*",
            "*COURIER*",
            "*POSTAGE*",
        ],
        priority=35
    )

    OPEX_TRAVEL = CategoryDefinition(
        code="OPEX-TRAVEL",
        name="Travel & Entertainment",
        section=CFSection.OPERATING,
        line_item="Other Operating Expenses",
        direction=CFDirection.OUTFLOW,
        description_patterns=[
            "*TRAVEL*",
            "*FLIGHT*",
            "*HOTEL*",
            "*ACCOMMODATION*",
            "*ENTERTAINMENT*",
            "*BUSINESS TRIP*",
        ],
        priority=36
    )

    OPEX_MAINT = CategoryDefinition(
        code="OPEX-MAINT",
        name="Maintenance & Repairs",
        section=CFSection.OPERATING,
        line_item="Other Operating Expenses",
        direction=CFDirection.OUTFLOW,
        description_patterns=[
            "*MAINTENANCE*",
            "*REPAIR*",
            "*SERVICE*",
            "*FIX*",
        ],
        priority=37
    )

    OPEX_IT = CategoryDefinition(
        code="OPEX-IT",
        name="IT & Software",
        section=CFSection.OPERATING,
        line_item="Other Operating Expenses",
        direction=CFDirection.OUTFLOW,
        description_patterns=[
            "*SOFTWARE*",
            "*LICENSE*",
            "*IT SERVICE*",
            "*SUBSCRIPTION*",
            "*MICROSOFT*",
            "*ORACLE*",
            "*SAP*",
        ],
        priority=38
    )

    OPEX_INSURANCE = CategoryDefinition(
        code="OPEX-INS",
        name="Insurance",
        section=CFSection.OPERATING,
        line_item="Other Operating Expenses",
        direction=CFDirection.OUTFLOW,
        description_patterns=[
            "*INSURANCE*",
            "*POLICY PREMIUM*",
            "*COVERAGE*",
        ],
        priority=32
    )

    OPEX_BANK_CHARGES = CategoryDefinition(
        code="OPEX-BANK",
        name="Bank Charges",
        section=CFSection.OPERATING,
        line_item="Other Operating Expenses",
        direction=CFDirection.OUTFLOW,
        description_patterns=[
            "*BANK CHARGE*",
            "*BANK FEE*",
            "*SERVICE CHARGE*",
            "*COMMISSION*",
            "*TRANSFER FEE*",
            "*SWIFT*",
        ],
        priority=40
    )


# =============================================================================
# SECTION 4: INVESTING ACTIVITIES CATEGORIES
# =============================================================================

class InvestingCategories:
    """Investing Activities Category Definitions."""

    # Time Deposits
    TD_PLACEMENT = CategoryDefinition(
        code="TD-PLACE",
        name="TD Placements",
        section=CFSection.INVESTING,
        line_item="Time Deposit Placements",
        direction=CFDirection.OUTFLOW,
        description_patterns=[
            "*TD PLACEMENT*",
            "*TIME DEPOSIT*",
            "*FIXED DEPOSIT*",
            "*TERM DEPOSIT*",
            "*FD PLACEMENT*",
            "*DEPOSIT PLACEMENT*",
        ],
        priority=3
    )

    TD_MATURITY = CategoryDefinition(
        code="TD-MAT",
        name="TD Maturities",
        section=CFSection.INVESTING,
        line_item="Time Deposit Maturities",
        direction=CFDirection.INFLOW,
        description_patterns=[
            "*TD MATURITY*",
            "*TD MAT*",
            "*MATURED DEPOSIT*",
            "*FD MATURITY*",
            "*DEPOSIT MATURED*",
        ],
        priority=3
    )

    TD_INTEREST = CategoryDefinition(
        code="TD-INT",
        name="TD Interest",
        section=CFSection.INVESTING,
        line_item="Interest Received on TDs",
        direction=CFDirection.INFLOW,
        description_patterns=[
            "*TD INTEREST*",
            "*DEPOSIT INTEREST*",
            "*INTEREST INCOME*",
            "*INT RECEIVED*",
            "*FD INTEREST*",
        ],
        priority=4
    )

    # Fixed Assets
    CAPEX_PURCHASE = CategoryDefinition(
        code="CAPEX-PUR",
        name="Capital Expenditure",
        section=CFSection.INVESTING,
        line_item="Purchase of Fixed Assets",
        direction=CFDirection.OUTFLOW,
        description_patterns=[
            "*CAPEX*",
            "*CAPITAL EXP*",
            "*FIXED ASSET*",
            "*EQUIPMENT*",
            "*MACHINERY*",
            "*VEHICLE*",
            "*FURNITURE*",
        ],
        priority=25
    )

    CAPEX_DISPOSAL = CategoryDefinition(
        code="CAPEX-DISP",
        name="Asset Disposals",
        section=CFSection.INVESTING,
        line_item="Sale of Fixed Assets",
        direction=CFDirection.INFLOW,
        description_patterns=[
            "*ASSET SALE*",
            "*DISPOSAL*",
            "*SOLD EQUIPMENT*",
            "*VEHICLE SALE*",
        ],
        priority=26
    )

    # Investments
    INV_PURCHASE = CategoryDefinition(
        code="INV-PUR",
        name="Investment Purchases",
        section=CFSection.INVESTING,
        line_item="Purchase of Investments",
        direction=CFDirection.OUTFLOW,
        description_patterns=[
            "*INVESTMENT*",
            "*SECURITIES*",
            "*SUKUK*",
            "*BOND*",
            "*EQUITY*",
            "*SHARE PURCHASE*",
        ],
        priority=27
    )

    INV_SALE = CategoryDefinition(
        code="INV-SALE",
        name="Investment Sales",
        section=CFSection.INVESTING,
        line_item="Sale of Investments",
        direction=CFDirection.INFLOW,
        description_patterns=[
            "*INVESTMENT SALE*",
            "*SECURITIES SOLD*",
            "*SUKUK MATURED*",
            "*BOND REDEMPTION*",
        ],
        priority=27
    )


# =============================================================================
# SECTION 5: FINANCING ACTIVITIES CATEGORIES
# =============================================================================

class FinancingCategories:
    """Financing Activities Category Definitions."""

    # Loan Activity
    LOAN_DRAWDOWN = CategoryDefinition(
        code="LOAN-DRAW",
        name="Loan Drawdowns",
        section=CFSection.FINANCING,
        line_item="Loan Drawdowns",
        direction=CFDirection.INFLOW,
        description_patterns=[
            "*LOAN DRAW*",
            "*FACILITY DRAW*",
            "*DRAWDOWN*",
            "*CREDIT FACILITY*",
            "*LOAN DISBURSEMENT*",
        ],
        priority=2
    )

    LOAN_PRINCIPAL = CategoryDefinition(
        code="LOAN-PRIN",
        name="Loan Principal Repayments",
        section=CFSection.FINANCING,
        line_item="Loan Repayments - Principal",
        direction=CFDirection.OUTFLOW,
        description_patterns=[
            "*LOAN PRINCIPAL*",
            "*PRINCIPAL REPAYMENT*",
            "*LOAN REPAY*",
            "*FACILITY REPAY*",
        ],
        priority=2
    )

    LOAN_INTEREST = CategoryDefinition(
        code="LOAN-INT",
        name="Loan Interest Payments",
        section=CFSection.FINANCING,
        line_item="Loan Repayments - Interest",
        direction=CFDirection.OUTFLOW,
        description_patterns=[
            "*LOAN INTEREST*",
            "*INT PAYMENT*",
            "*INTEREST EXPENSE*",
            "*FACILITY INTEREST*",
        ],
        priority=2
    )

    # Equity
    EQUITY_INJECTION = CategoryDefinition(
        code="EQUITY-IN",
        name="Capital Injection",
        section=CFSection.FINANCING,
        line_item="Capital Contributions",
        direction=CFDirection.INFLOW,
        description_patterns=[
            "*CAPITAL INJECT*",
            "*SHARE CAPITAL*",
            "*EQUITY*",
            "*SHAREHOLDER*",
            "*CAPITAL CONTRIBUTION*",
        ],
        priority=1
    )

    DIVIDEND_PAID = CategoryDefinition(
        code="DIV-OUT",
        name="Dividends Paid",
        section=CFSection.FINANCING,
        line_item="Dividends Paid",
        direction=CFDirection.OUTFLOW,
        description_patterns=[
            "*DIVIDEND*",
            "*DISTRIBUTION*",
            "*SHAREHOLDER PAYMENT*",
        ],
        priority=1
    )


# =============================================================================
# SECTION 6: INTERNAL TRANSFERS (Non-Cash Flow Impact)
# =============================================================================

class InternalCategories:
    """Internal Transfer Categories (No Cash Flow Impact)."""

    INTER_ACCOUNT = CategoryDefinition(
        code="INT-ACCT",
        name="Inter-Account Transfer",
        section=CFSection.OPERATING,  # Excluded from CF
        line_item="Internal Transfer",
        direction=CFDirection.INFLOW,  # Could be either
        description_patterns=[
            "*TRANSFER TO*",
            "*TRANSFER FROM*",
            "*INTERNAL*",
            "*INTER-ACCOUNT*",
            "*INTER ACCOUNT*",
            "*OWN ACCOUNT*",
        ],
        priority=1
    )

    FX_CONVERSION = CategoryDefinition(
        code="INT-FX",
        name="FX Conversion",
        section=CFSection.OPERATING,  # Excluded from CF
        line_item="FX Conversion",
        direction=CFDirection.INFLOW,
        description_patterns=[
            "*FX CONV*",
            "*CURRENCY CONV*",
            "*EXCHANGE*",
            "*FOREX*",
        ],
        priority=1
    )


# =============================================================================
# SECTION 7: UNCATEGORIZED (Catch-all)
# =============================================================================

UNCATEGORIZED = CategoryDefinition(
    code="UNCATEGORIZED",
    name="Uncategorized",
    section=CFSection.OPERATING,
    line_item="Other Operating Items",
    direction=CFDirection.OUTFLOW,
    description_patterns=[],
    priority=999
)


# =============================================================================
# SECTION 8: MASTER CATEGORY LIST
# =============================================================================

def get_all_categories() -> List[CategoryDefinition]:
    """
    Get all category definitions sorted by priority.

    Returns:
        List of CategoryDefinition objects sorted by priority (ascending).
    """
    categories = []

    # Operating categories
    for attr_name in dir(OperatingCategories):
        if not attr_name.startswith('_'):
            attr = getattr(OperatingCategories, attr_name)
            if isinstance(attr, CategoryDefinition):
                categories.append(attr)

    # Investing categories
    for attr_name in dir(InvestingCategories):
        if not attr_name.startswith('_'):
            attr = getattr(InvestingCategories, attr_name)
            if isinstance(attr, CategoryDefinition):
                categories.append(attr)

    # Financing categories
    for attr_name in dir(FinancingCategories):
        if not attr_name.startswith('_'):
            attr = getattr(FinancingCategories, attr_name)
            if isinstance(attr, CategoryDefinition):
                categories.append(attr)

    # Internal categories
    for attr_name in dir(InternalCategories):
        if not attr_name.startswith('_'):
            attr = getattr(InternalCategories, attr_name)
            if isinstance(attr, CategoryDefinition):
                categories.append(attr)

    # Add uncategorized at the end
    categories.append(UNCATEGORIZED)

    # Sort by priority
    categories.sort(key=lambda x: x.priority)

    return categories


def get_category_by_code(code: str) -> CategoryDefinition:
    """
    Get a category definition by its code.

    Args:
        code: Category code (e.g., "AR-COLLECT")

    Returns:
        CategoryDefinition if found, UNCATEGORIZED otherwise.
    """
    for cat in get_all_categories():
        if cat.code == code:
            return cat
    return UNCATEGORIZED


def get_categories_by_section(section: CFSection) -> List[CategoryDefinition]:
    """
    Get all categories for a specific Cash Flow section.

    Args:
        section: CFSection enum value

    Returns:
        List of CategoryDefinition objects for the section.
    """
    return [cat for cat in get_all_categories() if cat.section == section]


def get_categories_by_line_item(line_item: str) -> List[CategoryDefinition]:
    """
    Get all categories for a specific Cash Flow line item.

    Args:
        line_item: Cash Flow Statement line item name

    Returns:
        List of CategoryDefinition objects for the line item.
    """
    return [cat for cat in get_all_categories() if cat.line_item == line_item]


# =============================================================================
# SECTION 9: CASH FLOW LINE ITEMS STRUCTURE
# =============================================================================

CASH_FLOW_STRUCTURE: Dict[str, Dict] = {
    "Operating": {
        "header": "CASH FLOWS FROM OPERATING ACTIVITIES",
        "line_items": [
            "Collections from Customers",
            "Payments to Suppliers",
            "Strategic Supplier Payments",
            "Government Payments",
            "Payroll Disbursements",
            "Utilities & Rent",
            "Other Operating Expenses",
        ],
        "subtotal": "NET CASH FROM OPERATING ACTIVITIES",
    },
    "Investing": {
        "header": "CASH FLOWS FROM INVESTING ACTIVITIES",
        "line_items": [
            "Time Deposit Placements",
            "Time Deposit Maturities",
            "Interest Received on TDs",
            "Purchase of Fixed Assets",
            "Sale of Fixed Assets",
            "Purchase of Investments",
            "Sale of Investments",
        ],
        "subtotal": "NET CASH FROM INVESTING ACTIVITIES",
    },
    "Financing": {
        "header": "CASH FLOWS FROM FINANCING ACTIVITIES",
        "line_items": [
            "Loan Drawdowns",
            "Loan Repayments - Principal",
            "Loan Repayments - Interest",
            "Capital Contributions",
            "Dividends Paid",
        ],
        "subtotal": "NET CASH FROM FINANCING ACTIVITIES",
    },
}
