from decimal import Decimal
from typing import List, Optional

from pydantic import BaseModel, Field

# Treasury transactions export. Labels arrive already resolved in Spanish from
# valera and dates as Bogota wall-clock times ("YYYY-MM-DDTHH:MM"); this service
# only lays them out. Amounts are Decimal so cents survive the JSON hop.

MAX_TRANSACTION_ROWS = 30000


class TreasuryTransactionRow(BaseModel):
    reference: str
    status: str
    registeredAt: str
    decidedAt: Optional[str] = None
    companyName: str
    bankName: str
    accountNumber: str
    paymentMethod: str
    batchReference: str = ""
    isQr: bool = False
    payeeTaxId: str = ""
    payeeName: str = ""
    description: str = ""
    amount: Decimal
    currencyLabel: str
    treasurerName: str = ""


class TreasuryCurrencyTotal(BaseModel):
    currencyLabel: str
    count: int
    total: Decimal


class TreasuryTransactionsExportRequest(BaseModel):
    generatedAt: str
    generatedBy: str = ""
    filtersSummary: str = ""
    rows: List[TreasuryTransactionRow] = Field(..., max_items=MAX_TRANSACTION_ROWS)
    totals: List[TreasuryCurrencyTotal] = []


MAX_BALANCE_ROWS = 50000


class TreasuryBalanceRow(BaseModel):
    cutDate: str
    slot: str
    balanceGroupName: str
    companyName: str
    bankName: str
    accountNumber: str
    currencyLabel: str
    bankBalance: Decimal
    bookBalance: Optional[Decimal] = None
    copEquivalent: Optional[Decimal] = None


class TreasuryBalancesExportRequest(BaseModel):
    generatedAt: str
    generatedBy: str = ""
    filtersSummary: str = ""
    rows: List[TreasuryBalanceRow] = Field(..., max_items=MAX_BALANCE_ROWS)
