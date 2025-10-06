from datetime import date, datetime
from enum import Enum
from typing import List, Optional

from sqlmodel import Field, Relationship, SQLModel


class AssetType(str, Enum):
    stock = "stock"
    etf = "etf"


class AssetBase(SQLModel):
    symbol: str = Field(index=True, max_length=16)
    name: str = Field(max_length=128)
    type: AssetType = Field(default=AssetType.stock, description="stock or etf")


class Asset(AssetBase, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    created_at: datetime = Field(default_factory=datetime.utcnow)

    transactions: List["Transaction"] = Relationship(back_populates="asset")
    dividends: List["Dividend"] = Relationship(back_populates="asset")
    metrics: List["Metric"] = Relationship(back_populates="asset")


class AssetCreate(AssetBase):
    pass


class AssetRead(AssetBase):
    id: int
    created_at: datetime


class TransactionBase(SQLModel):
    date: date
    price_per_share: float = Field(gt=0)
    shares: float = Field(description="positive for buy, negative for sell")
    fees: float = 0.0


class Transaction(TransactionBase, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    asset_id: int = Field(foreign_key="asset.id", index=True)

    asset: Optional[Asset] = Relationship(back_populates="transactions")


class TransactionCreate(TransactionBase):
    asset_id: int


class TransactionRead(TransactionBase):
    id: int
    asset_id: int


class DividendBase(SQLModel):
    date_received: date
    amount_received: float = Field(ge=0)


class Dividend(DividendBase, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    asset_id: int = Field(foreign_key="asset.id", index=True)

    asset: Optional[Asset] = Relationship(back_populates="dividends")


class DividendCreate(DividendBase):
    asset_id: int


class DividendRead(DividendBase):
    id: int
    asset_id: int


class MetricBase(SQLModel):
    key: str = Field(index=True, max_length=64)
    value: Optional[str] = Field(
        default=None, description="placeholder for future metrics"
    )


class Metric(MetricBase, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    asset_id: int = Field(foreign_key="asset.id", index=True)

    asset: Optional[Asset] = Relationship(back_populates="metrics")


class MetricCreate(MetricBase):
    asset_id: int


class MetricRead(MetricBase):
    id: int
    asset_id: int
