from typing import List

from fastapi import APIRouter, HTTPException
from sqlmodel import select

from app.core.db import get_session
from app.models.models import (
    Asset,
    AssetCreate,
    AssetRead,
    Dividend,
    DividendCreate,
    DividendRead,
    Transaction,
    TransactionCreate,
    TransactionRead,
)


router = APIRouter(prefix="/assets", tags=["assets"])


@router.get("/", response_model=List[AssetRead])
def list_assets() -> List[AssetRead]:
    with get_session() as session:
        assets = session.exec(select(Asset).order_by(Asset.symbol)).all()
        return assets


@router.post("/", response_model=AssetRead)
def create_asset(asset: AssetCreate) -> AssetRead:
    with get_session() as session:
        exists = session.exec(select(Asset).where(Asset.symbol == asset.symbol)).first()
        if exists:
            raise HTTPException(status_code=400, detail="Asset symbol already exists")
        db_asset = Asset(**asset.model_dump())
        session.add(db_asset)
        session.commit()
        session.refresh(db_asset)
        return db_asset


@router.get("/{asset_id}", response_model=AssetRead)
def get_asset(asset_id: int) -> AssetRead:
    with get_session() as session:
        asset = session.get(Asset, asset_id)
        if not asset:
            raise HTTPException(status_code=404, detail="Asset not found")
        return asset


@router.delete("/{asset_id}")
def delete_asset(asset_id: int) -> dict:
    with get_session() as session:
        asset = session.get(Asset, asset_id)
        if not asset:
            raise HTTPException(status_code=404, detail="Asset not found")
        session.delete(asset)
        session.commit()
        return {"ok": True}


# Transactions
@router.get("/{asset_id}/transactions", response_model=List[TransactionRead])
def list_transactions(asset_id: int) -> List[TransactionRead]:
    with get_session() as session:
        stmt = (
            select(Transaction)
            .where(Transaction.asset_id == asset_id)
            .order_by(Transaction.date)
        )
        return session.exec(stmt).all()


@router.post("/{asset_id}/transactions", response_model=TransactionRead)
def create_transaction(asset_id: int, tx: TransactionCreate) -> TransactionRead:
    if tx.asset_id != asset_id:
        raise HTTPException(status_code=400, detail="asset_id mismatch")
    with get_session() as session:
        if not session.get(Asset, asset_id):
            raise HTTPException(status_code=404, detail="Asset not found")
        db_tx = Transaction(**tx.model_dump())
        session.add(db_tx)
        session.commit()
        session.refresh(db_tx)
        return db_tx


# Dividends
@router.get("/{asset_id}/dividends", response_model=List[DividendRead])
def list_dividends(asset_id: int) -> List[DividendRead]:
    with get_session() as session:
        stmt = (
            select(Dividend)
            .where(Dividend.asset_id == asset_id)
            .order_by(Dividend.date_received)
        )
        return session.exec(stmt).all()


@router.post("/{asset_id}/dividends", response_model=DividendRead)
def create_dividend(asset_id: int, div: DividendCreate) -> DividendRead:
    if div.asset_id != asset_id:
        raise HTTPException(status_code=400, detail="asset_id mismatch")
    with get_session() as session:
        if not session.get(Asset, asset_id):
            raise HTTPException(status_code=404, detail="Asset not found")
        db_div = Dividend(**div.model_dump())
        session.add(db_div)
        session.commit()
        session.refresh(db_div)
        return db_div
