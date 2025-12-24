# price.py
from fastapi import APIRouter, Header, HTTPException, Request
from pydantic import BaseModel
from typing import List
import os
import json

router = APIRouter()

PRICE_FILE_PATH = "/etc/secrets/price.json"

with open(PRICE_FILE_PATH, "r", encoding="utf-8") as f:
    PRICE_TABLE = json.load(f)

def get_unit_price(weight: float) -> float:
    for rule in PRICE_TABLE:
        if rule["min"] <= weight <= rule["max"]:
            return rule["price"]
    return 0.0


class Item(BaseModel):
    weight: float

class CalcRequest(BaseModel):
    items: List[Item]


@router.post("/calc")
async def calc_total(
    req: CalcRequest,
    request: Request,
    x_api_key: str = Header(None)
):
    print("HEADERS:", dict(request.headers))
    print("PARSED BODY:", req)

    if x_api_key != os.getenv("API_TOKEN"):
        raise HTTPException(status_code=403, detail="Forbidden")

    total = 0.0
    for item in req.items:
        price = get_unit_price(item.weight)
        total += price

    return {"total": total}
