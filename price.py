from fastapi import FastAPI, Header, HTTPException
from pydantic import BaseModel
from typing import List
import os
import json

app = FastAPI()

# 读取 Secret File
PRICE_FILE_PATH = "/etc/secrets/price.json"

with open(PRICE_FILE_PATH, "r", encoding="utf-8") as f:
    PRICE_TABLE = json.load(f)

def get_unit_price(weight: float) -> float:
    for rule in PRICE_TABLE:
        if rule["min"] <= weight <= rule["max"]:
            return rule["price"]
    return 0


# ===== 请求模型 =====
class Item(BaseModel):
    weight: float

class CalcRequest(BaseModel):
    items: List[Item]


# ===== Excel 调用的路由 =====
@app.post("/calc")
def calc_total(
    req: CalcRequest,
    x_api_key: str = Header(None)
):
    # Token 校验
    if x_api_key != os.getenv("API_TOKEN"):
        raise HTTPException(status_code=403, detail="Forbidden")

    total = 0.0
    for item in req.items:
        unit = get_unit_price(item.weight)
        total += unit * item.weight

    # 返回总价
    return {"total": total}
