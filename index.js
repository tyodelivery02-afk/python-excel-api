import express from "express";
import fs from "fs";

const app = express();
app.use(express.json());

// 读取 Secret File
const PRICE_TABLE = JSON.parse(
  fs.readFileSync("/etc/secrets/price.json", "utf8")
);

function getUnitPrice(weight) {
  const rule = PRICE_TABLE.find(
    r => weight >= r.min && weight <= r.max
  );
  return rule ? rule.price : 0;
}

app.post("/calc", (req, res) => {

  // Token 校验
  if (req.headers["x-api-key"] !== process.env.API_TOKEN) {
    return res.status(403).json({ error: "Forbidden" });
  }

  const items = req.body.items;
  if (!Array.isArray(items)) {
    return res.status(400).json({ error: "Invalid request" });
  }

  let total = 0;

  items.forEach(item => {
    const unit = getUnitPrice(item.weight);
    total += unit * item.weight;
  });

  // ✅ 只返回总价
  res.json({ total });
});

app.listen(3000);
