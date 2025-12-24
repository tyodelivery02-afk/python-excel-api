from fastapi import FastAPI, UploadFile, Form, Request
from fastapi.responses import Response, JSONResponse
from fastapi.exceptions import RequestValidationError

from process_excel import process_excel
from price import router as price_router

app = FastAPI()

# 注册 price.py 里的 /calc
app.include_router(price_router)


# 🔥 全局校验错误处理器
@app.exception_handler(RequestValidationError)
async def validation_exception_handler(request: Request, exc: RequestValidationError):
    print("❌ VALIDATION ERROR")
    print("URL:", request.url)
    print("HEADERS:", dict(request.headers))

    try:
        body = await request.body()
        print("RAW BODY:", body)
    except Exception as e:
        print("FAILED TO READ BODY:", e)

    print("DETAIL:", exc.errors())

    return JSONResponse(
        status_code=422,
        content={"detail": exc.errors()},
    )


@app.post("/process")
async def process(file: UploadFile, statsData: str = Form(...)):
    # Excel バイナリを読み込む
    input_bytes = await file.read()

    # Excel 処理
    output_bytes = process_excel(input_bytes, statsData)

    # 処理後 Excel を返す
    return Response(
        content=output_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": "attachment; filename=filled.xlsx"
        }
    )
