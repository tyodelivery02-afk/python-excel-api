from fastapi import FastAPI, UploadFile, Form
from fastapi.responses import Response
from process_excel import process_excel

app = FastAPI()

@app.post("/process")
async def process(file: UploadFile, statsData: str = Form(...)):
    # 读取 Excel 的二进制
    input_bytes = await file.read()

    # 调用处理函数
    output_bytes = process_excel(input_bytes, statsData)

    # 返回处理后的 Excel
    return Response(
        content=output_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": "attachment; filename=filled.xlsx"
        }
    )
