from fastapi import FastAPI, UploadFile
from src import files_handler
from uvicorn import run
from openpyxl import load_workbook
import logging

app = FastAPI()


@app.post("/uploadfiles/")
async def create_upload_files(files: list[UploadFile]):
    logging.info(files)
    res = await files_handler(files)
    return res

if __name__ == "__main__":
    run(app, host="0.0.0.0", port=8080)