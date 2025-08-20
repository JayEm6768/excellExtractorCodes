from fastapi import FastAPI, UploadFile
from fastapi.responses import FileResponse
import pandas as pd
import tempfile, os, zipfile

app = FastAPI()

@app.post("/api/process")
async def process(file: UploadFile):
    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = os.path.join(tmpdir, file.filename)
        with open(input_path, "wb") as f:
            f.write(await file.read())

        # >>> Place your Excel processing code here <<< 
        # (save outputs into tmpdir)

        zip_path = os.path.join(tmpdir, "results.zip")
        with zipfile.ZipFile(zip_path, "w") as zipf:
            for f in os.listdir(tmpdir):
                if f.endswith(".xlsx"):
                    zipf.write(os.path.join(tmpdir, f), f)

        return FileResponse(zip_path, filename="results.zip")
