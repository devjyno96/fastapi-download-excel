from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from .routers import download_excel

app = FastAPI(
    title="FastapiDownloadExcel",
    description="21.11.11 version",
    version="0.0.1",
)

origins = [
    "http://localhost",
    "http://localhost:8080",
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(download_excel.router)
