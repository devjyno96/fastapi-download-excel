from fastapi import APIRouter, Depends, status

from typing import Optional

from ..repository import download_excel as download_excel_repository

from fastapi.responses import FileResponse
from typing import List

router = APIRouter(
    prefix='/download_excel',
    tags=['Download Excel']
)


@router.get('/',
            status_code=status.HTTP_200_OK,
            response_class=FileResponse
            )
def download_excel():
    """
    """
    return download_excel_repository.download_excel()

@router.get('/search',
            status_code=status.HTTP_200_OK,
            response_class=FileResponse
            )
def get_search_result():
    """
    """
    return download_excel_repository.get_search_result_repo()
