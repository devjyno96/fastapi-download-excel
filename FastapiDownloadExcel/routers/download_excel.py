from fastapi import APIRouter, Depends, status

from typing import Optional

from ..repository import download_excel

from typing import List

router = APIRouter(
    prefix='/download_excel',
    tags=['Download Excel']
)


# Table Group
@router.get('/',
            status_code=status.HTTP_200_OK,
            )
def download_excel():
    """
    """
    return download_excel.test()
