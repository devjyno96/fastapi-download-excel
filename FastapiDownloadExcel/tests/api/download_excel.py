import time

import json

from FastapiDownloadExcel.tests.BaseTest.basetest import BaseTest


class Order_01_Agent_Setting_Test(BaseTest):

    def setUp(self):
        pass

    # Table Group
    def test_01_test(self):
        self.test_client.get('/download_excel')
        pass

    def test_02_search(self):
        self.test_client.get('/download_excel/search')
        pass

if __name__ == "__main__":
    BaseTest.main()
