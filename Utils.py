import shutil
import os
import os.path
from os import path
from openpyxl.worksheet.datavalidation import DataValidation


class Utils:

    @staticmethod
    def companyNameLookUpMethod(companyName):
        companyDict = {
            'ALL': 'ALL',
            'APL': 'APL',
            'ABI': 'ABI',
            'BHC': 'BHC',
            'CPH': 'CPH',
            "EPP": "EPP",
            'FFI': 'Foremost Farms, Inc.',
            'FTC': 'FTC',
            'GDC': 'GDC',
            'HII': 'HII',
            'LRC': 'LRC',
            'LTG': 'LTG',
            'DIR': 'LTG DIR',
            'MAC': 'MAC',
            'PAL': 'PAL',
            'PNB': 'PNB',
            'PMI': 'PMI',
            'RAP': 'RAP',
            'TYK': 'TYK',
            'TDI': 'TDI',
            'CHI': 'CHI',
            'SPV': 'SPV-AMC Group',
            'TMC': 'TMC',
            'UNI': 'UNI',
            'UER': 'UER',
            'VMC': 'VMC',
            'ZHI': 'ZHI',
            'STN': 'STN',
            'PAN': 'PAN',
            'ANA': 'ANA',
            'LTC': 'LTC',
            'OGC': 'OGC'
        }
        company_Code = ""
        for key, value in companyDict.items():
            if companyName.strip() == value:
                company_Code = key

        return company_Code

    @staticmethod
    def createSubCompanyFolder(companyCode, out):
        companyDir = os.path.join(out, companyCode)

        # creating new DIR base on company code
        if not path.exists(companyDir):
            os.mkdir(os.path.join(out, companyCode))

        pass
