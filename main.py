import os

import numpy as np
import openpyxl
import pandas as pd
import os.path
import shutil
import logging
import re
import dateutil.parser
from datetime import datetime
from pathlib import Path
from openpyxl.styles import Border, Side
# from Utils import Utils
from os import path
from Utils import Utils
# from dateutil.parser import parse
import itertools


def folderStructureCreation(dirPath):
    print("Checking Folder structure...")

    folderList = ['in', 'out', 'log', 'error', 'template']

    for folder in folderList:
        folderPath = os.path.join(dirPath, folder)

        # Check if folder exist
        if not path.exists(folderPath):
            os.makedirs(folderPath)

    if not os.listdir(inPath):
        print(inPath + ' is empty.', end="")
        quit()

    print("Done!")

    pass


def processError(errMessage, companyName, errorFileNAme):
    newdf = pd.DataFrame(columns=['row', 'error message'])

    # Error Index - remove duplicate in the list
    errMessage = list(dict.fromkeys(errMessage))

    for error in errMessage:
        if error is not None:
            err = error.split("-")

            newdf = newdf.append({'row': int(err[0].strip()), 'error message': err[1]},
                                 ignore_index=True)

    errMsg = []
    throwErrMsg = []

    # newdf.Col = pd.to_numeric(newdf['row'], errors='coerce')
    newdf.sort_values(by=['row'], inplace=True, ascending=True)
    # print(newdf)

    groups = newdf.groupby('row')
    for comp, records in groups:
        for j, row in records.iterrows():
            errMsg.append(row['error message'])

        throwErrMsg.append("Error: Row " + str(comp) + " - [" + ', '.join(errMsg) + "]")

        errMsg.clear()

    generateErrorLog(throwErrMsg, companyName, errorFileNAme)

    arrErr = []

    arrErr.extend(newdf['row'].tolist())

    return arrErr


def generateErrorLog(errMsg, companyCode, arg):
    util = Utils()

    if len(errMsg):
        util.createSubCompanyFolder(companyCode, errPath)
        f = open(
            errPath + "/" + companyCode + "/" + companyCode + "_" + arg + "_err_log_" + dateTime + ".txt",
            "a")
        for err in errMsg:
            f.writelines(err + "\n")

        errMsg.clear()
    pass


def dateValidator(row, fdose, sdose, param, columnName, vaccinated):
    arr_err = []

    # print(row)
    if columnName == '1st Dose Vaccination Date (MM/DD/YYYY)':
        # if (fdose != 'N/A' and fdose != ''):
        if (fdose != 'N/A' and fdose != '') and (param == "N/A" or not param):
            arr_err.append(str(row + 2) + '-1st Dose Vaccination Date (MM/DD/YYYY) column should not be blank or N/A.')
        else:
            if (param != 'N/A' and param != '' and pd.notna(param)) and (
                    fdose != '' and fdose != 'N/A' and pd.isna(fdose)):
                try:
                    datetime.strptime(str(param), '%Y-%m-%d %X')
                    # datetime.strptime(str(param), '%Y-%m-%d')
                except ValueError:
                    arr_err.append(str(row + 2) + '-' + columnName + 'column incorrect data format, should be '
                                                                     'MM/DD/YYYY.')
                    # print(str(row + 2) + '-' + str(param) + ' incorrect data format, should be '
                    #                                         'MM/DD/YYYY.')

        if fdose == 'N/A' and (param != '' and param != 'N/A' and pd.notna(param)):
            arr_err.append(str(row + 2) + '-1st Dose Vaccination Date (MM/DD/YYYY) column should be N/A.')

        if fdose == '' and param != '':
            arr_err.append(str(row + 2) + '-1st Dose Vaccination Date (MM/DD/YYYY) column should be N/A.')

        if fdose == '':
            arr_err.append(str(row + 2) + '-1st dose (LGU or LTGC) column should be N/A.')

    elif columnName == '2nd Dose Vaccination Date (MM/DD/YYYY)':

        # if sdose == '':
        #     arr_err.append(str(row + 2) + '-2nd dose (LGU or LTGC) column should be N/A.')

        if (sdose == 'N/A' and (param != '' and param != 'N/A' and pd.notna(param))) and vaccinated.strip() == 'no':
            arr_err.append(str(row + 2) + '-2nd Dose Vaccination Date (MM/DD/YYYY) column should be N/A.')

        if sdose == '' and (param != '' and param != 'N/A'):
            arr_err.append(str(row + 2) + '-2nd Dose Vaccination Date (MM/DD/YYYY) column should be N/A.')

        if sdose == 'N/A' and (param != '' and param != 'N/A' and pd.notna(param)):
            arr_err.append(str(row + 2) + '-2nd Dose Vaccination Date (MM/DD/YYYY) column should be N/A.')

        if (param != 'N/A' and param != '' and pd.notna(param)) and (sdose != '' and sdose != 'N/A'):
            try:
                datetime.strptime(str(param), '%Y-%m-%d %X')
                # datetime.strptime(str(param), '%Y-%m-%d')
            except ValueError:
                arr_err.append(str(row + 2) + '-' + columnName + 'column incorrect data format, should be '
                                                                 'MM/DD/YYYY.')
                # print(str(row + 2) + '-' + str(param) + ' incorrect data format, should be '
                #                                         'MM/DD/YYYY.')

    return arr_err


def validateDateFormat(df, companyName):
    arrYearValidate = ['1st Dose Vaccination Date (MM/DD/YYYY)', '2nd Dose Vaccination Date (MM/DD/YYYY)']
    errMessage = []

    for field in arrYearValidate:
        errMessage.extend(df.apply(lambda x: dateValidator(x.name,
                                                           x['1st dose \n (LGU or LTGC)'],
                                                           x['2nd dose \n (LGU or LTGC)'],
                                                           x[field],
                                                           field,
                                                           x['Vaccinated'], ), axis=1))

    errMessage.extend(df.apply(lambda x: validatedDate(x.name, x['1st Dose Date'], x['2nd Dose Date']),
                               axis=1))

    errMessage = list(itertools.chain.from_iterable(errMessage))

    # for i in errMessage:
    #     print(i)
    return processError(errMessage, companyName, 'ValidateDateFormat')


def CheckOtherEmptyField(row, field, columnName):
    arr_err = []
    typeofEmp = ['Employee', 'Third Party Service Provider']

    if columnName == 'Employee Number':
        if field == '':
            arr_err.append(str(row + 2) + '-' + columnName + ' column should not be blank.')

    if columnName == 'Vaccinated':
        if field == '' or field == 'N/A':
            arr_err.append(str(row + 2) + '-' + columnName + ' column should be Yes or No response.')

    if columnName == 'Company Name in EZ':
        if field == '':
            arr_err.append(str(row + 2) + '-' + columnName + ' column should not be blank.')

    if columnName == 'Type of Employee':
        istypeofEmp = field.lower() in list(map(lambda x: x.lower(), typeofEmp))

        if field == '':
            arr_err.append(str(row + 2) + '-' + columnName + ' column should not be blank. ')
        elif not istypeofEmp:
            arr_err.append(str(row + 2) + '-Type of Employee ' + field + ' is not in the list of employees.')

    return arr_err


def ValidateOtherEmptyField(df, companyName):
    errMessage = []
    arrOtherField = ['Employee Number', 'Company Name in EZ', 'Type of Employee', 'Vaccinated']



    for columnName in arrOtherField:
        errMessage.extend(df.apply(lambda x: CheckOtherEmptyField(x.name,
                                                                  x[columnName],
                                                                  columnName), axis=1))

    errMessage = list(itertools.chain.from_iterable(errMessage))

    # for i in errMessage:
    #     print(i)
    return processError(errMessage, companyName, 'otherColumn')


def VaccinatedYesNo(row, Vaccinated, param, columnName):
    arr_err = []

    vaccineBrand = ['AstraZeneca', 'Covaxin', 'Janssen by J&J', 'Moderna',
                    'Pfizer-BioNTech', 'Sinopharm', 'Sinovac', 'Sputnik V',
                    'N/A']

    if str(Vaccinated).lower() == 'yes':
        if columnName == 'Vaccine Brand':
            isOtherVaccineBrand = str(param).lower() in list(map(lambda x: x.lower(), vaccineBrand))

            if param == '':
                arr_err.append(str(row + 2) + '-Vaccine Brand column should not be Blank, put N/A instead.')
            elif not isOtherVaccineBrand:
                arr_err.append(str(row + 2) + '-Vaccine Brand ' + str(param) + ' is not in the list of brands.')

            if param == 'N/A':
                arr_err.append(str(row + 2) + '-Vaccine Brand column should not be N/A if Vaccinated column '
                                              'is Yes')
        if columnName == '1st dose \n (LGU or LTGC)':
            if param == '':
                arr_err.append(str(row + 2) + '-1st dose (LGU or LTGC) column should not be Blank, put N/A instead.')
            # if param == 'N/A':
            #     arr_err.append(str(row + 2) + '-1st dose (LGU or LTGC) column should not be N/A if '
            #                                   'Vaccinated column is Yes ')
        if columnName == '2nd dose \n (LGU or LTGC)':
            if param == '':
                arr_err.append(str(row + 2) + '-2nd dose (LGU or LTGC) column should not be Blank, put N/A instead.')
            # if param == 'N/A':
            #     arr_err.append(str(row + 2) + '-2nd dose (LGU or LTGC) column should not be N/A if '
            #                                   'Vaccinated column is Yes ')

    elif str(Vaccinated).lower() == 'no':
        if columnName == 'Vaccine Brand':
            if param == '':
                arr_err.append(str(row + 2) + '-Vaccine Brand column should be N/A if Vaccinated column is No')
            if param != 'N/A':
                arr_err.append(str(row + 2) + '-Vaccine Brand column should be N/A if Vaccinated column is No')

        if columnName == '1st dose \n (LGU or LTGC)':
            if param == '':
                arr_err.append(str(row + 2) + '-1st dose (LGU or LTGC) column should be N/A if Vaccinated column is No')
            if param != 'N/A':
                arr_err.append(str(row + 2) + '-1st dose (LGU or LTGC) column should be N/A if Vaccinated column is No')

        if columnName == '2nd dose \n (LGU or LTGC)':
            if param == '':
                arr_err.append(str(row + 2) + '-2nd dose (LGU or LTGC) column should be N/A if Vaccinated column is No')
            if param != 'N/A':
                arr_err.append(str(row + 2) + '-2nd dose (LGU or LTGC) column should be N/A if Vaccinated column is No')

    else:
        arr_err.append(str(row + 2) + '-2nd dose (LGU or LTGC) column invalid entry.')

    return arr_err


def ValidateVaccinatedField(df, companyName):
    errMessage = []
    arrVaccinated = ['Vaccine Brand', '1st dose \n (LGU or LTGC)', '2nd dose \n (LGU or LTGC)',
                     '1st Dose Vaccination Date (MM/DD/YYYY)']

    for columnName in arrVaccinated:
        errMessage.extend(df.apply(lambda x: VaccinatedYesNo(x.name,
                                                             x['Vaccinated'],
                                                             x[columnName],
                                                             columnName), axis=1))

    errMessage = list(itertools.chain.from_iterable(errMessage))
    for i in errMessage:
        print(i)
    return processError(errMessage, companyName, 'VaccinatedYesNo')


def ValidateVaccineYear(row, fdoseYear, sDoseYear, field, columnName):
    arr_err = []

    if columnName == '1st Dose Vaccination Date (MM/DD/YYYY)':
        if str(fdoseYear) != 'nan' and str(fdoseYear) != 'N/A':
            if (fdoseYear != '2020' and fdoseYear != '2021') and field != '':
                arr_err.append(str(row + 2) + '-1st Dose Vaccination Date (MM/DD/YYYY) column year is invalid.')
    elif columnName == '2nd Dose Vaccination Date (MM/DD/YYYY)':
        if str(sDoseYear) != 'nan' and str(fdoseYear) != 'N/A':
            if (sDoseYear != '2020' and sDoseYear != '2021') and field != '':
                arr_err.append(str(row + 2) + '-2nd Dose Vaccination Date (MM/DD/YYYY) column year is invalid.')

    return arr_err


def ValidateYear(df, companyName):
    arrYearValidate = ['1st Dose Vaccination Date (MM/DD/YYYY)', '2nd Dose Vaccination Date (MM/DD/YYYY)']
    errMessage = []
    for field in arrYearValidate:
        errMessage.extend(df.apply(lambda x: ValidateVaccineYear(x.name,
                                                                 x['1st Dose year'],
                                                                 x['2nd Dose year'],
                                                                 x[field],
                                                                 field), axis=1))
    errMessage = list(itertools.chain.from_iterable(errMessage))
    # for i in errMessage:
    #     print(i)
    return processError(errMessage, companyName, 'EmptyField')


def convertDateFormat(name, param):
    if pd.notna(param):
        return dateutil.parser.parse(str(param)).strftime("%Y-%m-%d")
    else:
        return param


def validatedDate(row, firstDoseDate, secondDoseDate):
    arr_err = []
    if pd.notna(firstDoseDate) and pd.notna(secondDoseDate):

        if not (firstDoseDate < secondDoseDate) and (firstDoseDate != secondDoseDate):
            arr_err.append(str(row + 2) + '-2nd Dose Vaccination Date (MM/DD/YYYY) column should not be less than 1st '
                                          'Dose Vaccination Date (MM/DD/YYYY), Date is invalid.')

        elif firstDoseDate == secondDoseDate:
            arr_err.append(str(row + 2) + '-1st Dose Vaccination Date (MM/DD/YYYY) column should not be equal to 2nd '
                                          'Dose Vaccination Date (MM/DD/YYYY), Date is invalid.')

    return arr_err


def getData(fileName):
    filePath = os.path.join(inPath, fileName)
    df = pd.read_excel(filePath, na_filter=False, dtype=str)

    # Get Filename Company
    fileName = os.path.splitext(fileName)[0]
    companyName = fileName.split("_")[1]

    rowIDErrors = []
    # rowIDErrors.extend(validateDateFormat(df, companyName))
    rowIDErrors.extend(ValidateOtherEmptyField(df, companyName))
    rowIDErrors.extend(ValidateVaccinatedField(df, companyName))

    # print(df['2nd Dose Vaccination Date (MM/DD/YYYY)'].map(type))
    # print(df)

    df['1st Dose Vaccination Date (MM/DD/YYYY)'] = pd.to_datetime(df['1st Dose Vaccination Date (MM/DD/YYYY)'],
                                                                  errors='coerce')
    df['2nd Dose Vaccination Date (MM/DD/YYYY)'] = pd.to_datetime(df['2nd Dose Vaccination Date (MM/DD/YYYY)'],
                                                                  errors='coerce')

    df['1st Dose Date'] = df.apply(lambda x: convertDateFormat(x.name, x['1st Dose Vaccination Date (MM/DD/YYYY)']),
                                   axis=1)
    df['2nd Dose Date'] = df.apply(lambda x: convertDateFormat(x.name, x['2nd Dose Vaccination Date (MM/DD/YYYY)']),
                                   axis=1)

    df['1st Dose Date'] = pd.to_datetime(df['1st Dose Date'], errors='coerce')
    df['2nd Dose Date'] = pd.to_datetime(df['2nd Dose Date'], errors='coerce')

    rowIDErrors.extend(validateDateFormat(df, companyName))
    print(df)

    df['1st Dose year'] = df['1st Dose Vaccination Date (MM/DD/YYYY)'].dt.strftime("%Y")
    df['2nd Dose year'] = df['2nd Dose Vaccination Date (MM/DD/YYYY)'].dt.strftime("%Y")
    df['1st Dose Vaccination Date (MM/DD/YYYY)'] = df['1st Dose Vaccination Date (MM/DD/YYYY)'].dt.strftime("%m/%d/%Y")
    df['2nd Dose Vaccination Date (MM/DD/YYYY)'] = df['2nd Dose Vaccination Date (MM/DD/YYYY)'].dt.strftime("%m/%d/%Y")

    rowIDErrors.extend(ValidateYear(df, companyName))
    df = df.replace(np.nan, 'N/A')

    # Error Index - remove duplicate in the list
    errorIndex = list(dict.fromkeys(rowIDErrors))

    # convert string element to int
    errorIndex = [int(i) for i in errorIndex]

    for i in range(len(errorIndex)):
        errorIndex[i] = int(errorIndex[i]) - 2

    # df.drop(df.index[errorIndex], inplace=True)

    # print(df)
    return df


def dropUnwantedColumns(master_df):
    master_df.drop(columns=['1st Dose year',
                            '2nd Dose year'], inplace=True)

    pass


def duplicateTemplateLTGC(tempLTGC_Path, out, outputFilename):
    companyDir = out + "/"
    srcFile = companyDir + outputFilename + ".xlsx"

    if not os.path.isfile(srcFile):
        shutil.copy(tempLTGC_Path, srcFile)

    return companyDir + outputFilename + ".xlsx"


if __name__ == '__main__':
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)
    pd.set_option('display.width', None)

    today = datetime.today()
    dateTime = today.strftime("%m%d%y_%H%M%S")

    # Windows OS Path
    dirPath = 'C:\\Users\\admin\\Documents\\Vaccine Project\\UpdatedCountVaccineesConso'

    inPath = os.path.join(dirPath, "in")
    outPath = os.path.join(dirPath, "out")
    errPath = os.path.join(dirPath, "error")
    logPath = os.path.join(dirPath, "log")
    backupPath = os.path.join(dirPath, "backup")
    # templateFilePath = os.path.join(dirPath, "template/template_v1.xlsx")
    templateFilePath = 'src/template/template_v1.xlsx'

    outFilename = 'Target VACCINEE Conso_' + dateTime

    # Folder Structure Creation
    folderStructureCreation(dirPath)

    logging.info("==============================================================")
    logging.info("Running Scpirt: REgistration of Moderna Order for HH Consolidation......")
    logging.info("==============================================================")

    # Get all Files in
    arrFilenames = os.listdir(inPath)
    arrdf = []

    for inFile in arrFilenames:
        if not inFile == ".DS_Store":
            print("Reading: " + inFile + "......")

            arrdf.append(getData(inFile))

    master_df = pd.concat(arrdf)

    dropUnwantedColumns(master_df)

    # print(master_df)

    # Create copy of template file and save it to out folder
    templateFile = duplicateTemplateLTGC(templateFilePath, outPath, outFilename)

    # Write df_master(consolidated/append data) to excel
    writer = pd.ExcelWriter(templateFile, engine='openpyxl', mode='a')
    writer.book = openpyxl.load_workbook(templateFile)
    writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)
    master_df.to_excel(writer, sheet_name="Sheet1", startrow=1, header=False, index=False)
    writer.save()
