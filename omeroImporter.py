from datetime import datetime
import pathlib
import os
import sys
import time
import json
import shutil
import re
import pandas as pd
import xlrd
import xlwings
import ezomero
from imageio.v2 import imread

# BACKBLAZE
import boto3  # REQUIRED! - Details here: https://pypi.org/project/boto3/
from botocore.exceptions import ClientError
from botocore.config import Config

# from dotenv import load_dotenv  # Project Must install Python Package:  python-dotenv

# EMAIL
import email
import smtplib
import ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# CRYPTO
from cryptography.fernet import Fernet

# OMERO
import ezomero as ezome
import omero.clients
from omero.gateway import (
    MapAnnotationWrapper,
    PlateWrapper,
    ScreenWrapper,
)
from omero.model import (
    ImageI,
    ScreenPlateLinkI,
    PlateI,
    WellI,
    WellSampleI,
    ScreenI,
)
from omero.gateway import BlitzGateway
from omero.rtypes import rint

dateFormatter = "%d-%m-%Y_%H-%M-%S"

outputPreviousImportedFileName = "OmeroImporter_previousImported.txt"
outputLogFileName = "OmeroImporter_log.txt"
# outputMetadataLogFileName = "OmeroImporter_metadata_log.txt"
outputImportedFileName = "OmeroImporter_imported.txt"
# outputMetadataFileName = "OmeroImporter_metadata.txt"
configFileFolder = "OmeroImporter"
configFileName = "OmeroImporter.cfg"
keyFileName = "OmeroImporter.key"

outputLogFilePath = None
# outputMetadataLogFilePath = None
outputImportedFilePath = None
# outputMetadataFilePath = None

parameters = None
p_omeroHostname = "hostName"
p_omeroPort = "port"
p_omeroUsername = "userName"
p_omeroPSW = "userPassword"
p_target = "target"
p_dest = "destination"
# p_headless = "headless"
p_delete = "hasDelete"
p_mma = "hasMMA"
p_b2 = "hasB2"
p_b2_endpoint = "b2Endpoint"
p_b2_bucketName = "b2BucketName"
p_b2_appKeyId = "b2AppKeyId"
p_b2_appKey = "b2AppKey"
p_userEmail = "userEmail"
p_adminsEmail = "adminsEmail"
p_emailFrom = "senderEmail"
p_emailFromPSW = "sendEmailPSW"
p_startTime = "startTime"
p_endTime = "endTime"
p_key = "key"

import_status = "import"
import_status_imported = "imported"
import_status_pimported = "previously imported"
import_status_found = "found"
import_status_id = "id"
import_path = "path"
import_annotate = "annotated"

metadata_plates = "plates"
metadata_wells = "wells"
metadata_site = "Site"
metadata_well_name = "Well_Name"
metadata_OME_Images = "OME_Images"
metadata_OME_Image_Name = "OME_Image_Name"
metadata_files = "Files"

metadata_file_name = "File_Name"
metadata_file_path = "File_Path"
metadata_image_mma = "MMA_File_Path"
metadata_C_index = "C"
metadata_Z_index = "Z"
metadata_T_index = "T"
metadate_Site_index = "Site"
metadata_image_tags1 = "Tags"
metadata_image_tags2 = "Τags"

excel_module_ome = "OMERO specific"

excel_screen = 0
excel_screenName = 'Screen_Name'

excel_plate = 1
excel_plateName = "Plate_Name"

excel_plateMap = 2

excel_WellImageMap = 3
excel_module = "Module"
excel_key = "Key"
excel_value = "Value"
excel_replaceNaN = "EMPTY-PD-VALUE"

excel_plateMapTableStartCell = "A14"
excel_datasetCell = "C10"


class WrappedException(Exception):
    def __init__(self, info, e):
        self.exception = e
        super().__init__(info)


# Return a boto3 client object for B2 service
def get_b2_client(endpoint, keyID, applicationKey):
    b2_client = boto3.client(
        service_name="s3",
        endpoint_url=endpoint,
        aws_access_key_id=keyID,
        aws_secret_access_key=applicationKey,
    )
    return b2_client


# Return a boto3 resource object for B2 service
def get_b2_resource(endpoint, keyID, applicationKey):
    b2 = boto3.resource(
        service_name="s3",
        endpoint_url=endpoint,
        aws_access_key_id=keyID,
        aws_secret_access_key=applicationKey,
        config=Config(
            signature_version="s3v4",
        ),
    )
    return b2


def upload_file(bucketName, filePath, fileName, b2, b2path=None):
    # filePath = directory + '/' + file
    remotePath = b2path
    if remotePath is None:
        remotePath = fileName
    else:
        remotePath = re.sub(r"\\", "/", remotePath)
    printToConsole("remotePath " + remotePath)
    try:
        response = b2.Bucket(bucketName).upload_file(filePath, remotePath)
    except ClientError as ce:
        raise
    return response


def mergeDictionaries(dict1, dict2):
    dict = deepCopyDictionary(dict1)
    mergedDict = deepMergeDictionaries(dict, dict2)
    return mergedDict


def deepCopyDictionary(dict1):
    newDict = {}
    for key in dict1:
        if isinstance(dict1[key], dict):
            newDict[key] = deepCopyDictionary(dict1[key])
        else:
            newDict[key] = dict1[key]
    return newDict


def deepMergeDictionaries(dict1, dict2):
    newDict = dict1
    for key in dict2:
        if isinstance(dict2[key], dict):
            if key in dict1:
                newDict[key] = deepMergeDictionaries(dict1[key], dict2[key])
            else:
                newDict[key] = deepCopyDictionary(dict2[key])
        else:
            newDict[key] = dict2[key]
    return newDict


def writeConfigFile(path, dict):
    configFilePath = os.path.join(path, configFileName)
    keyFilePath = os.path.join(path, keyFileName)
    try:
        with open(keyFilePath, "w") as f:
            try:
                f.write(str(dict[p_key]))
                f.close()
            except (FileNotFoundError, PermissionError, OSError) as e:
                printToConsole("Writing key file failed for " + keyFilePath)
                printToConsole(repr(e))
    except (IOError, OSError) as e:
        printToConsole("Writing key file failed for " + keyFilePath)
        printToConsole(repr(e))
    try:
        with open(configFilePath, "w") as f:
            try:
                for key in dict:
                    if key == p_key:
                        continue
                    value = dict[key]
                    f.write(str(key) + " = " + str(value))
                    f.write("\n")
                f.close()
            except (FileNotFoundError, PermissionError, OSError) as e:
                printToConsole("Writing config file failed for " + configFilePath)
                printToConsole(repr(e))
    except (IOError, OSError) as e:
        printToConsole("Writing config file failed for " + configFilePath)
        printToConsole(repr(e))


def readConfigFile(path):
    configFile = os.path.join(path, configFileName)
    configFilePath = pathlib.Path(configFile).resolve()
    keyFile = os.path.join(path, keyFileName)
    keyFilePath = pathlib.Path(keyFile).resolve()
    key = None
    params = {}
    if not configFilePath.exists() or not keyFilePath.exists():
        return params
    try:
        with open(keyFilePath, "r") as f:
            try:
                key = f.readline().strip()
                f.close()
                params[p_key] = key
            except (FileNotFoundError, PermissionError, OSError) as e:
                message = "Opening key file failed for " + keyFilePath
                writeToLog("ERROR: " + message)
                writeToLog(repr(e))
                printToConsole(message)
                printToConsole(repr(e))
    except (IOError, OSError) as e:
        message = "Reading key file failed for " + keyFilePath
        writeToLog("ERROR: " + message)
        writeToLog(repr(e))
        printToConsole(message)
        printToConsole(repr(e))
    try:
        with open(configFilePath, "r") as f:
            try:
                while True:
                    line = f.readline()
                    if not line:
                        break
                    data = line.strip()
                    if data.startswith("//") or data.startswith("#"):
                        continue
                    else:
                        tokens = data.split(" = ")
                        val = tokens[1]
                        params[tokens[0]] = val
                f.close()
                return params
            except (FileNotFoundError, PermissionError, OSError) as e:
                message = "Opening config file failed for " + configFilePath
                writeToLog("ERROR: " + message)
                writeToLog(repr(e))
                printToConsole(message)
                printToConsole(repr(e))
    except (IOError, OSError) as e:
        message = "Reading config file failed for " + configFilePath
        writeToLog("ERROR: " + message)
        writeToLog(repr(e))
        printToConsole(message)
        printToConsole(repr(e))


def writeCurrentImported(dict):
    try:
        with open(outputImportedFilePath, "w") as f:
            try:
                json.dump(dict, f)
            except (FileNotFoundError, PermissionError, OSError) as e:
                message = (
                    "Writing current imported file failed for " + outputImportedFilePath
                )
                writeToLog("ERROR: " + message)
                writeToLog(repr(e))
                printToConsole(message)
                printToConsole(repr(e))
    except (IOError, OSError) as e:
        message = "Writing current imported file failed for " + outputImportedFilePath
        writeToLog("ERROR: " + message)
        writeToLog(repr(e))
        printToConsole(message)
        printToConsole(repr(e))


def writePreviousImported(path, dict):
    importedFilePath = os.path.join(path, outputPreviousImportedFileName)
    try:
        with open(importedFilePath, "w") as f:
            try:
                json.dump(dict, f)
            except (FileNotFoundError, PermissionError, OSError) as e:
                message = (
                    "Writing previous imported file failed for " + importedFilePath
                )
                writeToLog("ERROR: " + message)
                writeToLog(repr(e))
                printToConsole(message)
                printToConsole(repr(e))
    except (IOError, OSError) as e:
        message = "Writing previous imported file failed for " + importedFilePath
        writeToLog("ERROR: " + message)
        writeToLog(repr(e))
        printToConsole(message)
        printToConsole(repr(e))


def readPreviousImportedFile(path):
    importedFilePath = os.path.join(path, outputPreviousImportedFileName)
    if not pathlib.Path(importedFilePath).resolve().exists():
        return None
    try:
        with open(importedFilePath, "r") as f:
            try:
                data = json.load(f)
                return data
            except (FileNotFoundError, PermissionError, OSError) as e:
                message = "Reading previous imported failed for " + importedFilePath
                writeToLog("ERROR: " + message)
                writeToLog(repr(e))
                printToConsole(message)
                printToConsole(repr(e))
    except (IOError, OSError) as e:
        message = "Opening previous imported failed for " + importedFilePath
        writeToLog("ERROR: " + message)
        writeToLog(repr(e))
        printToConsole(message)
        printToConsole(repr(e))


def initFiles(path):
    now = datetime.now()
    nowFormat = now.strftime(dateFormatter)

    global outputLogFilePath
    outputLogFilePath = os.path.join(path, nowFormat + "_" + outputLogFileName)

    try:
        open(outputLogFilePath, "x")
    except (IOError, OSError) as e:
        printToConsole("Creating log file failed for " + outputLogFilePath)
        printToConsole(repr(e))

    # global outputMetadataLogFilePath
    # outputMetadataLogFilePath = os.path.join(
    #     path, nowFormat + "_" + outputMetadataLogFileName
    # )
    # try:
    #     open(outputMetadataLogFilePath, "x")
    # except (IOError, OSError) as e:
    #     writeToLog("Creating log file failed for " + outputMetadataLogFilePath + "\n")
    #     writeToLog(repr(e) + "\n")

    global outputImportedFilePath
    outputImportedFilePath = os.path.join(
        path, nowFormat + "_" + outputImportedFileName
    )

    try:
        open(outputImportedFilePath, "x")
    except (IOError, OSError) as e:
        printToConsole("Creating log file failed for " + outputImportedFilePath)
        printToConsole(repr(e))

    # global outputMetadataFilePath
    # outputMetadataFilePath = os.path.join(
    #     path, nowFormat + "_" + outputMetadataFileName
    # )
    # try:
    #     open(outputMetadataFilePath, "x")
    # except (IOError, OSError) as e:
    #     writeToLog("Creating log file failed for " + outputMetadataFilePath + "\n")
    #     writeToLog(repr(e) + "\n")


def printToConsole(s):
    now = datetime.now()
    nowFormat = now.strftime(dateFormatter)
    print(nowFormat + " - " + s + "\n")


def writeToLog(s):
    now = datetime.now()
    nowFormat = now.strftime(dateFormatter)
    try:
        with open(outputLogFilePath, "a") as f:
            try:
                f.write(nowFormat + " : " + s)
                f.write("\n")
                f.close()
            except (FileNotFoundError, PermissionError, OSError) as e:
                printToConsole("Writing to log file failed for " + outputLogFilePath)
                printToConsole(repr(e))
    except (IOError, OSError) as e:
        printToConsole("Opening log file failed for " + outputLogFilePath)
        printToConsole(repr(e))


# def writeToMetadataLogFile(s):
#     now = datetime.now()
#     nowFormat = now.strftime(dateFormatter)
#     try:
#         with open(outputMetadataLogFilePath, "a") as f:
#             try:
#                 f.write(nowFormat + " : " + s)
#                 f.write("\n")
#                 f.close()
#             except (FileNotFoundError, PermissionError, OSError) as e:
#                 printToConsole(
#                     "Writing to log file failed for " + outputMetadataLogFilePath
#                 )
#                 printToConsole(repr(e))
#     except (IOError, OSError) as e:
#         printToConsole("Opening log file failed for " + outputMetadataLogFilePath)
#         printToConsole(repr(e))


# def writeToPreviousMetadataFile(imported):
#     now = datetime.now()
#     nowFormat = now.strftime(dateFormatter)
#     try:
#         with open(outputMetadataFilePath, "a") as f:
#             try:
#                 f.write("Files Metadata was written for:\n")
#                 for s in imported:
#                     f.write(s)
#                     f.write("\n")
#                 f.close()
#             except (FileNotFoundError, PermissionError, OSError) as e:
#                 printToConsole(
#                     "Writing to log file failed for " + outputMetadataFilePath
#                 )
#                 printToConsole(repr(e))
#     except (IOError, OSError) as e:
#         printToConsole("Opening log file failed for " + outputMetadataFilePath)
#         printToConsole(repr(e))


def sendErrorEmail(emailTo, adminsEmailTo, error, emailFrom, emailFromPSW):
    subject = "Omero Import error report"
    text = "Omero Importer job has been terminated due to the following error:\n"
    text += error + "\n\n"
    if emailTo != None and emailFrom != None and emailFromPSW != None:
        sendEmail(emailTo, subject, text, emailFrom, emailFromPSW)
    if adminsEmailTo != None and emailFrom != None and emailFromPSW != None:
        # printToConsole("emailFrom " + str(emailFrom))
        # printToConsole("emailFromPSW " + str(emailFromPSW))
        # printToConsole("emailTo " + str(emailTo))
        sendAdminEmail(adminsEmailTo, subject, text, emailFrom, emailFromPSW)


def sendCompleteEmail(
    emailTo, adminsEmailTo, hasNewImport, results, emailFrom, emailFromPSW
):
    subject = "Omero Importer job completion report"
    text = "Omero Importer job successfully complete.\n"
    if hasNewImport:
        text += "The following structure has been created:\n"
        for projectKey in results:
            projectData = results[projectKey]
            text += "Project: " + projectKey + " " + projectData[import_status]
            if import_annotate in projectData:
                if projectData[import_status] == import_status_pimported:
                    text += " , metadata updated\n"
                else:
                    text += " , metadata written\n"
            else:
                text += "\n"
            for datasetKey in projectData:
                datasetData = projectData[datasetKey]
                if not isinstance(datasetData, dict):
                    continue
                text += "Dataset: " + datasetKey + " " + datasetData[import_status]
                if import_annotate in datasetData:
                    if datasetData[import_status] == import_status_pimported:
                        text += " , metadata updated\n"
                    else:
                        text += " , metadata written\n"
                else:
                    text += "\n"
                for imageKey in datasetData:
                    imageData = datasetData[imageKey]
                    if not isinstance(imageData, dict):
                        continue
                    text += "Image: " + imageKey + " " + imageData[import_status]
                    if import_annotate in imageData:
                        if imageData[import_status] == import_status_pimported:
                            text += " , metadata updated\n"
                        else:
                            text += " , metadata written\n"
                    else:
                        text += "\n"
    else:
        text += "No new structure was created.\n"
    text += "\n"
    if hasNewImport:
        sendEmail(emailTo, subject, text, emailFrom, emailFromPSW)
        sendAdminEmail(adminsEmailTo, subject, text, emailFrom, emailFromPSW)


def sendEmail(emailTo, subject, text, emailFrom, emailFromPSW):

    body = text
    body += "\nThis is an automatic message from an unsupervised email address, please do not reply to this mail.\n"
    body += "If you need assistance contact caterina.strambio@umassmed.edu"

    message = MIMEMultipart("alternative")
    message["Subject"] = subject
    message["From"] = emailFrom
    if isinstance(emailTo, str):
        message["To"] = emailTo
    else:
        message["To"] = ", ".join(emailTo)

    message.attach(MIMEText(body, "plain"))

    email = message.as_string()

    # Log in to server using secure context and send email
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(emailFrom, emailFromPSW)
        server.sendmail(emailFrom, emailTo, email)


def sendAdminEmail(emailTo, subject, text, emailFrom, emailFromPSW):

    body = text
    body += "\nThis is an automatic message from an unsupervised email address, please do not reply to this mail.\n"
    body += "If you need assistance contact caterina.strambio@umassmed.edu"

    message = MIMEMultipart("alternative")
    message["Subject"] = subject
    message["From"] = emailFrom
    if isinstance(emailTo, str):
        message["To"] = emailTo
    else:
        message["To"] = ", ".join(emailTo)

    message.attach(MIMEText(body, "plain"))

    f1 = outputLogFilePath
    f1Path = pathlib.Path(f1)
    # Open file in binary mode
    with open(f1Path, "rb") as attachment:
        # Add file as application/octet-stream
        # Email client can usually download this automatically as attachment
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
    # Encode file in ASCII characters to send by email
    encoders.encode_base64(part)
    # Add header as key/value pair to attachment part
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {f1Path.name}",
    )
    # Add attachment to message and convert message to string
    message.attach(part)

    # f2 = outputMetadataLogFilePath
    # # Open file in binary mode
    # with open(f2, "rb") as attachment:
    #     # Add file as application/octet-stream
    #     # Email client can usually download this automatically as attachment
    #     part = MIMEBase("application", "octet-stream")
    #     part.set_payload(attachment.read())
    # # Encode file in ASCII characters to send by email
    # encoders.encode_base64(part)
    # # Add header as key/value pair to attachment part
    # part.add_header(
    #     "Content-Disposition",
    #     f"attachment; filename= {f2}",
    # )
    # # Add attachment to message and convert message to string
    # message.attach(part)

    f3 = outputImportedFilePath
    f3Path = pathlib.Path(f3)
    # Open file in binary mode
    with open(f3Path, "rb") as attachment:
        # Add file as application/octet-stream
        # Email client can usually download this automatically as attachment
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
    # Encode file in ASCII characters to send by email
    encoders.encode_base64(part)
    # Add header as key/value pair to attachment part
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {f3Path.name}",
    )
    # Add attachment to message and convert message to string
    message.attach(part)

    # f4 = outputMetadataFilePath
    # # Open file in binary mode
    # with open(f4, "rb") as attachment:
    #     # Add file as application/octet-stream
    #     # Email client can usually download this automatically as attachment
    #     part = MIMEBase("application", "octet-stream")
    #     part.set_payload(attachment.read())
    # # Encode file in ASCII characters to send by email
    # encoders.encode_base64(part)
    # # Add header as key/value pair to attachment part
    # part.add_header(
    #     "Content-Disposition",
    #     f"attachment; filename= {f4}",
    # )
    # # Add attachment to message and convert message to string
    # message.attach(part)

    email = message.as_string()

    # Log in to server using secure context and send email
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(emailFrom, emailFromPSW)
        server.sendmail(emailFrom, emailTo, email)


def readCSVFile(path):
    try:
        with open(path) as f:
            try:
                dataArray = []
                dataDict = {}
                isImageListFile = False
                keys = None
                imageIndex = 0
                while True:
                    line = f.readline()
                    if not line:
                        break
                    data = line.strip()
                    if "- PROJECT" in data or "- DATASET" in data:
                        continue
                    if "- IMAGE" in data:
                        isImageListFile = True
                        continue
                    if "Key,Value" in data:
                        continue
                    if "Image_Name" in data:
                        keys = data.split(",")
                        continue

                    if not isImageListFile:
                        dataSplit = data.split(",")
                        key = dataSplit[0]
                        value = dataSplit[1]
                        dataDict[key] = value
                    else:
                        dataSplit = data.split(",")
                        i = 0
                        dataArray.append({})
                        for key in keys:
                            value = dataSplit[i]
                            i = i + 1
                            if (
                                key == metadata_image_tags1
                                or key == metadata_image_tags2
                            ):
                                if value == "":
                                    tags = []
                                else:
                                    tags = value.split("#")
                                dataArray[imageIndex][key] = tags
                            else:
                                dataArray[imageIndex][key] = value
                        imageIndex = imageIndex + 1
                f.close()
                if isImageListFile:
                    return dataArray
                return dataDict
            except Exception as e:
                raise WrappedException("Read file failed for " + repr(path), e)
    except Exception as e:
        raise WrappedException("Open file failed for " + repr(path), e)


def collectMetadataFromCSV(path):
    data = {}
    projects = {}
    datasets = {}
    images = {}
    targetPath = pathlib.Path(path).resolve()
    for path in targetPath.iterdir():
        if path.is_dir():
            continue
        name = path.name
        if ".csv" not in name:
            continue
        nameParts = name.split("#")
        parts = len(nameParts)
        if parts == 1:
            # project
            projectName = name.replace(".csv", "")
            try:
                projectData = readCSVFile(path)
            except Exception as e:
                raise
            projects[projectName] = projectData
        elif parts == 2:
            # dataset
            projectName = nameParts[0]
            datasetName = nameParts[1].replace(".csv", "")
            try:
                datasetData = readCSVFile(path)
            except Exception as e:
                raise
            if projectName not in datasets:
                datasets[projectName] = {}
            datasets[projectName][datasetName] = datasetData
        elif parts == 3:
            # image list
            projectName = nameParts[0]
            datasetName = nameParts[1]
            try:
                imageData = readCSVFile(path)
            except Exception as e:
                raise
            if projectName not in images:
                images[projectName] = {}
            # images[projectName][datasetName] = []
            images[projectName][datasetName] = imageData

    for projectKey in projects:
        data[projectKey] = projects[projectKey]
    for projectKey in datasets:
        if metadata_plates not in data[projectKey]:
            data[projectKey][metadata_plates] = {}
        for datasetKey in datasets[projectKey]:
            data[projectKey][metadata_plates][datasetKey] = datasets[projectKey][
                datasetKey
            ]
    for projectKey in images:
        for datasetKey in images[projectKey]:
            if metadata_wells not in data[projectKey][metadata_plates][datasetKey]:
                data[projectKey][metadata_plates][datasetKey][metadata_wells] = {}
            data[projectKey][metadata_plates][datasetKey][metadata_wells] = images[
                projectKey
            ][datasetKey]
    printToConsole("Data:")
    printToConsole(str(data))
    return data


def parseImageListSpreadsheetData(ssData):
    data = []
    keys = ssData.columns
    size = len(ssData[keys[0]])
    for i in range(0, size):
        objectData = {}
        for key in keys:
            if key == excel_replaceNaN:
                continue
            value = ssData[key][i]
            if key == metadata_image_tags1 or key == metadata_image_tags2:
                if value == excel_replaceNaN:
                    value = []
                else:
                    values = value.split(",")
                    value = values
            if value == excel_replaceNaN:
                continue
            if isinstance(value, str):
                value = value.strip()
            objectData[key] = value
        if objectData != None and objectData != {}:
            data.append(objectData)
            # data[i] = objectData
    return data


def parseSpreadsheetData(ssData, objectNameKey):
    data = {}
    objectData = {}
    module = None
    objectName = None
    modules = ssData[excel_module].values

    keys = ssData[excel_key].values
    values = ssData[excel_value].values
    for i in range(0, len(keys)):
        if modules[i] != excel_replaceNaN:
            module = modules[i]
            if not module in objectData:
                objectData[module] = {}
        key = keys[i]
        if key == excel_replaceNaN:
            continue
        value = values[i]
        if value == excel_replaceNaN:
            continue
            # value = "NA"
        if isinstance(value, str):
            value = value.strip()
        if key == objectNameKey:
            objectName = value
        objectData[module][key] = value
    cleanObjectData = {k: v for k, v in objectData.items() if v != None and v != {}}
    data[objectName] = cleanObjectData
    return data


def collectMetadataFromExcel(path):
    data = {}
    excelPath = None
    targetPath = pathlib.Path(path).resolve()
    print("RESOLVED TARGET PATH: " + str(targetPath))
    for path in targetPath.iterdir():
        if path.is_dir():
            continue
        name = path.name
        if ".xlsx" in name or ".xlsm" in name:
            excelPath = path
            ssFileScreenData = pd.read_excel(
                path, sheet_name=excel_screen, header=8
            ).fillna(excel_replaceNaN)
            ssScreenData = parseSpreadsheetData(ssFileScreenData, excel_screenName)
            screenName = list(ssScreenData.keys())[0]


            if screenName not in data:
                # TODO IF PROJECT NAME CHANGED IN SAME PROJECT DIRECTORY IF OVERRIDE OVERRIDE IF NOT SEND ERROR EMAIL
                data.update(ssScreenData)

            if metadata_plates not in data[screenName]:
                data[screenName][metadata_plates] = {}
            ssFilePlateData = pd.read_excel(
                path, sheet_name=excel_plate, header=8
            ).fillna(excel_replaceNaN)
            ssPlateData = parseSpreadsheetData(ssFilePlateData, excel_plateName)
            plateName = list(ssPlateData.keys())[0]
            if plateName not in data[screenName][metadata_plates]:
                data[screenName][metadata_plates].update(ssPlateData)

            if metadata_wells not in data[screenName][metadata_plates][plateName]:
                data[screenName][metadata_plates][plateName][metadata_wells] = {}
            ssFileWellListData = pd.read_excel(
                path, sheet_name=excel_plateMap, header=12
            ).fillna(excel_replaceNaN)
            ssPlateMapData = parseImageListSpreadsheetData(ssFileWellListData)
            data[screenName][metadata_plates][plateName][
                metadata_wells
            ] = ssPlateMapData

            ssWellImageMapData = pd.read_excel(path, sheet_name=excel_WellImageMap).fillna(excel_replaceNaN)
            testData = parseImageListSpreadsheetData(ssWellImageMapData)
            
            # ADD OME_IMAGE NAMES AND FILE NAMES TO WELL DATA
            for entry in testData:
                wellName = entry.pop(metadata_well_name)
                ome_image_name = entry.pop(metadata_OME_Image_Name)
                for plate in data[screenName][metadata_plates].values():
                    for well in plate[metadata_wells]:
                        if well[metadata_well_name] == wellName:
                            if metadata_OME_Images not in well:
                                well[metadata_OME_Images] = []
                            # Check if the OME_Image already exists in the well
                            ome_image_entry = next((img for img in well[metadata_OME_Images] if img[metadata_OME_Image_Name] == ome_image_name), None)
                            if ome_image_entry is None:
                                # Create a new OME_Image entry
                                ome_image_entry = {
                                    metadata_OME_Image_Name: ome_image_name,
                                    metadata_files: []
                                }
                                well[metadata_OME_Images].append(ome_image_entry)
                            # Add the file information to the OME_Image entry
                            ome_image_entry[metadata_files].append(entry)
            
            # get number of sites in each well, and the z,c,t of the images
            site,zStep,channel,timepoint = 0,0,0,0
            for plate in data[screenName][metadata_plates].values():
                for well in plate[metadata_wells]:
                    if metadata_OME_Images in well:
                        for ome_image_entry in well[metadata_OME_Images]:
                            for image in ome_image_entry[metadata_files]:
                                site = max(site, image[metadata_site])
                                channel = max(channel, image[metadata_C_index])
                                zStep = max(zStep, image[metadata_Z_index])
                                timepoint = max(timepoint, image[metadata_T_index]) 

    #printToConsole("Data:")
    #printToConsole(str(data))
    return data, excelPath, site, zStep, channel, timepoint

# This function returns the data from an excel file, starting at a specific cell and reading onwards.
def readExcelTable(file):
    wb = xlwings.Book(file)

    sheet = wb.sheets[excel_plateMap]
    screenSheet = wb.sheets[excel_screen]

    #Define the range starting from the startCell until the end of the sheet
    data = sheet.range(excel_plateMapTableStartCell).expand().value
    screenData = screenSheet[excel_datasetCell].value.strip()
    
    wb.close()

    return data, screenData



def createWell(conn,plateId, row, col):
    well = WellI()
    well.setPlate(PlateI(plateId, False))
    well.setColumn(rint(col))
    well.setRow(rint(row))

    conn.getUpdateService().saveObject(well)


# This function converts wellId's to row-column format ex. E04 -> 5-4
def getWellCoords(input):
    col = int(input[1:])
    row = ord(input[0]) - ord('A') + 1
    return row-1,col-1

def main(argv, argc):
    if len(argv) > 1 and argv[1] == "-h":
        print("Help for Omero Importer CL")
        print("-cfg <options>, to create a global config file")
        print("options (* required):")
        print("*-H <hostname>")
        print("-p <port>, default is 4064")
        print("*-u <admin userName>")
        print("*-psw <admin password>")
        print("*-t <target>, target directory to launch the importer")
        print(
            "-d <destination>, destination directory where to move files after import (in this case if not specified copy does not happen)"
        )
        print("-del, to delete files after import and copy, default is false")
        print("-mma, to add microscope and acquisition settings file, default is false")
        print(
            "-b2 <endpoint#bucketName#appKeyId#appKey>, to use backblaze as destination for copy (conflict with -d)"
        )
        print(
            "-ts <hh:mm>, to specify the daily start time of the application, default is non-stop"
        )
        print(
            "-te <hh:mm>, to specify the time limit after which the application should auto terminate, default is non-stop"
        )
        print("*-sml <email address> to set up automatic email sender")
        print("*-smlp <password> to set up automatic email sender password")
        print(
            "*-aml <email address1#email address2:...> to set up automatic email to admin upon error or completion"
        )
        print("#####")
        print(
            "-ucfg <userDirectory> <options> to create a user config file in a specific directory, conflicting user options override global options"
        )
        print("options (* required):")
        print("*-u <user userName>")
        print("*-psw <user password>")
        print(
            "-d <destination>, destination directory where to move files after import (in this case if not specified copy does not happen)"
        )
        print("-del, to delete files after import and copy")
        print(
            "-b2 <endpoint#bucketName#appKeyId#appKey>, to use backblaze as destination for copy (conflict with -d)"
        )
        print("-mma, to add microscope and acquisition settings file")
        print(
            "*-ml <email address1:email address2:...> to set up automatic email upon error or completion"
        )
        quit()

    localPath = None
    try:
        localPath = pathlib.Path(__file__).parent.resolve(strict=True)
    except FileNotFoundError:
        localPath = pathlib.Path().resolve()
    initFiles(localPath)
    printToConsole("LOG FILE INIT")

    isCfg = False
    isUCfg = False

    # Both param
    destination_g = None
    hasDelete_g = False
    hasMMA_g = False
    hasB2_g = False
    b2Endpoint_g = None
    b2BucketName_g = None
    b2AppKeyId_g = None
    b2AppKey_g = None

    # Global param
    hostName = None
    port = 4064
    target = None
    startTimeHr = None
    startTimeMin = None
    endTimeHr = None
    endTimeMin = None
    adminsEmailTo = None
    emailFrom = None
    emailFromPSW = None

    # Well and Image params
    site = 0
    zStep = 0
    channel = 0
    timepoint = 0

    # User param
    userDirectoryPath = None
    userName_g = None
    userPSW_g = None
    emailTo = None
    isAdmin = False

    # Path to the Excel file
    excelFile = None

    for i in range(1, argc):
        arg = argv[i]
        if arg == "-cfg":
            isCfg = True
        elif arg == "-ucfg":
            isUCfg = True
            userDirectory = argv[i + 1]
            if userDirectory == None:
                error = "user directory cannot be undefined with the -ucfg option, application terminated."
                writeToLog("ERROR: " + error)
                printToConsole("ERROR: " + error)
                quit()
            try:
                userDirectoryPath = pathlib.Path(userDirectory).resolve()
                if not userDirectoryPath.exists():
                    # if not os.path.exists(tmpTarget):
                    error = (
                        "User directory "
                        + userDirectory
                        + " doesn't exists, application terminated."
                    )
                    writeToLog("ERROR: " + error)
                    printToConsole("ERROR: " + error)
                    quit()
                if not userDirectoryPath.is_dir():
                    # if not os.path.isdir(tmpTarget):
                    error = (
                        "User directory "
                        + userDirectory
                        + " is not a directory, application terminated."
                    )
                    writeToLog("ERROR: " + error)
                    printToConsole("ERROR: " + error)
                    quit()
                # target = tmpTarget
            except IOError as e:
                error = (
                    "Something went wrong trying to determine if user directory "
                    + userDirectory
                    + " exists and is a directory, application terminated."
                )
                writeToLog("ERROR: " + error)
                printToConsole("ERROR: " + error)
                quit()
        elif arg == "-H":
            hostName = argv[i + 1]
        elif arg == "-p":
            port = argv[i + 1]
        elif arg == "-u":
            userName_g = argv[i + 1]
        elif arg == "-psw":
            userPSW_g = argv[i + 1]
        elif arg == "-t":
            target = argv[i + 1]
        elif arg == "-d":
            destination_g = argv[i + 1]
        elif arg == "-del":
            hasDelete_g = True
        elif arg == "-mma":
            hasMMA_g = True
        elif arg == "-b2":
            hasB2_g = True
            b2Data = argv[i + 1]
            b2DataSplit = b2Data.split("#")
            if (len(b2DataSplit) < 4) or (len(b2DataSplit) > 4):
                error = (
                    "wrong number of arguments in -b2 option, application terminated."
                )
                writeToLog("ERROR: " + error)
                printToConsole("ERROR: " + error)
                quit()
            b2Endpoint_g = b2DataSplit[0]
            b2BucketName_g = b2DataSplit[1]
            b2AppKeyId_g = b2DataSplit[2]
            b2AppKey_g = b2DataSplit[3]
        elif arg == "-ml":
            mlData = argv[i + 1]
            mlDataSplit = mlData.split(":")
            if len(mlDataSplit) > 2:
                emailTo = mlDataSplit
            emailTo = mlData
        elif arg == "-aml":
            amlData = argv[i + 1]
            amlDataSplit = amlData.split(":")
            if len(amlDataSplit) > 2:
                adminsEmailTo = amlDataSplit
            adminsEmailTo = amlData
        elif arg == "-sml":
            emailFrom = argv[i + 1]
        elif arg == "-smlp":
            emailFromPSW = argv[i + 1]
        elif arg == "-ts":
            teData = argv[i + 1]
            teDataSplit = teData.split(":")
            if (len(teDataSplit) < 2) or (len(teDataSplit) > 2):
                error = (
                    "wrong number of arguments in -ts option, application terminated."
                )
                writeToLog("ERROR: " + error)
                printToConsole("ERROR: " + error)
                quit()
            startTimeHr = teDataSplit[0]
            startTimeMin = teDataSplit[1]
        elif arg == "-te":
            teData = argv[i + 1]
            teDataSplit = teData.split(":")
            if (len(teDataSplit) < 2) or (len(teDataSplit) > 2):
                error = (
                    "wrong number of arguments in -te option, application terminated."
                )
                writeToLog("ERROR: " + error)
                printToConsole("ERROR: " + error)
                quit()
            endTimeHr = teDataSplit[0]
            endTimeMin = teDataSplit[1]
        else:
            if not arg.startswith("-"):
                continue
            printToConsole(
                "Option "
                + arg
                + " not recognized, please use -h to review available options, application terminated."
            )
            quit()

    if isCfg:
        dict = {}
        key = Fernet.generate_key()
        f = Fernet(key)
        dict[p_key] = key.decode()
        dict[p_omeroHostname] = hostName
        dict[p_omeroPort] = port
        dict[p_target] = target
        dict[p_omeroUsername] = f.encrypt(bytes(userName_g, "utf-8")).decode()
        dict[p_omeroPSW] = f.encrypt(bytes(userPSW_g, "utf-8")).decode()
        if destination_g != None:
            dict[p_dest] = destination_g
        if hasDelete_g:
            dict[p_delete] = hasDelete_g
        if hasMMA_g:
            dict[p_mma] = hasMMA_g
        if hasB2_g:
            dict[p_b2] = hasB2_g
            dict[p_b2_endpoint] = f.encrypt(bytes(b2Endpoint_g, "utf8")).decode()
            dict[p_b2_bucketName] = f.encrypt(bytes(b2BucketName_g, "utf8")).decode()
            dict[p_b2_appKeyId] = f.encrypt(bytes(b2AppKeyId_g, "utf8")).decode()
            dict[p_b2_appKey] = f.encrypt(bytes(b2AppKey_g, "utf8")).decode()
        # if startTime != None:
        #     dict[p_startTime] = startTime
        if startTimeHr != None and startTimeMin != None:
            dict[p_startTime] = str(startTimeHr) + ":" + str(startTimeMin)
        if endTimeHr != None and endTimeMin != None:
            dict[p_endTime] = str(endTimeHr) + ":" + str(endTimeMin)
        dict[p_adminsEmail] = adminsEmailTo
        dict[p_emailFrom] = f.encrypt(bytes(emailFrom, "utf8")).decode()
        dict[p_emailFromPSW] = f.encrypt(bytes(emailFromPSW, "utf8")).decode()
        writeConfigFile(localPath, dict)
        message = "Global configuration file generated"
        writeToLog(message)
        printToConsole(message)
        quit()
    elif isUCfg:
        dict = {}
        key = Fernet.generate_key()
        f = Fernet(key)
        dict[p_key] = key.decode()
        dict[p_omeroUsername] = f.encrypt(bytes(userName_g, "utf-8")).decode()
        dict[p_omeroPSW] = f.encrypt(bytes(userPSW_g, "utf-8")).decode()
        if destination_g != None:
            dict[p_dest] = destination_g
        if hasDelete_g:
            dict[p_delete] = hasDelete_g
        if hasMMA_g:
            dict[p_mma] = hasMMA_g
        if hasB2_g:
            dict[p_b2] = hasB2_g
            dict[p_b2_endpoint] = f.encrypt(bytes(b2Endpoint_g, "utf8")).decode()
            dict[p_b2_bucketName] = f.encrypt(bytes(b2BucketName_g, "utf8")).decode()
            dict[p_b2_appKeyId] = f.encrypt(bytes(b2AppKeyId_g, "utf8")).decode()
            dict[p_b2_appKey] = f.encrypt(bytes(b2AppKey_g, "utf8")).decode()
        dict[p_userEmail] = f.encrypt(bytes(emailTo, "utf8")).decode()
        writeConfigFile(userDirectoryPath, dict)
        message = "User configuration file generated in " + userDirectory
        writeToLog(message)
        printToConsole(message)
        quit()

    # Read global parameters
    parameters = readConfigFile(localPath)
    if parameters == None:
        error = (
            "Reading global parameters from config file failed, application terminated."
        )
        writeToLog("ERROR: " + error)
        printToConsole("ERROR: " + error)
        quit()
    eKey = parameters[p_key]
    if eKey == None:
        error = "Reading encryption key for global parameters failed, application terminated."
        writeToLog("ERROR: " + error)
        printToConsole("ERROR: " + error)
        quit()
    f = Fernet(eKey)
    for key in parameters:
        if key.startswith("#"):
            continue
        value = parameters[key]
        if key == p_omeroUsername:
            userName_g = str(f.decrypt(value).decode())
            isAdmin = True
        if key == p_omeroPSW:
            userPSW_g = str(f.decrypt(value).decode())
        if key == p_omeroHostname:
            hostName = value
        if key == p_omeroPort:
            port = value
        if key == p_target:
            target = value
        if key == p_dest:
            destination_g = value
        if key == p_delete:
            hasDelete_g = value
        if key == p_mma:
            hasMMA_g = value
        if key == p_b2:
            hasB2_g = value
        if key == p_b2_endpoint:
            b2Endpoint_g = str(f.decrypt(value).decode())
        if key == p_b2_bucketName:
            b2BucketName_g = str(f.decrypt(value).decode())
        if key == p_b2_appKeyId:
            b2AppKeyId_g = str(f.decrypt(value).decode())
        if key == p_b2_appKey:
            b2AppKey_g = str(f.decrypt(value).decode())
        if key == p_adminsEmail:
            adminsEmailTo = value
        if key == p_emailFrom:
            emailFrom = str(f.decrypt(value).decode())
        if key == p_emailFromPSW:
            emailFromPSW = str(f.decrypt(value).decode())
        if key == p_startTime:
            vals = value.split(":")
            startTimeHr = vals[0]
            startTimeMin = vals[1]
        if key == p_endTime:
            vals = value.split(":")
            endTimeHr = vals[0]
            endTimeMin = vals[1]
    printToConsole("GLOBAL CONFIG READ")

    if emailFrom == None:
        error = "Automatic email sender must be set, application terminated."
        writeToLog("ERROR: " + error)
        printToConsole("ERROR: " + error)
        quit()
    if emailFromPSW == None:
        error = "Automatic email sender password must be set, application terminated."
        writeToLog("ERROR: " + error)
        printToConsole("ERROR: " + error)
        quit()

    if hostName == None:
        error = "Hostname must be set, application terminated."
        writeToLog("ERROR: " + error)
        printToConsole("ERROR: " + error)
        sendErrorEmail(emailTo, adminsEmailTo, error, emailFrom, emailFromPSW)
        quit()

    portI = None
    try:
        port = int(port)
        if int(port) == port:
            portI = int(port)
    except TypeError as e:
        error = "Port is not a valid number, application terminated."
        writeToLog("ERROR: " + error)
        writeToLog(repr(e))
        printToConsole("ERROR: " + error)
        printToConsole(repr(e))
        sendErrorEmail(emailTo, adminsEmailTo, error + repr(e), emailFrom, emailFromPSW)
        quit()

    if hasB2_g and (
        (b2AppKeyId_g == None) or (b2AppKey_g == None) or (b2BucketName_g == None)
    ):
        error = "Some bucket information for backblaze backup not been specified, application terminated."
        writeToLog("ERROR: " + error)
        printToConsole("ERROR: " + error)
        sendErrorEmail(emailTo, adminsEmailTo, error, emailFrom, emailFromPSW)
        quit()

    # if (endTimeHr == None) or (endTimeMin == None):
    #     error = "Hour or minute have not been specified."
    #     writeToLog("ERROR: " + error)
    #     printToConsole("ERROR: " + error)
    #     sendErrorEmail(emailTo, adminsEmailTo, error, emailFrom, emailFromPSW)
    #     quit()

    startTimeHrI = None
    startTimeMinI = None
    if startTimeHr != None:
        try:
            if int(startTimeHr) == startTimeHr:
                startTimeHrI = int(startTimeHr)
        except TypeError as e:
            error = "Hour value for start time is not a valid number, application terminated."
            writeToLog("ERROR: " + error)
            writeToLog(repr(e))
            printToConsole("ERROR: " + error)
            printToConsole(repr(e))
            sendErrorEmail(
                emailTo, adminsEmailTo, error + repr(e), emailFrom, emailFromPSW
            )
            quit()

    if startTimeMin != None:
        try:
            if int(startTimeMin) == startTimeMin:
                startTimeMinI = int(startTimeMin)
        except TypeError as e:
            error = "Minute value for start time is not a valid number, application terminated."
            writeToLog("ERROR: " + error)
            writeToLog(repr(e))
            printToConsole("ERROR: " + error)
            printToConsole(repr(e))
            sendErrorEmail(
                emailTo, adminsEmailTo, error + repr(e), emailFrom, emailFromPSW
            )
            quit()

    endTimeHrI = None
    endTimeMinI = None
    if endTimeHr != None:
        try:
            if int(endTimeHr) == endTimeHr:
                endTimeHrI = int(endTimeHr)
        except TypeError as e:
            error = (
                "Hour value for end time is not a valid number, application terminated."
            )
            writeToLog("ERROR: " + error)
            writeToLog(repr(e))
            printToConsole("ERROR: " + error)
            printToConsole(repr(e))
            sendErrorEmail(
                emailTo, adminsEmailTo, error + repr(e), emailFrom, emailFromPSW
            )
            quit()

    if endTimeMin != None:
        try:
            if int(endTimeMin) == endTimeMin:
                endTimeMinI = int(endTimeMin)
        except TypeError as e:
            error = "Minute value for end time is not a valid number, application terminated."
            writeToLog("ERROR: " + error)
            writeToLog(repr(e))
            printToConsole("ERROR: " + error)
            printToConsole(repr(e))
            sendErrorEmail(
                emailTo, adminsEmailTo, error + repr(e), emailFrom, emailFromPSW
            )
            quit()

    if target == None:
        error = "Target directory has not been specified, application terminated."
        writeToLog("ERROR: " + error)
        printToConsole("ERROR: " + error)
        sendErrorEmail(emailTo, adminsEmailTo, error, emailFrom, emailFromPSW)
        quit()
    try:
        targetPath = pathlib.Path(target).resolve()
        if not targetPath.exists():
            # if not os.path.exists(tmpTarget):
            error = "Target directory doesn't exists, application terminated."
            writeToLog("ERROR: " + error)
            printToConsole("ERROR: " + error)
            sendErrorEmail(emailTo, adminsEmailTo, error, emailFrom, emailFromPSW)
            quit()
        if not targetPath.is_dir():
            # if not os.path.isdir(tmpTarget):
            error = "Target directory is not a directory, application terminated."
            writeToLog("ERROR: " + error)
            printToConsole("ERROR: " + error)
            sendErrorEmail(emailTo, adminsEmailTo, error, emailFrom, emailFromPSW)
            quit()
        # target = tmpTarget
    except IOError as e:
        error = "Exception trying to determine if target directory exists and is a directory, application terminated."
        writeToLog("ERROR: " + error)
        writeToLog(repr(e))
        printToConsole("ERROR: " + error)
        printToConsole(repr(e))
        sendErrorEmail(emailTo, adminsEmailTo, error + repr(e), emailFrom, emailFromPSW)
        quit()

    if destination_g != None:
        try:
            destPath = pathlib.Path(destination_g).resolve()
            if not destPath.exists():
                error = "Destination directory doesn't exists, application terminated."
                writeToLog("ERROR: " + error)
                printToConsole("ERROR: " + error)
                sendErrorEmail(emailTo, adminsEmailTo, error, emailFrom, emailFromPSW)
                quit()
            if not destPath.is_dir():
                error = (
                    "Destination directory is not a directory, application terminated."
                )
                writeToLog("ERROR: " + error)
                printToConsole("ERROR: " + error)
                sendErrorEmail(emailTo, adminsEmailTo, error, emailFrom, emailFromPSW)
                quit()
            # destination = tmpDest
        except IOError as e:
            error = "Exception trying to determine if destination directory exists and is a directory, application terminated."
            writeToLog("ERROR: " + error)
            writeToLog(repr(e))
            printToConsole("ERROR: " + error)
            printToConsole(repr(e))
            sendErrorEmail(
                emailTo, adminsEmailTo, error + repr(e), emailFrom, emailFromPSW
            )
            quit()

    printToConsole("GLOBAL PARAMETERS CONFIG INIT")
    printToConsole(str(parameters))
    endTimePassed = False

    fullImportedData = readPreviousImportedFile(localPath)
    currentImportedData = {}
    targetPath = pathlib.Path(target).resolve()
    for userPath in targetPath.iterdir():
        if userPath.is_file():
            continue
        userFolder = userPath.name
        currentImportedData[userFolder] = {}
        userCurrentImportedData = currentImportedData[userFolder]
        userFullImportedData = None
        if fullImportedData != None and userFolder in fullImportedData:
            userFullImportedData = fullImportedData[userFolder]
        userName = None
        userPSW = None
        emailTo = None
        destination = None
        hasDelete = None
        hasMMA = None
        hasB2 = None
        b2Endpoint = None
        b2BucketName = None
        b2AppKeyId = None
        b2AppKey = None
        uParameters = readConfigFile(userPath)
        # Read user parameters
        # if uParameters == None:
        #     error = (
        #         "Reading user parameters from config file failed, user folder "
        #         + userPath
        #         + " skipped."
        #     )
        #     writeToLog("ERROR: " + error)
        #     printToConsole("ERROR: " + error)
        #     sendErrorEmail(emailTo, adminsEmailTo, error, emailFrom, emailFromPSW)
        #     continue
        if uParameters != None and uParameters != {}:
            eKey = uParameters[p_key]
            if eKey == None:
                error = (
                    "Reading encryption key for user parameters failed, user folder "
                    + userPath
                    + " skipped."
                )
                writeToLog("ERROR: " + error)
                printToConsole("ERROR: " + error)
                sendErrorEmail(emailTo, adminsEmailTo, error, emailFrom, emailFromPSW)
                continue

            f = Fernet(eKey)
            for key in uParameters:
                if key.startswith("#"):
                    continue
                value = uParameters[key]
                if key == p_omeroUsername:
                    userName = str(f.decrypt(value).decode())
                if key == p_omeroPSW:
                    userPSW = str(f.decrypt(value).decode())
                if key == p_userEmail:
                    emailTo = str(f.decrypt(value).decode())
                if key == p_dest:
                    destination = value
                if key == p_delete:
                    hasDelete = value
                if key == p_mma:
                    hasMMA = value
                if key == p_b2:
                    hasB2 = value
                if key == p_b2_endpoint:
                    b2Endpoint = str(f.decrypt(value).decode())
                if key == p_b2_bucketName:
                    b2BucketName = str(f.decrypt(value).decode())
                if key == p_b2_appKeyId:
                    b2AppKeyId = str(f.decrypt(value).decode())
                if key == p_b2_appKey:
                    b2AppKey = str(f.decrypt(value).decode())

        if userName == None:
            userName = userName_g
        if userName == None:
            error = (
                "No Omero admin or user Username has not been specified, user folder "
                + str(userPath)
                + " skipped."
            )
            writeToLog("ERROR: " + error)
            printToConsole("ERROR: " + error)
            sendErrorEmail(emailTo, adminsEmailTo, error, emailFrom, emailFromPSW)
            continue

        if userPSW == None:
            userPSW = userPSW_g
        if userPSW == None:
            error = (
                "No Omero admin or user Password has not been specified, user folder "
                + str(userPath)
                + " skipped."
            )
            writeToLog("ERROR: " + error)
            printToConsole("ERROR: " + error)
            sendErrorEmail(emailTo, adminsEmailTo, error, emailFrom, emailFromPSW)
            continue

        # if emailTo == None:
        #     error = (
        #         "User email has not been specified, user folder "
        #         + str(userPath)
        #         + " skipped."
        #     )
        #     writeToLog("ERROR: " + error)
        #     printToConsole("ERROR: " + error)
        #     sendErrorEmail(emailTo, adminsEmailTo, error, emailFrom, emailFromPSW)
        #     continue

        if destination != None:
            try:
                destPath = pathlib.Path(destination).resolve()
                if not destPath.exists():
                    error = (
                        "User destination directory doesn't exists, user folder "
                        + str(userPath)
                        + " skipped."
                    )
                    writeToLog("ERROR: " + error)
                    printToConsole("ERROR: " + error)
                    sendErrorEmail(
                        emailTo, adminsEmailTo, error, emailFrom, emailFromPSW
                    )
                    continue
                if not destPath.is_dir():
                    error = (
                        "User destination directory is not a directory, user folder "
                        + str(userPath)
                        + " skipped."
                    )
                    writeToLog("ERROR: " + error)
                    printToConsole("ERROR: " + error)
                    sendErrorEmail(
                        emailTo, adminsEmailTo, error, emailFrom, emailFromPSW
                    )
                    continue
                # destination = tmpDest
            except IOError as e:
                error = (
                    "Exception trying to determine if user destination directory exists and is a directory, user folder "
                    + str(userPath)
                    + " skipped."
                )
                writeToLog("ERROR: " + error)
                writeToLog(repr(e))
                printToConsole("ERROR: " + error)
                printToConsole(repr(e))
                sendErrorEmail(
                    emailTo, adminsEmailTo, error + repr(e), emailFrom, emailFromPSW
                )
                continue
        else:
            destination = destination_g

        if hasDelete == None:
            hasDelete = hasDelete_g

        if hasMMA == None:
            hasMMA = hasMMA_g

        if hasB2 != None:
            if hasB2 and (
                (b2AppKeyId == None) or (b2AppKey == None) or (b2BucketName == None)
            ):
                error = (
                    "Some user bucket information for backblaze backup have not been specified, user folder "
                    + str(userPath)
                    + " skipped."
                )
                writeToLog("ERROR: " + error)
                printToConsole("ERROR: " + error)
                sendErrorEmail(emailTo, adminsEmailTo, error, emailFrom, emailFromPSW)
                quit()
        else:
            hasB2 = hasB2_g
            b2Endpoint = b2Endpoint_g
            b2BucketName = b2BucketName_g
            b2AppKeyId = b2AppKeyId_g
            b2AppKey = b2AppKey_g

        printToConsole("USER PARAMETERS CONFIG INIT")
        printToConsole(str(uParameters))

        # Call function to return reference to B2 service
        b2 = None
        if hasB2:
            b2 = get_b2_resource(b2Endpoint, b2AppKeyId, b2AppKey)
        # Call function to return reference to B2 service
        # b2_client = get_b2_client(b2Endpoint, b2AppKeyId, b2AppKey)

        # conn = BlitzGateway(
        #     userName, userPSW, host=hostName, port=portI, secure=True
        # )
        # conn.connect()
        conn = ezome.connect(
            host=hostName,
            port=portI,
            user=userName,
            password=userPSW,
            group="",
            secure=True,
        )

        omeConnUser = conn.getUser()
        omeConnUserName = omeConnUser.getName()
        if conn == None:
            error = "Connection error"
            writeToLog("ERROR: " + error)
            printToConsole("ERROR: " + error)
            sendErrorEmail(emailTo, adminsEmailTo, error, emailFrom, emailFromPSW)
            quit()
        printToConsole("Connected")
        printToConsole(str(conn))
        conn.c.enableKeepAlive(60)
        # conn.close(True)
        # quit()

        omeUserName = None
        userConn = None
        # userConn = None
        if omeConnUserName.lower() != userFolder.lower() and isAdmin:
            if conn.isFullAdmin():
                omeUser = conn.getObject(
                    "Experimenter", attributes={"omeName": userFolder}
                )
                omeUserName = omeUser.getName()
                omeUserPSW = omeUser.getLdap()
                quit()
                # userConn = conn.suConn(omeUser.getName())
                if emailTo == None:
                    emailTo = omeUser.getEmail()
                # message = "Admin switched to user " + omeUser.getName()
                # writeToLog(message)
                # printToConsole(message)
            else:
                error = (
                    "Cannot find user "
                    + userFolder
                    + ", current connection doesn't have proper admin rights."
                )
                printToConsole(error)
                writeToLog(error)
                continue  

        # find group_id using group name ?
        # userConn.SERVICE_OPTS.setOmeroGroup(group_id)
        # session = userConn.getSession()
        # Explore User Projects
        hasNewImport = False
        for projectPath in userPath.iterdir():
            if projectPath.is_file():
                continue
            # projectCFolder = os.path.join(userFolder, projectPath.name)
            if omeUserName != None:
                userConn = conn.suConn(omeUserName)
            else:
                userConn = conn
            data = None
            namespace = omero.constants.metadata.NSCLIENTMAPANNOTATION
            try:
                data,excelFile,site,zStep,channel,timepoint = collectMetadataFromExcel(projectPath)
            except WrappedException as e:
                error = e.message
                writeToLog(error)
                writeToLog(repr(e.exception))
                sendErrorEmail(
                    emailTo,
                    adminsEmailTo,
                    error + "\n" + repr(e.exception),
                    emailFrom,
                    emailFromPSW,
                )
            except Exception as e:
                writeToLog(repr(e))
                sendErrorEmail(
                    emailTo,
                    adminsEmailTo,
                    repr(e),
                    emailFrom,
                    emailFromPSW,
                )
            if data == None or data == {}:
                continue

            for screenKey in data:
                screen = data[screenKey]
                screenFullImportedData = None
                screenCurrentImportedData = None
                if userFullImportedData != None and screenKey in userFullImportedData:
                    screenFullImportedData = userFullImportedData[screenKey]
                if screenKey not in userCurrentImportedData:
                    currentImportedData[userFolder][screenKey] = {}
                screenCurrentImportedData = currentImportedData[userFolder][screenKey]
                screenID = None
                omeProject = None

                screenQName = os.path.join(userFolder, screenKey)
                screenCurrentImportedData[import_path] = screenQName

                # continue from here
                if screenFullImportedData == None:
                    omeScreen = userConn.getObject(
                        "Screen", attributes={"name": screenKey}
                    )
                    if omeScreen == None:
                        newScreen = ScreenWrapper(userConn, ScreenI())
                        newScreen.setName(screenKey)
                        newScreen.save()
                        omeScreen = newScreen
                        screenCurrentImportedData[import_status] = (
                            import_status_imported
                        )
                        screenID = newScreen._obj.id.val
                        writeToLog(
                            "Screen created for "
                            + screenQName
                            + " ("
                            + str(screenID)
                            + ")"
                        )
                        hasNewImport = True
                    else:
                        ID = omeScreen._obj.id.val
                        screenCurrentImportedData[import_status] = import_status_found
                        writeToLog(
                            "Screen found for "
                            + screenQName
                            + " ("
                            + str(screenID)
                            + ")"
                        )
                else:
                    screenID = screenFullImportedData[import_status_id]
                    omeScreen = userConn.getObject("Screen", screenID)
                    screenCurrentImportedData[import_status] = import_status_pimported
                    writeToLog(
                        "Screen previously imported for "
                        + screenQName
                        + " ("
                        + str(screenID)
                        + ")"
                    )
                screenCurrentImportedData[import_status_id] = screenID

                # TODO should we always update the annotation? or only if not previously imported?
                # ATM only if not previously imported

                screenKeyValueData = []

                for moduleKey in screen:
                    if moduleKey == metadata_plates or moduleKey == excel_module_ome:
                        continue
                    screenKeyValueData.append([moduleKey, ""])
                    for screenAnnKey in screen[moduleKey]:
                        value = screen[moduleKey][screenAnnKey]
                        dataSplit = None
                        if "description" not in screenAnnKey.lower() and isinstance(
                            value, str
                        ):
                            dataSplit = value.split(",")
                        if dataSplit != None and len(dataSplit) > 1:
                            for i in range(0, len(dataSplit)):
                                screenKeyValueData.append(
                                    [screenAnnKey + "_" + str(i), str(dataSplit[i])]
                                )
                        else:
                            screenKeyValueData.append([screenAnnKey, str(value)])
                if (
                    screenFullImportedData == None
                    or import_annotate not in screenFullImportedData
                    # or projectFullImportedData[import_annotate] == False
                ):
                    if len(screenKeyValueData) > 0:
                        newScreenMapAnn = MapAnnotationWrapper(userConn)
                        newScreenMapAnn.setNs(namespace)
                        newScreenMapAnn.setValue(screenKeyValueData)
                        # newProjMapAnn.setNs(moduleKey)
                        # newProjMapAnn.setValue(projectKeyValueData[moduleKey])
                        newScreenMapAnn.save()
                        omeScreen.linkAnnotation(newScreenMapAnn)
                        screenCurrentImportedData[import_annotate] = (
                            newScreenMapAnn._obj.id.val
                        )

                        writeToLog(
                            "Annotation created for "
                            + screenQName
                            + " ("
                            + str(screenID)
                            + ")"
                        )

                if omeUserName != None:
                    userConn.close()

                
                for plateKey in screen[metadata_plates]:
                    plate = screen[metadata_plates][plateKey]

                    if omeUserName != None:
                        userConn = conn.suConn(omeUserName)
                        # userConn.c.enableKeepAlive(60)
                    else:
                        userConn = conn

                    plateFullImportedData = None
                    plateCurrentImportedData = None
                    if (
                        screenFullImportedData != None
                        and plateKey in screenFullImportedData
                    ):
                        plateFullImportedData = screenFullImportedData[plateKey]
                    if plateKey not in screenCurrentImportedData:
                        screenCurrentImportedData[plateKey] = {}
                    plateCurrentImportedData = screenCurrentImportedData[plateKey]
                    plateID = None
                    omePlate = None
                    plateQName = os.path.join(screenQName, plateKey)
                    plateCurrentImportedData[import_path] = plateQName
                    if plateFullImportedData == None:
                        plIDs = ezome.get_plate_ids(userConn, screen=screenID)
                        omePlate = None
                        for plID in plIDs:
                            omePL = userConn.getObject("Plate", plID)
                            if omePL.getName() == plateKey:
                                omePlate = omePL

                        if omePlate == None:
                            newPlate = PlateWrapper(userConn, PlateI())
                            newPlate.setName(plateKey)
                            newPlate.save()
                            omePlate = newPlate
                            plateID = newPlate._obj.id.val
                            link = ScreenPlateLinkI()
                            link.setChild(PlateI(plateID, False))
                            link.setParent(ScreenI(screenID, False))
                            userConn.getUpdateService().saveObject(link)
                            plateCurrentImportedData[import_status] = (
                                import_status_imported
                            )
                            writeToLog(
                                "Plate created for "
                                + plateQName
                                + " ("
                                + str(plateID)
                                + ")"
                            )
                            hasNewImport = True
                        else:
                            plateID = omePlate._obj.id.val
                            plateCurrentImportedData[import_status] = (
                                import_status_found
                            )
                            writeToLog(
                                "Plate found for "
                                + plateQName
                                + " ("
                                + str(plateID)
                                + ")"
                            )
                    else:
                        plateID = plateFullImportedData[import_status_id]
                        omePlate = userConn.getObject("Plate", plateID)
                        plateCurrentImportedData[import_status] = (
                            import_status_pimported
                        )
                        writeToLog(
                            "Plate previously imported for "
                            + plateQName
                            + " ("
                            + str(plateID)
                            + ")"
                        )

                    plateCurrentImportedData[import_status_id] = plateID

                    plateKeyValueData = []

                    for moduleKey in plate:
                        if (
                            moduleKey == metadata_wells
                            or moduleKey == excel_module_ome
                        ):
                            continue
                        plateKeyValueData.append([moduleKey, ""])
                        for plAnnKey in plate[moduleKey]:
                            value = plate[moduleKey][plAnnKey]
                            dataSplit = None
                            if "description" not in plAnnKey.lower() and isinstance(
                                value, str
                            ):
                                dataSplit = value.split(",")
                            if dataSplit != None and len(dataSplit) > 1:
                                for i in range(0, len(dataSplit)):
                                    plateKeyValueData.append(
                                        [plAnnKey + "_" + str(i), str(dataSplit[i])]
                                    )
                            else:
                                plateKeyValueData.append([plAnnKey, str(value)])
                    if (
                        plateFullImportedData == None
                        or import_annotate not in plateFullImportedData
                        # or datasetFullImportedData[import_annotate] == False
                    ):
                        if len(plateKeyValueData) > 0:
                            newPlMapAnn = MapAnnotationWrapper(userConn)
                            newPlMapAnn.setNs(namespace)
                            newPlMapAnn.setValue(plateKeyValueData)
                            newPlMapAnn.save()
                            omePlate.linkAnnotation(newPlMapAnn)
                            plateCurrentImportedData[import_annotate] = (
                                newPlMapAnn._obj.id.val
                            )
                            writeToLog(
                                "Annotation created for "
                                + plateQName
                                + " ("
                                + str(plateID)
                                + ")"
                            )
                            hasNewImport = True
                        else:
                            if len(plateKeyValueData) > 0:
                                plMapAnnID = plateFullImportedData[import_annotate]
                                newPlMapAnn = userConn.getObject(
                                    "MapAnnotation", plMapAnnID
                                )
                                newPlMapAnn.setValue(plateKeyValueData)
                            writeToLog(
                                "Annotation updated for "
                                + plateQName
                                + " ("
                                + str(plateID)
                                + ")"
                            )
                            hasNewImport = True

                    if omeUserName != None:
                        userConn.close()
















                    # **for each row in the Plate-Map
                    for well in plate[metadata_wells]:
                        wellName = well[metadata_well_name]

                        if omeUserName != None:
                            userConn = conn.suConn(omeUserName)
                            # userConn.c.enableKeepAlive(60)
                        else:
                            userConn = conn
                        
                        wellFullImportedData = None
                        wellCurrentImportedData = None
                        if (
                            plateFullImportedData != None
                            and wellName in plateFullImportedData
                        ):
                            wellFullImportedData = plateFullImportedData[wellName]

                        if wellName not in plateCurrentImportedData:
                            plateCurrentImportedData[wellName] = {}
                        wellCurrentImportedData = plateCurrentImportedData[wellName]

                        wellID = None
                        omeWell = None
                        if wellFullImportedData == None:
                            wIDs = ezome.get_well_ids(userConn, plate=plateID)
                            omeWell = None
                            for wID in wIDs:
                                omeW = userConn.getObject("Well", wID, opts={"load_images": True})
                                if omeW.getName() == wellName:
                                    omeWell = omeW
                            
                            if omeWell == None:
                                row,col = getWellCoords(wellName)

                                createWell(userConn,plateID, row, col)


                                for w in userConn.getObject("Plate",plateID).listChildren():
                                    if w.getRow() == row and w.getColumn() == col:
                                        wellID = w.getId()

                                omeWell = userConn.getObject("Well", wellID, opts={"load_images": True})


                                wellCurrentImportedData[import_status] = (
                                    import_status_imported
                                )
                                writeToLog(
                                    "Well created for "
                                    + wellName
                                    + " ("
                                    + str(wellID)
                                    + ")"
                                )
                                hasNewImport = True
                            
                            else:
                                wellID = omeWell.getId()
                                wellCurrentImportedData[import_status] = (
                                    import_status_found
                                )
                                writeToLog(
                                    "Well found for "
                                    + wellName
                                    + " ("
                                    + str(wellID)
                                    + ")"
                                )
                        else:
                            wellID = wellFullImportedData[import_status_id]
                            omeWell = userConn.getObject("Well", wellID, opts={"load_images": True})
                            wellCurrentImportedData[import_status] = (
                                import_status_pimported
                            )
                            writeToLog(
                                "Well previously imported for "
                                + wellName
                                + " ("
                                + str(wellID)
                                + ")"
                            )
                        
                        wellCurrentImportedData[import_status_id] = wellID

                        wellKeyValueData = []

                        for wellAnnKey in well:
                            if (
                                wellAnnKey == metadata_file_name
                            ):
                                continue
                            wellKeyValueData.append([wellAnnKey, str(well[wellAnnKey])])
                        if (
                            wellFullImportedData == None
                            or import_annotate not in wellFullImportedData
                            # or imageFullImportedData[import_annotate] == False
                        ):
                            if len(wellKeyValueData) > 0:
                                omeWell = userConn.getObject("Well", wellID, opts={"load_images": True})
                                newwellMapAnn = MapAnnotationWrapper(userConn)
                                newwellMapAnn.setNs(namespace)
                                newwellMapAnn.setValue(wellKeyValueData)
                                newwellMapAnn.save()
                                omeWell.linkAnnotation(newwellMapAnn)
                                writeToLog(
                                    "Annotation created for "
                                    + wellName
                                    + " ("
                                    + str(wellID)
                                    + ")"
                                )
                            wellCurrentImportedData[import_annotate] = (
                                newwellMapAnn._obj.id.val
                            )
                            hasNewImport = True
                        else:
                            if len(wellKeyValueData) > 0:
                                wellMapAnnID = wellFullImportedData[import_annotate]
                                newwellMapAnn = userConn.getObject(
                                    "MapAnnotation", wellMapAnnID
                                )
                                newwellMapAnn.setValue(wellKeyValueData)
                            writeToLog(
                                "Annotation updated for "
                                + wellName
                                + " ("
                                + str(wellID)
                                + ")"
                            )
                            hasNewImport = True

                        #now we combine images and add them to each well
                        if metadata_OME_Images not in well:
                            continue
                        
                        wellObj = omeWell._obj
                        for imagekey in well[metadata_OME_Images]:
                            omeImageName = imagekey[metadata_OME_Image_Name]
                            
                            if omeUserName != None:
                                userConn = conn.suConn(omeUserName)
                                # userConn.c.enableKeepAlive(60)
                            else:
                                userConn = conn

                            imgFullImportedData = None
                            imgCurrentImportedData = None
                            if (
                                wellFullImportedData != None
                                and omeImageName in wellFullImportedData
                            ):
                                imgFullImportedData = wellFullImportedData[omeImageName]

                            if omeImageName not in wellCurrentImportedData:
                                wellCurrentImportedData[omeImageName] = {}
                            imgCurrentImportedData = wellCurrentImportedData[omeImageName]

                            imageId = None
                            omeImage = None
                            if imgFullImportedData == None:
                                imgIDs = ezome.get_image_ids(userConn)
                                omeImage = None
                                for imgID in imgIDs:
                                    omeI = userConn.getObject("Image", imgID)
                                    if omeI.getName() == omeImageName:
                                        omeImage = omeI
                                
                                #if the image doesnt exist, create it
                                if omeImage == None:
                                    files = imagekey[metadata_files]
                                    imageList = []
                                    for s in range (1,site+1):
                                        for c in range(1,channel+1):
                                            for t in range(1, timepoint+1):
                                                for z in range(1,zStep+1):
                                                    for file in files:
                                                        if file[metadata_Z_index] == z and file[metadata_T_index] == t and file[metadata_C_index] == c and file[metadate_Site_index] == s:
                                                            imageList.append(file[metadata_file_path])

                                    #store each image after they go through imread()
                                    imageDict = {}
                                    for i, image in enumerate(imageList):
                                        imageDict[i+1] = imread(image)
                                    
                                    #add these images back into a list
                                    planes = list(imageDict.values())

                                    def planeGen():
                                        for p in planes:
                                            yield p
                                    
                                    omeImage = userConn.createImageFromNumpySeq(
                                        planeGen(),omeImageName, zStep,channel,timepoint
                                    )
                                    imageId = omeImage.getId()

                                    imgCurrentImportedData[import_status] = (
                                        import_status_imported
                                    )

                                    imgCurrentImportedData[import_status_id] = imageId

                                    #now that the image/site is created, we can add it to the well
                                    wellSample = WellSampleI()
                                    wellSample.setImage(ImageI(imageId,False))
                                    wellObj.addWellSample(wellSample)

                                    writeToLog(
                                        "Image created for "
                                        + omeImageName
                                        + " ("
                                        + str(imageId)
                                        + ")"
                                    )
                                    hasNewImport = True
                                
                                else:
                                    imageId = omeImage._obj.id.val
                                    imgCurrentImportedData[import_status] = (
                                        import_status_found
                                    )
                                    writeToLog(
                                        "Image found for "
                                        + omeImageName
                                        + " ("
                                        + str(imageId)
                                        + ")"
                                    )   
                            else:
                                imageId = imgFullImportedData[import_status_id]
                                omeImage = userConn.getObject("Image", imageId)
                                imgCurrentImportedData[import_status] = (
                                    import_status_pimported
                                )
                                writeToLog(
                                    "Image previously imported for "
                                    + omeImageName
                                    + " ("
                                    + str(imageId)
                                    + ")"
                                )
                        #update the well once all sites are added to it


                        update_service = userConn.getUpdateService()
                        update_service.saveObject(wellObj)


                        imgCurrentImportedData[import_status_id] = imageId                                                      

                        if (
                            startTimeHr != None
                            and startTimeMin != None
                            and endTimeHr != None
                            and endTimeMin != None
                        ):
                            now = datetime.now()
                            if now.hour < startTimeHr and now.hour > endTimeHr:
                                endTimePassed = True
                            if now.hour == endTimeHr and now.minute > endTimeMin:
                                endTimePassed = True
                            if now.hour == startTimeHr and now.minute < startTimeMin:
                                endTimePassed = True

                        if omeUserName != None:
                            userConn.close()

                        if endTimePassed:
                            break
                    if endTimePassed:
                        break
                if endTimePassed:
                    break
            if endTimePassed:
                break

        # sendCompleteEmail(
        #     emailTo,
        #     adminsEmailTo,
        #     hasNewImport,
        #     currentImportedData[userFolder],
        #     emailFrom,
        #     emailFromPSW,
        # )
        if endTimePassed:
            break

    conn.close()
    printToConsole("Close connection")
    

    writeCurrentImported(currentImportedData)
    if fullImportedData != None:
        mergedImportedData = mergeDictionaries(currentImportedData, fullImportedData)
    else:
        mergedImportedData = deepCopyDictionary(currentImportedData)
    printToConsole("mergedImportedData " + str(mergedImportedData))
    writePreviousImported(localPath, mergedImportedData)


if __name__ == "__main__":
    main(sys.argv, len(sys.argv))
