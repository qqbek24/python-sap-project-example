
### Only for read !!! 
### To use it, you need additional custom libraries for bot and sap.


# Description

Task returns an Excel file (table) with processed documents from cockpit. Document can be posted or verified. Last column in Excel contains actual status of document. The file is sent to client.


| Data                     | Values                                                                                                                                                                                                       |
| ------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| **Developer:**           | Jakub Koziorowski                          |
| **Business Analyst:**    |                                        |
| **Business Owners:**     |                          |


## Prod environment:


| Parameter Name       | Value         |
| :------------------- | :------------ |
| **SAP System Name:** |           |
| **Client:**          |          |
| **Language:**        |           |
| **Transaction:**     | cockpit/1; XK03   |
| **Variant:**         |      |
| **SAP PCE user:**    |         |

cockpit/1 - list of scanned documents waiting for processing

## Test environment:


| Parameter Name   | Value         |
| :--------------- | :------------ |
| **SAP System Name:** |           |
| **Client:**          |          |
| **Language:**        |           |
| **Transaction:** | cockpit/1; XK03   |
| **Variant:**     |        |
| **SAP ACE user:**     |         |

## Trigger

Jenkins schedule or on demand via email trigger (Excel list must be enclosed)

## Input

Initial Input:

Case email trigger:

- list of documents for processing as xlsx attachment
- number of days for date range (last character in email subject)


## Output

- Excel file with documents and process status as **File name: "Report_datetime_MM_FR.xlsx"**
- Email with the report and number of processed and not processed documents
- Email a

## Process

1. Download list of documents from cockpit/1 transaction - **Variant: "mm_fra2"**
2. Extract list of documents basing on vendor list orders - **Layout: " "**
3. Process each document from step 2 according to process_item() from process_sap.py
5. Send final report with comments to client


## Error handling

1. In a case when an error occurred bot sends an email with a message what happened.
2. In case of any other error, recommendation is sent to contact RPA team (except standard log mail to operator/bot account mail)

