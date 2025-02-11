# RPA-SAP

This Python script automates the scheduling of sales orders (SOs) in the SAP system, covering the steps of updating amounts in VA02, planning in CN33, allocating dates in CJ20N, and running MRP in MD51. In addition, the script sends automatic emails to the responsible MRP planners after scheduling.

## Description

The ProgramadorSAP script simplifies the process of scheduling orders in SAP, reducing the time and manual effort required. It interacts with the SAP GUI through the `win32com` library, reads data from Excel spreadsheets, performs consistency checks, and executes the VA02, CN33, CJ20N, and MD51 transactions in an automated manner. Additionally, the script sends emails to MRP planners with information about the scheduled orders, using the `win32com` library to interact with Outlook.

## Features

* **Input data reading:** Reads data from Excel spreadsheets containing information about the orders to be scheduled, including EC, work, branch, sales order, origin, PEP element, status, and total cost.
* **Consistency check:** Checks if the input spreadsheets are consistent, looking for missing, duplicate, or null volume ECs.
* **Updating amounts in VA02:** Updates the amount of sales orders in the VA02 transaction of SAP.
* **Planning in CN33:** Performs order planning in the CN33 transaction of SAP.
* **Date allocation in CJ20N:** Allocates planning dates for order components in the CJ20N transaction of SAP.
* **MRP execution in MD51:** Executes Materials Requirements Planning (MRP) for orders in the MD51 transaction of SAP.
* **PEP Element file generation:** Generates a text file containing the PEP Elements of the scheduled orders.
* **Order conversion and release (COHV):** Converts and releases planned orders through the COHV transaction.
* **Email sending:** Sends automatic emails to responsible MRP planners, informing them about the scheduled orders and attaching relevant information.
* **Error handling:** Handles errors during the steps of interaction with SAP and email sending, recording them in the spreadsheets and informing the user.

## How to use

1. **Pre-requisites:**
    * Python 3 installed.
    * `win32com`, `pandas`, `xlwings`, `glob`, `os`, `datetime` and `traceback` libraries installed (`pip install pywin32 pandas xlwings`).
    * SAP GUI installed and configured.
    * Access to transactions VA02, CN33, CJ20N, MD51 and COHV in SAP.
    * Input Excel spreadsheets and configuration files in the paths specified in the code.

2. **Configuration:**
    * Adjust the paths of the files and folders in the variables `path_ec`, `path_ec_script`, `file_ec`, `file_ec_script`, `file_macro`, `file_pep` and `file_itens` at the beginning of the script, if necessary.
    * Check and adjust the names of spreadsheets and tabs, if necessary.
    * Configure the connection parameters with SAP, such as the instance name and system number.
    * Set the SAP user profile to be used.
    * Adjust the email recipients in the `email_cc` variable in the `EnviarEmail` class.

3. **Execution:**
    * Run the Python script.
    * The script will interact with SAP, read the data from the spreadsheets, perform the scheduling operations and send the emails.
    * Track progress and errors through the console.

## Dependencies

* python >= 3.0
* pywin32
* pandas
* xlwings
* glob
* os
* datetime
* traceback

## Contribution

Contributions are welcome! Feel free to open issues and pull requests.

## Author

Gustavo Nunes Ferraz