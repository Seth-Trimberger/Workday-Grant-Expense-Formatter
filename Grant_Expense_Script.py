# import section
import numpy as np
import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox


# class used to store header data from the top of the excel file
class OutPutClass:
    # this sets up all the default values for the header info
    def __init__(self):
        self.Award = ""
        self.Grant = ""
        self.Principal_Investigator = ""
        self.CostCenter = ""
        self.CostCenterHierarchies = ""
        self.AccountingStartDate = ""
        self.AccountingEndDate = ""
        self.BudgetStartDate = ""
        self.BudgetEndDate = ""
        self.TransactionStartDate = ""
        self.TransactionEndDate = ""


# this reads the first 12 lines of the excel file and pulls out the header info
def ReadInFirst12Lines(filename):
    print("reading in the frist 12 lines")

    df = pd.read_excel(filename, skiprows=1, nrows=12, header=None)

    output = OutPutClass()

    # loop through header rows and match labels to class properties
    for index, row in df.iterrows():
        key = str(row[0]).strip().replace(" ", "").replace("_", "")
        value = "" if pd.isna(row[1]) else str(row[1]).strip()

        if key.lower() == "award":
            output.Award = value
        elif key.lower() == "grant":
            output.Grant = value
        elif key.lower() == "principalinvestigator":
            output.Principal_Investigator = value
        elif key.lower() == "costcenter":
            output.CostCenter = value
        elif key.lower() == "costcenterhierarchies":
            output.CostCenterHierarchies = value
        elif key.lower() == "accountingstartdate":
            output.AccountingStartDate = value
        elif key.lower() == "accountingenddate":
            output.AccountingEndDate = value
        elif key.lower() == "budgetstartdate":
            output.BudgetStartDate = value
        elif key.lower() == "budgetenddate":
            output.BudgetEndDate = value
        elif key.lower() == "transactionstartdate":
            output.TransactionStartDate = value
        elif key.lower() == "transactionenddate":
            output.TransactionEndDate = value

    return output


# this builds the output file name using pi last name and accounting start date
def CreateOutPutFileName(outputData: OutPutClass):
    # hard coded folder where the new files will be saved
    OutputDirectory = r"OutPut"

    # grab pi name and try to pull just the last name
    pi_full = outputData.Principal_Investigator.strip()

    if pi_full == "":
        pi_last = "unknownpi"
    else:
        pi_last = pi_full.split()[-1]

    # grab the accounting start date and format it safely
    raw_date = outputData.AccountingStartDate

    if raw_date == "":
        formatted_date = "unknowndate"
    else:
        try:
            date_obj = pd.to_datetime(raw_date)
            formatted_date = date_obj.strftime("%Y-%m-%d")
        except:
            formatted_date = str(raw_date).replace("/", "-")

    # build the final file name string
    filename = f"expense detail repot {pi_last} {formatted_date}.xlsx"

    # combine folder and file name
    full_path = os.path.join(OutputDirectory, filename)

    return full_path


# this reads the actual transaction data and reformats it
def ReadInExcellFiles(filename):
    df = pd.read_excel(filename, skiprows=12)

    # these are the object class codes we only care about
    CodesToFilter = ["030", "032", "033", "060", "160"]
    pattern = r"\b(?:" + "|".join(CodesToFilter) + r")\b"

    filtered = df[df['Object Class'].astype(str).str.contains(pattern, na=False)].copy()

    out = pd.DataFrame()

    # map excel columns into the new output format
    out["grant number"] = filtered["Grant"].astype(str).str[:10]
    out["accounting date"] = filtered["Accounting Date"]

    supplier = filtered["Supplier"].astype(str)
    merchant = filtered["Merchant"].astype(str)

    supplier = supplier.replace("nan", "")
    merchant = merchant.replace("nan", "")

    out["supplier"] = np.where(supplier != "", supplier, np.where(merchant != "", merchant, ""))
    out["amount"] = filtered["Amount"]
    out["spend catergory"] = filtered["Spend Category"]

    # use memos as the description field
    if "Memos" in filtered.columns:
        out["description"] = filtered["Memos"]
    else:
        out["description"] = ""

    # pull the long document column if found
    doc_col = "Initiating Spend Transaction of Facilities And Administration or Award Revenue Operational Journal"

    if doc_col in filtered.columns:
        out["document"] = filtered[doc_col]
    else:
        out["document"] = ""

    # split all rows into separate dataframes based on grant number
    grant_groups = {
        grant: subdf.reset_index(drop=True)
        for grant, subdf in out.groupby("grant number")
    }

    return grant_groups


# this opens a file picker and allows selecting multiple excel files
def SelectExcellFiles():
    thisRoot = tk.Tk()
    thisRoot.withdraw()

    filePaths = filedialog.askopenfilenames(
        initialdir=r"OutPut",
        title="select the excell files to run: ",
        filetypes=[("excell files", ".xlsx")]
    )

    thisRoot.destroy()
    return list(filePaths)


# this writes each grant to its own sheet in the excel output file
def WriteGrantsToExcel(grant_groups, output_filename):
    with pd.ExcelWriter(output_filename, engine="xlsxwriter") as writer:
        # loop through each grant group and save as a sheet
        for grant, GrantDataFrame in grant_groups.items():

            if pd.isna(grant) or str(grant).strip() == "":
                sheet_name = "unknown_grant"
            else:
                sheet_name = str(grant)

            GrantDataFrame.to_excel(writer, sheet_name=sheet_name, index=False)


# main driver that controls the full program flow
def main():
    FilesToOpen = SelectExcellFiles()

    if not FilesToOpen:
        print("no files selcted. exiting.")
        return

    # loop through each selected excel file and process it
    for FileToOpen in FilesToOpen:
        print("processing:", FileToOpen)

        HeaderData = ReadInFirst12Lines(FileToOpen)
        OutputFilePath = CreateOutPutFileName(HeaderData)

        # check if output file already exists and skip if it does
        if os.path.exists(OutputFilePath):
            messagebox.showerror(
                "file alreayd exists",
                f"the output file already exists:\n\n{OutputFilePath}\n\nthis file will be skiped."
            )
            continue

        grant_groups = ReadInExcellFiles(FileToOpen)


        WriteGrantsToExcel(grant_groups, OutputFilePath)
        print("done prossesing:", FileToOpen)


main()
