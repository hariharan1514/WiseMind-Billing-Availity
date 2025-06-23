import pandas as pd
from datetime import datetime, timedelta
import os


### Check the file path ###
parentFolderPath = r"Z:\Wisemind\Charge Entry -Billing\Billing Dates"
pathTemp_Date = datetime.today().strftime('%m%d%Y')
configSheetPath = r"Z:\Wisemind\Charge Entry -Billing\Automation Config File\ConfigSheet.xlsx"

attendanceDatafilepath = (parentFolderPath + "\\" +pathTemp_Date[4:] + "\\" +datetime.today().strftime("%m %b'%Y") + "\\" +pathTemp_Date + "\\" +f"Detailed Attendance - {pathTemp_Date}.xlsx")
allClientsDatafilepath = (parentFolderPath + "\\" +pathTemp_Date[4:] + "\\" +datetime.today().strftime("%m %b'%Y") + "\\" +pathTemp_Date + "\\" +f"All Claims - {pathTemp_Date}.xlsx")

if configSheetPath:
    configsheet = pd.read_excel(configSheetPath, sheet_name=1)

    # Scrubbing Filter Criterias
    remove_status = configsheet['Status'].dropna().tolist()
    noShow_CPTs = configsheet["No Show CPT's"].dropna().astype(int).tolist()
    lateCancel_CPTs = configsheet["Late Cancel CPT's"].dropna().astype(int).tolist()
    kept_CPTs = configsheet["Kept CPT's"].dropna().astype(int).tolist()
    availityPayor = configsheet['Availity Payors'].dropna().tolist()
    exception_ClientID = configsheet['Exception Client ID'].dropna().tolist()

    if os.path.exists(attendanceDatafilepath) and os.path.exists(allClientsDatafilepath):

        ### Read the Scrubbing file and All Claims file ###
        scrubbing_df = pd.read_excel(attendanceDatafilepath)
        allClaims_df = pd.read_excel(allClientsDatafilepath)

        # Remove the Yes label in Is Billed Column
        removeYes_df = scrubbing_df[scrubbing_df['Is Billed'] != "Yes"]

        exception_df = scrubbing_df[scrubbing_df['Is Billed'] == "Yes"]

        # In the “Status” Column we need to select only “Canceled, Rescheduled, Upcoming” labels and remove those rows in a data.
        statusCancelation_df = removeYes_df[~removeYes_df['Status'].isin(remove_status)]

        exception_df = exception_df[~(exception_df['Is Billed'] == "Yes") | (~exception_df['Status'].isin(remove_status))]


        # In the “Status” Column we need to select only “No Show” label and “Service Type” column we need to unselect only “Late Cancelation” labels and remove remaining rows in a data.
        noShow_CPTs_Pattern = r'^(' + '|'.join(f"{code}:" for code in noShow_CPTs) + r')'
        cancellation993_df = statusCancelation_df[(statusCancelation_df['Status'] != "No Show") | (statusCancelation_df['Service Type'].str.match(noShow_CPTs_Pattern))]

        exception_df = exception_df[~(exception_df['Status'] == "No Show") | (~exception_df['Service Type'].str.match(noShow_CPTs_Pattern))]

        # In the “Status” column we need to select only “Late Cancel” label and “Service Type” column need to Unselect “all CPT’s & 993 cancellation” and remove remaining rows in a data."
        lateCancel_CPTs_Pattern = r'^(' + '|'.join(f"{code}:" for code in lateCancel_CPTs) + r')'
        lateCancelation_CPTs_df = cancellation993_df[(cancellation993_df['Status'] != "Late Cancel") | (cancellation993_df['Service Type'].str.match(lateCancel_CPTs_Pattern))]

        exception_df = exception_df[~(exception_df['Status'] == "Late Cancel") | (~exception_df['Service Type'].str.match(lateCancel_CPTs_Pattern))]

        ### In the “Status” column we need to select only “Kept” label and “Service label” column we need to select all labels except “all CPT’s & 993: cancellation & 999: Individual Self Pay” and remove remaining rows. ###
        kept_CPTs_Pattern = r'^(' + '|'.join(f"{code}:" for code in kept_CPTs) + r')'
        kept_CPTs_df = lateCancelation_CPTs_df[(lateCancelation_CPTs_df['Status'] != "Kept") | (lateCancelation_CPTs_df['Service Type'].str.match(kept_CPTs_Pattern))]
        print(len(kept_CPTs_df))

        exception_df = exception_df[~(exception_df['Status'] == "Kept") | (exception_df['Service Type'].str.match(kept_CPTs_Pattern))]

        exceptionFileSavePath = attendanceDatafilepath.replace(f'Detailed Attendance - {pathTemp_Date}.xlsx',f"Exception Scrubbing File - {pathTemp_Date}.xlsx")
        exception_df.to_excel(exceptionFileSavePath, index=False)

        # Need to create one column named as “Payor Name” and capture the Payers using of VLOOKUP function (=VLOOKUP (Scrubbing Data Client ID Number column, all claims’ Data Client ID column, all claims’ data Primary Insurance: Provider Name Col index number,0))

        allClaims_df = allClaims_df[['Client Id','Primary Insurance: Provider Name']]
        vLookup_df = kept_CPTs_df.merge(allClaims_df,how='left',left_on='Client ID Number',right_on='Client Id')
        vLookup_df.rename(columns={'Primary Insurance: Provider Name': 'Payor Name'}, inplace=True)
        vLookup_df.drop(columns=['Client Id'],inplace=True)

        #Change the UnitedHealthcare label to Optum
        # united_health = ["United Healthcare", "UnitedHealthcare"]
        vLookup_df.loc[vLookup_df['Payor Name'].str.contains('United Healthcare', na=False), 'Payor Name'] = 'Optum'
        vLookup_df.loc[vLookup_df['Payor Name'].str.contains('UnitedHealthcare', na=False), 'Payor Name'] = 'Optum'

        # manualFileSavePath = attendanceDatafilepath.replace(f'Detailed Attendance - {pathTemp_Date}.xlsx',
        #                                                     f"Manual Scrubbing File - {pathTemp_Date}.xlsx")
        # vLookup_df.to_excel(manualFileSavePath, index=False)

        InactivePattern = r"(Inactive)"
        inactiveRemove_df = vLookup_df[~vLookup_df['Staff Member(s)'].str.contains(InactivePattern)]

        exceptionClients_df = inactiveRemove_df[~inactiveRemove_df['Client ID Number'].isin(exception_ClientID)]
        manualFileSavePath = attendanceDatafilepath.replace(f'Detailed Attendance - {pathTemp_Date}.xlsx',f"Manual Scrubbing File - {pathTemp_Date}.xlsx")
        exceptionClients_df.to_excel(manualFileSavePath, index=False)

        exceptionClients_df = inactiveRemove_df[~inactiveRemove_df['Client ID Number'].isin(exception_ClientID)]

        # In PayorName Columns remove Blanks
        removeBlanks_df = exceptionClients_df[exceptionClients_df['Payor Name'].notna()]

        # filter last Date only
        today = datetime.today()
        day_name = today.strftime("%A")
        print(day_name)
        if day_name == 'Monday':
            # Get dates for last Thursday (3), Friday (4), Saturday (5)
            last_thursday = today - timedelta(days=(today.weekday() - 3) % 7)
            last_friday = today - timedelta(days=(today.weekday() - 4) % 7)
            last_saturday = today - timedelta(days=(today.weekday() - 5) % 7)

            filterDate = [last_thursday.strftime('%m/%d/%Y'), last_friday.strftime('%m/%d/%Y'),last_saturday.strftime('%m/%d/%Y')]
            removePreviousDate_df = removeBlanks_df[removeBlanks_df['Date/Time'].str[:10].isin(filterDate)]

        else:
            filterDate = datetime.today() - timedelta(days=2)
            filterDate = filterDate.strftime('%m/%d/%Y')
            removePreviousDate_df = removeBlanks_df[removeBlanks_df['Date/Time'].str.match(filterDate)]

        # Remove SelfPay Payors
        remove_selfpay_payors = removePreviousDate_df[~removePreviousDate_df['Payor Name'] == "SelfPay"]

        # # Drop Availity Payors
        # dropAvailityPayor_df = removePreviousDate_df[~removePreviousDate_df['Payor Name'].isin(availityPayor)]

        automationFileSavePath = attendanceDatafilepath.replace(f'Detailed Attendance - {pathTemp_Date}.xlsx',f"Straightforward Billing Case - {pathTemp_Date}.xlsx")
        remove_selfpay_payors.to_excel(automationFileSavePath,index=False)
        print("Scrubbing Process Completed !")

    else:
        print("\n")
        print("First, run the 'WiseMind_Billing_Phase1.py' script, and then run the Scrubbing script.")
else:
    print("Configure file is missing !.")