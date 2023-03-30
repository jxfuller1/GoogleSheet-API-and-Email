import pygsheets
from datetime import datetime, date, timedelta
#import datetime
import time
import pandas as pd
import os
import smtplib
from email.message import EmailMessage

#note, parts_possible list and ones_completed list and are the ones i want to pass to email and also
#tostring_count_completed and tostring_count_notcompleted


#for getting previous day of week accounting for weekdays
def prev_weekday(adate):
    adate -= timedelta(days=1)
    while adate.weekday() > 4: # Mon-Fri are 0-4
        adate -= timedelta(days=1)
    return adate

#for setting time of day to send out email
def check_time(hour, minute, current_time):
    if "AM" in current_time.upper():
        if "06" in hour:
            if "50" in minute:
                return True
    else:
        return False


# Create the Client
# Enter the name of the downloaded KEYS
# file in service_account_file\
client = pygsheets.authorize(service_account_file="path to .json file key from google service account")

# Sample command to verify successful
# authorization of pygsheets
# Prints the names of the spreadsheet
# shared with or owned by the service
# account
#print(client.spreadsheet_titles())

while True:
    #only iterate loops every XX time checking for time of day to run script and send out email
    time.sleep(60)
    #get current time and am/pm
    now = datetime.now()
    #turn current time into a string
    current_time = now.strftime("%H:%M:%S %p")

    #split time in order to separate hour and minute to send to a function later
    split_time = current_time.split(":")
    hour = split_time[0]
    minute = split_time[1]

    #send hour / time to function to see ifit's the correct time to execute rest of program
    clock_time = check_time(hour, minute, current_time)

    #if correct time of day run script
    if clock_time == True:

        # opens a spreadsheet by its name/title
        spreadsht = client.open("name of google sheet to be read")

        #read sheet
        wks = spreadsht.worksheet_by_title("Sheet1")

        #gets previous weekday and accounts for weekeends, getting previous date as i want to collect the date for the prevoius day
        Previous_Date = prev_weekday(date.today()).strftime("%m/%d/%Y")

        #split into month day year
        month, day, year = Previous_Date.split("/")

        #for removing 0's in from of the day number for comparison later
        if len(day) == 2:
            if "0" in day[0]:
                day = day.replace("0", "")

        #for removing 0's in from the month number for comparison later
        if len(month) == 2:
            if "0" in month[0]:
                month = month.replace("0", "")

        #setup list for passing to email later
        ones_completed = []

        #pass google sheet data as a dataframe and read the data i want from it
        df = wks.get_as_df()

        #start loop  to capture the data i want out of the google sheet
        k = 0
        while k < len(df):
            #looks for this symbol in this column as it contains dates i want to iterate through
            if "/" in str(df.iloc[k, 2]):  #read dates in google sheet
                #get date
                part_date = str(df.iloc[k, 2])
                #print(part_date)

                #split day from googlesheet into a list
                total_day = part_date.split("/")

                #put day and month in separate values
                day_part = total_day[1]
                month_part = total_day[0]
                #print(day_part, month_part)

                #if previous day/month match up to what is in cell on google sheet collect data
                #these are ones completed on the previous day
                if month in month_part and day in day_part:
                    #if month/day matches get part number, description and job number
                    part_number = str(df.iloc[k, 0][0:10])
                    part_descript = str(df.iloc[k, 1])
                    part_job = str(df.iloc[k, 3])
                    #pass the information to the ones_complete list for passing to email
                    ones_completed.append("          " + part_number + " " + part_job + "   " + part_descript)
            k +=1

        #if none found in above loop add this to the list, have spaces in it for correct formatting for email
        if len(ones_completed) == 0:
            ones_completed.append("          append something")

        #print(ones_completed)

        #folder where data for inspections for each day are stored
        excel_path = "//base pathway to excel files"

        #get all files and pass into list in folder
        readfolder = os.listdir(excel_path)

        #set a switch variable if there's no excel made
        no_excel = False

        #loop for iterating all folders to look for previous day, the files are named with the date, im looking
        #for a folder with the date im trying to collect data for
        h = 0
        while h < len(readfolder):
            #if excel has the - in the name split it up as the dates are split with a - symbol
            if "-" in readfolder[h]:
                #split name of folder
                split_folder_name = readfolder[h].split("-")
                #get month/date into own variables for checking if they match the dates im looking to collect data for
                done_month = split_folder_name[0]
                done_day = split_folder_name[1]
                #check if dates match , then it's the excel i want to read down below
                if month in done_month and day in done_day:
                    #excel path variables
                    full_excel_path = excel_path + readfolder[h]
                    #set variable to true that the excel file exists
                    no_excel = True
            h +=1

        #setup lists for checking/passing to email later
        all_ones_done = []
        parts_possible = []

        #if excel file exists
        if no_excel == True:
            #if excel exists, read as a dataframe
            df_excel = pd.read_excel(full_excel_path)

            #append all parts done the previous day to a list for easy lookup
            #loops for collecting certain data out of the excel into a list for checking later
            k = 0
            while k < len(df_excel):
                if "nan" not in str(df_excel.iloc[k, 9]):  # read part numbers in excel
                    part_numbers_done = str(df_excel.iloc[k, 9][0:10])
                    job_number_done = str(df_excel.iloc[k, 0])
                    part_descript_done = str(df_excel.iloc[k, 10])
                    all_ones_done.append(part_numbers_done + " " + job_number_done + "   " + part_descript_done)
                k += 1

            #iterate through parts that were comopleted and compare against google sheet to see if one on there that was done by QP instead of FAI made
            k = 0
            while k < len(all_ones_done):
                split_ones_done = all_ones_done[k].split(" ")

                #for every item on all_ones_done list, iterate through the googlesheet to see if
                #that item is on it and if it has an FAI done or not by checking data in certain cells
                p = 0
                while p < len(df):
                    if "AK" not in split_ones_done[0]:  #only look at PK's
                        if split_ones_done[0] in str(df.iloc[p, 0]):  # read part number on google sheet to compare against ones done
                            if len(str(df.iloc[p, 2])) == 0:
                                #if FAI made, append to list that has additional spaces in it
                                #this additional spaces is just for setting up formatting in the email
                                parts_possible.append("          " + all_ones_done[k])
                                #print(str(df.iloc[p, 0]) + " this is zero")
                    p+=1
                k +=1

            #if no matches and list is 0, set the list to something for passing into email later
            if len(parts_possible) == 0:
                parts_possible.append("          None!")
        else:
            #if excel doesn't exist append this to list for sending in email
            parts_possible.append("          Data Not Available")

        #this is for finding how many on list done vs how many to go on google sheet using logic
        k = 0
        count_completed = 0
        count_notcompleted = 0
        while k < len(df):
            #if date entered into cell , that means FAI made, and count it for passing to email later
            if "/" in str(df.iloc[k, 2]):
                count_completed +=1
            #if no date in cell and different cells contain certain data, then count it for passing to email later
            if len(str(df.iloc[k, 2])) == 0:
                if "PK" in str(df.iloc[k, 0]):
                    if "X" not in str(df.iloc[k, 6]):
                        count_notcompleted +=1
            k +=1

        #set the counts as a string instead of a integer
        tostring_count_completed = str(count_completed)
        tostring_count_notcompleted = str(count_notcompleted - 4)  #minusing 4 from this count to return correct value

        #print(ones_completed, parts_possible, tostring_count_completed, tostring_count_notcompleted)


        #sent the strings up to put all the data into one string for the email
        tostringones_completed = '\n'.join(ones_completed)
        tostringparts_possible = '\n'.join(parts_possible)
        count_completed_mail = "          Completed:  " + tostring_count_completed
        count_notcompleted_mail = "          Remaining:  " + tostring_count_notcompleted

        #complete message to send in the email.... im just sending a huge string to the email
        #rather than setting up some HTML code to do the formatting... which would be better....
        #but im too lazy to figure it out at the moment
        total_email_msg = "Auto Generated E-mail for " + Previous_Date + "\n" + "=================================================" + "\n\n\n" + "headline:\n\n" + tostringones_completed + "\n\n\n" + \
                          "headline:\n\n" + tostringparts_possible + "\n\n\n" + "headline:\n\n" + count_completed_mail + "\n" + count_notcompleted_mail

        #send the emails out
        msg = EmailMessage()
        msg.set_content(total_email_msg)
        msg['Subject'] = 'subject of email'
        msg['From'] = "email.com"
        msg['To'] = ["email.com"]

        # Send the message via our own SMTP server.
        #port 465 is specifically for GMail using smtplib
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.login("email.com", "password to email")
        server.send_message(msg)
        server.quit()

