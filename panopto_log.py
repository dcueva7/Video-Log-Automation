#!/usr/bin/env python3
import pandas as pd
import gspread 
import gspread_dataframe as gd
from gspread_dataframe import get_as_dataframe, set_with_dataframe


def create_dfs():
    """
    Convert Input Excel File Into a Python Dataframe.
    Read input excel.
    Instanstiate "Days" dataframe.

    return: 6 dataframes; input_df, monday_df, tuesday_df, wednesday_df, thursday_df, friday_df
    """

    input_df = pd.read_excel(r'/Users/cuevs/Downloads/Summer_list.xlsx')
    monday_df = pd.DataFrame(columns=['Course', 'Professor', 'Start Time', 'End Time', 'Room'])
    tuesday_df = pd.DataFrame(columns=['Course', 'Professor', 'Start Time', 'End Time', 'Room'])
    wednesday_df = pd.DataFrame(columns=['Course', 'Professor', 'Start Time', 'End Time', 'Room'])
    thursday_df = pd.DataFrame(columns=['Course', 'Professor', 'Start Time', 'End Time', 'Room']) 
    friday_df = pd.DataFrame(columns=['Course', 'Professor', 'Start Time', 'End Time', 'Room'])
    return input_df, monday_df, tuesday_df, wednesday_df, thursday_df, friday_df 

def populate_dfs(input_df, monday_df, tuesday_df, wednesday_df, thursday_df, friday_df):
    # Iterate through input dataframe
    for row_count in range(len(input_df)):
        mode = str(input_df.loc[row_count,'Mode'])
        if "Lec" in mode:
            course = f"{input_df.loc[row_count,'Course']}"
        else:
            course = f"{input_df.loc[row_count,'Course']}-{mode}"
        professor = f"{input_df.loc[row_count,'Professor']}"
        start_time = str(input_df.loc[row_count, 'Start'])
        room = str(input_df.loc[row_count, 'Room'])
        end_time = str(input_df.loc[row_count,'End'])
        row = [course, professor, start_time, end_time, room]
        days = str(input_df.loc[row_count, 'Days'])
        week_df = (monday_df, tuesday_df, wednesday_df, thursday_df, friday_df)
        day_selection(row, days, week_df)
    return monday_df, tuesday_df, wednesday_df, thursday_df, friday_df 

def day_selection(row, days, week_df):
    monday_df, tuesday_df, wednesday_df, thursday_df, friday_df = week_df
    if "Th" == days:
        thursday_df.loc[len(thursday_df.index)] = row
    else:
        if "M" in days:
            monday_df.loc[len(monday_df.index)] = row
        if "T" in days:
            tuesday_df.loc[len(tuesday_df.index)] = row
        if "W" in days:
            wednesday_df.loc[len(wednesday_df.index)] = row
        if "Th" in days:
            thursday_df.loc[len(thursday_df.index)] = row
        if "F" in days:    
            friday_df.loc[len(friday_df.index)] = row

def pop_sheet(mon, tu, wed, thu, fri, gsheet):
    pass
    #gsheet.update_cells('A5', 


def main():
    #create dataframes
    input_df, monday_df, tuesday_df, wednesday_df, thursday_df, friday_df = create_dfs()
    
    #populate dataframes
    monday_df, tuesday_df, wednesday_df, thursday_df, friday_df = populate_dfs(input_df, monday_df, tuesday_df, wednesday_df, thursday_df, friday_df)
    
    #populate google sheet
    
    # Google API Credentials
    sa = gspread.service_account(filename ='service_acct.json')
    # open gsheets  
    sh = sa.open("DENTSC Spring 2023 Panopto Posting Log")
    wks = sh.worksheet("SUMMERWK1 5/24-4/29")

    set_with_dataframe(wks, monday_df, 6, include_column_header=False )
    #gsheet.update_cell()
    set_with_dataframe(wks, tuesday_df, 45, include_column_header=False )
    set_with_dataframe(wks, wednesday_df,91, include_column_header=False )
    set_with_dataframe(wks, thursday_df, 130, include_column_header=False )
    set_with_dataframe(wks, friday_df, 177, include_column_header=False )

if __name__ == '__main__':
    main()