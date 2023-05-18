# Daniel Cueva (dcueva@usc.edu)

# Last revised - 12/16/2022

from operator import index
import pandas as pd
import gspread
import gspread_dataframe as gd
from gspread_dataframe import get_as_dataframe, set_with_dataframe
from datetime import datetime as dt, timedelta


def filter_excel():
    """
    Importing excel sheet into pandas dataframe
    Filtering excel sheet by DEN room

    """
    # importing excel sheet into dataframe
    input_df = pd.read_excel('/Users/cuevs/Downloads/VSOE-20232 from 4-25-23.xlsx',dtype={'CourseNumber': str})

    # creating dataframe to host filtered content with desired columns    
    filtered = pd.DataFrame(
        columns=['Course', 'Professor', 'Start', 'End', 'Room', 'Days', 'Mode'])

    # list of den rooms
    den_rooms = ['OHE100B', 'OHE100C', 'OHE100D', 'OHE114', 'OHE120', 'OHE122', 'OHE132', 'OHE136', 'SCT1501K',
                 'SCT1501M', 'RTH105', 'RTH109', 'RTH115', 'RTH217', 'SGM123', 'SGM124', 'DEN@Viterbi']

    count = 0
    # iterating through dataframe
    for row in range(len(input_df)):
        room = input_df.loc[row,'Room']
        courseDept =  str(input_df.loc[row, 'CourseDept'])
        courseNum = str(input_df.loc[row, 'CourseNumber'])
        
        # checking if course is held in DEN room
        if room in den_rooms:
            #if room is in RTH or SGM check if it has a DEN section
            if ('RTH' in str(room) or 'SGM' in str(room)) and not(sgm_rthFilter(courseDept=courseDept,courseNum=courseNum,input_df=input_df)):
                    continue

            # checking for professor name (some classes don't have professor listed)
            if input_df.loc[row,'Instructors'] != '':
                professor = str(input_df.loc[row,'Instructors']).split(',')[0]
            else:
                professor = ''

            # populating filtered data frame with filtered content
            cells = [str(courseDept) + str(courseNum), professor, input_df.loc[row, 'BegTime'], input_df.loc[row, 'EndTime'], input_df.loc[row, 'Room'], input_df.loc[row, 'Days'], input_df.loc[row, 'Mode']]
            filtered.loc[count] = cells
            count += 1

            days = str(filtered.loc[count-1,'Days'])
            length = len(days)

            #if course happens on multiple days, create additional row for each day
            if ((length > 1) and (str(days) != 'Th')):
                if (days[0] == 'T') and (days[1] == 'h'):
                    filtered.loc[count-1,'Days'] = 'Th'
                    start = 2
                else:
                    filtered.loc[count-1,'Days'] = days[0]
                    start = 1

                for item in range(start,length):
                    if days[item] == 'M':
                        filtered.loc[count] = filtered.loc[count-1]
                        filtered.loc[count,'Days'] = 'M'
                        count += 1

                    elif (days[item] == 'T') and (days[item+1] != 'h'):
                        filtered.loc[count] = filtered.loc[count-1]
                        filtered.loc[count,'Days'] = 'T'
                        count += 1

                    elif days[item] == 'W':
                        filtered.loc[count] = filtered.loc[count-1]
                        filtered.loc[count,'Days'] = 'W'
                        count += 1

                    elif (days[item] == 'T') and (days[item+1] == 'h'):
                        filtered.loc[count] = filtered.loc[count-1]
                        filtered.loc[count,'Days'] = 'Th'
                        count += 1

                    elif days[item] == 'F':
                        filtered.loc[count] = filtered.loc[count-1]
                        filtered.loc[count,'Days'] = 'F'
                        count += 1


    return filtered

def remove_duplicates(df):

    length = range(len(df))

    #length = range(0,358)
    rows_to_delete = []

    for row in length:
        course = df.loc[row, 'Course']
        start = df.loc[row, 'Start']
        end = df.loc[row, 'End']
        day = df.loc[row,'Days']
        mode = df.loc[row, 'Mode']
        room = df.loc[row,'Room']
        
        for item in length: #len(df)
            #check to make sure not comparing same row
            if row != item:
                if((df.loc[item, 'Course'] == course) and (df.loc[item, 'Start'] == start) and (df.loc[item, 'End'] == end) and (df.loc[item, 'Days'] == day) and (df.loc[item, 'Mode'] == mode) and (df.loc[item, 'Room'] == 'DEN@Viterbi')):
                    rows_to_delete.append(item)  

    df.drop(rows_to_delete,axis=0,inplace=True)

    return df

def sgm_rthFilter(courseDept, courseNum, input_df):

    # check if course has a DEN section
    for row in range(len(input_df)):
        if ((input_df.loc[row,'CourseDept'] == courseDept) and (input_df.loc[row,'CourseNumber'] == courseNum) and (input_df.loc[row,'Room'] == 'DEN@Viterbi')):
            return True

    return False

#code to populate on-air log
def populate(wks,df):

    #Monday 7am starts at (11,2) [row,column] add 25 for next time frame

    print(type(wks.cell(711,2).value))


    
    # time = wks.cell(11,2).value
    # time_range = dt.strptime(time,'%I:%M %p').time()
    
    # start_time = dt.strptime(str(df.loc[0,'Start']),'%I:%M%p')

    # print(start_time + timedelta(seconds=1800))    

def main():

    # filter excel and remove duplicates
    df = filter_excel()
    filtered = remove_duplicates(df)
    
    # Google API Credentials
    sa = gspread.service_account(filename='service_acct.json')

    # open gsheets
    sh = sa.open("DENTSC Spring 2023 Panopto Posting Log")
    wks = sh.worksheet("summer_list")

    

    # create spreadsheet with final SOC
    set_with_dataframe(wks, filtered)


if __name__ == '__main__':
    main()  