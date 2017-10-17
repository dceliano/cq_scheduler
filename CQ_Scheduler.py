#Developer - Dom Celiano
#Start Date - 6 Feb 15
#Program to read in 2 files (1 of sitter free periods and 1 of current week's preferences/non-preferences) and create a CQ schedule based on those values

from xlrd import open_workbook
NUM_PREFERENCES = 4
NUM_CANT_SITS = 3
CURRENT_GO = 'F'
WEEK = 'M' #the first day of the week

class Sitter():
    def __init__(self, name, preferences, non_sits):
        self.name = name
        self.preferences = preferences #array
        self.non_sits = non_sits
    def setStatus(self, status):
        #true if LOS and false if not
        self.status = status
    def setFreePeriods(self, free_periods):
        self.free_periods = free_periods
class Day():
    """There will be 1 of these objects for each day"""
    def __init__(self, day):
        self.day = day
        self.six = [0,0,0] #0630 - will be in format: (num preferred, num free, num can't sit)
        self.seven = [0,0,0]
        self.eight = [0,0,0]
        self.nine = [0,0,0]
        self.ten = [0,0,0]
        self.eleven = [0,0,0]
        self.twelve = [0,0,0]
        self.thirteen = [0,0,0]
        self.fourteen = [0,0,0]
        self.fifteen = [0,0,0]
        self.twentytwo = [0,0,0]
        self.twentyfour = [0,0,0]
    def setDayInfo():
        pass
        
def main():
    sitters = [] #array of sitter objects
    loadSitterInfo(sitters)
    
    
    days = [Day('mon'), Day('tue'), Day('wed'), Day('thu'), Day('fri'), Day('sat'), Day('sun')]
    
    for i in range(0, len(sitters)):
        for r in range(0, NUM_PREFERENCES):
            for t in range(0, len(days)):
                if(sitters[i].preferences[r][0:3] == days[t].day):
                    if(sitters[i].preferences[r][3:7] == '0630'):
                        days[t].six[0] += 1 #add one to the preferred
                    elif(sitters[i].preferences[r][3:7] == '0730'):
                        days[t].seven[0] += 1
                    elif(sitters[i].preferences[r][3:7] == '0830'):
                        days[t].eight[0] += 1
                    elif(sitters[i].preferences[r][3:7] == '0930'):
                        days[t].nine[0] += 1
                    elif(sitters[i].preferences[r][3:7] == '1030'):
                        days[t].ten[0] += 1
                    elif(sitters[i].preferences[r][3:7] == '1130'):
                        days[t].eleven[0] += 1
                    elif(sitters[i].preferences[r][3:7] == '1230'):
                        days[t].twelve[0] += 1
                    elif(sitters[i].preferences[r][3:7] == '1330'):
                        days[t].thirteen[0] += 1
                    elif(sitters[i].preferences[r][3:7] == '1430'):
                        days[t].fourteen[0] += 1
                    elif(sitters[i].preferences[r][3:7] == '1530'):
                        days[t].fifteen[0] += 1
                    elif(sitters[i].preferences[r][3:7] == '2200'):
                        days[t].twentytwo[0] += 1
                    elif(sitters[i].preferences[r][3:7] == '2400'):
                        days[t].twentyfour[0] += 1
        print("%s, LOS Status: %s, Free Periods: %s\nPreferences: %s Can't Sit: %s\n" % (sitters[i].name, sitters[i].status, sitters[i].free_periods, sitters[i].preferences, sitters[i].non_sits))
        
    text_file = open("Preferences.txt", "w")
    for i in range(0, len(sitters)):
        text_file.write("%s, LOS Status: %s, Free Periods: %s\nPreferences: %s Can't Sit: %s\n\n" % (sitters[i].name, sitters[i].status, sitters[i].free_periods, sitters[i].preferences, sitters[i].non_sits))
    text_file.close()
 
    #for i in range(0, len(days)):
     #   print("Num who prefer %s at 2200: %d" % (days[i].day, days[i].twentytwo[0]))
        
def loadSitterInfo(sitters):
    wb1 = open_workbook('preferences.xls')
    for s in wb1.sheets():
        for row in range(s.nrows):
            values = [] #an array of strings
            for col in range(s.ncols):
                values.append(s.cell(row,col).value)
            if values[0] != 'Timestamp': #if not the first row
                preferences = [] #inside this array, preferences will be stored as (day)(time) i.e. sat0630 is reveille, fri1130 is lunch
                for i in range(2, NUM_PREFERENCES + 2):
                    if values[i] == 'Monday' or values[i] == 'Monday (M)' or values[i] == 'Monday (T)':
                        preferences.append('mon')
                    elif values[i] == 'Tuesday' or values[i] == 'Tuesday (M)' or values[i] == 'Tuesday (T)':
                        preferences.append('tue')
                    elif values[i] == 'Wednesday' or values[i] == 'Wednesday (M)' or values[i] == 'Wednesday (T)':
                        preferences.append('wed')
                    elif values[i] == 'Thursday' or values[i] == 'Thursday (M)' or values[i] == 'Thursday (T)':
                        preferences.append('thu')
                    elif values[i] == 'Friday' or values[i] == 'Friday (M)' or values[i] == 'Friday (T)':
                        preferences.append('fri')
                    elif values[i] == 'Saturday':
                        preferences.append('sat')
                    else:
                        preferences.append('sun')
                start_value = NUM_PREFERENCES + 2
                print(start_value)
                for i in range(start_value, 2 + (2*NUM_PREFERENCES)):
                    if values[i] == '0630-0730':
                        preferences[i-start_value] = preferences[i-start_value] + '0630'
                    elif values[i] == '0730-0830 (1st)':
                        preferences[i-start_value] = preferences[i-start_value] + '0730'
                    elif values[i] == '0830-0930 (2nd)':
                        preferences[i-start_value] = preferences[i-start_value] + '0830'
                    elif values[i] == '0930-1030 (3rd)':
                        preferences[i-start_value] = preferences[i-start_value] + '0930'
                    elif values[i] == '1030-1130 (4th)':
                        preferences[i-start_value] = preferences[i-start_value] + '1030'
                    elif values[i] == '1130-1230 (Lunch)':
                        preferences[i-start_value] = preferences[i-start_value] + '1130'
                    elif values[i] == '1230-1330 (5th)':
                        preferences[i-start_value] = preferences[i-start_value] + '1230'
                    elif values[i] == '1330-1430 (6th)':
                        preferences[i-start_value] = preferences[i-start_value] + '1330'
                    elif values[i] == '1430-1530 (7th)':
                        preferences[i-start_value] = preferences[i-start_value] + '1430'
                    elif values[i] == '1530-1700':
                        preferences[i-start_value] = preferences[i-start_value] + '1530'
                    elif values[i] == 'Taps':
                        if preferences[i-start_value] == 'fri' or preferences[i-start_value] == 'sat':
                            preferences[i-start_value] = preferences[i-start_value] + '2400'
                        else:
                            preferences[i-start_value] = preferences[i-start_value] + '2200'
                non_sits = []
                for i in range(2 + (2*NUM_PREFERENCES), 2 + (2*NUM_PREFERENCES) + NUM_CANT_SITS):
                    if values[i] == 'Monday':
                        non_sits.append('mon')
                    elif values[i] == 'Tuesday':
                        non_sits.append('tue')
                    elif values[i] == 'Wednesday':
                        non_sits.append('wed')
                    elif values[i] == 'Thursday':
                        non_sits.append('thu')
                    elif values[i] == 'Friday':
                        non_sits.append('fri')
                    elif values[i] == 'Saturday':
                        non_sits.append('sat')
                    elif values[i] == 'Sunday':
                        non_sits.append('sun')
                    else:
                        pass
                start_value = 2 + (2*NUM_PREFERENCES) + NUM_CANT_SITS
                print(start_value)
                for i in range(start_value, 2 + (2*NUM_PREFERENCES) + (2*NUM_CANT_SITS)):
                    if values[i] == '0630-0730':
                        non_sits[i-start_value] = non_sits[i-start_value] + '0630'
                    elif values[i] == '0730-0830 (1st)':
                        non_sits[i-start_value] = non_sits[i-start_value] + '0730'
                    elif values[i] == '0830-0930 (2nd)':
                        non_sits[i-start_value] = non_sits[i-start_value] + '0830'
                    elif values[i] == '0930-1030 (3rd)':
                        non_sits[i-start_value] = non_sits[i-start_value] + '0930'
                    elif values[i] == '1030-1130 (4th)':
                        non_sits[i-start_value] = non_sits[i-start_value] + '1030'
                    elif values[i] == '1130-1230 (Lunch)':
                        non_sits[i-start_value] = non_sits[i-start_value] + '1130'
                    elif values[i] == '1230-1330 (5th)':
                        non_sits[i-start_value] = non_sits[i-start_value] + '1230'
                    elif values[i] == '1330-1430 (6th)':
                        non_sits[i-start_value] = non_sits[i-start_value] + '1330'
                    elif values[i] == '1430-1530 (7th)':
                        non_sits[i-start_value] = non_sits[i-start_value] + '1430'
                    elif values[i] == '1530-1700':
                        non_sits[i-start_value] = non_sits[i-start_value] + '1530'
                    elif values[i] == 'Taps':
                        if non_sits[i-start_value] == 'fri' or non_sits[i-start_value] == 'sat':
                            non_sits[i-start_value] = non_sits[i-start_value] + '2400'
                        else:
                            non_sits[i-start_value] = non_sits[i-start_value] + '2200'
                    else:
                        pass
                sitters.append(Sitter(values[1], preferences, non_sits))
    
    wb2 = open_workbook('freeperiods.xls')
    for s in wb2.sheets():
        for row in range(s.nrows):
            values = [] #an array of strings - the entire row
            for col in range(s.ncols):
                values.append(s.cell(row,col).value)
            if row == 0:
                top_row = values
            if (row != 0 and values[0] == ''): #we have reached the end of the names
                break;
            if(values[0] != ''): #while there is still a value in the cell
                free_periods = []
                for col in range(1, 14):
                    if values[col] == 'F':
                        free_time = getFreeTime(top_row[col])
                        free_periods.append(free_time)
                    elif values[col] == 'G':
                        if values[14] != CURRENT_GO: #check the gym go
                            free_time = getFreeTime(top_row[col])
                            free_periods.append(free_time)
                for sitter in sitters:
                    if sitter.name == values[0]:
                        sitter.setFreePeriods(free_periods)
                        if values[15] == 'x' or values[15] == 'X':
                            sitter.setStatus(True)
                        else:
                            sitter.setStatus(False)
          
def getFreeTime(free_period):
    if free_period == 'M1':
        free_time = 'M0730'
    elif free_period == 'M2':
        free_time = 'M0830'
    elif free_period == 'M3':
        free_time = 'M0930'
    elif free_period == 'M4':
        free_time = 'M1030'
    elif free_period == 'M6':
        free_time = 'M1330'
    elif free_period == 'M7':
        free_time = 'M1430'
    elif free_period == 'T1':
        free_time = 'M0730'
    elif free_period == 'T2':
        free_time = 'M0830'
    elif free_period == 'T3':
        free_time = 'M0930'
    elif free_period == 'T4':
        free_time = 'M1030'
    elif free_period == 'T5':
        free_time = 'M1230'
    elif free_period == 'T6':
        free_time = 'M1330'
    elif free_period == 'T7':
        free_time = 'M1430'
    return free_time
    
    
######## Main Program ########
if __name__ == "__main__":
    main()