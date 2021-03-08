# Rota power flow
import math
import numpy as np
import csv  # this will be used to read a csv file for the load function
from matplotlib import pyplot as plt
from pprint import pprint

# variables that frame graph function
timeDays = 1.0    # user provides the number of days for the scenerio to run
timeStep = 2.0    # minutes
overallMin = 0.0  # overall minutes in the scenerio
dataPoints = 0.0  # overall number of data points
time = 0.0        # current time in scenerio

# constants
hoursInDay = 24.0  # number of hours in a day
minutesInHour = 60.0  # number of minutes in an hour
minutesInDay = 1440.0  # number of minutes in a day
batteryDischargeTime = 30.0  # 30 mins to drian the batteries


def timeInitializer(timeDays, hoursInDay, minutesInHour, timeStep):
    # This function establishes the total amount of time in minutes for the scenerio
    # timeDuration is the number of days multiplied by number of hours then converted to min
    timeDuration = (timeDays * hoursInDay)*minutesInHour

    return timeDuration


def timeSim(time, timeTable, timeHoursTable):
    # This function calculates values for variables timeHours and time, then appends them to timeTable and timeHoursTable.
    # This function is used in the for loop on line 230 to generate and store the values for time and timehours

    # appends the current value of time to timeTable
    timeTable.append(time)
    # calculates the value for time hours
    timeHours = time/60.0
    # appends the current timeHours to timeHoursTable
    timeHoursTable.append(timeHours)
    return timeHours


def readLoadData(loadDataFile):
    # This function reads data from a .csv file which would contain the load data for the site of the microgrid

    # opens a file and assigns it to a variable name(loadFile)
    with open(loadDataFile, 'r') as loadFile:
        # using the csv.reader() method to read from the load file, a varable named loadFileReader is used to reference the method.
        loadFileReader = csv.reader(loadFile)
        # next is used to skip the 1st line in the file which is the title or header line of the spreadsheet
        next(loadFileReader)
        # timeList is a local list used to store the time data from the .csv file
        timeList = []
        # loadList is a local list used to store the load data from the .csv file
        loadList = []
        # for loop used to read each line of the .csv file (loadFileReader)
        for line in loadFileReader:
            # appends the first item(also casted as a float) in column 0 (time) to the local list timeList
            timeList.append(float(line[0]))
            # appends the first item(also casted as a float) in column 1 (load) to the local list loadList
            loadList.append(float(line[1]))
    return [timeList, loadList]


def pvSim(pvSimTable, time, pvIn, pvf, pvw):
    # This function calculates the data for the list pvSimTable[]

    # the is the formula that calculates to value for PV(the power that the solar panels produce)
    # =MAX(0,-PV_in*SIN(PVF*PI()*A3/1440+PVW))
    pv = (pvIn * (-1.0)) * math.sin(pvf * math.pi * (time/minutesInDay) + pvw)

    if pv < 0:
        # else pv = 0 because pv cannot have a negative value.
        pv = 0
        # appends the pv value to pvSimTable

    pvSimTable.append(pv)

    return pv


def loadSim(time, loadSimTable, ppk, loadData):
    # this function takes load data and appends it to the list loadSimTable
    # parameters are time(current time in scenerio), loadSimTable, ppk, loadData is the data from the previous readLoadData function

    # for loop that uses range of the length of the table loadData[0](this column was picked because all columns in the .csv file are the same length)
    # and loops around the list while appending the data
    for i in range(len(loadData[0])):
        # if current scenerio time is less then or equal to the loadData[0][i] (current iteration of the time column)
        if time <= loadData[0][i]:
            # load equals to the table loadData[1][i] (current iteration of the load column)
            load = loadData[1][i]*ppk
            # appends the value of load to the loadSimTable
            loadSimTable.append(load)
            return load
    # after the for loop runs, it will not account for the last piece of load data, this new definition of load accounts for that
    load = loadData[1][-1]*ppk
    # appends the last piece of load data in the loadData[1] list
    loadSimTable.append(load)
    return load


def gridSim(time, gridOffAt, load, pv, genSet, bess, gridSimTable):
    # gridSim gathers the values for the list gridSimTable[], it determines the output power for the grid based on when the grid fault occurs

    # if statement that says if time is less than gridOffAt, the grid value will will be produced by the following equation
    if time < gridOffAt:
        # =-(F3+C3+E3-D3), equation for grid from the excel sheet

        grid = -1*(genSet + pv + bess - load)

    # else there is a grid fault and the grid is not producing power therefore its value is 0
    else:
        grid = 0

    # appends the grid value to gridSimTable
    gridSimTable.append(grid)
    return grid


def genSetSim(time, pv, load, grid, genSet, gridOffAt, genSetSimTable, slackBus):
    # This function calculates the values for the generator output based on either the gridOffAt variable or the slackBus variable

    # if time is less than the gridOfAtt value, then the grid is on and the generators are not needed so its output is 0
    if time < gridOffAt:
        genSet = 0

    # if the slackBus value is 2, then the generator is on, and genSet is calculated by the following equation
    else:
        if slackBus == 2:
            genSet = -1*(pv + grid - load)
        else:
            genSet = 0
    # appends the value from genSet to the table genSetSimTable
    genSetSimTable.append(genSet)
    return genSet


def bessSim(genSet, pv, load, grid, slackBus, soc_prev, soc_u, roc, bessSimTable):
    # this function calculates the value of bess for which the values soc and grid are dependent on
    # the parameter soc_prev is used to provide a value for soc in order for bessSim() to run due to both functions being dependent on eachother

    # if slackBus is equal to 0, then bess equals zero because if the grid is on then there is no need to discharge the bateries
    if slackBus == 0:
        bess = 0
    # If slackBus is equal to 1 then the system is in battery mode and the value for bess is calculated with the following function
    elif slackBus == 1:

        bess = -1*(genSet + pv + grid - load)

    # else slackBus is at a value of 2 and generator is on
    else:
        # if the value for soc_prev is less than the value for soc_u bess is calculated from the following equation
        if soc_prev < soc_u:
            bess = -1*(roc)
        # else the battery is fully charged and therefore the bess value is 0
        else:
            bess = 0
    # appends the value of bess to the table bessSimTable
    bessSimTable.append(bess)

    return bess


def socSim(bess, bessPu, socSimTable):
    # This function calculates the value for soc (state of charge)

    # if the length of the socSimTable is equal to 0, the equation for soc should be calculated by the following
    # equation where it adds a value of 1 instead of the last value from socSimTable
    if len(socSimTable) == 0:
        # =-E3*2*60/3600/BESS_pu+1
        soc = -1*((bess*2.0*60.0)/3600.0) / bessPu + 1
    # else it is assumed that there are values in the list socSimTable and we will use those in the equation for soc
    else:

        soc = -1*bess*2.0*60.0/3600.0 / bessPu + socSimTable[-1]
    # appends the value of soc to the socSimTable
    socSimTable.append(soc)
    return soc


def slackBusSim(time, genSet, genSetMin, dod, gridOffAt, genSetActivationDelay, soc, soc_u, slackBus, slackBusSimTable):
    # This function uses if statements to determine the condition of the microgrid in terms of a slackBus value
    # genSetActivationDelay is == 30 due to the developers being briefed that the system had 30 mins of battery once the grid goes down

    # if time is less than gridOffAt, slackBus is equal to 0
    if time < gridOffAt:
        slackBus = 0
    # if time is less than gridOffAt time plus genSetActivationDelay(30 mins), that would trigger slackBus to be 1.
    # The previous if statement would not apply during this iteration of the loop because time would be past the gridoffAt value
    elif time < gridOffAt + genSetActivationDelay:
        slackBus = 1
    else:

        # slackBus is equal to 2, and soc(state of charge) is greater than the value of soc_u(.99) and genSet is less than genSetMin (this would cause wet stacking)
        # change slackBus to 1 because the battery is fully charged and can be used again as the primary source of power for the system
        if slackBus == 2 and soc > soc_u and genSet < genSetMin:
            slackBus = 1
        # if slackBus equals 1 (we are running on batteries) and soc is less than depth of discharge(dod) change slackbus to 2
        else:
            if slackBus == 1 and soc < dod:
                slackBus = 2
            else:
                pass
    # appends the value of slackBus to slackBusSimTable
    slackBusSimTable.append(slackBus)
    return slackBus


def rotaSim(ppk, genRat, pvIn, bessPu, roc, pvw, pvf, dod, soc_u, genSetMin, days, gridOffAt, genSetActivationDelay, loadDataFile):
    # This function serves as our main function for the program. It does multiple tasks.
    # Task 1- uses a for loop to generate the time step sequence (clock) for the scenerio
    # Task 2- calls functions within the for loop to generate values and place those values in the lists below
    # Task 3- returns the lists which hold all the generated data for the scenerio

    # these are the inputed values from the GUI that the user would input
    '''
    Args:
        ppk (double): peakLoad
    genRat = 1.25  # Genset rating
    pvIn = 1.04  # power in from solar panels
    bessPu = .6  # battery rating
    roc = .23  # BESS rate of charge
    pvw = .3925  # Pv width
    pvf = 2.4  # Pv Frequency
    dod = .5  # depth of discharge
    soc_u = .99  # BESS upper limit
    genSetMin = .6  # Genset Min
    days = 1.0
    gridOffAt = 60.0  # this will toggle slackBus to slackBus = 1
    genSetActivationDelay = 30

    '''
    # Tables that store all the returned values from the functions in the program
    pvSimTable = []        # a list of all PV values
    loadSimTable = []      # a list of all laod values
    gridSimTable = []      # a list of all grid values
    genSetSimTable = []    # a list of all genSet values
    bessSimTable = []      # a list of all bess values
    socSimTable = []       # a lst of all SOC values
    slackBusSimTable = []  # a list of all slackBus values
    timeTable = []         # a list of all the time values
    timeHoursTable = []    # a list of all the timeHours values

    loadData = readLoadData(loadDataFile)

    # prints time and timeHours as a heading for the table of values used for developer testing, will be removed in final draft
    print("time timehours pv load bess genSet grid soc slackBus")
    # grabs timeDuration from the function timeInitializer (line 50) for local use
    timeDuration = timeInitializer(
        timeDays, hoursInDay, minutesInHour, timeStep)

    # calls the function loadTriggers to gain the time values for which the load should change
    # this will be deleted in the final version of the client program where the user will be able to upload a .csv file containing load profile
    # loadTriggers(timeDuration)

    # slackBus is set to zero to provide an initial value so the function slackBus() can run
    slackBus = 0
    # provides an initial value for genSet so slackBusSim() can run
    genSet = 0
    # provides an initial value for soc so slackBusSim() can run
    soc = 0
    # provides an initial value for bess so socSim() can run
    bess = 0
    # provides an initial value for grid so bessSim() can run
    grid = 0

    # for loop to perform timestep operations and appends the time to the timeTable list as well as the timeHours list
    time = 0.0
    while time <= timeDuration:
        time = time + timeStep
        # creates variable timeHours which stores the value from timeSim()
        timeHours = timeSim(time, timeTable, timeHoursTable)

        # creates variable slackBus which stores value from slackBusSim()
        slackBus = slackBusSim(time, genSet, genSetMin, dod, gridOffAt, genSetActivationDelay,
                               soc, soc_u, slackBus, slackBusSimTable)

        # creates variable load which stores value from loadSim()
        load = loadSim(time, loadSimTable, ppk, loadData)

        # creates variable pv which stores value for pvSim()
        pv = pvSim(pvSimTable, time, pvIn, pvf, pvw)

        # provides a value for soc needed to run bessSim()
        soc_prev = 0
        # if the length of the socSimTable has atleast one value in it
        # set the variable soc_prev to the last value in the list so bessSim() can run
        if len(socSimTable) != 0:
            soc_prev = socSimTable[-1]
         # create a variable grid that stores the return value of gridSim()
        grid = gridSim(time, gridOffAt, load, pv, genSet, bess, gridSimTable)

        # create a variable genSet that stores the return value of genSetSim()
        genSet = genSetSim(time, pv, load, grid, genSet,
                           gridOffAt, genSetSimTable, slackBus)

        # create variable bess that stores the return value from besSim()
        bess = bessSim(genSet, pv, load, grid, slackBus,
                       soc_prev, soc_u, roc, bessSimTable)
        # create a variable soc that stores the return value of socSim()
        soc = socSim(bess, bessPu, socSimTable)

        # prints the current values of time, timeHours, pv, load, slackBus for troubleshooting purposes for the developer to test values, will be deleted in final draft
        print(time, '%.1f' % timeHours, pv, load,
              bess, genSet, grid, soc, slackBus)
    # returns the lists that were generated during the function call
    return [timeTable, timeHoursTable, pvSimTable, loadSimTable, bessSimTable, genSetSimTable, gridSimTable, socSimTable, slackBusSimTable]

# used for developer visual purposes, will be deleted upon final draft


def rotaGraph(tables, title):
    fig, axs = plt.subplots(2)
    # rotaGraph function contains the scripts to generate a graph using pyploy from matplotlib
    axs[0].plot(tables[1], tables[3])
    axs[0].plot(tables[1], tables[2])
    axs[0].plot(tables[1], tables[6])
    axs[0].plot(tables[1], tables[5])
    axs[0].plot(tables[1], tables[7])
    plt.title(title)
    # axs[0].xlabel('Time in Minutes')
    # axs[0].ylabel('axs[0] in MW')
    axs[0].legend(["LOAD", "SOLAR (PV)", "grid", "genSet", "SOC"])

    axs[1].plot(tables[1], tables[4])
    axs[1].plot(tables[1], tables[8])
    axs[1].legend(["bess", "slackBus"])

    plt.show()


# writes the data from rotaSim() into a .csv file, can be used as a member method in the front end program if client wants to see the numerical data
# the parameters are tables= output tables, headers will be a list of the headers in the file, and fileName is what you want to name the output file
def writeCsv(tables, headers, fileName):
    # this part of the functiong deals exclusively with writing the headers of the file( the column names )
    f = open(fileName, "w")
    cols = len(headers)
    for j in range(cols):
        header = headers[j]
        f.write(header)
        if j != cols - 1:  # if j is not the last item in the list add a comma
            f.write(", ")
    f.write('\n')  # adds new line

    rows = len(tables[0])  # in this case i is the row
    cols = len(tables)    # j is the column
    for i in range(rows):
        for j in range(cols):
            f.write(format(tables[j][i], '.60g'))
            if j != cols - 1:
                f.write(", ")
        f.write('\n')
    f.close()

   # the below commented out section is commented out because i am calling the function in another file to demonstrate the ability to use this program as a library that can be imported into another python program

    # rotaGraph(rotaSim(1,  1.25,  1.04, .6, .23, .3925, 2.4, .5, .99, .6, 1, 60, 30),
    #          "Rota powerflow - Case 1 Nighttime grid fault, then continue with Feeder 3 only")  # calls the rotaSim function
    # rotaGraph()  # calls the rotaGraph function
