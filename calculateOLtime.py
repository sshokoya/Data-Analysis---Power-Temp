# calculate OL time when FET temperature and power varies over period
# written august 23, 2022

import csv
from turtle import shape
import pandas as pd
import datetime
import math
import numpy as np

unit = input("Please enter unit1 or unit6: ")
if unit == "unit1" or unit == "unit 1" or unit == "Unit1" or unit == "Unit 1":
    power_file = r"C:\Users\SESA603713\Documents\powerOutput M1 test2.xlsx" #dewetron excel export
    power_dataM1 = pd.read_excel(power_file, sheet_name="Data1", usecols=(0, 4), skiprows=2) #Time(s) and M1_AC1 (Grid)/P_L1_H1; RMS
    VacM1 = pd.read_excel(power_file, sheet_name="Data1", usecols=(0,1), skiprows=2) # Time(s) and M1_AC1 (Grid) Vac; RMS
if unit == "unit6" or unit == "unit 6" or unit == "Unit6" or unit == "Unit 6":
    power_file = r"C:\Users\SESA603713\Documents\powerOutput S3 test2.xlsx" #dewetron excel export 
    power_dataM1 = pd.read_excel(power_file, sheet_name="Data1", usecols=(0, 6), skiprows=2) #Time(s) and S3_AC1 (Grid)/P_L1_H1; RMS
    VacM1 = pd.read_excel(power_file, sheet_name="Data1", usecols=(0,2), skiprows=2) # Time(s) and S3_AC1 (Grid) Vac; RMS
file_address = r"C:\Users\SESA603713\Documents\OL temp data 08-24.xlsx" #temp data template

#data = pd.read_excel(file_address, usecols=(0,3,4,7,8,11,12,15,16,19,20,23), skiprows=2)
unit1_temp = pd.read_excel(file_address, usecols=(0,3),skiprows=2) 
unit2_temp = pd.read_excel(file_address, usecols=(4,7), skiprows=2)
unit3_temp = pd.read_excel(file_address, usecols=(8,11), skiprows=2)
unit4_temp = pd.read_excel(file_address, usecols=(12,15), skiprows=2)
unit5_temp = pd.read_excel(file_address, usecols=(16,19), skiprows=2)
unit6_temp = pd.read_excel(file_address, usecols=(20,23), skiprows=2)
temp_array1 = np.array(unit1_temp)
temp_array2 = np.array(unit2_temp)
temp_array3 = np.array(unit3_temp)
temp_array4 = np.array(unit4_temp)
temp_array5 = np.array(unit5_temp)
temp_array6 = np.array(unit6_temp)
power_array1 = np.array(power_dataM1)
VacM1np = np.array(VacM1)

P_o = 6800.00 #Watts
SurgePower = 7350.00
V_o = 230.00 #Volts
Ir = SurgePower/V_o # RatedOutputCurrent
Freq = 50.00 #ouput frequency Hz
Occ = 2.00; #overloadChecks per cycle
Ocf = Occ * Freq
Sr = 2.00 #surgeRatio
MaxSurgeTime = 100.00 #seconds

#temp derating factors
tempDerate40 = 1
tempDerate4060 = 13/(2**4)
tempDerate6075 = 10/16
tempDerate75 = 7/16

sumTime = 0.00
sumTime2 = 0.00
sumTime3 = 0.00
sumTime4 = 0.00

if unit == "unit1" or unit == "unit 1" or unit == "Unit1" or unit == "Unit 1":
    for i in range(1, len(temp_array1)-1): #start from second item in array not first due to i-1 calc
        #calculate time spent <40C
        if temp_array1[i-1][1] < 40.00: 
            sumTime = sumTime + (temp_array1[i][0] - temp_array1[i-1][0]) 
        #calculate time spent 40-60C
        if 40.00 <= temp_array1[i][1] <= 60.00: 
            sumTime2 = sumTime2 + (temp_array1[i][0] - temp_array1[i-1][0])  # sum the difference
        #calculate time spent 60-75C
        if 60.00 < temp_array1[i][1] <= 75.00: 
            sumTime3 = sumTime3 + (temp_array1[i][0] - temp_array1[i-1][0])  # sum the difference
        #calculate time spent >75C
        if temp_array1[i][1] > 75.00: 
            sumTime4 = sumTime4 + (temp_array1[i+1][0] - temp_array1[i][0])  # sum the difference
if unit == "unit6" or unit == "unit 6" or unit == "Unit6" or unit == "Unit 6":
    for i in range(1, len(temp_array6)-1): #start from second item in array not first due to i-1 calc
        #calculate time spent <40C
        if temp_array6[i-1][1] < 40.00: 
            sumTime = sumTime + (temp_array6[i][0] - temp_array6[i-1][0]) 
        #calculate time spent 40-60C
        if 40.00 <= temp_array6[i][1] <= 60.00: 
            sumTime2 = sumTime2 + (temp_array6[i][0] - temp_array6[i-1][0])  # sum the difference
        #calculate time spent 60-75C
        if 60.00 < temp_array6[i][1] <= 75.00: 
            sumTime3 = sumTime3 + (temp_array6[i][0] - temp_array6[i-1][0])  # sum the difference
        #calculate time spent >75C
        if temp_array6[i][1] > 75.00: 
            sumTime4 = sumTime4 + (temp_array6[i+1][0] - temp_array6[i][0])  # sum the difference


# get % time spent in each temp range
totalTime = sumTime + sumTime2 + sumTime3 + sumTime4
timeratio40 = sumTime/totalTime
timeratio4060 = sumTime2/totalTime
timeratio6075 = sumTime3/totalTime
timeratio75 = sumTime4/totalTime
print("The total test time is", totalTime)
print("The time sum for <40C is", sumTime)
print("The time sum for 40-60C is", sumTime2)
print("The time sum for 61-75C is", sumTime3)
print("The time sum for >75C is", sumTime4)

# import power output data
for x in range(len(VacM1np)):
        if VacM1np[x][1] > 205.00: 
            timetoStart = VacM1np[x][0] #seconds till unit starts outputting 230VAC in data, i.e. start of test
            break

powerOutput = [] #empty regular list
if unit == "unit1" or unit == "unit 1" or unit == "Unit1" or unit == "Unit 1":
    for y in range(len(temp_array1)):
        #calculate power array that corresponds to temperature readings 
        for x in range(len(power_array1)): #start from first item in array
            if math.isclose((power_array1[x][0] - timetoStart), temp_array1[y][0], rel_tol=1e-6): #if the time of test matches for both arrays use that power ouput value from unit in new array to calcualte OL time - subtract the time till start of invert mode (X seconds)
                arr = np.array([power_array1[x][0]-timetoStart,power_array1[x][1] ]) #input the entire row of data (2cols)
                powerOutput.append(arr)
                break
if unit == "unit6" or unit == "unit 6" or unit == "Unit6" or unit == "Unit 6":
    for y in range(len(temp_array6)):
        #calculate power array that corresponds to temperature readings 
        for x in range(len(power_array1)): #start from first item in array
            #use math.isclose to get rid of floating point errors (loop was missing some values due to == comparator)
            if math.isclose((power_array1[x][0] - timetoStart), temp_array6[y][0], rel_tol=1e-6): #if the time of test matches for both arrays use that power output value from unit in new array to calcualte OL time - subtract the time till start of invert mode (X seconds)
                arr = np.array([power_array1[x][0]-timetoStart,power_array1[x][1] ]) #input the entire row of data (2cols)
                powerOutput.append(arr)
                break

powerOutputnp = np.array(powerOutput) #transformed into a numpy array

# calculations
eMax = ((Sr*P_o - SurgePower)/V_o)**2 * MaxSurgeTime * Ocf #hardcoded - does not change
a = np.full((1, len(powerOutputnp)), 1)
b = np.full((1, len(powerOutputnp)), V_o)
V_o_np = np.vstack((a,b)) #watch for brackets - now have 2xsomethign matrix
loadCurrentnp = powerOutputnp/(V_o_np.T) #everything in the first col of A is divided by the first col of X, and same with 2nd columns
samplesToTrip = []
timeToTrip = []
minToTip = []
for row in range(len(powerOutputnp)):
    sampleSurgeEnp = np.array([(loadCurrentnp[row][1] - Ir)**2  * np.sign(loadCurrentnp[row][1]-Ir)]) #A^2
    samplesToTrip.append(eMax/sampleSurgeEnp) # number of samples to trip
    timeToTrip.append(samplesToTrip[row]/Ocf) #seconds
    minToTip.append(timeToTrip[row]/60) #minutes

samplesToTripnp = np.array(samplesToTrip)
timeToTripnp = np.array(timeToTrip)
finalMatrix = np.append(powerOutputnp, timeToTripnp, axis=1)
if unit == "unit1" or unit == "unit 1" or unit == "Unit1" or unit == "Unit 1":
    finalMatrix = np.append(finalMatrix, temp_array1, axis=1)
if unit == "unit6" or unit == "unit 6" or unit == "Unit6" or unit == "Unit 6": 
    finalMatrix = np.append(finalMatrix, temp_array6, axis=1)
np.set_printoptions(suppress=True, #format final matrix in terminal correctly
   formatter={'float_kind':'{:16.3f}'.format}, linewidth=130)
print(finalMatrix)

#sum total time to OL trip by getting sum then avg Power for each Temp range, then sum each Ol_time for each individual temp range - not as accurate
k1 = 0
k2 = 0
k3 = 0
k4 = 0
sumPower1 = 0.00
sumPower2 = 0.00
sumPower3 = 0.00
sumPower4 = 0.00

for i in range(1, len(finalMatrix)): #skip zero time readings
    if finalMatrix[i,4] < 40.00 and finalMatrix[i][1] > 4000:
        sumPower1 += finalMatrix[i,1]
        k1 += 1 
    if 40.00 <= finalMatrix[i,4] <= 60.00 and finalMatrix[i][1] > 4000:
        sumPower2 += finalMatrix[i,1]
        k2 += 1 
    if 60.00 < finalMatrix[i,4] <= 75.00 and finalMatrix[i][1] > 4000:
        sumPower3 += finalMatrix[i,1]
        k3 += 1 
    if finalMatrix[i,4] > 75.00 and finalMatrix[i][1] > 4000:
        sumPower4 += finalMatrix[i,1]
        k4 += 1 

avgPower40 = sumPower1/(k1)
print("The average power for <40C is", avgPower40)
loadCurrent = avgPower40/(V_o) #constant power & current
sampleSurgeE1 = (loadCurrent - Ir)**2  * np.sign(loadCurrent-Ir) #A^2
samplesToTrip1 = (eMax/sampleSurgeE1) # number of samples to trip
timeToTrip40 = samplesToTrip1/Ocf #seconds

avgPower4060 = sumPower2/(k2)
print("The average power for 40-60C is", avgPower4060)
loadCurrent = avgPower4060/(V_o) #constant power & current
sampleSurgeE2 = (loadCurrent - Ir)**2  * np.sign(loadCurrent-Ir) #A^2
samplesToTrip2 = (eMax/sampleSurgeE2) # number of samples to trip
timeToTrip4060 = samplesToTrip2/Ocf #seconds

avgPower6075 = sumPower3/(k3)
print("The average power for 60-75C is", avgPower6075)
loadCurrent = avgPower6075/(V_o) #constant power & current
sampleSurgeE3 = (loadCurrent - Ir)**2  * np.sign(loadCurrent-Ir) #A^2
samplesToTrip3 = (eMax/sampleSurgeE3) # number of samples to trip
timeToTrip6075 = samplesToTrip3/Ocf #seconds

avgPower75 = sumPower4/(k4)
print("The average power for >75C is", avgPower75)
loadCurrent = avgPower75/(V_o) #constant power & current
sampleSurgeE4 = (loadCurrent - Ir)**2  * np.sign(loadCurrent-Ir) #A^2
samplesToTrip4 = (eMax/sampleSurgeE4) # number of samples to trip
timeToTrip75 = samplesToTrip4/Ocf #seconds

OL_time = (timeToTrip40*tempDerate40*timeratio40) + (timeToTrip4060*tempDerate4060*timeratio4060) + (timeToTrip6075*tempDerate6075*timeratio6075) + (timeToTrip75*tempDerate75*timeratio75)
#print("The total overload time could be", OL_time/(60), "minutes, with average powers stated above." )

#method 2: add to OL accumulator for each power output and corresponding trip time from temp data - not as accurate
OL_acc = 0.00
counter = 0
for i in range(1, len(finalMatrix)): #skip zero time readings
    tripTime = finalMatrix[i][2] #time to trip column
        #calculate time spent <40C
    if finalMatrix[i][4] < 40.00 and finalMatrix[i][1] > 4000: #temp less than 40C and powerOutput > 4000W
        OL_acc += (tripTime *  timeratio40 * tempDerate40)
        counter += 1
    #calculate time spent 40-60C
    if 40.00 <= finalMatrix[i][4] <= 60.00 and finalMatrix[i][1] > 4000: 
        OL_acc += (tripTime *  timeratio4060 * tempDerate4060)
        counter += 1
    #calculate time spent 60-75C
    if 60.00 < finalMatrix[i][4] <= 75.00 and finalMatrix[i][1] > 4000: 
        OL_acc += (tripTime *  timeratio6075 * tempDerate6075)
        counter += 1
    #calculate time spent >75C
    if finalMatrix[i][4] > 75.00 and finalMatrix[i][1] > 4000: 
        OL_acc += (tripTime *  timeratio75 * tempDerate75)
        counter += 1

# print("The total overload time could be", OL_acc/(counter*60), "minutes" )
    
# alternative method 3: sum total power over full OL period then get average over period, and calc time to trip based on that (incl temp deration)
# most accurate method thus far 
powerSum = 0.00
samples = 0
for i in range(1, len(finalMatrix)): #skip zero time readings
    if finalMatrix[i][1] > 4000: # remove outliers from average
        powerSum += finalMatrix[i][1]
        samples += 1
avgPower = powerSum/samples
OL_acc = 0.00

loadCurrent = avgPower/(V_o) #constant power & current
sampleSurgeE2 = (loadCurrent - Ir)**2  * np.sign(loadCurrent-Ir) #A^2
samplesToTrip2 = (eMax/sampleSurgeE2) # number of samples to trip
timeToTrip2 = samplesToTrip2/Ocf #seconds
minToTip = timeToTrip2/60 #min

#calculate time to trip <40C
OL_acc += (timeToTrip2 * timeratio40 * tempDerate40)
#calculate time to trip 40-60C
OL_acc += (timeToTrip2 * timeratio4060 * tempDerate4060)
#calculate time to trip 60-75C
OL_acc += (timeToTrip2 * timeratio6075 * tempDerate6075)
#calculate time to trip >75C
OL_acc += (timeToTrip2 * timeratio75 * tempDerate75)

print("The total overload time should be", OL_acc/(60), "minutes, where the average output power over the whole test is", round(avgPower), "W")


