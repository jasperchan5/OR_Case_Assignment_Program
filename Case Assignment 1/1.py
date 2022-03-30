# Import libraries
import gurobipy
import pandas as pd
import numpy as np
import os
pre = os.path.dirname(os.path.realpath(__file__))
fname = 'OR110-1_case01.xlsx'
path = os.path.join(pre, fname)

# Data reading
instance_1, instance_2, instance_3 = pd.read_excel(path,"Instance 1"), pd.read_excel(path,"Instance 2"), pd.read_excel(path,"Instance 3")
# print(instance_1)
# print(instance_2)
# print(instance_3)
# Define start time: 7:30 a.m.

df = pd.DataFrame(instance_1)

# Processing type
tempProcessType = []
process_type = []
processIdx = 1
while df.columns[processIdx] != 'Splitting Timing':
    col_name = df.columns[processIdx]
    tempProcessType.append(instance_1[col_name])
    processIdx += 1
processIdx += 1
for i in range(1,len(tempProcessType[0])):
    concat = []
    for j in range(0,len(tempProcessType)):
        if pd.isna(tempProcessType[j][i]) == False:
            concat.append(tempProcessType[j][i])
    process_type.append(concat)
print("\nProcess type:\n",process_type)

# Processing time
tempProcessTime = []
processing_time = []
while df.columns[processIdx] != 'Due Time':
    col_name = df.columns[processIdx]
    tempProcessTime.append(instance_1[col_name])
    processIdx += 1
for i in range(1,len(tempProcessTime[0])):
    concat = []
    for j in range(0,len(tempProcessTime)):
        if pd.isna(tempProcessTime[j][i]) == False:
            concat.append(tempProcessTime[j][i])
    processing_time.append(concat)
print("\nProcessing time:\n",processing_time)

# Splitting timing
splitting_timing = np.array(instance_1["Splitting Timing"])
splitting_timing = splitting_timing[~np.isnan(splitting_timing)]
print("\nSplitting timing:\n",splitting_timing)

# Due time
start_time = 470
tempDue = list(instance_1['Due Time'])
tempDue.remove(np.nan)
due_time = np.array([])
for i in tempDue:
    # Subtract the due time with start time so that we can count from 0
    converted = i.hour * 60 + i.minute - start_time
    due_time = np.append(due_time,converted)
print("\nDue time:\n",due_time)

# For problem 1 we define finish time as the sum of processing time
finish_time = np.array([])
for i in processing_time:
    finish_time = np.append(finish_time,sum(i))
print("\nFinish time:\n",finish_time)

# Solving

model_1 = gurobipy.Model("model_1")