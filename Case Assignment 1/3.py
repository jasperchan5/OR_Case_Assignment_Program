# Import libraries
import gurobipy as gb
import pandas as pd
import numpy as np
import os
pre = os.path.dirname(os.path.realpath(__file__))
fname = 'OR110-1_case01.xlsx'
path = os.path.join(pre, fname)

# Data reading
data = [pd.read_excel(path,"Instance 1"), pd.read_excel(path,"Instance 2"), pd.read_excel(path,"Instance 3")]
for instance in data:
    print(instance)
    df = pd.DataFrame(instance)

    # Processing type
    tempProcessType = []
    process_type = []
    processIdx = 1
    while df.columns[processIdx] != 'Splitting Timing':
        col_name = df.columns[processIdx]
        tempProcessType.append(instance[col_name])
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
        tempProcessTime.append(instance[col_name])
        processIdx += 1
    for i in range(1,len(tempProcessTime[0])):
        concat = []
        for j in range(0,len(tempProcessTime)):
            if pd.isna(tempProcessTime[j][i]) == False:
                concat.append(tempProcessTime[j][i]*60)
        processing_time.append(concat)
    print("\nProcessing time:\n",processing_time)

    # Splitting timing
    splitting_timing = np.array(instance["Splitting Timing"])
    splitting_timing = splitting_timing[~np.isnan(splitting_timing)]
    print("\nSplitting timing:\n",splitting_timing)

    # Due time
    # Define start time: 7:30 a.m.
    start_time = 470
    tempDue = list(instance['Due Time'])
    tempDue.remove(np.nan)
    due_time = []
    for i in tempDue:
        # Subtract the due time with start time so that we can count from 0
        converted = i.hour * 60 + i.minute - start_time
        due_time.append(converted)
    print("\nDue time:\n",due_time)
