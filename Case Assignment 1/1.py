# Import libraries
from cmath import nan
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

# Due time
start_time = 470
tempDue = list(instance_1['Due Time'])
tempDue.remove(np.nan)
due_time = np.array([])
for i in tempDue:
    # Subtract the due time with start time so that we can count from 0
    converted = i.hour * 60 + i.minute - start_time
    due_time = np.append(due_time,converted)
# print(due_time)

# Splitting timimg
splitting_timing = np.array(instance_1["Splitting Timing"])
splitting_timing = splitting_timing[~np.isnan(splitting_timing)]
# print(splitting_timing)

# Solving

model_1 = gurobipy.Model("model_1")