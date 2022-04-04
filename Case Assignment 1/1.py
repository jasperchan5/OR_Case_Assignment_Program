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
for instance in len(0,1):
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

    # For problem 1 we define finish time as the sum of processing time
    finish_time = []
    for i in processing_time:
        finish_time.append(sum(i))
    print("\nFinish time:\n",finish_time)


    # Solving
    model_1 = gb.Model("model_1")

    # Define job and machine
    job = range(1,13)
    machine = range(2,6)

    # Define x-jk: job j be assigned to machine k
    x = []
    for j in range(0,len(job)):
        x_2 = []
        for k in range(0,len(machine)):
            x_2.append(model_1.addVar(lb = 0, vtype = gb.GRB.BINARY, name = "x" + str(j+1) + str(k+1)))
        x.append(x_2)

    # Define y-ijk: job j be done after job i on the same machine k
    y = []
    for i in range(0,len(job)):
        for j in range(i,len(job)):
            y_2 = []
            for k in range(0,len(machine)):
                y_2.append(model_1.addVar(lb = 0, vtype = gb.GRB.BINARY, name = "y" + str(i+1) + str(j+1) + str(k+1)))
            y.append(y_2)

    # Define tj: job j is tardy or not
    t = []
    for j in range(0,len(job)):
        t.append(model_1.addVar(lb = 0, vtype = gb.GRB.BINARY, name = "t" + str(j+1)))
        
    # Objective function
    model_1.setObjective(
        gb.quicksum(t[j] for j in range(0,len(job)))
        , gb.GRB.MINIMIZE) 

    # Add constraints and name them
    M = 999999
    model_1.addConstrs((finish_time[j] - due_time[j] <= M*t[j] for j in range(0,len(job))), "tardyOrNot")

    # x_ik + x_jk >= 2*(y_ij + y_ji) forall i, j 
    for i in range(0,len(job)):
        for j in range(i,len(job)):
            for k in range(0,len(machine)):   
                model_1.addConstr((x[i][k] + x[j][k] >= 2*(y[i][j]+y[j][i])), "B")


    # x_ik + x_jk - 1 <= y_ij + y_ji forall i, j, k
    for i in range(0, len(job)):
        for j in range(i, len(job)):
            for k in range(0, len(machine)):
                model_1.addConstr((x[i][k] - x[j][k] - 1) <= (y[i][j] + y[j][i]), "C")  
    
    # y_ij + y_ji <= 1 forall i, j
    for i in range(0, job):
        for j in range(i, job):
            model_1.addConstr((y[i][j] + y[j][i]) <= 1, "D")
    
    #sigma(x_jk) = 1 forall j
    for k in range(0, machine):
        for j in range(0, job):
            model_1.addConstr((gb.quicksum(x[j][k]) == 1), "E")
    
    # f_i + p_j - f_j <= M(1 - y_ij)
    for i in range(0, job):
        for j in range(i, job):
            model_1.addConstr((finish_time[i] + processing_time[j] - finish_time[j]) <= M * (1 - y[i][j]), "F")
    
    # f_j >= p_j + C
    for j in range(0, job):
        model_1.addConstr(finish_time[j] >= (processing_time[j] + start_time), "G")
    
    # x_j1 <= Typej
    for j in range(0, job):
        model_1.addConstr((x[j][0] <= process_type[j]), "H")

    model_1.optimize()

    print("Result:")

    for i in t:
        print(i.varName, '=', i.x)
    # head of the result table


    print("z* =", model_1.objVal)    # print objective value