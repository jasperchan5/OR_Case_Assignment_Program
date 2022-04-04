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
for i_count, instance in enumerate(data):
    # print(instance)
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
    # print("\nProcess type:\n",process_type)

    # Processing time
    tempProcessTime = []
    processing_time = [] # 2 dimensional array
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
    # print("\nProcessing time:\n",processing_time)

    # Splitting timing
    splitting_timing = np.array(instance["Splitting Timing"])
    splitting_timing = splitting_timing[~np.isnan(splitting_timing)]
    # print("\nSplitting timing:\n",splitting_timing)

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
    # print("\nDue time:\n",due_time)

    # Solving
    model_1 = gb.Model("model_1")

    # Define job and machine
    job = [range(1,13),range(1,12),range(1,11)]
    machine = range(2,6)

    # Define x-jk: job j be assigned to machine k
    x = model_1.addVars(job[i_count], machine, vtype = gb.GRB.BINARY,name="x")
    
    # Define y-ijk: job j be done after job i on the same machine k
    y = model_1.addVars(job[i_count], job[i_count], vtype = gb.GRB.BINARY,name="y")

    # Define tj: job j is tardy or not
    t = model_1.addVars(job[i_count], vtype = gb.GRB.BINARY,name="t")
    
    # Define fj: finish time of job j
    f = model_1.addVars(job[i_count], vtype = gb.GRB.INTEGER,name="f")
        
    # Objective function
    model_1.setObjective(
        gb.quicksum(t)
        , gb.GRB.MINIMIZE) 
    
    # Add constraints and name them
    jobLen, machineLen = len(job[i_count])+1, len(machine)+1
    print(jobLen)
    M = 999999
    for j in range(1,jobLen):
        model_1.addConstr(M*t[j] >= f[j] - due_time[j-1])

    # x_ik + x_jk >= 2*(y_ij + y_ji) forall i, j 
    for i in range(1,jobLen):
        for j in range(i+1,jobLen):
            for k in range(2,machineLen):
                model_1.addConstr((x[i,k] + x[j,k] >= 2*(y[i,j]+y[j,i])), "B")


    # x_ik + x_jk - 1 <= y_ij + y_ji forall i, j, k
    for i in range(1, jobLen):
        for j in range(i, jobLen):
            for k in range(2, machineLen):
                model_1.addConstr((x[i,k] - x[j,k] - 1) <= (y[i,j] + y[j,i]), "C")  
    
    # y_ij + y_ji <= 1 forall i, j
    for i in range(1, jobLen):
        for j in range(i, jobLen):
            model_1.addConstr((y[i,j] + y[j,i]) <= 1, "D")
    
    #sigma(x_jk) = 1 forall j
    model_1.addConstr((gb.quicksum(x) == 1), "E")
    
    # f_i + p_j - f_j <= M(1 - y_ij)
    for i in range(1, jobLen):
        for j in range(i, jobLen):
            model_1.addConstr((f[i] + sum(processing_time[j-1]) - f[j]) <= M * (1 - y[i,j]), "F")
    
    # f_j >= p_j + C
    for j in range(1, jobLen):
        model_1.addConstr(f[j] >= (sum(processing_time[j-1]) + start_time), "G")
    
    # x_j1 <= Typej
    # for j in range(1, jobLen):
    #     model_1.addConstr((x[j,1] <= process_type[j-1]), "H")

    model_1.optimize()

    print("Result:")
    # print("\nDistribution:")
    # for xx in x:
    #     print(x[xx].varName,'=',x[xx].X)
    # print("\nDistribution_2:")
    # for yy in y:
    #     print(y[yy].varName,'=',y[yy].X)
    print("\nTardy count:")
    for tt in t:
        print(t[tt].varName,'=',t[tt].X)
    print("\nFinish time:")
    for ff in f:
        print(f[ff].varName,'=',f[ff].X)
    # head of the result table


    print("\nz* =", model_1.ObjVal)    # print objective value