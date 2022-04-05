# Import libraries
import gurobipy as gb
import pandas as pd
import numpy as np
import os
pre = os.path.dirname(os.path.realpath(__file__))
fname = 'OR110-1_case01.xlsx'
path = os.path.join(pre, fname)
def P3_Solve():
    tardy_result = []
    makespan_result = []
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
            # if len(concat) == 1 and concat[0] == 'Boiling':
            #     process_type.append(1)
            # else:
            #     process_type.append(0)
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
        job_count = len(processing_time)
        for i in splitting_timing:
            if i:
                job_count += 1

        # Solving
        model_1 = gb.Model("model_1")
        model_1.setParam('OutputFlag', 0)  # Also dual_subproblem.params.outputflag = 0
        # Define job and machine
        machine = range(1,6)
        # print(job_count)
        jobLen, machineLen = job_count+1, len(machine)+1
        # Define x-jk: job j be assigned to machine k
        x = model_1.addVars(range(1,jobLen), machine, vtype = gb.GRB.BINARY,name="x")
        
        # Define y-ijk: job j be done after job i on the same machine k
        y = model_1.addVars(range(1,jobLen), range(1,jobLen), vtype = gb.GRB.BINARY,name="y")

        # Define tj: job j is tardy or not
        t = model_1.addVars(range(1,jobLen), vtype = gb.GRB.BINARY,name="t")
        
        # Define fj: finish time of job j
        f = model_1.addVars(range(1,jobLen), vtype = gb.GRB.INTEGER,name="f")
            
        # Define idxj: splitting index of job j
        
        # Objective function
        model_1.setObjective(
            gb.quicksum(t)
            , gb.GRB.MINIMIZE) 
        # Due time
        # Define start time: 7:30 a.m.
        start_time = 450
        tempDue = list(instance['Due Time'])
        tempDue.remove(np.nan)
        due_time = []
        for i in tempDue:
            # Subtract the due time with start time so that we can count from 0
            converted = i.hour * 60 + i.minute - start_time
            due_time.append(converted)
        # print("\nDue time:\n",due_time)
        
        # Handle splitting processes
        BIG_DUE = 999999
        new_processing_time = []
        new_process_type = []
        new_due_time = []
        for p_count, split_index in enumerate(splitting_timing):
            # print(p_count,split_index)
            if split_index == 0:
                new_processing_time.append(sum(processing_time[p_count]))
                new_process_type.append(1)
                for t1 in process_type[p_count]:
                    if t1 != "Boiling":
                        new_process_type[-1] = 0
                        break
                new_due_time.append(due_time[p_count])
                continue
            
            new_due_time.append(BIG_DUE)
            new_due_time.append(due_time[p_count])
            temp_pro_1 = processing_time[p_count][0:int(split_index)]
            # print(temp_pro_1)
            temp_pro_2 = processing_time[p_count][int(split_index):]
            new_processing_time.append(sum(temp_pro_1))
            new_processing_time.append(sum(temp_pro_2))
            # the coder is sleepy here
            model_1.addConstr(f[len(new_processing_time)-1] + new_processing_time[-1] - f[len(new_processing_time)] <= 0)
            temp_type_1 = process_type[p_count][0:int(split_index)]
            temp_type_2 = process_type[p_count][int(split_index):]
            new_process_type.append(1)
            for t1 in temp_type_1:
                if t1 != "Boiling":
                    new_process_type[-1] = 0
                    break
            new_process_type.append(1)
            for t2 in temp_type_2:
                if t2 != "Boiling":
                    new_process_type[-1] = 0
                    break
            
                
        # print("\nNew processing time:",new_processing_time)
        # print("\nNew process type:",new_process_type)
            

        
        # Add constraints and name them
        

        # fj - dj <= M*tj
        M = 99999
        for j in range(1,jobLen):
            model_1.addConstr(f[j] - new_due_time[j-1] <= M*t[j], "A")

        # for i in range(1,jobLen):
        #     model_1.addConstr(y[i,i] == 0)

        # x_ik + x_jk >= 2*(y_ij + y_ji) forall i, j 
        # for k in range(2,machineLen):
        #     for i in range(1,jobLen):
        #         for j in range(1,jobLen):
        #             if i != j:
        #                 model_1.addConstr((x[i,k] + x[j,k] >= 2*(y[i,j]+y[j,i])), "B")


        # x_ik + x_jk - 1 <= y_ij + y_ji forall i, j, k
        for k in range(1, machineLen):
            for i in range(1, jobLen):
                for j in range(1, jobLen):
                    if i != j:
                        model_1.addConstr((x[i,k] + x[j,k] - 1) <= (y[i,j] + y[j,i]), "C")  
        
        # y_ij + y_ji <= 1 forall i, j
        for i in range(1, jobLen):
            for j in range(1, jobLen):
                if i != j:
                    model_1.addConstr((y[i,j] + y[j,i]) <= 1, "D")
        
        #sigma(x_jk) = 1 forall j
        for j in range(1,jobLen):
            model_1.addConstr((x[j,1] + x[j,2] + x[j,3] + x[j,4] + x[j,5]) == 1, "E")
        
        # f_i + p_j - f_j <= M(1 - y_ij)
        for i in range(1, jobLen):
            for j in range(1, jobLen):
                if i != j:
                    model_1.addConstr((f[i] + new_processing_time[j-1] - f[j]) <= M*(1-y[i,j]), "F")
        
        # f_j >= p_j + C
        for j in range(1, jobLen):
            model_1.addConstr(f[j] >= (new_processing_time[j-1]), "G")
        
        # x_j1 <= Typej
        for j in range(1, jobLen):
            model_1.addConstr((x[j,1] <= new_process_type[j-1]), "H")

        model_1.optimize()

        # print("Result:")
        # print("\nDistribution:")
        # for xx in x:
        #     if x[xx].X == 1:
        #         print(x[xx].varName,'=',x[xx].X)
        # print("\nDistribution_2:")
        # for yy in y:
        #     if y[yy].X == 1:
        #         print(y[yy].varName,'=',y[yy].X)
        # print("\nTardy count:")
        # for tt in t:
        #     print(t[tt].varName,'=',t[tt].X)
        # print("\nFinish time:")
        # for ff in f:
        #     print(f[ff].varName,'=',f[ff].X)

        # head of the result table

        t_result = model_1.ObjVal
        
        tardy_result.append(t_result)
        # print("\n==========\nz* =", t_result)    # print objective value
        # print("==========")
        w = model_1.addVar(vtype = gb.GRB.INTEGER,name="w")
        
        # Objective function
        model_1.setObjective(w, gb.GRB.MINIMIZE) 
        
        # Add constraints and name them
        
        # sum(tj) <= t
        model_1.addConstr(t.sum() <= t_result, "H")
        
        # w >= fj
        for j in range(1, jobLen):
            model_1.addConstr(w >= f[j], "I")
        
        model_1.optimize()
        w_result = model_1.ObjVal # Minute
        real_w_hour, real_w_min = int(w_result/60) ,int(w_result%60)
        # print("\nz* =", f"{real_w_hour}:{real_w_min}")    # print objective value
        
        makespan_result.append([real_w_hour,real_w_min])
    return tardy_result, makespan_result
        
P3_Solve()