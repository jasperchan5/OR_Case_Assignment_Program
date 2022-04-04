from p_1 import P1_Solve
from p_2 import P2_Solve
from p_3 import P3_Solve
import pandas as pd
t1, r1 = P1_Solve()
t2, r2 = P2_Solve()
t3, r3 = P3_Solve()
print("Tardy P1:",t1)
print("Tardy P2:",t2)
print("Tardy P3:",t3)
print("Makespan P1",r1)
print("Makespan P2",r2)
print("Makespan P3",r3)
