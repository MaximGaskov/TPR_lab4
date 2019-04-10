import pandas as pd

MU_ARR =  (0, 1, 1, 0)

df_input_1 = pd.read_excel("input.xlsx", parse_cols=list(range(1,13)), nrows=3, skiprows=1)
print(df_input_1)

df_input_2 = pd.read_excel("input.xlsx", parse_cols=list(range(1,13)), nrows=3, skiprows=6)
print(df_input_2)

inputList_E1_K1 = []
inputList_E1_K2 = []
inputList_E1_K3 = []

inputList_E2_K1 = []
inputList_E2_K2 = []
inputList_E2_K3 = []

for row in df_input_1.itertuples():
    inputList_E1_K1.append(row[1:5])
    inputList_E1_K2.append(row[5:9])
    inputList_E1_K3.append(row[9:13])

for row in df_input_2.itertuples():
    inputList_E2_K1.append(row[1:5])
    inputList_E2_K2.append(row[5:9])
    inputList_E2_K3.append(row[9:13])



EXP_ESTIM_PAIRS = ((inputList_E1_K1[0], inputList_E2_K1[0]), (inputList_E1_K2[0], inputList_E2_K2[0]), (inputList_E1_K3[0], inputList_E2_K3[0]), # first obj
                   (inputList_E1_K1[1], inputList_E2_K1[1]), (inputList_E1_K2[1], inputList_E2_K2[1]), (inputList_E1_K3[1], inputList_E2_K3[1]), # second obj
                   (inputList_E1_K1[2], inputList_E2_K1[2]), (inputList_E1_K2[2], inputList_E2_K2[2]), (inputList_E1_K3[2], inputList_E2_K3[2])) # third obj

def merge_to_new_trapezium(cross_table, table_OK):
    
    top_array = []
    bottom_array = []

    for mu,u in cross_table:
        if mu == 1:
            top_array.append(u)
        elif mu == 0:
            bottom_array.append(u)

    top_left_point = min(top_array)
    top_right_point = max(top_array)

    bottom_left_point = 0
    bottom_right_point = float("inf")

    for mu,u in cross_table:
        if mu == 0:
            if u < top_left_point and u > bottom_left_point:
                    bottom_left_point = u
            if u > top_right_point and u < bottom_right_point:
                    bottom_right_point = u


    centrA = round((bottom_left_point + 2 * top_left_point + 2 * top_right_point + bottom_right_point) / 6, 6)
    table_OK.append(centrA)

    print("merged trapezium:", bottom_left_point, top_left_point, top_right_point, bottom_right_point, sep="  ")
    print("centr(A): ", centrA)
    

def count_cross_table(expert_data1, expert_data2, k_num, o_num):
    
    cross_table = [[],[],[],[]]
    cross_table_arr = []
    ct_index = 0
    for j in range(3, -1, -1):   
        for i in range(0,4):
            cross_table[ct_index].append(str(min(MU_ARR[i], MU_ARR[j])) + ", " + str(round((expert_data1[i] + expert_data2[j])/2, 2)))
            cross_table_arr.append((min(MU_ARR[i], MU_ARR[j]), round((expert_data1[i] + expert_data2[j])/2, 2)))
        ct_index += 1
    df_cross_table = pd.DataFrame(cross_table,
                        index=['u4 эксперт 2', 'u3 эксперт 2', 'u2 эксперт 2', 'u1 эксперт 2'], 
                        columns=['u1 эксперт 1', 'u2 эксперт 1', 'u3 эксперт 1', 'u4 эксперт 1'])
    print(df_cross_table)

    file_name_excel = "table_O" + str(o_num) + "_K" + str(k_num) + ".xlsx"
    df_cross_table.to_excel(file_name_excel)

    return cross_table_arr



# START
table_OK = []

k_counter = 1

for i in range(0,9):
   
    if i % 3 == 0:
        o_counter = i // 3 + 1
        print("______________________________________________")
        print("Object " + str(o_counter) )
        print("______________________________________________")
        k_counter = 1

    exp1_estim, exp2_estim = EXP_ESTIM_PAIRS[i]
    print("|| K", k_counter, "||")
    print("expert 1 estimate: ", exp1_estim)
    print("expert 2 estimate: ", exp2_estim, "\n")
    ctable = count_cross_table(exp1_estim, exp2_estim, k_counter, o_counter)
    print()
    merge_to_new_trapezium(ctable, table_OK)
    print("\n")
    k_counter += 1

table_OK2D = [[], [], []]
for i in range(0, 3):
    for j in range(0, 3):
        table_OK2D[i].append(table_OK[3*i + j])

print("Table of objects and criteria:")

df_table_OK = pd.DataFrame(table_OK2D, index=['O1', 'O2', 'O3'], columns=['K1', 'K2', 'K3'])

print(df_table_OK)

df_table_OK.to_excel("O_K_table.xlsx")
