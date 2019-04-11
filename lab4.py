import pandas as pd

MU_ARR =  (0, 1, 1, 0)

df_input_1 = pd.read_excel("input.xlsx", parse_cols=list(range(1,13)), nrows=3, skiprows=1)

df_input_2 = pd.read_excel("input.xlsx", parse_cols=list(range(1,13)), nrows=3, skiprows=6)

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

def contains_with_bigger_mu(cross_table, val):

    for i in range(len(cross_table)):
        for mu,u in cross_table[i]:
            if u == val and mu > 0:
                return True
    return False

def merge_to_new_trapezium(cross_table, table_OK):
    
    top_array = []
    bottom_array = []


    for i in range(len(cross_table)):
        for mu,u in cross_table[i]:
            if mu == 1:
                top_array.append(u)
            elif mu == 0:
                bottom_array.append(u)

    top_left_point = min(top_array)
    top_right_point = max(top_array)

    bottom_left_point = 0
    bottom_right_point = float("inf")

    for i in range(len(cross_table)):
        for mu,u in cross_table[i]:
            if mu == 0:
                if u < top_left_point and u > bottom_left_point and not contains_with_bigger_mu(cross_table, u):
                        bottom_left_point = u
                if u > top_right_point and u < bottom_right_point and not contains_with_bigger_mu(cross_table, u):
                        bottom_right_point = u


    centrA = round((bottom_left_point + 2 * top_left_point + 2 * top_right_point + bottom_right_point) / 6, 6)
    table_OK.append(centrA)

    print("merged trapezium:", bottom_left_point, top_left_point, top_right_point, bottom_right_point, sep="  ")
    print("centr(A): ", centrA)
    

def count_cross_table(expert_data1, expert_data2, k_num, o_num):
    
    extended_expert_data1 = []
    extended_expert_data2 = []

    ed1_delta1 = int(round(expert_data1[1] - expert_data1[0], 1) * 10)
    ed1_delta2 = int(round(expert_data1[2] - expert_data1[1], 1) * 10)
    ed1_delta3 = int(round(expert_data1[3] - expert_data1[2], 1) * 10)

    extended_expert_data1.append((0, expert_data1[0]))
    for i in range(ed1_delta1 - 1):
        extended_expert_data1.append((round(1/ed1_delta1 * (i + 1), 3), round(extended_expert_data1[-1][1] + 0.1, 1)))
    extended_expert_data1.append((1, expert_data1[1]))
    for i in range(ed1_delta2 - 1):
        extended_expert_data1.append((1, round(extended_expert_data1[-1][1] + 0.1, 1)))
    extended_expert_data1.append((1, expert_data1[2]))
    for i in range(ed1_delta3 - 1):
        extended_expert_data1.append((round(1/ed1_delta3 * (ed1_delta3 - 1 - i), 3), round(extended_expert_data1[-1][1] + 0.1, 1)))
    extended_expert_data1.append((0, expert_data1[3]))


    ed2_delta1 = int(round(expert_data2[1] - expert_data2[0], 1) * 10)
    ed2_delta2 = int(round(expert_data2[2] - expert_data2[1], 1) * 10)
    ed2_delta3 = int(round(expert_data2[3] - expert_data2[2], 1) * 10)


    extended_expert_data2.append((0, expert_data2[0]))
    for i in range(ed2_delta1 - 1):
        extended_expert_data2.append((round(1/ed2_delta1 * (i + 1), 3), round(extended_expert_data2[-1][1] + 0.1, 1)))
    extended_expert_data2.append((1, expert_data2[1]))
    for i in range(ed2_delta2 - 1):
        extended_expert_data2.append((1, round(extended_expert_data2[-1][1] + 0.1, 1)))
    extended_expert_data2.append((1, expert_data2[2]))
    for i in range(ed2_delta3 - 1):
        extended_expert_data2.append((round(1/ed2_delta3 * (ed2_delta3 - 1 - i), 3), round(extended_expert_data2[-1][1] + 0.1, 1)))
    extended_expert_data2.append((0, expert_data2[3]))


    cross_table = []
    cross_table_result = []

    for i in range(len(extended_expert_data2)-1, -1, -1):
        row = []
        row_res = []
        for j in range(0,len(extended_expert_data1)):
            row.append(str(min(extended_expert_data1[j][0], extended_expert_data2[i][0])) + ", " 
                + str(round((extended_expert_data1[j][1] + extended_expert_data2[i][1])/2, 2)))
            row_res.append((min(extended_expert_data1[j][0], extended_expert_data2[i][0]), 
                round((extended_expert_data1[j][1] + extended_expert_data2[i][1])/2, 2)))
        cross_table.append(row)
        cross_table_result.append(row_res)

    cols_df = []
    index_df = []

    for i in range(1, len(extended_expert_data1) + 1):
        cols_df.append("u{} эксперт 1".format(i))
    for i in range(len(extended_expert_data2), 0, -1):
        index_df.append("u{} эксперт 2".format(i))


    print("expert 1 estimate: ", expert_data1)
    print("extended: ", extended_expert_data1)
    print("expert 2 estimate: ", expert_data2)
    print("extended: ", extended_expert_data2, "\n")

    df_cross_table = pd.DataFrame(cross_table, columns=cols_df, index=index_df)
    print(df_cross_table)
    print()

    file_name_excel = "table_O" + str(o_num) + "_K" + str(k_num) + ".xlsx"
    df_cross_table.to_excel(file_name_excel)


    return cross_table_result

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
    print()

    ctable = count_cross_table(exp1_estim, exp2_estim, k_counter, o_counter)
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
