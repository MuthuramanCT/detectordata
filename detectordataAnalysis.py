import pandas as pd
import numpy as np

df = pd.ExcelFile('detectordata.xlsx')
writer = pd.ExcelWriter('output.xlsx')
for name in df.sheet_names:
    if name != 'Graphs' and name != 'Aggregates_Entry':
        data = pd.read_excel('detectordata.xlsx', sheet_name=name)
        total_rows = data.shape[0]
        processdata = data[['\'C_q_Lkw__wert', '\'C_q_Pkw__wert', '\'C_v_Lkw__wert', '\'C_v_Pkw__wert']].copy()
        total_rows = processdata.shape[0]
        iterator = int((total_rows / 5)+1)
        columns = ['Type', '\'C_q_Lkw__wert', '\'C_q_Pkw__wert', '\'C_v_Lkw__wert', '\'C_v_Pkw__wert']
        output_table = pd.DataFrame(columns=columns)
        temp = {}
        temp1 = pd.DataFrame()
        for i in range(0, iterator):
            five_rows = processdata.loc[(i*5):((i*5)+4)]
            five_rows = five_rows.replace(0, np.NaN)
            C_q_Lkw__wert = five_rows['\'C_q_Lkw__wert'].sum()
            C_q_Pkw__wert = five_rows['\'C_q_Pkw__wert'].sum()
            C_v_Lkw__wert = five_rows['\'C_v_Lkw__wert'].mean()
            C_v_Pkw__wert = five_rows['\'C_v_Pkw__wert'].mean()
            temp = {'Type': 'Aggregate', '\'C_q_Lkw__wert': [C_q_Lkw__wert], '\'C_q_Pkw__wert': [C_q_Pkw__wert],
                    '\'C_v_Lkw__wert': [C_v_Lkw__wert], '\'C_v_Pkw__wert': [C_v_Pkw__wert]}
            temp1 = pd.DataFrame.from_dict(temp)
            five_rows['Type'] = 'Value'
            output_table = output_table.append([five_rows, temp1], ignore_index=True)
        output_table = output_table.replace(np.NaN, 0)
        output_table.to_excel(writer,sheet_name=name, index=False)
writer.save()

