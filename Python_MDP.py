# Загрузка библиотек
import pandas as pd
import win32com.client
import singleton

rastr = win32com.client.Dispatch('Astra.Rastr')

reg_shab = "C:/Users/Руслан/PycharmProjects/pythonProject/resources/режим.rg2"
reg = "C:/Users/Руслан/PycharmProjects/pythonProject/resources/regimeMDP.rg2"
sech_shab = "C:/Users/Руслан/PycharmProjects/pythonProject/resources/сечения.sch"
sech = "C:/Users/Руслан/PycharmProjects/pythonProject/resources/flowgate.json"
fault = "C:/Users/Руслан/PycharmProjects/pythonProject/resources/faults.json"
traject = "C:/Users/Руслан/PycharmProjects/pythonProject/resources/vector.csv"
traject_shab = "C:/Users/Руслан/PycharmProjects/pythonProject/resources/траектория утяжеления.ut2"

singleton.trajectory_loading(traject, traject_shab)
singleton.flowgate_loading(sech, sech_shab)
faults = singleton.faults_loading(fault)

result_data = pd.DataFrame(columns=['Criteria', 'MDP'])

result_criteria_1 = singleton.criteria1_20percent_nofault(
    reg, reg_shab, 0, traject_shab, sech_shab)
result_criteria_2 = singleton.criteria2_voltage_nofault(
    reg, reg_shab, 0, traject_shab, sech_shab)
result_criteria_3 = singleton.criteria3_8percent_fault(
    reg, reg_shab, 0, traject_shab, sech_shab,
    faults)
result_criteria_4 = singleton.criteria4_voltage_fault(
    reg, reg_shab, 0, traject_shab, sech_shab,
    faults)
result_criteria_5 = singleton.criteria5_current_nofault(
    reg, reg_shab, 0, traject_shab, sech_shab)
result_criteria_6 = singleton.criteria6_current_fault(
    reg, reg_shab, 0, traject_shab, sech_shab,
    faults)

result_data = result_data.append(result_criteria_1, ignore_index=True)
result_data = result_data.append(result_criteria_2, ignore_index=True)
result_data = result_data.append(result_criteria_3, ignore_index=True)
result_data = result_data.append(result_criteria_4, ignore_index=True)
result_data = result_data.append(result_criteria_5, ignore_index=True)
result_data = result_data.append(result_criteria_6, ignore_index=True)
print(result_data)