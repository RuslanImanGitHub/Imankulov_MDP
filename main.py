# Загрузка библиотек
import pandas as pd
import win32com.client
import functions

rastr = win32com.client.Dispatch('Astra.Rastr')

reg_shab = "resources/режим.rg2"
reg = "resources/regimeMDP.rg2"
sech_shab = "resources/сечения.sch"
sech = "resources/flowgate.json"
fault = "resources/faults.json"
traject = "resources/vector.csv"
traject_shab = "resources/траектория утяжеления.ut2"

functions.trajectory_loading(traject, traject_shab)
functions.flowgate_loading(sech, sech_shab)
faults = functions.faults_loading(fault)

result_criteria_1 = functions.criteria1_20percent_nofault(
    reg, reg_shab, 0, traject_shab, sech_shab)
result_criteria_2 = functions.criteria2_voltage_nofault(
    reg, reg_shab, 0, traject_shab, sech_shab)
result_criteria_3 = functions.criteria3_8percent_fault(
    reg, reg_shab, 0, traject_shab, sech_shab,
    faults)
result_criteria_4 = functions.criteria4_voltage_fault(
    reg, reg_shab, 0, traject_shab, sech_shab,
    faults)
result_criteria_5 = functions.criteria5_current_nofault(
    reg, reg_shab, 0, traject_shab, sech_shab)
result_criteria_6 = functions.criteria6_current_fault(
    reg, reg_shab, 0, traject_shab, sech_shab,
    faults)
final_result = result_criteria_1, result_criteria_2, result_criteria_3,\
               result_criteria_4, result_criteria_5, result_criteria_6

result_data = pd.DataFrame(final_result)
print(result_data)
