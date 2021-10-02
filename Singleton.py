import pandas as pd
import sys
import win32com.client
import PySimpleGUI as sg

rastr = win32com.client.Dispatch('Astra.Rastr')


def trajectory_loading(trajectory_file, trajectory_shabl) -> None:
    """
    считывает файл траектории из формата .csv
    преобразовывает его к виду, в котором
    данная таблица находится в rastrwin3
    с использованием pandas.dataframe
    и загружает траекторию утяжеления в rastrwin3
    :param trajectory_file: путь к файлу траекории в формате .csv
    :param trajectory_shabl: путь к файлу шаблона траектории rastrwin3
    """
    # Загрузка траектории утяжеления
    # Подготовка данных к загрузке в Растр
    rastr.Save('Trajectory.ut2', trajectory_shabl)
    rastr.Load(1, 'Trajectory.ut2', trajectory_shabl)
    Trajectory = pd.read_csv(trajectory_file)
    # Выделение траектории для нагрузки
    LoadTrajectory = Trajectory[Trajectory['variable'] == 'pn']
    LoadTrajectory = LoadTrajectory.rename(
        columns={
            'variable': 'pn',
            'value': 'pn_value',
            'tg': 'pn_tg'})
    # LoadTrajectory.to_csv('LoadTrajectory.csv', index=False)
    # Выделение траектории для генерации
    GenTrajectory = Trajectory[Trajectory['variable'] == 'pg']
    GenTrajectory = GenTrajectory.rename(
        columns={
            'variable': 'pg',
            'value': 'pg_value',
            'tg': 'pg_tg'})
    # GenTrajectory.to_csv('GenTrajectory.csv', index=False)
    # Создаем единый датарейм для исключения ошибок повторения узлов в
    # траектории утяжеления
    FinishedTrajectory = pd.merge(left=GenTrajectory, right=LoadTrajectory,
                                  left_on='node', right_on='node', how='outer')
    FinishedTrajectory = FinishedTrajectory.fillna(0)
    # Загрузка траектории в Растр итерациями
    i = 0
    for index, row in FinishedTrajectory.iterrows():
        rastr.Tables('ut_node').AddRow()
        rastr.Tables('ut_node').Cols('ny').SetZ(i, row['node'])
        if pd.notnull(row['pg']):
            rastr.Tables('ut_node').Cols('pg').SetZ(i, row['pg_value'])
            rastr.Tables('ut_node').Cols('tg').SetZ(i, row['pg_tg'])
            if pd.notnull(row['pn']):
                rastr.Tables('ut_node').Cols('pn').SetZ(i, row['pn_value'])
                rastr.Tables('ut_node').Cols('tg').SetZ(i, row['pn_tg'])
        i = i + 1
    # Код для проверки заполнения таблицы
    rastr.Save('Trajectory.ut2', trajectory_shabl)


def flowgate_loading(flowgate_file, flowgate_shabl) -> None:
    """
    считывает файл с сечением из формата .json
    с использованием pandas.dataframe
    и загружает сечение в таблицу сечений и группы линий в rastrwin3
    :param flowgate_file: путь к файлу сечения в формате .json
    :param flowgate_shabl: путь к файлу шаблона сечения rastrwin3
    """
    # Загрузка сечения
    flowgate = pd.read_json(flowgate_file)
    flowgate = flowgate.T
    rastr.Save('Flowgate.sch', flowgate_shabl)
    rastr.Load(1, 'Flowgate.sch', flowgate_shabl)
    i = 0
    serial_number_of_flowgate = 1
    position_of_flowgate = 0
    rastr.Tables('sechen').AddRow()
    rastr.Tables('sechen').Cols('ns').SetZ(
        position_of_flowgate, serial_number_of_flowgate)
    for index, row in flowgate.iterrows():
        rastr.Tables('grline').AddRow()
        rastr.Tables('grline').Cols('ns').SetZ(i, serial_number_of_flowgate)
        rastr.Tables('grline').Cols('ip').SetZ(i, row['ip'])
        rastr.Tables('grline').Cols('iq').SetZ(i, row['iq'])
        i = i + 1
    # Код для проверки заполнения таблицы
    rastr.Save('Flowgate.sch', flowgate_shabl)


def faults_loading(faults_file) -> pandas.Dataframe:
    """
    считывает файл с авариями в формате .json
    и загружает их в pandas.dataframe
    :param faults_file: путь к файлу с авариями в формате .json
    """
    # Загрузка нормативных возмущений
    fault = pd.read_json(faults_file)
    fault = fault.T
    return fault


def loading_regime(reg, reg_shab, trajectory_shabl, flowgate_shabl) -> None:
    """
    загружает файлы режима, траектории и сечения
    и увеличивает предельное число шагов утяжеления до 200
    :param reg: путь к файлу режима
    :param reg_shab: путь к файлу шаблона режима rastrwin3
    :param trajectory_shabl: путь к файлу шаблона траектории rastrwin3
    :param flowgate_shabl: путь к файлу шаблона сечения rastrwin3
    """
    rastr.Load(1, reg, reg_shab)
    rastr.Load(1, 'Trajectory.ut2', trajectory_shabl)
    rastr.Load(1, 'Flowgate.sch', flowgate_shabl)
    rastr.Tables('ut_common').Cols('iter').SetZ(0, 200)


def ut() -> None:
    """
    выполняет инициализацию утяжеления и
    в случае успеха утяжеляет режим до предела
    """
    if rastr.ut_utr('i') > 0:
        rastr.ut_utr('')


def ut_control(V, I, P) -> None:
    """
    включает контроль параметров для утяжеления и
    позволяет ввыбрать какие параметры контролировать для утяжеления
    0 - параметр включен
    1 - параметр отключен
    :param V: контроль напряжения при утяжелении
    :param I: контроль тока при утяжелении
    :param P: контроль резервов мощности при утяжелении
    """
    rastr.Tables('ut_common').Cols('enable_contr').SetZ(0, 1)
    rastr.Tables('ut_common').Cols('dis_v_contr').SetZ(0, V)
    rastr.Tables('ut_common').Cols('dis_i_contr').SetZ(0, I)
    rastr.Tables('ut_common').Cols('dis_p_contr').SetZ(0, P)


def criteria1(
        reg,
        reg_shab,
        position_of_flowgate,
        trajectory_shabl,
        flowgate_shabl) -> pandas.Dataframe:
    """
    осуществляет расчет МДП по первому критерию
    :param reg: путь к файлу режима
    :param reg_shab: путь к файлу шаблона режима rastrwin3
    :param position_of_flowgate: номер сечения в таблице сечений в rastrwin3
    :param trajectory_shabl: путь к файлу шаблона траектории rastrwin3
    :param flowgate_shabl: путь к файлу шаблона сечения rastrwin3
    """
    # Расчет МДП по критерию 1
    # Коэффициент запаса статичекой апериодической устойчивости в нормальной
    # схеме
    result_data = pd.DataFrame(columns=['Criteria', 'MDP'])
    loading_regime(reg, reg_shab, trajectory_shabl, flowgate_shabl)
    ut()

    P_limit = rastr.Tables('sechen').Cols('psech').Z(position_of_flowgate)
    mdp_1 = abs(P_limit) * 0.8 - 30
    result_criteria_1 = {'Criteria': '20% запас в норм схеме', 'MDP': mdp_1}
    result_data = result_data.append(result_criteria_1, ignore_index=True)
    return result_data


def criteria2(
        reg,
        reg_shab,
        position_of_flowgate,
        result_data,
        trajectory_shabl,
        flowgate_shabl) -> pandas.Dataframe:
    """
    осуществляет расчет МДП по второму критерию
    :param reg: путь к файлу режима
    :param reg_shab: путь к файлу шаблона режима rastrwin3
    :param position_of_flowgate: номер сечения в таблице сечений в rastrwin3
    :param result_data: датафрейм в который заносятся записи по наименьшему МДП посчитанному по критериям
    :param trajectory_shabl: путь к файлу шаблона траектории rastrwin3
    :param flowgate_shabl: путь к файлу шаблона сечения rastrwin3
    """
    # Расчет МДП по критерию 2
    # Коэффициент запаса по напряжению в узлах нагрузки в нормальной схеме
    loading_regime(reg, reg_shab, trajectory_shabl, flowgate_shabl)
    # Включим контроль по напряжению и отключим по всем остальным критериям
    ut_control(0, 1, 1)
    ut()

    P_limit_2 = rastr.Tables('sechen').Cols('psech').Z(position_of_flowgate)
    mdp_2 = abs(P_limit_2) - 30
    result_criteria_2 = {
        'Criteria': 'запас по напряжению в норм схеме',
        'MDP': mdp_2}
    result_data = result_data.append(result_criteria_2, ignore_index=True)
    return result_data


def criteria3(
        reg,
        reg_shab,
        position_of_flowgate,
        result_data,
        trajectory_shabl,
        flowgate_shabl,
        faults) -> pandas.Dataframe:
    """
    осуществляет расчет МДП по второму критерию
    :param reg: путь к файлу режима
    :param reg_shab: путь к файлу шаблона режима rastrwin3
    :param position_of_flowgate: номер сечения в таблице сечений в rastrwin3
    :param result_data: датафрейм в который заносятся записи по наименьшему МДП посчитанному по критериям
    :param trajectory_shabl: путь к файлу шаблона траектории rastrwin3
    :param flowgate_shabl: путь к файлу шаблона сечения rastrwin3
    :param faults: датафрейм с авариями
    """
    # Расчет МДП по критерию 3
    # Коэффициент запаса статичекой апериодической устойчивости в
    # послеаварийном режиме
    loading_regime(reg, reg_shab, trajectory_shabl, flowgate_shabl)
    vetv_table = rastr.Tables('vetv')
    prelim_data_3 = pd.DataFrame(columns=['Fault-node_index', 'MDP'])
    i = 0
    while (i < vetv_table.Size):
        current_ip = vetv_table.Cols('ip').Z(i)
        current_iq = vetv_table.Cols('iq').Z(i)
        current_np = vetv_table.Cols('np').Z(i)
        for index, row in faults.iterrows():
            if (current_ip == row['ip'] and
                    current_iq == row['iq'] and
                    current_np == row['np']):
                rastr.Load(1, reg, reg_shab)
                vetv_table = rastr.Tables('vetv')
                vetv_table.Cols('sta').SetZ(i, row['sta'])
                rastr.Commit
                rastr.rgm('p')
                ut()

                vetv_table.Cols('sta').SetZ(i, 0)
                rastr.rgm('p')
                P_limit_3_prelim = rastr.Tables('sechen').Cols(
                    'psech').Z(position_of_flowgate)
                mdp_3_prelim = abs(P_limit_3_prelim) * 0.92 - 30
                prelim_criteria_3 = {
                    'Fault-node_index': i, 'MDP': mdp_3_prelim}
                prelim_data_3 = prelim_data_3.append(
                    prelim_criteria_3, ignore_index=True)
                rastr.Rollback
        i = i + 1

    mdp_3 = abs(prelim_data_3['MDP'].min())
    result_criteria_3 = {
        'Criteria': '8% запас в послеаварийной схеме',
        'MDP': mdp_3}
    result_data = result_data.append(result_criteria_3, ignore_index=True)
    return result_data


def criteria4(
        reg,
        reg_shab,
        position_of_flowgate,
        result_data,
        trajectory_shabl,
        flowgate_shabl,
        faults) -> pandas.Dataframe:
    """
    осуществляет расчет МДП по второму критерию
    :param reg: путь к файлу режима
    :param reg_shab: путь к файлу шаблона режима rastrwin3
    :param position_of_flowgate: номер сечения в таблице сечений в rastrwin3
    :param result_data: датафрейм в который заносятся записи по наименьшему МДП посчитанному по критериям
    :param trajectory_shabl: путь к файлу шаблона траектории rastrwin3
    :param flowgate_shabl: путь к файлу шаблона сечения rastrwin3
    :param faults: датафрейм с авариями
    """
    # Расчет МДП по критерию 4
    # Коэффициент запаса по напряжению в узлах нагрузки в послеаварийном режиме
    loading_regime(reg, reg_shab, trajectory_shabl, flowgate_shabl)
    prelim_data_4 = pd.DataFrame(columns=['Fault-node_index', 'MDP'])
    j = 0
    while (j < rastr.Tables('vetv').Size):
        current_ip = rastr.Tables('vetv').Cols('ip').Z(j)
        current_iq = rastr.Tables('vetv').Cols('iq').Z(j)
        current_np = rastr.Tables('vetv').Cols('np').Z(j)
        for index, row in faults.iterrows():
            if (current_ip == row['ip'] and
                    current_iq == row['iq'] and
                    current_np == row['np']):
                rastr.Load(1, reg, reg_shab)
                rastr.Tables('ut_common').Cols('iter').SetZ(0, 200)
                # Включим контроль по напряжению и отключим по всем остальным
                # критериям
                ut_control(0, 1, 1)
                rastr.Tables('vetv').Cols('sta').SetZ(j, row['sta'])
                rastr.Commit
                rastr.rgm('p')
                ut()

                P_limit_4_prelim = rastr.Tables('sechen').Cols(
                    'psech').Z(position_of_flowgate)
                mdp_4_prelim = abs(P_limit_4_prelim) - 30
                prelim_criteria_4 = {
                    'Fault-node_index': j, 'MDP': mdp_4_prelim}
                prelim_data_4 = prelim_data_4.append(
                    prelim_criteria_4, ignore_index=True)
                rastr.Rollback
        j = j + 1

    mdp_4 = abs(prelim_data_4['MDP'].min())
    result_criteria_4 = {
        'Criteria': 'запас по напряжению в послеаварийной схеме',
        'MDP': mdp_4}
    result_data = result_data.append(result_criteria_4, ignore_index=True)
    return result_data


def criteria5(
        reg,
        reg_shab,
        position_of_flowgate,
        result_data,
        trajectory_shabl,
        flowgate_shabl) -> pandas.Dataframe:
    """
    осуществляет расчет МДП по второму критерию
    :param reg: путь к файлу режима
    :param reg_shab: путь к файлу шаблона режима rastrwin3
    :param position_of_flowgate: номер сечения в таблице сечений в rastrwin3
    :param result_data: датафрейм в который заносятся записи по наименьшему МДП посчитанному по критериям
    :param trajectory_shabl: путь к файлу шаблона траектории rastrwin3
    :param flowgate_shabl: путь к файлу шаблона сечения rastrwin3
    """
    # Расчет МДП по критерию 5
    # Допустимая токовая нагрузка в нормальной
    loading_regime(reg, reg_shab, trajectory_shabl, flowgate_shabl)
    # Включим контроль по току и отключим по всем остальным критериям
    ut_control(1, 0, 1)
    # Поместим значения тока оборудования в нужный столбец и отметим все ветви
    # для контроля напряжения
    i = 0
    while i < rastr.Tables('vetv').Size:
        rastr.Tables('vetv').Cols('i_dop').SetZ(
            i, rastr.Tables('vetv').Cols('i_dop_r').Z(i))
        if rastr.Tables('vetv').Cols('i_dop').Z(i) != 0:
            rastr.Tables('vetv').Cols('contr_i').SetZ(i, 1)
        i += 1
    rastr.rgm('p')
    ut()

    P_limit_5_final = rastr.Tables('sechen').Cols(
        'psech').Z(position_of_flowgate)
    mdp_5 = abs(P_limit_5_final) - 30
    result_criteria_5 = {
        'Criteria': 'токовая загрузка в норм схеме',
        'MDP': mdp_5}
    result_data = result_data.append(result_criteria_5, ignore_index=True)
    return result_data


def criteria6(
        reg,
        reg_shab,
        position_of_flowgate,
        result_data,
        trajectory_shabl,
        flowgate_shabl,
        faults) -> pandas.Dataframe:
    """
    осуществляет расчет МДП по второму критерию
    :param reg: путь к файлу режима
    :param reg_shab: путь к файлу шаблона режима rastrwin3
    :param position_of_flowgate: номер сечения в таблице сечений в rastrwin3
    :param result_data: датафрейм в который заносятся записи по наименьшему МДП посчитанному по критериям
    :param trajectory_shabl: путь к файлу шаблона траектории rastrwin3
    :param flowgate_shabl: путь к файлу шаблона сечения rastrwin3
    :param faults: датафрейм с авариями
    """
    # Расчет МДП по критерию 6
    # Допустимая токовая нагрузка в послеаварийной схеме
    loading_regime(reg, reg_shab, trajectory_shabl, flowgate_shabl)
    prelim_data_6 = pd.DataFrame(columns=['line_index', 'MDP'])
    j = 0
    while (j < rastr.Tables('vetv').Size):
        current_ip = rastr.Tables('vetv').Cols('ip').Z(j)
        current_iq = rastr.Tables('vetv').Cols('iq').Z(j)
        current_np = rastr.Tables('vetv').Cols('np').Z(j)
        for index, row in faults.iterrows():
            if (current_ip == row['ip'] and
                    current_iq == row['iq'] and
                    current_np == row['np']):
                rastr.Load(1, reg, reg_shab)
                rastr.Tables('ut_common').Cols('iter').SetZ(0, 200)
                # Включим контроль по току и отключим по всем остальным
                # критериям
                ut_control(1, 0, 1)
                rastr.Tables('vetv').Cols('sta').SetZ(j, row['sta'])
                rastr.Commit
                rastr.rgm('p')
                # Поместим значения тока оборудования в нужный столбец и
                # отметим все ветви для контроля напряжения
                i = 0
                while i < rastr.Tables('vetv').Size:
                    rastr.Tables('vetv').Cols('i_dop').SetZ(
                        i, rastr.Tables('vetv').Cols('i_dop_r').Z(i))
                    if rastr.Tables('vetv').Cols('i_dop').Z(i) != 0:
                        rastr.Tables('vetv').Cols('contr_i').SetZ(i, 1)
                    i += 1
                rastr.rgm('p')
                ut()
                P_limit_6 = rastr.Tables('sechen').Cols(
                    'psech').Z(position_of_flowgate)
                prelim_criteria_6 = {'line_index': j, 'MDP': P_limit_6}
                prelim_data_6 = prelim_data_6.append(
                    prelim_criteria_6, ignore_index=True)
                rastr.Rollback
        j = j + 1

    P_limit_6_final = abs(prelim_data_6['MDP'].abs().min())
    mdp_6 = abs(P_limit_6_final) - 30
    result_criteria_6 = {
        'Criteria': 'токовая загрузка в послеаварийной схеме',
        'MDP': mdp_6}
    result_data = result_data.append(result_criteria_6, ignore_index=True)
    result_data.head(6)
    print(result_data)
    return result_data
