import pandas as pd
import win32com.client
import csv

rastr = win32com.client.Dispatch('Astra.Rastr')


def csv_to_list_of_dicts(path: str) -> [dict]:
    """
    записывает файл csv в список словарей
    :param path: пусть к файлу с траектрией утяжеления
    :return list_of_nodes_dict: траекторию утяжеления
    """
    list_of_nodes_dict = []
    with open(path, newline='') as csv_data:
        nodes_dict = csv.DictReader(csv_data)
        # Creating empty list and adding dictionaries (rows)
        for row in nodes_dict:
            list_of_nodes_dict.append(row)
    return list_of_nodes_dict


def add_entry_in_traject(node_number: int, tg_flag: int) -> int:
    """ 
    создает ряд в таблице утяжеления и выставляет в ряду номер узла и tg
    :param node_number: номер узла
    :param tg_flag: учет тангенса tg
    :return index: номер строки в таблице утяжеления
    """
    index = rastr.Tables('ut_node').size
    rastr.Tables('ut_node').AddRow()
    rastr.Tables('ut_node').Cols('ny').SetZ(index, node_number)
    rastr.Tables('ut_node').Cols('tg').SetZ(index, tg_flag)
    return index


def set_node_tr_param(node_number: int,
                      param: str,
                      value: float) -> None:
    """ 
    заполнение таблицы траетории утяжеления
    :param node_number - параметр узла
    :param param: Параметр утяжеления pg/pn
    :param value: Приращение pg/pn
    """
    rastr.Tables('ut_node').Cols(param).SetZ(node_number, value)


def trajectory_loading(trajectory_file, trajectory_shabl: str) -> None:
    """
    считывает файл траектории из формата .csv
    преобразовывает его к виду, в котором
    данная таблица находится в rastrwin3
    и загружает траекторию утяжеления в rastrwin3
    :param trajectory_file: путь к файлу траекории в формате .csv
    :param trajectory_shabl: путь к файлу шаблона траектории rastrwin3
    """
    rastr.Save('Trajectory.ut2', trajectory_shabl)
    rastr.Load(1, 'Trajectory.ut2', trajectory_shabl)
    list_of_dicts = csv_to_list_of_dicts(trajectory_file)
    present_entries = {}
    for row in list_of_dicts:
        node = row.get('node', 0)
        if node not in present_entries:
            node_number = add_entry_in_traject(node, row.get('tg', 0))
            present_entries[node] = node_number
        else:
            node_number = present_entries[node]
        variable = row.get('variable', 'pn')
        set_node_tr_param(
            node_number, variable, float(row.get('value', 0)))
    rastr.Save('Trajectory.ut2', trajectory_shabl)


def flowgate_loading(flowgate_file: str, flowgate_shabl: str) -> None:
    """
    считывает файл с сечением из формата .json
    с использованием pandas.dataframe
    и загружает сечение в таблицу сечений и группы линий в rastrwin3
    :param flowgate_file: путь к файлу сечения в формате .json
    :param flowgate_shabl: путь к файлу шаблона сечения rastrwin3
    """
    # Загрузка сечения
    flowgate = pd.read_json(flowgate_file, orient="index")
    flowgate = flowgate.reset_index(drop=True)
    rastr.Save('Flowgate.sch', flowgate_shabl)
    rastr.Load(1, 'Flowgate.sch', flowgate_shabl)
    rastr.Tables('sechen').AddRow()
    rastr.Tables('sechen').Cols('ns').SetZ(0, 1)
    for index, row in flowgate.iterrows():
        rastr.Tables('grline').AddRow()
        rastr.Tables('grline').Cols('ns').SetZ(index, 1)
        rastr.Tables('grline').Cols('ip').SetZ(index, row['ip'])
        rastr.Tables('grline').Cols('iq').SetZ(index, row['iq'])

    # Код для проверки заполнения таблицы
    rastr.Save('Flowgate.sch', flowgate_shabl)


def faults_loading(faults_file: str) -> pd.DataFrame:
    """
    считывает файл с авариями в формате .json
    и загружает их в pandas.dataframe
    :param faults_file: путь к файлу с авариями в формате .json
    """
    # Загрузка нормативных возмущений
    fault = pd.read_json(faults_file, orient="index")
    return fault


def loading_regime(reg: str,
                   reg_shab: str,
                   trajectory_shabl: str,
                   flowgate_shabl: str) -> None:
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


def ut_control(v: int, i: int, p: int) -> None:
    """
    включает контроль параметров для утяжеления и
    позволяет ввыбрать какие параметры контролировать для утяжеления
    По умолчанию все параметры равны 1
    0 - параметр включен
    1 - параметр отключен
    :param v: контроль напряжения при утяжелении
    :param i: контроль тока при утяжелении
    :param p: контроль резервов мощности при утяжелении
    """
    rastr.Tables('ut_common').Cols('enable_contr').SetZ(0, 1)
    rastr.Tables('ut_common').Cols('dis_v_contr').SetZ(0, v)
    rastr.Tables('ut_common').Cols('dis_i_contr').SetZ(0, i)
    rastr.Tables('ut_common').Cols('dis_p_contr').SetZ(0, p)


def ddtn_or_adtn_current_control(flag: bool) -> None:
    """
    Осуществляет перестановку параметоров ДДТН и АДТН
    и выделение линий для контроля в них тока
    Параметр flag:
    False - доаварийный режим (контроль ДДТН)
    True - аварийный режим (контроль АДТН)
    :param flag
    """
    if flag is False:
        for i in range(0, rastr.Tables('vetv').Size):
            rastr.Tables('vetv').Cols('i_dop').SetZ(
                i, rastr.Tables('vetv').Cols('i_dop_r').Z(i))
            if rastr.Tables('vetv').Cols('i_dop').Z(i) != 0:
                rastr.Tables('vetv').Cols('contr_i').SetZ(i, 1)
    else:
        for i in range(0, rastr.Tables('vetv').Size):
            rastr.Tables('vetv').Cols('i_dop').SetZ(
                i, rastr.Tables('vetv').Cols('i_dop_r_av').Z(i))
            if rastr.Tables('vetv').Cols('i_dop').Z(i) != 0:
                rastr.Tables('vetv').Cols('contr_i').SetZ(i, 1)


def set_voltage(voltage: float) -> None:
    """
    Осуществляет настройку минимального напряжения для утяжеления
    :param voltage: запас по напряжению
    """
    rastr.Tables('com_regim').Cols('dv_min').SetZ(0, float(voltage))


def criteria1_20percent_nofault(
        reg: str,
        reg_shab: str,
        position_of_flowgate: int,
        trajectory_shabl: str,
        flowgate_shabl: str) -> dict:
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
    loading_regime(reg, reg_shab, trajectory_shabl, flowgate_shabl)
    ut()

    p_limit = rastr.Tables('sechen').Cols('psech').Z(position_of_flowgate)
    mdp_1 = abs(p_limit) * 0.8 - 30
    result_criteria_1 = {'Criteria': '20% запас в норм схеме', 'MDP': mdp_1}
    return result_criteria_1


def criteria2_voltage_nofault(
        reg: str,
        reg_shab: str,
        position_of_flowgate: int,
        trajectory_shabl: str,
        flowgate_shabl: str) -> dict:
    """
    осуществляет расчет МДП по второму критерию
    :param reg: путь к файлу режима
    :param reg_shab: путь к файлу шаблона режима rastrwin3
    :param position_of_flowgate: номер сечения в таблице сечений в rastrwin3
    :param trajectory_shabl: путь к файлу шаблона траектории rastrwin3
    :param flowgate_shabl: путь к файлу шаблона сечения rastrwin3
    """
    # Расчет МДП по критерию 2
    # Коэффициент запаса по напряжению в узлах нагрузки в нормальной схеме
    loading_regime(reg, reg_shab, trajectory_shabl, flowgate_shabl)
    # Включим контроль по напряжению и отключим по всем остальным критериям
    ut_control(v=0, i=1, p=1)
    # set_voltage(voltage=0.7/(1-0.15))
    ut()

    p_limit_2 = rastr.Tables('sechen').Cols('psech').Z(position_of_flowgate)
    mdp_2 = abs(p_limit_2) - 30
    result_criteria_2 = {
        'Criteria': 'запас по напряжению в норм схеме',
        'MDP': mdp_2}
    return result_criteria_2


def criteria3_8percent_fault(
        reg: str,
        reg_shab: str,
        position_of_flowgate: int,
        trajectory_shabl: str,
        flowgate_shabl: str,
        faults: [pd.DataFrame]) -> dict:
    """
    осуществляет расчет МДП по второму критерию
    :param reg: путь к файлу режима
    :param reg_shab: путь к файлу шаблона режима rastrwin3
    :param position_of_flowgate: номер сечения в таблице сечений в rastrwin3
    :param trajectory_shabl: путь к файлу шаблона траектории rastrwin3
    :param flowgate_shabl: путь к файлу шаблона сечения rastrwin3
    :param faults: датафрейм с авариями
    """
    # Расчет МДП по критерию 3
    # Коэффициент запаса статичекой апериодической устойчивости в
    # послеаварийном режиме
    prelim_data_3 = pd.DataFrame(columns=['Fault-node_index', 'MDP'])

    for index, row in faults.iterrows():
        loading_regime(reg, reg_shab, trajectory_shabl, flowgate_shabl)
        vetv = rastr.Tables('vetv')
        vetv.SetSel(f'ip={row["ip"]}&iq={ row["iq"]}&np={row["np"]}')
        vetv.Cols('sta').Calc(str(row['sta']))
        rastr.rgm('p')
        ut()
        # Расчитаем переток в послеваварийном режиме с запасом в 8%
        mdp_p_av = abs(rastr.Tables('sechen').Cols('psech')
                       .Z(position_of_flowgate))
        mdp_8_persent = mdp_p_av * 0.92
        # Получим число шагов утяжеления, за которые переток достиг предела
        steps = rastr.GetToggle()
        # Выставим шаг, при котором переток равен перетоку в
        # послеваварийном режиме с запасом в 8%
        j = 0
        while mdp_p_av > mdp_8_persent:
            steps.MoveOnPosition(len(steps.GetPositions()) - j)
            mdp_p_av = abs(rastr.Tables('sechen').Cols('psech')
                           .Z(position_of_flowgate))
            j += 1
        # Включаем ветвь и смотрим переток в доаварийном режиме
        vetv.Cols('sta').Calc(not row['sta'])
        rastr.rgm('p')
        p_limit_3_prelim = rastr.Tables('sechen').Cols(
            'psech').Z(position_of_flowgate)
        mdp_3_prelim = abs(p_limit_3_prelim) - 30
        prelim_criteria_3 = {
            'Fault-node_index': 1, 'MDP': mdp_3_prelim}
        prelim_data_3 = prelim_data_3.append(
            prelim_criteria_3, ignore_index=True)

    mdp_3 = abs(prelim_data_3['MDP'].min())
    result_criteria_3 = {
        'Criteria': '8% запас в послеаварийной схеме',
        'MDP': mdp_3}
    return result_criteria_3


def criteria4_voltage_fault(
        reg: str,
        reg_shab: str,
        position_of_flowgate: int,
        trajectory_shabl: str,
        flowgate_shabl: str,
        faults: [pd.DataFrame]) -> dict:
    """
    осуществляет расчет МДП по второму критерию
    :param reg: путь к файлу режима
    :param reg_shab: путь к файлу шаблона режима rastrwin3
    :param position_of_flowgate: номер сечения в таблице сечений в rastrwin3
    :param trajectory_shabl: путь к файлу шаблона траектории rastrwin3
    :param flowgate_shabl: путь к файлу шаблона сечения rastrwin3
    :param faults: датафрейм с авариями
    """
    # Расчет МДП по критерию 4
    # Коэффициент запаса по напряжению в узлах нагрузки в послеаварийном режиме
    prelim_data_4 = pd.DataFrame(columns=['Fault-node_index', 'MDP'])

    for index, row in faults.iterrows():
        loading_regime(reg, reg_shab, trajectory_shabl, flowgate_shabl)
        # Включим контроль по напряжению и отключим по всем остальным
        # критериям
        ut_control(v=0, i=1, p=1)
        vetv = rastr.Tables('vetv')
        vetv.SetSel(f'ip={row["ip"]}&iq={ row["iq"]}&np={row["np"]}')
        vetv.Cols('sta').Calc(str(row['sta']))
        rastr.rgm('p')
        # set_voltage(voltage=0.7 / (1 - 0.1))
        ut()

        p_limit_4_prelim = rastr.Tables('sechen').Cols(
            'psech').Z(position_of_flowgate)
        mdp_4_prelim = abs(p_limit_4_prelim) - 30
        prelim_criteria_4 = {
            'Fault-node_index': 1, 'MDP': mdp_4_prelim}
        prelim_data_4 = prelim_data_4.append(
            prelim_criteria_4, ignore_index=True)

    mdp_4 = abs(prelim_data_4['MDP'].min())
    result_criteria_4 = {
        'Criteria': 'запас по напряжению в послеаварийной схеме',
        'MDP': mdp_4}
    return result_criteria_4


def criteria5_current_nofault(
        reg: str,
        reg_shab: str,
        position_of_flowgate: int,
        trajectory_shabl: str,
        flowgate_shabl: str) -> dict:
    """
    осуществляет расчет МДП по второму критерию
    :param reg: путь к файлу режима
    :param reg_shab: путь к файлу шаблона режима rastrwin3
    :param position_of_flowgate: номер сечения в таблице сечений в rastrwin3
    :param trajectory_shabl: путь к файлу шаблона траектории rastrwin3
    :param flowgate_shabl: путь к файлу шаблона сечения rastrwin3
    """
    # Расчет МДП по критерию 5
    # Допустимая токовая нагрузка в нормальной
    loading_regime(reg, reg_shab, trajectory_shabl, flowgate_shabl)
    # Включим контроль по току и отключим по всем остальным критериям
    ut_control(v=1, i=0, p=1)
    # Поместим значения тока оборудования в нужный столбец и отметим все ветви
    # для контроля напряжения
    ddtn_or_adtn_current_control(False)
    rastr.rgm('p')
    ut()

    p_limit_5_final = rastr.Tables('sechen').Cols(
        'psech').Z(position_of_flowgate)
    mdp_5 = abs(p_limit_5_final) - 30
    result_criteria_5 = {
        'Criteria': 'токовая загрузка в норм схеме',
        'MDP': mdp_5}
    return result_criteria_5


def criteria6_current_fault(
        reg: str,
        reg_shab: str,
        position_of_flowgate: int,
        trajectory_shabl: str,
        flowgate_shabl: str,
        faults: [pd.DataFrame]) -> dict:
    """
    осуществляет расчет МДП по второму критерию
    :param reg: путь к файлу режима
    :param reg_shab: путь к файлу шаблона режима rastrwin3
    :param position_of_flowgate: номер сечения в таблице сечений в rastrwin3
    :param trajectory_shabl: путь к файлу шаблона траектории rastrwin3
    :param flowgate_shabl: путь к файлу шаблона сечения rastrwin3
    :param faults: датафрейм с авариями
    """
    # Расчет МДП по критерию 6
    # Допустимая токовая нагрузка в послеаварийной схеме
    prelim_data_6 = pd.DataFrame(columns=['line_index', 'MDP'])

    for index, row in faults.iterrows():
        loading_regime(reg, reg_shab, trajectory_shabl, flowgate_shabl)
        # Включим контроль по току и отключим по всем остальным
        # критериям
        ut_control(v=1, i=0, p=1)
        vetv = rastr.Tables('vetv')
        vetv.SetSel(f'ip={row["ip"]}&iq={ row["iq"]}&np={row["np"]}')
        vetv.Cols('sta').Calc(row['sta'])
        rastr.rgm('p')
        # Поместим значения тока оборудования в нужный столбец и
        # отметим все ветви для контроля напряжения
        ddtn_or_adtn_current_control(True)
        rastr.rgm('p')
        ut()

        p_limit_6 = rastr.Tables('sechen').Cols(
            'psech').Z(position_of_flowgate)
        prelim_criteria_6 = {'line_index': 1, 'MDP': p_limit_6}
        prelim_data_6 = prelim_data_6.append(
            prelim_criteria_6, ignore_index=True)

    p_limit_6_final = abs(prelim_data_6['MDP'].abs().min())
    mdp_6 = abs(p_limit_6_final) - 30
    result_criteria_6 = {
        'Criteria': 'токовая загрузка в послеаварийной схеме',
        'MDP': mdp_6}
    return result_criteria_6
