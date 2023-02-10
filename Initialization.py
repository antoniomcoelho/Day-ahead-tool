from xlrd import *
from numpy import *
from pyomo.environ import *
from math import *
import pandas as pd
import csv
from time import *
from copy import *




def initialization():
    ####################################################################################################################
    ####################################################################################################################
    ####################################################################################################################
    ####################################################################################################################
    # Networks
    ####################################################################################################################
    ####################################################################################################################
    ####################################################################################################################
    ####################################################################################################################

    ####################################################################################################################
    # Electricity network
    ####################################################################################################################
    # _________________ Branches__________________________________________________________________________________
    # Array with vectors identifying the starting bus, end bus, R (p.u.) and X (p.u.)

    name_folder = os.getcwd()

    book_networks = name_folder + "\\Input Data - Networks.xlsx"
    book_resources = name_folder + "\\Input Data - Resources.xlsx"
    book_other = name_folder + "\\Input Data - Other.xlsx"
    wb_networks = open_workbook(book_networks)
    wb_resources = open_workbook(book_resources)
    wb_other = open_workbook(book_other)

    # ____________________________Electricity prices________________________________________________
    xl_sheet = wb_networks.sheet_by_index(0)
    b_network_g = xl_sheet.cell(3, 2).value
    b_network_h = xl_sheet.cell(4, 2).value


    xl_sheet = wb_networks.sheet_by_index(1)
    branch_w = []
    for i in range(0, xl_sheet.nrows - 2):
        prov = []
        for j in range(0, 4):
            prov.append(float(xl_sheet.cell(2 + i, 1 + j).value))
        branch_w.append(prov)
    branch_w = array(branch_w)




    # _________________ Building number in each bus __________________________________________________________________
    # Vector with the buildings identified in each node
    # Example, a network with 23 buses and several buildings in each bus

    xl_sheet = wb_networks.sheet_by_index(2)
    load_in_bus_w = []
    for i in range(0, xl_sheet.nrows - 2):
        prov = []
        for j in range(0, xl_sheet.ncols - 2):
            if xl_sheet.cell(2 + i, 2 + j).value == '':
                break
            else:
                prov.append(int(xl_sheet.cell(2 + i, 2 + j).value))
        load_in_bus_w.append(prov)


    # _________________ Other parameters _______________________________________________________________________________
    # ref_bus - bus with reference generator
    # MVA_base - MVA system base
    # v_min - minimum voltage limit (p.u.)
    # v_max - maximum voltage limit (p.u.)
    # v_ref - voltage of reference bus (p.u.)
    # I_max - maximum current limit (A)


    xl_sheet = wb_networks.sheet_by_index(3)
    ref_bus = xl_sheet.cell(1, 2).value
    v_ref = xl_sheet.cell(2, 2).value
    MVA_base = xl_sheet.cell(3, 2).value
    v_max = xl_sheet.cell(4, 2).value
    v_min = xl_sheet.cell(5, 2).value
    I_max = xl_sheet.cell(6, 2).value

    other_w = {'ref_bus': ref_bus, 'MVA_base': MVA_base, 'v_min': v_min, 'v_max': v_max, 'v_ref': v_ref, 'I_max': I_max}

    # __________________________________________________________________________________________________________________
    electrical_network = {'branch_w': branch_w, 'load_in_bus_w': load_in_bus_w, 'other_w': other_w}





    ####################################################################################################################
    # Gas network
    ####################################################################################################################

    # _________________ Pipelines __________________________________________________________________________________
    # From, To, R, X

    xl_sheet = wb_networks.sheet_by_index(4)
    branch_g = []
    for i in range(0, xl_sheet.nrows - 2):
        prov = []
        for j in range(0, 4):
            prov.append(float(xl_sheet.cell(2 + i, 1 + j).value))
        branch_g.append(prov)
    branch_g = array(branch_g)


    # _________________ Building number in each bus __________________________________________________________________
    # Vector with the buildings identified in each node

    xl_sheet = wb_networks.sheet_by_index(5)
    load_in_bus_g = []
    for i in range(0, xl_sheet.nrows - 2):
        prov = []
        for j in range(0, xl_sheet.ncols - 2):
            if xl_sheet.cell(2 + i, 2 + j).value == '':
                break
            else:
                prov.append(int(xl_sheet.cell(2 + i, 2 + j).value))
        load_in_bus_g.append(prov)


    # _________________ Other parameters __________________________________________________________________________________
    # p_max - maximum pressure (bar)
    # ref_gas - reference bus

    xl_sheet = wb_networks.sheet_by_index(6)
    ref_gas = xl_sheet.cell(1, 2).value
    p_max = xl_sheet.cell(2, 2).value

    other_g = {'p_max': p_max * p_max, 'ref_gas': ref_gas, 'b_network': b_network_g}

    # __________________________________________________________________________________________________________________
    gas_network = {'branch_g': branch_g, 'load_in_bus_g': load_in_bus_g, 'other_g': other_g}





    ####################################################################################################################
    # Heat network
    ####################################################################################################################

    # _________________ Pipelines __________________________________________________________________________________
    # From, To, Length (m), Diameter (mm), Heat transfer coefficient (W/C)

    xl_sheet = wb_networks.sheet_by_index(7)
    branch_h = []
    for i in range(0, xl_sheet.nrows - 2):
        prov = []
        for j in range(0, 5):
            prov.append(float(xl_sheet.cell(2 + i, 1 + j).value))
        branch_h.append(prov)
    branch_h = array(branch_h)


    # _________________ Building number in each bus __________________________________________________________________
    # Vector with the buildings identified in each node

    xl_sheet = wb_networks.sheet_by_index(8)
    load_in_bus_h = []
    for i in range(0, xl_sheet.nrows - 2):
        prov = []
        for j in range(0, xl_sheet.ncols - 2):
            if xl_sheet.cell(2 + i, 2 + j).value == '':
                break
            else:
                prov.append(int(xl_sheet.cell(2 + i, 2 + j).value))
        load_in_bus_h.append(prov)


    # _________________ Building number in each bus __________________________________________________________________
    # Vector with information of heat generators connected to the district heating

    xl_sheet = wb_networks.sheet_by_index(9)
    gen_dh = []
    for i in range(0, xl_sheet.ncols - 2):
        type = xl_sheet.cell(1, 2 + i).value
        bus_elec = xl_sheet.cell(2, 2 + i).value
        bus_gas = xl_sheet.cell(3, 2 + i).value
        bus_dh = xl_sheet.cell(4, 2 + i).value
        P_max = xl_sheet.cell(5, 2 + i).value
        rend = xl_sheet.cell(6, 2 + i).value
        rend_chp_elec = xl_sheet.cell(7, 2 + i).value
        rend_chp_heat = xl_sheet.cell(8, 2 + i).value
        prov = {'type': type, 'bus_elec': bus_elec, 'bus_gas': bus_gas, 'bus_dh': bus_dh, 'P_max': P_max, 'rend': rend},
        gen_dh.append({'type': type, 'bus_elec': bus_elec, 'bus_gas': bus_gas, 'bus_dh': bus_dh, 'P_max': P_max, 'rend': rend,
                       'rend_elec': rend_chp_elec, 'rend_heat': rend_chp_heat})

    buses_generator = []
    for i in range(0, xl_sheet.ncols - 2):
        buses_generator.append(int(xl_sheet.cell(4, 2 + i).value))

    # _________________ Other parameters __________________________________________________________________________________
    # Cp - Specific heat of water (J/(kg.C))
    # Ta - Ambience temperature (C)
    # m_min - minimum mass flow rate (kg/s)
    # m_max - maximum mass flow rate (kg/s)
    # Ts_min, Ts_max - minimum and maximum supply nodal and pipeline temperature (C)
    # Tr_min, Tr_max - minimum and maximum return nodal and pipeline temperature (C)
    # friction - pipeline friction
    # p_min, p_max - minimum and maximum pipeline pressure (Pa)
    # T_load, T_gen - temperature of supply nodes with loads and generators (C)

    xl_sheet = wb_networks.sheet_by_index(10)
    Cp = xl_sheet.cell(1, 2).value
    Ta = xl_sheet.cell(2, 2).value
    m_max = xl_sheet.cell(3, 2).value
    m_min = xl_sheet.cell(4, 2).value
    Ts_min = xl_sheet.cell(5, 2).value
    Ts_max = xl_sheet.cell(6, 2).value
    Tr_min = xl_sheet.cell(7, 2).value
    Tr_max = xl_sheet.cell(8, 2).value
    friction = xl_sheet.cell(9, 2).value
    p_min = xl_sheet.cell(10, 2).value
    p_max = xl_sheet.cell(11, 2).value
    T_load = xl_sheet.cell(12, 2).value
    T_gen = xl_sheet.cell(13, 2).value

    other_h = {'Cp': Cp, 'Ta': Ta, 'm_max': m_max, 'm_min': m_min, 'Ts_min': Ts_min, 'Ts_max': Ts_max, 'Tr_min': Tr_min, 'Tr_max': Tr_max,
               'friction': friction, 'p_min': p_min, 'p_max': p_max, 'T_load': T_load, 'T_gen': T_gen, 'b_network': b_network_h}

    # __________________________________________________________________________________________________________________
    heat_network = {'branch_h': branch_h, 'load_in_bus_h': load_in_bus_h, 'buses_gen': buses_generator, 'gen_dh': gen_dh,
                    'other_h': other_h}








    ####################################################################################################################
    ####################################################################################################################
    ####################################################################################################################
    ####################################################################################################################
    # Buildings / Consumers / Resources
    ####################################################################################################################
    ####################################################################################################################
    ####################################################################################################################
    ####################################################################################################################
    # _________________ Reference generator __________________________________________________________________________________

    # _________________ Electrical load __________________________________________________________________________________

    # Vector with electrical load consumption of each building in MW for 24h
    # Buildings numbers should correspond to this vector position, i.e., building ID 1 should be in position 0,
    # building ID 2 should be in position 1, and so on

    # Example of loads for each building for 24h
    # load_w = [[load_building_1], [load_building_2], ..., [load_building_x]]
    #

    xl_sheet = wb_resources.sheet_by_index(0)
    load_w = []
    for i in range(0, xl_sheet.nrows - 2):
        prov = []
        for t in range(0, 24):
            prov.append(xl_sheet.cell(2 + i, 2 + t).value)
        load_w.append(prov)


    # _________________ Gas load __________________________________________________________________________________

    # Vector with electrical load consumption of each building in MW for 24h
    # Buildings numbers should correspond to this vector position, i.e., building ID 1 should be in position 0,
    # building ID 2 should be in position 1, and so on

    # Example of loads for each building for 24h
    # load_g = [[load_building_1], [load_building_2], ..., [load_building_x]]
    xl_sheet = wb_resources.sheet_by_index(1)
    load_g = []
    for i in range(0, xl_sheet.nrows - 2):
        prov = []
        for t in range(0, 24):
            prov.append(xl_sheet.cell(2 + i, 2 + t).value)
        load_g.append(prov)


    # _________________ Heat load __________________________________________________________________________________

    # Vector with electrical load consumption of each building in MW for 24h
    # Buildings numbers should correspond to this vector position, i.e., building ID 1 should be in position 0,
    # building ID 2 should be in position 1, and so on

    # Example of loads for each building for 24h
    # load_h = [[load_building_1], [load_building_2], ..., [load_building_x]]
    xl_sheet = wb_resources.sheet_by_index(2)
    load_h = []
    for i in range(0, xl_sheet.nrows - 2):
        prov = []
        for t in range(0, 24):
            prov.append(xl_sheet.cell(2 + i, 2 + t).value)
        load_h.append(prov)


    # _________________ Building resources __________________________________________________________________________________
    # Vector with characteristics of each building, including resources installed (heat pump, gas boiler,
    # district heating, PV, storage, gas consumtion) and the respective limits

    # Example with two buildings

    # Buildings with hp/ without PV and storage/ without thermal modelling
    xl_sheet = wb_resources.sheet_by_index(3)
    resouces_aggregator = []
    for i in range(0, xl_sheet.ncols - 3):
        installed_hp = xl_sheet.cell(2, 3 + i).value
        max_power_hp = xl_sheet.cell(3, 3 + i).value
        effic_hp = xl_sheet.cell(4, 3 + i).value
        installed_gb = xl_sheet.cell(5, 3 + i).value
        max_power_gb = xl_sheet.cell(6, 3 + i).value
        effic_gb = xl_sheet.cell(7, 3 + i).value
        installed_pv = xl_sheet.cell(8, 3 + i).value
        max_power_pv = xl_sheet.cell(9, 3 + i).value
        installed_sto = xl_sheet.cell(10, 3 + i).value
        max_power_sto = xl_sheet.cell(11, 3 + i).value
        max_soc = xl_sheet.cell(12, 3 + i).value
        min_soc = xl_sheet.cell(13, 3 + i).value
        initial_soc = xl_sheet.cell(14, 3 + i).value
        effic_sto = xl_sheet.cell(15, 3 + i).value
        installed_dh = xl_sheet.cell(16, 3 + i).value
        max_power_dh = xl_sheet.cell(17, 3 + i).value
        installed_g_load = xl_sheet.cell(18, 3 + i).value
        building_R = xl_sheet.cell(19, 3 + i).value
        building_C = xl_sheet.cell(20, 3 + i).value
        building_init_temp = xl_sheet.cell(21, 3 + i).value
        installed_EVs = xl_sheet.cell(22, 3 + i).value

        xl_sheet2 = wb_resources.sheet_by_index(4)
        xl_sheet3 = wb_resources.sheet_by_index(5)
        prov_temp_max = []
        prov_temp_min = []
        for t in range(0, 24):
            prov_temp_max.append(xl_sheet2.cell(2 + t, 2 + i).value)
            prov_temp_min.append(xl_sheet3.cell(2 + t, 2 + i).value)

        resouces_aggregator.append({'installed': {'hp': installed_hp, 'gb': installed_gb, 'dh': installed_dh, 'PV': installed_pv, 'sto': installed_sto, 'g_out': installed_g_load, "EVs": installed_EVs},
                         'limits':  {'hp': max_power_hp, 'gb': max_power_gb, 'PV': max_power_pv, 'sto_P': max_power_sto, 'sto_soc_max': max_soc,
                                     'sto_soc_min': min_soc, 'sto_soc_init': initial_soc, 'dh': max_power_dh},
                         'rend': {'hp': effic_hp, 'gb': effic_gb, 'sto': effic_sto},
                         'thermal': {'R': building_R, 'C': building_C, 'T_init': building_init_temp,
                                     'T_max': prov_temp_max,
                                     'T_min': prov_temp_min},
                       'consumption': {'elec': [], 'gas': [], 'heat': []}})



    # ________________ Fuel station ________________
    xl_sheet2 = wb_resources.sheet_by_index(6)
    elec_bus = xl_sheet2.cell(1, 2).value
    gas_bus = xl_sheet2.cell(2, 2).value
    P2G_max_P = xl_sheet2.cell(3, 2).value
    P2G_rend = xl_sheet2.cell(4, 2).value
    sto_hy_soc_max = xl_sheet2.cell(5, 2).value
    sto_hy_soc_min = xl_sheet2.cell(6, 2).value
    sto_hy_soc_init = xl_sheet2.cell(7, 2).value
    sto_hy_max_P = xl_sheet2.cell(8, 2).value
    hy_LHV = xl_sheet2.cell(9, 2).value
    c_water = xl_sheet2.cell(10, 2).value
    c_oxygen = xl_sheet2.cell(11, 2).value

    xl_sheet2 = wb_resources.sheet_by_index(7)
    prov_hydrogen_load = []
    for t in range(0, 24):
        prov_hydrogen_load.append(xl_sheet2.cell(2, 2 + t).value)

    resource_fuel_station = {'bus_elec': elec_bus, 'bus_gas': gas_bus, 'P2G_max_P': P2G_max_P, 'P2G_rend': P2G_rend,
                             'sto_hy_soc_max': sto_hy_soc_max, 'sto_hy_soc_min': sto_hy_soc_min, 'sto_hy_soc_init': sto_hy_soc_init,
                             'sto_hy_max_P': sto_hy_max_P,'hy_LHV': hy_LHV, 'c_water': c_water, 'c_oxygen': c_oxygen,
                             'load': prov_hydrogen_load}





    # ________________ EVs ________________
    xl_sheet2 = wb_resources.sheet_by_index(8)
    number_EVs = int(xl_sheet2.cell(1, 2).value)
    time_arrival = []
    time_departure = []
    soc_arrival = []
    soc_departure = []
    soc_max = []
    building_number = []
    for i in range(0, number_EVs):
        time_arrival.append(xl_sheet2.cell(4 + i, 2).value)
        time_departure.append(xl_sheet2.cell(4 + i, 3).value)
        soc_arrival.append(xl_sheet2.cell(4 + i, 4).value)
        soc_departure.append(xl_sheet2.cell(4 + i, 5).value)
        soc_max.append(xl_sheet2.cell(4 + i, 6).value)
        building_number.append(xl_sheet2.cell(4 + i, 7).value)

    EVs = {"time_arrival": time_arrival, "time_departure": time_departure, "soc_arrival": soc_arrival,
           "soc_departure" : soc_departure, "soc_max": soc_max, "building_number": building_number}


    resources = {'load_w': load_w, 'load_g': load_g, 'load_h': load_h, 'resouces_aggregator': resouces_aggregator, 'fuel_station': resource_fuel_station, "EVs": EVs}
    ####################################################################################################################
    ####################################################################################################################
    ####################################################################################################################
    ####################################################################################################################
    # Other
    ####################################################################################################################
    ####################################################################################################################
    ####################################################################################################################
    ####################################################################################################################

    ####################################################################################################################
    # Weather data
    ####################################################################################################################

    xl_sheet = wb_other.sheet_by_index(0)
    # _________________ Solar profile __________________________________________________________________________________
    # Vector with solar profile (0-1)
    # _________________ Outside temperature ____________________________________________________________________________
    # Vector with outside temperature in Celsius degrees

    profile_solar = []
    temperature_outside = []
    for t in range(0, 24):
        profile_solar.append(xl_sheet.cell(2, 2 + t).value * 2.5)
        temperature_outside.append(xl_sheet.cell(3, 2 + t).value)

    weather = {'profile_solar': profile_solar, 'temperature_outside': temperature_outside}



    ####################################################################################################################
    # Prices
    ####################################################################################################################

    # w - forecasted day-ahead energy electricity market price
    # g - forecasted day-ahead energy gas market price
    # w_sec - forecasted day-ahead secondary band price
    # w_up - forecasted day-ahead secondary upward activation price
    # w_down - forecasted day-ahead secondary downward activation price
    # ratio_up - forecasted day-ahead upward secondary reserve activation ratio
    # ratio_down - forecasted day-ahead downward secondary reserve activation ratio
    xl_sheet = wb_other.sheet_by_index(1)
    prices_w = []
    prices_g = []
    prices_band = []
    prices_up = []
    prices_down = []
    ratio_up = []
    ratio_down = []
    prices_g_sec = []
    prices_hy = []
    prices_hy_sec = []
    prices_water = []
    prices_oxygen = []
    for t in range(0, 24):
        prices_w.append(xl_sheet.cell(2, 2 + t).value)
        prices_g.append(xl_sheet.cell(3, 2 + t).value)
        prices_band.append(xl_sheet.cell(4, 2 + t).value)
        prices_up.append(xl_sheet.cell(5, 2 + t).value)
        prices_down.append(xl_sheet.cell(6, 2 + t).value)
        ratio_up.append(xl_sheet.cell(7, 2 + t).value)
        ratio_down.append(xl_sheet.cell(8, 2 + t).value)
        prices_g_sec.append(xl_sheet.cell(9, 2 + t).value)
        prices_hy.append(xl_sheet.cell(10, 2 + t).value)
        prices_hy_sec.append(xl_sheet.cell(11, 2 + t).value)
        prices_water.append(xl_sheet.cell(12, 2 + t).value)
        prices_oxygen.append(xl_sheet.cell(13, 2 + t).value)

    prices = {"w": prices_w,
             "g": prices_g,
             "w_sec": prices_band,
             "w_up": prices_up,
             "w_down": prices_down,
             "ratio_up": ratio_up,
             "ratio_down": ratio_down,
             "g_sec": prices_g_sec,
             "hy": prices_hy,
             "hy_sec": prices_hy_sec,
             "water": prices_water,
             "oxygen": prices_oxygen}



    ####################################################################################################################
    # ADMM
    ####################################################################################################################

    xl_sheet = wb_other.sheet_by_index(2)
    ro = xl_sheet.cell(1, 2).value
    criteria_final = xl_sheet.cell(2, 2).value
    admm = {'ro': ro, 'criteria_final': criteria_final}


    return electrical_network, gas_network, heat_network, resources, weather, prices, admm



if __name__ == '__main__':
    main()