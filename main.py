from xlrd import *
from numpy import *
from pyomo.environ import *
from math import *
import pandas as pd
import csv
from time import *
from copy import *
import multiprocessing as mp
from functools import partial



from Initialization import initialization
from create_variables import *
from run_model import *
from optimization import *
from power_flow import *
from power_flow_gas import *
from power_flow_heat import *
from calculate_criteria import *
from print_results import *
from save_results import *


def main():
    # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    #    Initialize case and settings
    # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    h = 24
    n_scenarios = 3
    b_prints = 1

    electrical_network, gas_network, heat_network, resources, weather, prices, admm = initialization()


    # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    #    Initial organization
    # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

    branch_w = electrical_network['branch_w']
    load_in_bus_w = electrical_network['load_in_bus_w']
    other_w = electrical_network['other_w']

    branch_g = gas_network['branch_g']
    load_in_bus_g = gas_network['load_in_bus_g']
    other_g = gas_network['other_g']

    branch_h = heat_network['branch_h']
    load_in_bus_h = heat_network['load_in_bus_h']
    gen_dh = heat_network['gen_dh']
    other_h = heat_network['other_h']

    load_w = resources['load_w']
    load_g = resources['load_g']
    load_h = resources['load_h']
    resources_agr = resources['resouces_aggregator']
    number_buildings = len(resources_agr)
    for i in range(0, len(resources_agr)):
        resources_agr[i]['consumption']['elec'] = load_w[i]
        resources_agr[i]['consumption']['gas'] = load_g[i]
        resources_agr[i]['consumption']['heat'] = load_h[i]
    fuel_station = resources['fuel_station']
    EVs = resources['EVs']

    profile_solar = weather['profile_solar']
    temperature_outside = weather['temperature_outside']
    number_EVs = len(EVs['building_number'])


    ro = {'ro': admm['ro']}
    ro = {'ro': admm['ro'] * 0.01}
    criteria_final = admm['criteria_final']

    print("")
    print('...Initialization completed...')
    print("")

    # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    #    ADMM initialization
    # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    criteria_h = []
    criteria_dual_h = []

    pi = {'w': [], 'g': [], 'h': [], 'w_up': [], 'g_up': [], 'h_up': [], 'w_down': [], 'g_down': [], 'h_down': [],
          'hy': [], 'hy_up': [], 'hy_down': []}

    pi['w_up'] = [[0 for i in range(0, h)] for i in range(0, len(load_in_bus_w))]
    pi['w_down'] = [[0 for i in range(0, h)] for i in range(0, len(load_in_bus_w))]
    pi['g'] = [[0 for i in range(0, h)] for i in range(0, len(load_in_bus_g))]
    pi['g_up'] = [[0 for i in range(0, h)] for i in range(0, len(load_in_bus_g))]
    pi['g_down'] = [[0 for i in range(0, h)] for i in range(0, len(load_in_bus_g))]
    pi['h'] = [[0 for i in range(0, h)] for i in range(0, len(load_in_bus_h))]
    pi['h_up'] = [[0 for i in range(0, h)] for i in range(0, len(load_in_bus_h))]
    pi['h_down'] = [[0 for i in range(0, h)] for i in range(0, len(load_in_bus_h))]
    pi['hy'] = [[0 for i in range(0, h)] for i in range(0, len(load_in_bus_g))]
    pi['hy_up'] = [[0 for i in range(0, h)] for i in range(0, len(load_in_bus_g))]
    pi['hy_down'] = [[0 for i in range(0, h)] for i in range(0, len(load_in_bus_g))]

    Pa = {'w': [], 'g': [], 'h': [], 'w_up': [], 'g_up': [], 'h_up': [], 'w_down': [], 'g_down': [], 'h_down': [],
          'hy': [], 'hy_up': [], 'hy_down': []}

    Pa['w_up'] = [[1 for i in range(0, h)] for i in range(0, len(load_in_bus_w))]
    Pa['w_down'] = [[1 for i in range(0, h)] for i in range(0, len(load_in_bus_w))]
    Pa['g'] = [[1 for i in range(0, h)] for i in range(0, len(load_in_bus_g))]
    Pa['g_up'] = [[1 for i in range(0, h)] for i in range(0, len(load_in_bus_g))]
    Pa['g_down'] = [[1 for i in range(0, h)] for i in range(0, len(load_in_bus_g))]
    Pa['h'] = [[1 for i in range(0, h)] for i in range(0, len(load_in_bus_h))]
    Pa['h_up'] = [[1 for i in range(0, h)] for i in range(0, len(load_in_bus_h))]
    Pa['h_down'] = [[1 for i in range(0, h)] for i in range(0, len(load_in_bus_h))]
    Pa['hy'] = [[1 for i in range(0, h)] for i in range(0, len(load_in_bus_g))]
    Pa['hy_up'] = [[1 for i in range(0, h)] for i in range(0, len(load_in_bus_g))]
    Pa['hy_down'] = [[1 for i in range(0, h)] for i in range(0, len(load_in_bus_g))]


    P_dso_w_up = []
    P_dso_w_down = []
    P_dso_g = []
    P_dso_g_up = []
    P_dso_g_down = []
    P_dso_h = []
    P_dso_h_up = []
    P_dso_h_down = []
    P_dso_hy = []
    P_dso_hy_up = []
    P_dso_hy_down = []

    time_h = {'aggregator': [], 'dso_w': [], 'dso_g': [], 'dso_h': []}

    # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    #    Run aggregator model
    # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    flag = 0
    k = 0
    iter = 0
    m3_h = []
    m2_h_up = []
    start = time()
    count = 0
    ro_tot = []
    costs_all = []
    print("")
    print('... Starting ADMM ...')
    print("")
    while flag == 0:
        if b_prints:
            print("")
            print("#####################################################################################################")
            print("#     Iteration ", iter)
            print("#####################################################################################################")


            print("")
            print("______ Run Model Aggregator ______")

        m = ConcreteModel()
        m.c1 = ConstraintList()

        m = create_variables(m, h, number_buildings, load_in_bus_w, load_in_bus_g, load_in_bus_h, number_EVs)

        m = run_model_buildings(m, h, number_buildings, profile_solar, temperature_outside, resources_agr,
                                load_in_bus_w, load_in_bus_g, load_in_bus_h, gen_dh, other_h, fuel_station, EVs, prices)

        m, time_h, cost = optimization_aggregator(m, h, prices, 0, iter, pi, ro, Pa, load_in_bus_w, load_in_bus_g, load_in_bus_h,
                                    other_g, other_h, b_prints, time_h, fuel_station, number_buildings)

        costs_all.append(cost)
        pd.DataFrame(costs_all).to_csv("costs.csv")

        if b_prints: 1
        #print_results_aggregator(m, h, load_in_bus_w, load_in_bus_g, load_in_bus_h)


        '''for t in range(0, h):
            n = 0
            print('hour', t)
            print(m.P_P2G_E[12, t].value)

            print(m.P_P2G_hy[12, t].value, "= HV ", m.P_P2G_hy_HV[12, t].value, "net ", m.P_P2G_hy_net[12, t].value, "sto", m.P_sto_hy_ch[0, t].value)

            print(m.y_soc_sto_hy[0, t + 1].value, "= soc", m.y_soc_sto_hy[0, t].value ,"ch", m.y_sto_hy_ch[0, t].value, "dis", m.y_sto_hy_dis[0, t].value)
            print(m.y_sto_hy_dis[0, t].value, "= HV", m.y_sto_hy_HV[0, t].value, "net", m.P_sto_hy_net[0, t].value/39.72)

            print("Reserves", m.U_P2G_E[12, t].value, m.D_P2G_E[12, t].value)

            for n in range(0, len(load_in_bus_g)):
                1
                #print(n, m.y_soc_sto_hy[n, t].value)
            print("sum net t", m.P_P2G_hy_net[12, t].value + m.P_sto_hy_net[0, t].value )

            print("Ph", m.P_hy[0, t].value, m.P_hy_up[0, t].value, m.P_hy_down[0, t].value)
            print("")

        print("sum net", sum(m.P_P2G_hy_net[12, t].value + m.P_sto_hy_net[0, t].value for t in range(0, h)))
        print("sum net P2G", sum(m.P_P2G_hy_net[12, t].value for t in range(0, h)))
        print("sum net sto", sum(m.P_sto_hy_net[0, t].value for t in range(0, h)))'''
        for t in range(0, h):
            '''for i in range(0, len(load_in_bus_g)):
                print(m.P_aggr_hy[t].value, m.P_aggr_hy_down[i, t].value, m.P_aggr_hy_up[i, t].value)
            for i in range(0, len(load_in_bus_w)):
                print(m.P_P2G_E[i, t].value, m.D_P2G_E[i, t].value, m.U_P2G_E[i, t].value)'''

            '''print(m.P_aggr_hy[t].value)
            for i in range(0, len(load_in_bus_g)):
                print(m.P_aggr_hy_up[i, t].value, m.P_aggr_hy_down[i, t].value)
                print(m.P_hy[i, t].value, m.P_hy_up[i, t].value, m.P_hy_down[i, t].value)
                print(m.y_soc_sto_hy[i, t].value, m.P_sto_hy_ch[i, t].value, m.y_sto_hy_ch[i, t].value, m.y_sto_hy_dis[i, t].value, m.b_y_sto_hy_ch[i, t].value,
                      m.b_y_sto_hy_dis[i, t].value, m.y_sto_hy_HV[i, t].value, m.P_sto_hy_net[i, t].value, m.D_sto_hy[i, t].value, m.U_sto_hy[i, t].value)
                print(m.P_dso_hy_up[0, i, t].value, m.P_dso_hy_down[0, i, t].value)
            for i in range(0, len(load_in_bus_w)):
                print(m.P_P2G_E[i, t].value, m.U_P2G_E[i, t].value, m.D_P2G_E[i, t].value, m.P_P2G_hy[i, t].value, m.U_P2G_hy[i, t].value,
                      m.D_P2G_hy[i, t].value, m.P_P2G_hy_HV[i, t].value, m.P_P2G_hy_net[i, t].value, m.U_P2G_hy_net[i, t].value, m.D_P2G_hy_net[i, t].value)'''




        sleep(0)

        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        #    Run electricity DSO model
        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        if b_prints:
            print("")
            print("______ Run Electricity DSO Flow Model ______")

        P_dso_w = []
        P_dso_old_w_up = deepcopy(P_dso_w_up)
        P_dso_w_up = []
        P_dso_old_w_down = deepcopy(P_dso_w_down)
        P_dso_w_down = []
        m1_h = []
        m1_h_up = []
        m1_h_down = []
        results_m1 = []

        for s in range(1, 3): #        for s in range(1, 1):
            if b_prints:
                print("")
                if s == 0:
                    print("______ Scenario Energy ______")
                elif s == 1:
                    print("______ Scenario Up ______")
                elif s == 2:
                    print("______ Scenario Down ______")

            for t in range(0, h):
                m1 = ConcreteModel()
                m1.c1 = ConstraintList()
                m1 = create_variables_power_flow(m1, h, branch_w, load_in_bus_w)

                m1 = power_flow_elec(m1, s, branch_w, load_in_bus_w, other_w)


                if s == 0:
                    m1, time_h = optimization_dso(m1, m, t, pi, ro, load_in_bus_w, b_prints, time_h)
                    m1_h.append(m1)

                    P_dso_prov_w = []
                    for i in range(0, len(load_in_bus_w)):
                        P_dso_prov_w.append(m1.P_dso[k, i, 0].value)
                    P_dso_w.append(P_dso_prov_w)



                elif s == 1:
                    m1, time_h = optimization_dso_up(m1, m, t, pi, ro, load_in_bus_w, b_prints, time_h)
                    m1_h_up.append(deepcopy(m1))

                    P_dso_prov_w_up = []
                    for i in range(0, len(load_in_bus_w)):
                        P_dso_prov_w_up.append(m1.P_dso_up[k, i, 0].value)
                    P_dso_w_up.append(P_dso_prov_w_up)

                elif s == 2:
                    m1, time_h = optimization_dso_down(m1, m, t, pi, ro, load_in_bus_w, b_prints, time_h)
                    m1_h_down.append(deepcopy(m1))

                    P_dso_prov_w_down = []
                    for i in range(0, len(load_in_bus_w)):
                        P_dso_prov_w_down.append(deepcopy(m1.P_dso_down[k, i, 0].value))
                    P_dso_w_down.append(P_dso_prov_w_down)

        results_m1.append(m1_h_down)
        results_m1.append(m1_h_up)
        results_m1.append(m1_h_down)


        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        #    Run gas DSO model
        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        if other_g['b_network']:
            if b_prints:
                print("")
                print("______ Run Gas DSO Flow Model ______")

            P_dso_old_g = deepcopy(P_dso_g)
            P_dso_g = []
            P_dso_old_g_up = deepcopy(P_dso_g_up)
            P_dso_g_up = []
            P_dso_old_g_down = deepcopy(P_dso_g_down)
            P_dso_g_down = []

            P_dso_old_hy = deepcopy(P_dso_hy)
            P_dso_hy = []
            P_dso_old_hy_up = deepcopy(P_dso_hy_up)
            P_dso_hy_up = []
            P_dso_old_hy_down = deepcopy(P_dso_hy_down)
            P_dso_hy_down = []

            m2_old = deepcopy(m2_h_up)
            m2_h = []
            m2_h_up = []
            m2_h_down = []
            results_m2 = []

            for s in range(1, 3):
                if b_prints:
                    print("")
                    if s == 0:
                        print("______ Scenario Energy ______")
                    elif s == 1:
                        print("______ Scenario Up ______")
                    elif s == 2:
                        print("______ Scenario Down ______")

                for t in range(0, h):
                    m2 = ConcreteModel()
                    m2.c1 = ConstraintList()

                    m2 = create_variables_power_flow_gas(m2, branch_g, load_in_bus_g, t, iter, m2_old)
                    m2 = power_flow_gas(m2, m, t, s, branch_g, load_in_bus_g, other_g)

                    if s == 0:
                        m2, time_h = optimization_dso_gas(m2, m, t, pi, ro, load_in_bus_g, b_prints, time_h)
                        m2_h.append(m2)

                        P_dso_prov_g = []
                        P_dso_prov_hy = []
                        for i in range(0, len(load_in_bus_g)):
                            P_dso_prov_g.append(m2.P_dso_gas[i, 0].value)
                            P_dso_prov_hy.append(m2.P_dso_hy[i, 0].value)
                        P_dso_g.append(P_dso_prov_g)
                        P_dso_hy.append(P_dso_prov_hy)

                    elif s == 1:
                        m2, time_h = optimization_dso_gas_up(m2, m, t, pi, ro, load_in_bus_g, b_prints, time_h)
                        m2_h_up.append(m2)

                        P_dso_prov_g_up = []
                        P_dso_prov_hy_up = []
                        for i in range(0, len(load_in_bus_g)):
                            P_dso_prov_g_up.append(m2.P_dso_gas_up[i, 0].value)
                            P_dso_prov_hy_up.append(m2.P_dso_hy_up[i, 0].value)
                        P_dso_g_up.append(P_dso_prov_g_up)
                        P_dso_hy_up.append(P_dso_prov_hy_up)

                        print(t, "WI", m2.WI[0, 0].value, ", HHV_mix", m2.HHV_mix[0, 0].value, "P_hy_net", m2.P_dso_hy_up[0, 0].value/3.54,
                              "perc", (m2.P_dso_hy_up[0, 0].value/3.54)/(m2.P_dso_hy_up[0, 0].value/3.54 + m2.Pgen_gas[0, 0].value))

                    elif s == 2:
                        m2, time_h = optimization_dso_gas_down(m2, m, t, pi, ro, load_in_bus_g, b_prints, time_h)
                        m2_h_down.append(m2)

                        P_dso_prov_g_down = []
                        P_dso_prov_hy_down = []
                        for i in range(0, len(load_in_bus_g)):
                            P_dso_prov_g_down.append(m2.P_dso_gas_down[i, 0].value)
                            P_dso_prov_hy_down.append(m2.P_dso_hy_down[i, 0].value)
                        P_dso_g_down.append(P_dso_prov_g_down)
                        P_dso_hy_down.append(P_dso_prov_hy_down)

                        print(t, "WI", m2.WI[0, 0].value, ", HHV_mix", m2.HHV_mix[0, 0].value, "P_hy_net",
                              m2.P_dso_hy_down[0, 0].value / 3.54,
                              "perc", (m2.P_dso_hy_down[0, 0].value/3.54)/(m2.P_dso_hy_down[0, 0].value/3.54 + m2.Pgen_gas[0, 0].value))
            results_m2.append(m2_h_down)
            results_m2.append(m2_h_up)
            results_m2.append(m2_h_down)

        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        #    Run heat DSO model
        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

        if other_h['b_network']:
            if b_prints:
                print("")
                print("______ Run Heat DSO Flow Model ______")

            P_dso_old_h = deepcopy(P_dso_h)
            P_dso_h = []
            P_dso_old_h_up = deepcopy(P_dso_h_up)
            P_dso_h_up = []
            P_dso_old_h_down = deepcopy(P_dso_h_down)
            P_dso_h_down = []
            m3_old = deepcopy(m3_h)
            m3_h = []
            m3_h_up = []
            m3_h_down = []
            results_m3 = []

            for s in range(0, 3):
                if b_prints:
                    print("")
                    if s == 0:
                        print("______ Scenario Energy ______")
                    elif s == 1:
                        print("______ Scenario Up ______")
                    elif s == 2:
                        print("______ Scenario Down ______")

                for t in range(0, h):
                    flag_error = 1

                    while flag_error:
                        m3 = ConcreteModel()
                        m3.c1 = ConstraintList()

                        m3 = create_variables_power_flow_heat(m3, m3_old, 1, t, iter, branch_h, load_in_bus_h, number_buildings)
                        m3 = power_flow_heat(m3, m, t, s, branch_h, load_h, load_in_bus_h, resources_agr,  heat_network, other_h)

                        if s == 0:
                            m3, flag_error, time_h = optimization_dso_heat(m3, m, t, pi, ro, load_in_bus_h, b_prints, time_h)

                            if flag_error == 0:
                                m3_h.append(m3)

                                P_dso_prov_gen = []
                                for i in range(0, len(load_in_bus_h)):
                                    P_dso_prov_gen.append(m3.P_dso_heat[i, 0].value)
                                P_dso_h.append(P_dso_prov_gen)

                        elif s == 1:
                            m3, flag_error, time_h = optimization_dso_heat_up(m3, m, t, pi, ro, load_in_bus_h, b_prints, time_h)

                            if flag_error == 0:
                                m3_h_up.append(m3)

                                P_dso_prov_gen_up = []
                                for i in range(0, len(load_in_bus_h)):
                                    P_dso_prov_gen_up.append(m3.P_dso_heat_up[i, 0].value)
                                P_dso_h_up.append(P_dso_prov_gen_up)

                        elif s == 2:
                            m3, flag_error, time_h = optimization_dso_heat_down(m3, m, t, pi, ro, load_in_bus_h, b_prints, time_h)

                            if flag_error == 0:
                                m3_h_down.append(m3)

                                P_dso_prov_gen_down = []
                                for i in range(0, len(load_in_bus_h)):
                                    P_dso_prov_gen_down.append(m3.P_dso_heat_down[i, 0].value)
                                P_dso_h_down.append(P_dso_prov_gen_down)
            results_m3.append(m3_h)
            results_m3.append(m3_h_up)
            results_m3.append(m3_h_down)


        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        #    Update pi
        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

        pi = update_pi(m, m1_h_up, m2_h, m2_h_up, m2_h_down, m1_h_down, m3_h, m3_h_up, m3_h_down,
                       k, h, pi, ro, load_in_bus_w, load_in_bus_g, load_in_bus_h, other_g, other_h)



        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        #    Stop criteria
        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        if b_prints:
            print("")
            print("______ Stop criteria calculation ______")

        criteria, criteria_dual = calculate_criteria(m, m1_h_up, m1_h_down, m2_h, m2_h_up, m2_h_down,
                                                     m3_h, m3_h_up, m3_h_down, h, k, iter,
                                                     load_in_bus_w, load_in_bus_g, load_in_bus_h, ro,
                                                     P_dso_old_w_up, P_dso_old_w_down,
                                                     P_dso_old_g, P_dso_old_g_up, P_dso_old_g_down,
                                                     P_dso_old_h, P_dso_old_h_up, P_dso_old_h_down,
                                                     P_dso_old_hy, P_dso_old_hy_up, P_dso_old_hy_down,
                                                     other_g, other_h)

        if b_prints: print("Primal residual =", criteria, ", Dual residual =", criteria_dual)
        criteria_h.append(criteria)
        criteria_dual_h.append(criteria_dual)

        print("save_criteria")
        save_criteria(criteria_h, criteria_dual_h)

        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        #    Save Results
        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

        save_time(time_h)

        if iter % 10 == 0:
            save_results_aggr3(m, resources_agr, h, load_in_bus_w, load_in_bus_g, load_in_bus_h, resources_agr, prices,
                               EVs)
            save_results_file(m, resources_agr, h, load_in_bus_w, load_in_bus_g, load_in_bus_h, resources_agr, prices,
                              fuel_station,
                              results_m1, branch_w, results_m2, branch_g, results_m3, branch_h)
            save_results_dso_w(m1_h_up, h, 1)
            save_results_dso_w(m1_h_down, h, 2)
            #save_results_dso_g(m2_h, h, 0)
            save_results_dso_g(m2_h_up, h, 1)
            save_results_dso_g(m2_h_down, h, 2)
            if other_h['b_network']:
                save_results_dso_h(m3_h, h, 0)
                save_results_dso_h(m3_h_up, h, 1)
                save_results_dso_h(m3_h_down, h, 2)
            save_results_aggr(m)

        criteria_final = 0.1
        # ((36+36+22)*2+27*3) * 24 = 6456
        criteria_final = 0.001 * 82
        if iter > 0 and criteria <= criteria_final and criteria >= -criteria_final and criteria_dual >= -criteria_final \
                and criteria_dual <= criteria_final:
            1
        if iter == 200:

            save_results_aggr3(m, resources_agr, h, load_in_bus_w, load_in_bus_g, load_in_bus_h, resources_agr, prices,
                               EVs)

            save_results_dso_w(m1_h_up, h, 1)
            save_results_dso_w(m1_h_down, h, 2)
            #save_results_dso_g(m2_h, h, 0)
            save_results_dso_g(m2_h_up, h, 1)
            save_results_dso_g(m2_h_down, h, 2)
            if other_h['b_network']:
                save_results_dso_h(m3_h, h, 0)
                save_results_dso_h(m3_h_up, h, 1)
                save_results_dso_h(m3_h_down, h, 2)

            end = time()

            print("")
            print("...ADMM converged...")
            print("Time of optimization:", end - start, "s")
            print("")

            save_results_aggr(m)

            flag = 1

        else:
            iter = iter + 1

        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        #    Update ro
        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        ro_tot.append(ro['ro'])
        pd.DataFrame(ro_tot).to_csv('ro_values.csv')
        print('ro', ro['ro'])
        if iter > 1 and count >= 5 and criteria_dual < criteria_final and ro['ro'] < 2.5:
            ro['ro'] = ro['ro'] * 2
            count = 0
        else:
            count = count + 1

        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        #    Create new variables
        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

        # ___________ Electricity ____________
        Pa_w_up = []
        for i in range(0, len(load_in_bus_w)):
            Pa_temp_w_up = []
            for t in range(0, h):
                Pa_temp_w_up.append(m1_h_up[t].P_dso_up[k, i, 0].value)
            Pa_w_up.append(Pa_temp_w_up)
            del (Pa_temp_w_up)
        Pa['w_up'] = deepcopy(Pa_w_up)

        Pa_w_down = []
        for i in range(0, len(load_in_bus_w)):
            Pa_temp_w_down = []
            for t in range(0, h):
                Pa_temp_w_down.append(m1_h_down[t].P_dso_down[k, i, 0].value)
            Pa_w_down.append(Pa_temp_w_down)
            del (Pa_temp_w_down)
        Pa['w_down'] = deepcopy(Pa_w_down)


        # ___________ Gas ____________
        if other_g['b_network']:
            Pa_g_up = []
            for i in range(0, len(load_in_bus_g)):
                Pa_temp_g_up = []
                for t in range(0, h):
                    Pa_temp_g_up.append(m2_h_up[t].P_dso_gas_up[i, 0].value)
                Pa_g_up.append(Pa_temp_g_up)
                del (Pa_temp_g_up)
            Pa['g_up'] = deepcopy(Pa_g_up)

            Pa_g_down = []
            for i in range(0, len(load_in_bus_g)):
                Pa_temp_g_down = []
                for t in range(0, h):
                    Pa_temp_g_down.append(m2_h_down[t].P_dso_gas_down[i, 0].value)
                Pa_g_down.append(Pa_temp_g_down)
                del (Pa_temp_g_down)
            Pa['g_down'] = deepcopy(Pa_g_down)


        # ___________ Heat ____________
        if other_h['b_network']:
            Pa_h = []
            for i in range(0, len(load_in_bus_h)):
                Pa_temp_h = []
                for t in range(0, h):
                    Pa_temp_h.append(m3_h[t].P_dso_heat[i, 0].value)
                Pa_h.append(Pa_temp_h)
                del (Pa_temp_h)
            Pa['h'] = deepcopy(Pa_h)

            Pa_h_up = []
            for i in range(0, len(load_in_bus_h)):
                Pa_temp_h_up = []
                for t in range(0, h):
                    Pa_temp_h_up.append(m3_h_up[t].P_dso_heat_up[i, 0].value)
                Pa_h_up.append(Pa_temp_h_up)
                del (Pa_temp_h_up)
            Pa['h_up'] = deepcopy(Pa_h_up)

            Pa_h_down = []
            for i in range(0, len(load_in_bus_h)):
                Pa_temp_h_down = []
                for t in range(0, h):
                    Pa_temp_h_down.append(m3_h_down[t].P_dso_heat_down[i, 0].value)
                Pa_h_down.append(Pa_temp_h_down)
                del (Pa_temp_h_down)
            Pa['h_down'] = deepcopy(Pa_h_down)

        # ___________ Hydrogen ____________
        if other_g['b_network']:
            Pa_hy_up = []
            for i in range(0, len(load_in_bus_g)):
                Pa_temp_hy_up = []
                for t in range(0, h):
                    Pa_temp_hy_up.append(m2_h_up[t].P_dso_hy_up[i, 0].value)
                Pa_hy_up.append(Pa_temp_hy_up)
                del (Pa_temp_hy_up)
            Pa['hy_up'] = deepcopy(Pa_hy_up)

            Pa_hy_down = []
            for i in range(0, len(load_in_bus_g)):
                Pa_temp_hy_down = []
                for t in range(0, h):
                    Pa_temp_hy_down.append(m2_h_down[t].P_dso_hy_down[i, 0].value)
                Pa_hy_down.append(Pa_temp_hy_down)
                del (Pa_temp_hy_down)
            Pa['hy_down'] = deepcopy(Pa_hy_down)

    save_results_file(m, resources_agr, h, load_in_bus_w, load_in_bus_g, load_in_bus_h, resources_agr, prices, fuel_station,
                      results_m1, branch_w, results_m2, branch_g, results_m3, branch_h)


if __name__ == '__main__':
    main()