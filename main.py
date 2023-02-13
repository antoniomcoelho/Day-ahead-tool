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



from Initialization import *
from create_variables import *
from run_model import *
from optimization import *
from power_flow import *
from power_flow_gas import *
from power_flow_heat import *
from calculate_criteria import *
from print_results import *
from save_results import *
from run_electricity_DSO_model import *
from run_gas_DSO_model import *
from run_heat_DSO_model import *


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

    pi['w'] = [[0 for i in range(0, h)] for i in range(0, len(load_in_bus_w))]
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

        print_aggregator_status(iter)

        print_aggregator_status(h, number_buildings, load_in_bus_w, load_in_bus_g, load_in_bus_h, number_EVs, profile_solar, temperature_outside, resources_agr,
                                gen_dh, other_g, other_h, fuel_station, EVs, prices, iter, pi, ro, Pa, time_h)





        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        #    Run electricity DSO model
        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        print_dso_problem_name("Electricity")

        P_dso_old_w_up = deepcopy(P_dso_w_up)
        P_dso_old_w_down = deepcopy(P_dso_w_down)

        #results_m1, P_dso_w_up, P_dso_w_down, m1_h_up, m1_h_down = run_electricity_DSO_model(m, h, branch_w, load_in_bus_w, other_w, pi, ro, time_h)

        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        #    Run gas DSO model
        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

        if other_g['b_network']:
            print_dso_problem_name("Gas")

            P_dso_old_g = deepcopy(P_dso_g)
            P_dso_old_g_up = deepcopy(P_dso_g_up)
            P_dso_old_g_down = deepcopy(P_dso_g_down)

            P_dso_old_hy = deepcopy(P_dso_hy)
            P_dso_old_hy_up = deepcopy(P_dso_hy_up)
            P_dso_old_hy_down = deepcopy(P_dso_hy_down)

            #results_m2, m2_h_up, P_dso_g_up, P_dso_hy_up, m2_h_down, P_dso_g_down, P_dso_hy_down = run_gas_DSO_model(m, m2_h_up, h, branch_g, load_in_bus_g, other_g, pi, ro, time_h, iter)


        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        #    Run heat DSO model
        # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

        if other_h['b_network']:
            print_dso_problem_name("Heat")

            P_dso_old_h = deepcopy(P_dso_h)
            P_dso_old_h_up = deepcopy(P_dso_h_up)
            P_dso_old_h_down = deepcopy(P_dso_h_down)

            results_m3, m3_h, P_dso_h, m3_h_up, P_dso_h_up, m3_h_down, P_dso_h_down, time_h = \
                run_heat_DSO_model(m, m3_h, h, branch_h, load_in_bus_h, other_h, pi, ro, time_h, iter, number_buildings, load_h, resources_agr, heat_network)




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



def print_aggregator_status(iter):
    print("")
    print("#####################################################################################################")
    print("#     Iteration ", iter)
    print("#####################################################################################################")

    print("")
    print("______ Run Model Aggregator ______")

def print_dso_problem_name(name):
    print("")
    print(f'______ Run {name} DSO Flow Model ______')


if __name__ == '__main__':
    main()