from xlrd import *
from numpy import *
from pyomo.environ import *
#from math import *
from time import *

import pandas as pd
import logging


def optimization_aggregator(m, h, prices, k, iter, pi0, ro0, Pa, load_in_bus_w, load_in_bus_g, load_in_bus_h,
                            other_g, other_h, b_prints, time_h, fuel_station0, number_buildings):

    prices_w = prices['w']
    prices_w_sec = prices['w_sec']
    prices_w_up = prices['w_up']
    prices_w_down = prices['w_down']
    prices_g = prices['g']
    prices_g_sec = prices['g_sec']
    ratio_up = prices['ratio_up']
    ratio_down = prices['ratio_down']
    prices_hy = prices['hy']
    prices_hy_sec = prices['hy_sec']
    prices_H2O = prices['water']
    prices_O2 = prices['oxygen']
    prices_H2O_sec = prices_H2O
    prices_O2_sec = prices_O2

    price_CO2 = 25
    CO2_license = 2.8

    c_H2O = fuel_station0['c_water']
    c_O2 = fuel_station0['c_oxygen']

    ro = ro0['ro']

    pi_w_up = pi0['w_up']
    pi_w_down = pi0['w_down']
    Pa0_w_up = Pa['w_up']
    Pa0_w_down = Pa['w_down']



    for t in range(0, h):
        for i in range(0, len(load_in_bus_w)):
            m.c1.add(m.P_dso[k, i, t] == m.P_w[i, t])
            m.c1.add(m.P_dso_up[k, i, t] == m.P_w_up[i, t])
            m.c1.add(m.P_dso_down[k, i, t] == m.P_w_down[i, t])

    for t in range(0, h):
        for i in range(0, len(load_in_bus_g)):
            m.c1.add(m.P_dso_gas_up[k, i, t] == m.P_g_up[i, t])
            m.c1.add(m.P_dso_gas_down[k, i, t] == m.P_g_down[i, t])

            m.c1.add(m.P_dso_hy_up[k, i, t] == m.P_hy_up[i, t])
            m.c1.add(m.P_dso_hy_down[k, i, t] == m.P_hy_down[i, t])

    for t in range(0, h):
        for i in range(0, len(load_in_bus_h)):
            m.c1.add(m.P_dso_heat[k, i, t] == m.P_h[i, t])
            m.c1.add(m.P_dso_heat_up[k, i, t] == m.P_h_up[i, t])
            m.c1.add(m.P_dso_heat_down[k, i, t] == m.P_h_down[i, t])


    buses_with_fuel_station_hy = []
    buses_with_fuel_station_w = []
    fuel_station = []
    for k in range(0, 1):
        buses_with_fuel_station_w.append(fuel_station0['bus_elec'])
        buses_with_fuel_station_hy.append(fuel_station0['bus_gas'])
        fuel_station.append(fuel_station0)


    for n in range(0, len(load_in_bus_w)):
        if n in buses_with_fuel_station_w:
            for k in range(0, len(fuel_station)):
                if fuel_station[k]['bus_elec'] == n:
                    for t in range(0, h):
                        print("Prenew", n, t)
                        m.c1.add(m.P_renew[t] == (m.P_P2G_E[n, t]) - sum(m.P_PV[j, t] for j in range(0, number_buildings)) +
                                (m.D_P2G_E[n, t] + sum(m.P_PV_down[j, t] for j in range(0, number_buildings))) * ratio_down[t] -
                                (m.U_P2G_E[n, t] + sum(m.P_PV_up[j, t] for j in range(0, number_buildings))) * ratio_up[t])

    flag_test = 0
    if flag_test == 1:
        data = load('test.npz')
        pi_w_up = data['pi_w_up']
        pi_w_down = data['pi_w_down']
        Pa0_w_up = data['Pa0_w_up']
        Pa0_w_down = data['Pa0_w_down']
        pi_h = data['pi_h']
        pi_h_up = data['pi_h_up']
        pi_h_down = data['pi_h_down']
        Pa0_h = data['Pa0_h']
        Pa0_h_up = data['Pa0_h_up']
        Pa0_h_down = data['Pa0_h_down']
        pi_g = data['pi_g']
        pi_g_up = data['pi_g_up']
        pi_g_down = data['pi_g_down']
        pi_hy_up = data['pi_hy_up']
        pi_hy_down = data['pi_hy_down']
        Pa0_g = data['Pa0_g']
        Pa0_g_up = data['Pa0_g_up']
        Pa0_g_down = data['Pa0_g_down']
        Pa0_hy_up = data['Pa0_hy_up']
        Pa0_hy_down = data['Pa0_hy_down']


    if iter == flag_test:
        m.value = Objective(expr= (sum(m.P_aggr_w[t] * prices_w[t] + m.P_aggr_g[t] * prices_g[t] - m.P_aggr_hy[t] * prices_hy[t] + m.P_renew[t] * 0.5
                                       for t in range(0, h))
                                + sum(sum(prices_H2O[t] * m.P_P2G_E[j, t] * c_H2O - prices_O2[t] * m.P_P2G_E[j, t] * c_O2
                                          for j in range(0, len(load_in_bus_w))) for t in range(0, h))
                                  -
                                  sum(sum(prices_w_sec[t] * (m.P_aggr_w_up[i, t] + m.P_aggr_w_down[i, t])
                                          for i in range(0, len(load_in_bus_w))) for t in range(0, h))
                                +
                                sum(sum(prices_w_down[t] * ratio_down[t] * m.P_aggr_w_down[i, t] -
                                        prices_w_up[t] * ratio_up[t] * m.P_aggr_w_up[i, t]
                                        for i in range(0, len(load_in_bus_w))) for t in range(0, h))
                                   +
                                sum(sum(-prices_g_sec[t] * ratio_down[t] * m.P_aggr_g_down[i, t] +
                                          prices_g_sec[t] * ratio_up[t] * m.P_aggr_g_up[i, t]
                                          for i in range(0, len(load_in_bus_g))) for t in range(0, h))
                                   +
                                   sum(sum(-prices_hy_sec[t] * ratio_down[t] * m.P_aggr_hy_down[i, t] +
                                           prices_hy_sec[t] * ratio_up[t] * m.P_aggr_hy_up[i, t]
                                           for i in range(0, len(load_in_bus_g))) for t in range(0, h))
                                   +
                                   sum(sum(prices_H2O_sec[t] * ratio_down[t] * m.D_P2G_E[i, t] * c_H2O -
                                           prices_H2O_sec[t] * ratio_up[t] * m.U_P2G_E[i, t] * c_H2O
                                           for i in range(0, len(load_in_bus_w))) for t in range(0, h))
                                   +
                                   sum(sum(-prices_O2_sec[t] * ratio_down[t] * m.D_P2G_E[i, t] * c_O2 +
                                           prices_O2_sec[t] * ratio_up[t] * m.U_P2G_E[i, t] * c_O2
                                           for i in range(0, len(load_in_bus_w))) for t in range(0, h))
                                   )/1000
                                  +
                                  sum(sum(m.P_chp_w[n, t] + m.P_chp_w_up[n, t] * ratio_up[t] - m.P_chp_w_down[n, t] * ratio_down[t]
                                          for n in range(0, len(load_in_bus_h))) for t in
                                      range(0, h)) / 1000 * 0.2 * price_CO2
                                  +
                                  (sum(sum(m.P_chp_h[n, t] + m.P_chp_w_up[n, t] * ratio_up[t] - m.P_chp_w_down[n, t] * ratio_down[t]
                                           for n in range(0, len(load_in_bus_h))) for t in range(0, h)) / 1000 * 0.2
                                   - CO2_license) * price_CO2
                            , sense=minimize)

    else:
        aux_gas_network = 0
        aux_heat_network = 0
        if other_g['b_network']:
            if flag_test == 0:
                pi_g = pi0['g']
                pi_g_up = pi0['g_up']
                pi_g_down = pi0['g_down']
                pi_hy_up = pi0['hy_up']
                pi_hy_down = pi0['hy_down']
                Pa0_g = Pa['g']
                Pa0_g_up = Pa['g_up']
                Pa0_g_down = Pa['g_down']
                Pa0_hy_up = Pa['hy_up']
                Pa0_hy_down = Pa['hy_down']



            aux_gas_network = sum(sum(pi_g_up[i][t] * (m.P_dso_gas_up[k, i, t] - Pa0_g_up[i][t])
                                      for i in range(0, len(load_in_bus_g))) +
                                ro / 2 * sum((m.P_dso_gas_up[k, i, t] - Pa0_g_up[i][t]) ** 2
                                             for i in range(0, len(load_in_bus_g))) for t in range(0, h)) \
                              + \
                              sum(sum(pi_g_down[i][t] * (m.P_dso_gas_down[k, i, t] - Pa0_g_down[i][t])
                                      for i in range(0, len(load_in_bus_g))) +
                                ro / 2 * sum((m.P_dso_gas_down[k, i, t] - Pa0_g_down[i][t]) ** 2
                                             for i in range(0, len(load_in_bus_g))) for t in range(0, h))

            aux_gas_network = aux_gas_network + sum(sum(pi_hy_up[i][t] * (m.P_dso_hy_up[k, i, t] - Pa0_hy_up[i][t])
                                      for i in range(0, len(load_in_bus_g))) +
                                ro / 2 * sum((m.P_dso_hy_up[k, i, t] - Pa0_hy_up[i][t]) ** 2
                                             for i in range(0, len(load_in_bus_g))) for t in range(0, h)) \
                              + \
                              sum(sum(pi_hy_down[i][t] * (m.P_dso_hy_down[k, i, t] - Pa0_hy_down[i][t])
                                      for i in range(0, len(load_in_bus_g))) +
                                ro / 2 * sum((m.P_dso_hy_down[k, i, t] - Pa0_hy_down[i][t]) ** 2
                                             for i in range(0, len(load_in_bus_g))) for t in range(0, h))



        if other_h['b_network']:
            if flag_test == 0:
                pi_h = pi0['h']
                pi_h_up = pi0['h_up']
                pi_h_down = pi0['h_down']
                Pa0_h = Pa['h']
                Pa0_h_up = Pa['h_up']
                Pa0_h_down = Pa['h_down']


            aux_heat_network =  sum(sum(pi_h[i][t] * (m.P_dso_heat[k, i, t] - Pa0_h[i][t]) for i in
                                         range(0, len(load_in_bus_h))) +
                                     ro / 2 * sum((m.P_dso_heat[k, i, t] - Pa0_h[i][t]) ** 2 for i in
                                                  range(0, len(load_in_bus_h)))
                                     for t in range(0, h)) \
                                + \
                                sum(sum(pi_h_up[i][t] * (m.P_dso_heat_up[k, i, t] - Pa0_h_up[i][t]) for i in
                                         range(0, len(load_in_bus_h))) +
                                     ro / 2 * sum((m.P_dso_heat_up[k, i, t] - Pa0_h_up[i][t]) ** 2 for i in
                                                  range(0, len(load_in_bus_h)))
                                     for t in range(0, h)) \
                                + \
                                sum(sum(pi_h_down[i][t] * (m.P_dso_heat_down[k, i, t] - Pa0_h_down[i][t]) for i in
                                         range(0, len(load_in_bus_h))) +
                                     ro / 2 * sum((m.P_dso_heat_down[k, i, t] - Pa0_h_down[i][t]) ** 2 for i in
                                                  range(0, len(load_in_bus_h)))
                                     for t in range(0, h))

        print("Save results")

        if flag_test == 0:
            pi_w_up = pi0['w_up']
            pi_w_down = pi0['w_down']
            Pa0_w_up = Pa['w_up']
            Pa0_w_down = Pa['w_down']
        '''savez('test.npz', pi_h=pi_h, pi_h_up=pi_h_up, pi_h_down=pi_h_down, Pa0_h=Pa0_h, Pa0_h_up=Pa0_h_up, Pa0_h_down=Pa0_h_down,
                    pi_g=pi_g,pi_g_up=pi_g_up,pi_g_down=pi_g_down,pi_hy_up=pi_hy_up,pi_hy_down=pi_hy_down,Pa0_g=Pa0_g,Pa0_g_up=Pa0_g_up,Pa0_g_down=Pa0_g_down,Pa0_hy_up=Pa0_hy_up,Pa0_hy_down=Pa0_hy_down,
                    pi_w_up=pi_w_up,pi_w_down=pi_w_down,Pa0_w_up=Pa0_w_up,Pa0_w_down=Pa0_w_down)
        '''


        m.value = Objective(expr = (sum(m.P_aggr_w[t] * prices_w[t] + m.P_aggr_g[t] * prices_g[t] - m.P_aggr_hy[t] * prices_hy[t]
                                       for t in range(0, h))
                                + sum(sum(prices_H2O[t] * m.P_P2G_E[j, t] * c_H2O - prices_O2[t] * m.P_P2G_E[j, t] * c_O2
                                          for j in range(0, len(load_in_bus_w))) for t in range(0, h))
                                  -
                                  sum(sum(prices_w_sec[t] * (m.P_aggr_w_up[i, t] + m.P_aggr_w_down[i, t])
                                          for i in range(0, len(load_in_bus_w))) for t in range(0, h))
                                +
                                sum(sum(prices_w_down[t] * ratio_down[t] * m.P_aggr_w_down[i, t] -
                                        prices_w_up[t] * ratio_up[t] * m.P_aggr_w_up[i, t]
                                        for i in range(0, len(load_in_bus_w))) for t in range(0, h))
                                   +
                                sum(sum(-prices_g_sec[t] * ratio_down[t] * m.P_aggr_g_down[i, t] +
                                          prices_g_sec[t] * ratio_up[t] * m.P_aggr_g_up[i, t]
                                          for i in range(0, len(load_in_bus_g))) for t in range(0, h))
                                   +
                                   sum(sum(-prices_hy_sec[t] * ratio_down[t] * m.P_aggr_hy_down[i, t] +
                                           prices_hy_sec[t] * ratio_up[t] * m.P_aggr_hy_up[i, t]
                                           for i in range(0, len(load_in_bus_g))) for t in range(0, h))
                                   +
                                   sum(sum(-prices_H2O_sec[t] * ratio_down[t] * m.D_P2G_E[i, t] * c_H2O +
                                           prices_H2O_sec[t] * ratio_up[t] * m.U_P2G_E[i, t] * c_H2O
                                           for i in range(0, len(load_in_bus_w))) for t in range(0, h))
                                   +
                                   sum(sum(-prices_O2_sec[t] * ratio_down[t] * m.D_P2G_E[i, t] * c_O2 +
                                           prices_O2_sec[t] * ratio_up[t] * m.U_P2G_E[i, t] * c_O2
                                           for i in range(0, len(load_in_bus_w))) for t in range(0, h))
                                   )/1000
                                  +
                                  sum(sum(m.P_chp_w[n, t] + m.P_chp_w_up[n, t] * ratio_up[t] - m.P_chp_w_down[n, t] * ratio_down[t]
                                          for n in range(0, len(load_in_bus_h))) for t in
                                      range(0, h)) / 1000 * 0.2 * price_CO2
                                  +
                                  (sum(sum(m.P_chp_h[n, t] + m.P_chp_w_up[n, t] * ratio_up[t] - m.P_chp_w_down[n, t] * ratio_down[t]
                                           for n in range(0, len(load_in_bus_h))) for t in range(0, h)) / 1000 * 0.2
                                   - CO2_license) * price_CO2



                                 +

                                 sum(sum(pi_w_up[i][t] * (m.P_dso_up[k, i, t] - Pa0_w_up[i][t]) for i in
                                         range(0, len(load_in_bus_w))) +
                                     ro / 2 * sum((m.P_dso_up[k, i, t] ** 2 - 2 * m.P_dso_up[k, i, t] * Pa0_w_up[i][
                                         t] + Pa0_w_up[i][t] ** 2) for i in
                                                  range(0, len(load_in_bus_w)))
                                     for t in range(0, h))

                                 +

                                 sum(sum(pi_w_down[i][t] * (m.P_dso_down[k, i, t] - Pa0_w_down[i][t]) for i in
                                         range(0, len(load_in_bus_w))) +
                                     ro / 2 * sum(
                                         (m.P_dso_down[k, i, t] ** 2 - 2 * m.P_dso_down[k, i, t] * Pa0_w_down[i][
                                             t] + Pa0_w_down[i][t] ** 2) for i in
                                         range(0, len(load_in_bus_w)))
                                     for t in range(0, h))

                                + aux_gas_network

                                + aux_heat_network

                            , sense=minimize)


    solver = SolverFactory("cplex")

    results = solver.solve(m, tee=False)



    cost = value((sum(m.P_aggr_w[t] * prices_w[t] + m.P_aggr_g[t] * prices_g[t] - m.P_aggr_hy[t] * prices_hy[t] + m.P_renew[t] * 0.5
                                       for t in range(0, h))
                                + sum(sum(prices_H2O[t] * m.P_P2G_E[j, t] * c_H2O - prices_O2[t] * m.P_P2G_E[j, t] * c_O2
                                          for j in range(0, len(load_in_bus_w))) for t in range(0, h))
                                  -
                                  sum(sum(prices_w_sec[t] * (m.P_aggr_w_up[i, t] + m.P_aggr_w_down[i, t])
                                          for i in range(0, len(load_in_bus_w))) for t in range(0, h))
                                +
                                sum(sum(prices_w_down[t] * ratio_down[t] * m.P_aggr_w_down[i, t] -
                                        prices_w_up[t] * ratio_up[t] * m.P_aggr_w_up[i, t]
                                        for i in range(0, len(load_in_bus_w))) for t in range(0, h))
                                   +
                                sum(sum(-prices_g_sec[t] * ratio_down[t] * m.P_aggr_g_down[i, t] +
                                          prices_g_sec[t] * ratio_up[t] * m.P_aggr_g_up[i, t]
                                          for i in range(0, len(load_in_bus_g))) for t in range(0, h))
                                   +
                                   sum(sum(-prices_hy_sec[t] * ratio_down[t] * m.P_aggr_hy_down[i, t] +
                                           prices_hy_sec[t] * ratio_up[t] * m.P_aggr_hy_up[i, t]
                                           for i in range(0, len(load_in_bus_g))) for t in range(0, h))
                                   +
                                   sum(sum(prices_H2O_sec[t] * ratio_down[t] * m.D_P2G_E[i, t] * c_H2O -
                                           prices_H2O_sec[t] * ratio_up[t] * m.U_P2G_E[i, t] * c_H2O
                                           for i in range(0, len(load_in_bus_w))) for t in range(0, h))
                                   +
                                   sum(sum(-prices_O2_sec[t] * ratio_down[t] * m.D_P2G_E[i, t] * c_O2 +
                                           prices_O2_sec[t] * ratio_up[t] * m.U_P2G_E[i, t] * c_O2
                                           for i in range(0, len(load_in_bus_w))) for t in range(0, h))
                                   )/1000
                                  +
                                  sum(sum(m.P_chp_w[n, t] + m.P_chp_w_up[n, t] * ratio_up[t] - m.P_chp_w_down[n, t] * ratio_down[t]
                                          for n in range(0, len(load_in_bus_h))) for t in
                                      range(0, h)) / 1000 * 0.2 * price_CO2
                                  +
                                  (sum(sum(m.P_chp_h[n, t] + m.P_chp_w_up[n, t] * ratio_up[t] - m.P_chp_w_down[n, t] * ratio_down[t]
                                           for n in range(0, len(load_in_bus_h))) for t in range(0, h)) / 1000 * 0.2
                                   - CO2_license) * price_CO2)

    if (results.solver.status == SolverStatus.ok) and ( results.solver.termination_condition == TerminationCondition.optimal):
        if b_prints:
            print("Flow optimized")
        time_h['aggregator'].append(results.solver.time)
    else:
        if b_prints:
            print("Did no converge")

    print("Cost", (sum(m.P_aggr_w[t].value * prices_w[t] + m.P_aggr_g[t].value * prices_g[t] - m.P_aggr_hy[t].value * prices_hy[t]
                                       for t in range(0, h))
                                + sum(sum(prices_H2O[t] * m.P_P2G_E[j, t].value * c_H2O - prices_O2[t] * m.P_P2G_E[j, t].value * c_O2
                                          for j in range(0, len(load_in_bus_w))) for t in range(0, h))
                                  -
                                  sum(sum(prices_w_sec[t] * (m.P_aggr_w_up[i, t].value + m.P_aggr_w_down[i, t].value)
                                          for i in range(0, len(load_in_bus_w))) for t in range(0, h))
                                +
                                sum(sum(prices_w_down[t] * ratio_down[t] * m.P_aggr_w_down[i, t].value -
                                        prices_w_up[t] * ratio_up[t] * m.P_aggr_w_up[i, t].value
                                        for i in range(0, len(load_in_bus_w))) for t in range(0, h))
                                   +
                                sum(sum(-prices_g_sec[t] * ratio_down[t] * m.P_aggr_g_down[i, t].value +
                                          prices_g_sec[t] * ratio_up[t] * m.P_aggr_g_up[i, t].value
                                          for i in range(0, len(load_in_bus_g))) for t in range(0, h))
                                   +
                                   sum(sum(-prices_hy_sec[t] * ratio_down[t] * m.P_aggr_hy_down[i, t].value +
                                           prices_hy_sec[t] * ratio_up[t] * m.P_aggr_hy_up[i, t].value
                                           for i in range(0, len(load_in_bus_g))) for t in range(0, h))
                                   +
                                   sum(sum(-prices_H2O_sec[t] * ratio_down[t] * m.D_P2G_E[i, t].value * c_H2O +
                                           prices_H2O_sec[t] * ratio_up[t] * m.U_P2G_E[i, t].value * c_H2O
                                           for i in range(0, len(load_in_bus_w))) for t in range(0, h))
                                   +
                                   sum(sum(-prices_O2_sec[t] * ratio_down[t] * m.D_P2G_E[i, t].value * c_O2 +
                                           prices_O2_sec[t] * ratio_up[t] * m.U_P2G_E[i, t].value * c_O2
                                           for i in range(0, len(load_in_bus_w))) for t in range(0, h))
                                   )/1000
                                  +
                                  sum(sum(m.P_chp_w[n, t].value + m.P_chp_w_up[n, t].value * ratio_up[t] - m.P_chp_w_down[n, t].value * ratio_down[t]
                                          for n in range(0, len(load_in_bus_h))) for t in
                                      range(0, h)) / 1000 * 0.2 * price_CO2
                                  +
                                  (sum(sum(m.P_chp_h[n, t].value + m.P_chp_h_up[n, t].value * ratio_up[t] - m.P_chp_h_down[n, t].value * ratio_down[t]
                                           for n in range(0, len(load_in_bus_h))) for t in range(0, h)) / 1000 * 0.2
                                   - CO2_license) * price_CO2)

    print("Cost w", (sum(m.P_aggr_w[t].value * prices_w[t] for t in range(0, h)) ))
    print("Cost g", (sum(m.P_aggr_g[t].value * prices_g[t] for t in range(0, h)) ))
    print("")

    print("Cost band", sum(sum(prices_w_sec[t] * (m.P_aggr_w_up[i, t].value + m.P_aggr_w_down[i, t].value)
                                          for i in range(0, len(load_in_bus_w))) for t in range(0, h)))
    print("Cost sec down", sum(sum(prices_w_down[t] * ratio_down[t] * m.P_aggr_w_down[i, t].value
                                        for i in range(0, len(load_in_bus_w))) for t in range(0, h)))
    print("Cost sec up", sum(sum( prices_w_up[t] * ratio_up[t] * m.P_aggr_w_up[i, t].value
                                        for i in range(0, len(load_in_bus_w))) for t in range(0, h)))
    print("Cost g down", sum(sum(-prices_g_sec[t] * ratio_down[t] * m.P_aggr_g_down[i, t].value
                                          for i in range(0, len(load_in_bus_g))) for t in range(0, h)))
    print("Cost g up", sum(sum( prices_g_sec[t] * ratio_up[t] * m.P_aggr_g_up[i, t].value
                                          for i in range(0, len(load_in_bus_g))) for t in range(0, h)))
    print("Cost hy", (sum(m.P_aggr_hy[t].value * prices_hy[t] for t in range(0, h)) ))


    print("Cost hy sec", sum(sum(-prices_hy_sec[t] * ratio_down[t] * m.P_aggr_hy_down[i, t].value +
                                           prices_hy_sec[t] * ratio_up[t] * m.P_aggr_hy_up[i, t].value
                                           for i in range(0, len(load_in_bus_g))) for t in range(0, h)))


    print("Cost O2", sum(sum( - prices_O2[t] * m.P_P2G_E[j, t].value * c_O2
                                          for j in range(0, len(load_in_bus_w))) for t in range(0, h)))
    print("Cost water", sum(sum(prices_H2O[t] * m.P_P2G_E[j, t].value * c_H2O
                                          for j in range(0, len(load_in_bus_w))) for t in range(0, h)))

    print("Cost water down up", sum(sum(-prices_H2O_sec[t] * ratio_down[t] * m.D_P2G_E[i, t].value * c_H2O +
                                           prices_H2O_sec[t] * ratio_up[t] * m.U_P2G_E[i, t].value * c_H2O
                                           for i in range(0, len(load_in_bus_w))) for t in range(0, h)))
    print("Cost O2 down up", sum(sum(-prices_O2_sec[t] * ratio_down[t] * m.D_P2G_E[i, t].value * c_O2 +
                                           prices_O2_sec[t] * ratio_up[t] * m.U_P2G_E[i, t].value * c_O2
                                           for i in range(0, len(load_in_bus_w))) for t in range(0, h)))

    print("Cost CO2", sum(sum(m.P_chp_w[n, t].value + m.P_chp_w_up[n, t].value * ratio_up[t] - m.P_chp_w_down[n, t].value * ratio_down[t]
                                          for n in range(0, len(load_in_bus_h))) for t in
                                      range(0, h)) / 1000 * 0.2 * price_CO2)
    print("Cost CO2 2", (sum(sum(m.P_chp_h[n, t].value + m.P_chp_h_up[n, t].value * ratio_up[t] - m.P_chp_h_down[n, t].value * ratio_down[t]
                                           for n in range(0, len(load_in_bus_h))) for t in range(0, h)) / 1000 * 0.2
                                   - CO2_license) * price_CO2)


    return m, time_h, cost




def solve_using_ipopt(m):
    solver_name = 'ipopt'
    #solverpath_exe = 'C:\\Users\\amcoelho\\.conda\\pkgs\\ipopt-3.11.1-2\\Library\\bin\\ipopt'
    logging.getLogger('pyomo.core').setLevel(logging.ERROR)
    solver = SolverFactory('ipopt')
    solver.options['max_iter'] = 1000000
    results = solver.solve(m, tee=False)

    return m, results


def print_status_flow(results, time_h, t):
    if (results.solver.status == SolverStatus.ok) and (
            results.solver.termination_condition == TerminationCondition.optimal):
        print("Flow optimized - hour", t)
        time_h['dso_w'].append(results.solver.time)
    else:
        print("Did not converge")

        return time_h



def optimization_dso(m, m0, t, pi0, ro0, load_in_bus_w, time_h):
    pi_w = pi0['w']
    ro = ro0['ro']
    k = 0
    m.value = Objective(expr=
                        sum(pi_w[i][t] * (m0.P_dso[k, i, t].value - m.P_dso[k, i, 0])
                            for i in range(0, len(load_in_bus_w))) +
                        ro / 2 * (sum(((m0.P_dso[k, i, t].value - m.P_dso[k, i, 0])) ** 2
                                            for i in range(0, len(load_in_bus_w))))
                        , sense=minimize)

    solver_name = 'ipopt'
    solverpath_exe = 'C:\\Users\\amcoelho\\.conda\\pkgs\\ipopt-3.11.1-2\\Library\\bin\\ipopt'
    solver = SolverFactory(solver_name, executable=solverpath_exe)
    solver.options['max_iter'] = 1000000

    results = solver.solve(m, tee=False)
    if (results.solver.status == SolverStatus.ok) and ( results.solver.termination_condition == TerminationCondition.optimal):
        print("Flow optimized - hour", t)
        time_h['dso_w'].append(results.solver.time)
    else:
        print("Did not converge")



    return m, time_h



def optimization_dso_up(m, m0, t, pi0, ro0, load_in_bus_w, time_h):
    pi_w_up = pi0['w_up']
    ro = ro0['ro']
    k = 0

    m.value = Objective(expr=
                        sum(pi_w_up[i][t] * (m0.P_dso_up[k, i, t].value - m.P_dso_up[k, i, 0])
                            for i in range(0, len(load_in_bus_w))) +
                        ro / 2 * (sum(((m0.P_dso_up[k, i, t].value - m.P_dso_up[k, i, 0])) ** 2
                                            for i in range(0, len(load_in_bus_w))))
                        , sense=minimize)

    m, results = solve_using_ipopt(m)
    time_h = print_status_flow(results, time_h, t)

    return m, time_h




def optimization_dso_down(m, m0, t, pi0, ro0, load_in_bus_w, time_h):
    pi_w_down = pi0['w_down']
    ro = ro0['ro']
    k = 0

    m.value = Objective(expr=
                        sum(pi_w_down[i][t] * (m0.P_dso_down[k, i, t].value - m.P_dso_down[k, i, 0])
                            for i in range(0, len(load_in_bus_w))) +
                        ro / 2 * (sum(((m0.P_dso_down[k, i, t].value - m.P_dso_down[k, i, 0])) ** 2
                                            for i in range(0, len(load_in_bus_w))))
                        , sense=minimize)

    m, results = solve_using_ipopt(m)
    time_h = print_status_flow(results, time_h, t)

    return m, time_h




def optimization_dso_gas(m, m0, t, pi0, ro0, load_in_bus_g, b_prints, time_h):
    ro = ro0['ro']
    pi_g = pi0['g']
    k = 0

    m.value = Objective(expr=
                        sum(pi_g[i][t] * (m0.P_dso_gas[k, i, t].value - m.P_dso_gas[i, 0])
                            for i in range(0, len(load_in_bus_g))) +
                        ro / 2 * (sum((((m0.P_dso_gas[k, i, t].value) ** 2
                                              - 2 * (m0.P_dso_gas[k, i, t].value) * (m.P_dso_gas[i, 0]))
                                             + (m.P_dso_gas[i, 0])** 2)
                                            for i in range(0, len(load_in_bus_g))))
                        , sense=minimize)

    solver = SolverFactory('ipopt')
    solver.options['max_iter'] = 1000000

    results1 = solver.solve(m, tee=False, load_solutions=False)
    if (results1.solver.status == SolverStatus.ok) and (
            results1.solver.termination_condition == TerminationCondition.optimal):
        results = m.solutions.load_from(results1)
        time_h['dso_g'].append(results1.solver.time)

    if b_prints:
        if (results1.solver.status == SolverStatus.ok) and (
                results1.solver.termination_condition == TerminationCondition.optimal):
            print("Flow optimized - hour", t)
        else:
            print("Did not converge")

    return m, time_h


def optimization_dso_gas_up(m, m0, t, pi0, ro0, load_in_bus_g, b_prints, time_h):
    k = 0
    ro = ro0['ro']
    pi_g_up = pi0['g_up']
    pi_hy_up = pi0['hy_up']

    for i in range(0, len(load_in_bus_g)):
        1
        #print("hy", t, i, m0.P_hy_up[i, t].value, m0.P_dso_hy_up[k, i, t].value)
        #print("hy", t, i, m0.P_hy_down[i, t].value, m0.P_dso_hy_down[k, i, t].value)

    v = 1
    m.value = Objective(expr=
                        sum(pi_g_up[i][t] * (m0.P_dso_gas_up[k, i, t].value * v - m.P_dso_gas_up[i, 0] * v)
                            for i in range(0, len(load_in_bus_g))) +
                        ro / 2 * (sum((((m0.P_dso_gas_up[k, i, t].value * v) ** 2
                                              - 2 * (m0.P_dso_gas_up[k, i, t].value * v) * (m.P_dso_gas_up[i, 0] * v))
                                             + (m.P_dso_gas_up[i, 0] * v)** 2)
                                            for i in range(0, len(load_in_bus_g)))) +

                        sum(pi_hy_up[i][t] * (m0.P_dso_hy_up[k, i, t].value * v - m.P_dso_hy_up[i, 0] * v)
                            for i in range(0, len(load_in_bus_g))) +
                        ro / 2 * (sum((((m0.P_dso_hy_up[k, i, t].value * v) ** 2
                                        - 2 * (m0.P_dso_hy_up[k, i, t].value * v) * (m.P_dso_hy_up[i, 0] * v))
                                       + (m.P_dso_hy_up[i, 0] * v) ** 2)
                                      for i in range(0, len(load_in_bus_g))))

                        , sense=minimize)

    solver = SolverFactory('ipopt')
    solver.options['max_iter'] = 1000000


    results1 = solver.solve(m, tee=False, load_solutions=False)
    if (results1.solver.status == SolverStatus.ok) and (
            results1.solver.termination_condition == TerminationCondition.optimal):
        results = m.solutions.load_from(results1)
        time_h['dso_g'].append(results1.solver.time)

    if b_prints:
        if (results1.solver.status == SolverStatus.ok) and (
                results1.solver.termination_condition == TerminationCondition.optimal):
            print("Flow optimized - hour", t)
        else:
            print("Did not converge")

    return m, time_h



def optimization_dso_gas_down(m, m0, t, pi0, ro0, load_in_bus_g, b_prints, time_h):
    k = 0
    ro = ro0['ro']
    pi_g_down = pi0['g_down']
    pi_hy_down = pi0['hy_down']

    for i in range(0, len(load_in_bus_g)):
        1
        #print("hy", t, i, m0.P_hy_down[i, t].value, m0.P_dso_hy_down[k, i, t].value)
        #print(t, i, m0.P_dso_gas_down[0, i, t].value)
    v = 1
    m.value = Objective(expr=
                        sum(pi_g_down[i][t] * (m0.P_dso_gas_down[k, i, t].value * v - m.P_dso_gas_down[i, 0] * v)
                            for i in range(0, len(load_in_bus_g))) +
                        ro / 2 * (sum((((m0.P_dso_gas_down[k, i, t].value * v) ** 2
                                              - 2 * (m0.P_dso_gas_down[k, i, t].value * v) * (m.P_dso_gas_down[i, 0] * v))
                                             + (m.P_dso_gas_down[i, 0] * v)** 2)
                                            for i in range(0, len(load_in_bus_g))))
                        +
                        sum(pi_hy_down[i][t] * (m0.P_dso_hy_down[k, i, t].value * v - m.P_dso_hy_down[i, 0] * v)
                            for i in range(0, len(load_in_bus_g))) +
                        ro / 2 * (sum((((m0.P_dso_hy_down[k, i, t].value * v) ** 2
                                        - 2 * (m0.P_dso_hy_down[k, i, t].value * v) * (m.P_dso_hy_down[i, 0] * v))
                                       + (m.P_dso_hy_down[i, 0] * v) ** 2)
                                      for i in range(0, len(load_in_bus_g))))
                        , sense=minimize)

    solver = SolverFactory('ipopt')
    solver.options['max_iter'] = 100000000

    results1 = solver.solve(m, tee=False, load_solutions=False)
    if (results1.solver.status == SolverStatus.ok) and (
            results1.solver.termination_condition == TerminationCondition.optimal):
        results = m.solutions.load_from(results1)
        time_h['dso_g'].append(results1.solver.time)

    if b_prints:
        if (results1.solver.status == SolverStatus.ok) and (
                results1.solver.termination_condition == TerminationCondition.optimal):
            print("Flow optimized - hour", t)
        else:
            print("Did not converge")

    return m, time_h




def optimization_dso_heat(m, m0, t, pi0, ro0, load_in_bus_h, b_prints, time_h):
    ro = ro0['ro']
    pi_h = pi0['h']
    k = 0

    for i in range(0, len(load_in_bus_h)):
        print(i, m0.P_dso_heat[k, i, t].value)

    m.value = Objective(expr=
                            sum(pi_h[i][t] * (m0.P_dso_heat[k, i, t].value - m.P_dso_heat[i, 0]) for i in range(0, len(load_in_bus_h))) +
                              ro/2 * (sum(((m0.P_dso_heat[k, i, t].value - m.P_dso_heat[i, 0])) ** 2 for i in range(0, len(load_in_bus_h))))
                        , sense=minimize)

    solver = SolverFactory('ipopt')
    solver.options['max_iter'] = 1000000

    results1 = solver.solve(m, tee=False, load_solutions=False)

    if (results1.solver.status == SolverStatus.ok) and \
            (results1.solver.termination_condition == TerminationCondition.optimal):
        results = m.solutions.load_from(results1)
        if b_prints: print("Flow optimized - hour", t)
        flag_error = 0
        time_h['dso_h'].append(results1.solver.time)
    else:
        if b_prints: print("Repeat solving...")
        flag_error = 1


    return m, flag_error, time_h


def optimization_dso_heat_up(m, m0, t, pi0, ro0, load_in_bus_h, b_prints, time_h):
    ro = ro0['ro']
    pi_h_up = pi0['h_up']
    k = 0

    m.value = Objective(expr=
                            sum(pi_h_up[i][t] * (m0.P_dso_heat_up[k, i, t].value - m.P_dso_heat_up[i, 0]) for i in range(0, len(load_in_bus_h))) +
                              ro/2 * (sum(((m0.P_dso_heat_up[k, i, t].value - m.P_dso_heat_up[i, 0])) ** 2 for i in range(0, len(load_in_bus_h))))
                        , sense=minimize)

    solver = SolverFactory('ipopt')
    solver.options['max_iter'] = 1000000

    results1 = solver.solve(m, tee=False, load_solutions=False)

    if (results1.solver.status == SolverStatus.ok) and \
            (results1.solver.termination_condition == TerminationCondition.optimal):
        results = m.solutions.load_from(results1)
        if b_prints: print("Flow optimized - hour", t)
        flag_error = 0
        time_h['dso_h'].append(results1.solver.time)
    else:
        if b_prints: print("Repeat solving...")
        flag_error = 1

    return m, flag_error, time_h


def optimization_dso_heat_down(m, m0, t, pi0, ro0, load_in_bus_h, b_prints, time_h):
    ro = ro0['ro']
    pi_h_down = pi0['h_down']
    k = 0

    m.value = Objective(expr=
                            sum(pi_h_down[i][t] * (m0.P_dso_heat_down[k, i, t].value - m.P_dso_heat_down[i, 0]) for i in range(0, len(load_in_bus_h))) +
                              ro/2 * (sum(((m0.P_dso_heat_down[k, i, t].value - m.P_dso_heat_down[i, 0])) ** 2 for i in range(0, len(load_in_bus_h))))
                        , sense=minimize)

    solver = SolverFactory('ipopt')
    solver.options['max_iter'] = 1000000

    results1 = solver.solve(m, tee=False, load_solutions=False)

    if (results1.solver.status == SolverStatus.ok) and \
            (results1.solver.termination_condition == TerminationCondition.optimal):
        results = m.solutions.load_from(results1)
        if b_prints: print("Flow optimized - hour", t)
        flag_error = 0
        time_h['dso_h'].append(results1.solver.time)
    else:
        if b_prints: print("Repeat solving...")
        flag_error = 1


    return m, flag_error, time_h

