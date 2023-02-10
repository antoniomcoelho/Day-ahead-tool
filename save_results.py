from xlrd import *
from numpy import *
from pyomo.environ import *
from math import *
import pandas as pd
import xlwt
from time import *


def save_results_aggr(m):

    DF_P_w = pd.DataFrame()
    DF_P_g = pd.DataFrame()
    DF_P_aggr_w_up = pd.DataFrame()
    DF_P_aggr_w_down = pd.DataFrame()
    for v in m.component_objects(Var, active=True):
        for index in v:  # index[0, j, i, t]; v = ve_in_hp,...
            if v.name == 'P_w':
                DF_P_w.at[index] = value(v[index])
            if v.name == 'P_g':
                DF_P_g.at[index] = value(v[index])
            if v.name == 'P_aggr_w_up':
                DF_P_aggr_w_up.at[index] = value(v[index])
            if v.name == 'P_aggr_w_down':
                DF_P_aggr_w_down.at[index] = value(v[index])

    print("...Save data...")
    with pd.ExcelWriter("Bids_aggregator.xls") as writer:
        DF_P_w.to_excel(writer, sheet_name='Electricity energy bid')
        DF_P_g.to_excel(writer, sheet_name='Gas energy bid')
        DF_P_aggr_w_up.to_excel(writer, sheet_name='Upward secondary band')
        DF_P_aggr_w_down.to_excel(writer, sheet_name='Downward secondary band')


    return 0



def save_criteria(criteria, criteria_dual):

    criteria = pd.DataFrame(criteria)

    criteria_dual = pd.DataFrame(criteria_dual)


    print("...Save criteria...")
    with pd.ExcelWriter("criteria_values.xls") as writer:
        criteria.to_excel(writer, sheet_name='criteria')

        criteria_dual.to_excel(writer, sheet_name='criteria_dual')


    return 0


def save_results_file(m, load_initial, h, load_in_bus_w, load_in_bus_g, load_in_bus_h, buildings, prices, fuel_station0,
                      results_m1, branch, results_m2, branch_g, results_m3, branch_h):

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

    flag_gas = 1
    flag_heat = 1

    # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    #       DA bids
    # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    book = xlwt.Workbook()
    sh1 = book.add_sheet("Bids")
    sh1.write(1, 0, "Energy bid")
    sh1.write(2, 0, "Upward bid")
    sh1.write(3, 0, "Downward bid")
    sh1.write(4, 0, "Pg")
    sh1.write(5, 0, "Phy")
    sh1.write(6, 0, "CO2")
    sh1.write(7, 0, "Water")
    sh1.write(8, 0, "Oxygen")
    sh1.write(9, 0, "Guarantees")
    for t in range(0, h):
        sh1.write(0, t + 1, t)
        sh1.write(1, t + 1, m.P_aggr_w[t].value/1000)
        sh1.write(2, t + 1, sum(m.P_aggr_w_up[i, t].value for i in range(0, len(load_in_bus_w)))/1000)
        sh1.write(3, t + 1, sum(m.P_aggr_w_down[i, t].value for i in range(0, len(load_in_bus_w)))/1000)
        sh1.write(4, t + 1, sum(m.P_g[n, t].value for n in range(0, len(load_in_bus_g)))/1000)
        sh1.write(5, t + 1, sum(m.P_hy[n, t].value for n in range(0, len(load_in_bus_g)))/1000)
        sh1.write(6, t + 1, (sum(m.P_chp_w[n, t].value for n in range(0, len(load_in_bus_h))) ) / 1000 * 0.2)
        sh1.write(9, t + 1, (m.P_renew[t].value * 0.5 )/ 1000)

    sh1.write(7, 1, sum(sum(m.P_P2G_E[j, t].value * c_H2O for j in range(0, len(load_in_bus_w))) for t in range(0, h)))
    sh1.write(8, 1, sum(sum(m.P_P2G_E[j, t].value * c_O2 for j in range(0, len(load_in_bus_w))) for t in range(0, h)))
    #sh1.write(9, 1, sum(m.P_renew[t].value * 0.5 for t in range(0, h)))/1000

    # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    #       DMERs
    # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

    sh1 = book.add_sheet("DMERs sum")

    sh1.write(0, 0, "Psto demand")
    sh1.write(1, 0, "Psto supply")
    sh1.write(2, 0, "Psto up")
    sh1.write(3, 0, "Psto down")
    sh1.write(4, 0, "Psto gas")
    sh1.write(5, 0, "Psto hy")
    n_line = 6
    sh1.write(n_line, 0, "Php demand")
    n_line = n_line + 1
    sh1.write(n_line, 0, "Php supply")
    n_line = n_line + 1
    sh1.write(n_line, 0, "Php up")
    n_line = n_line + 1
    sh1.write(n_line, 0, "Php down")
    n_line = n_line + 1
    sh1.write(n_line, 0, "Php gas")
    n_line = n_line + 1
    sh1.write(n_line, 0, "Php hy")

    n_line = n_line + 1
    sh1.write(n_line, 0, "Ppv demand")
    n_line = n_line + 1
    sh1.write(n_line, 0, "Ppv supply")
    n_line = n_line + 1
    sh1.write(n_line, 0, "Ppv up")
    n_line = n_line + 1
    sh1.write(n_line, 0, "Ppv down")
    n_line = n_line + 1
    sh1.write(n_line, 0, "Ppv gas")
    n_line = n_line + 1
    sh1.write(n_line, 0, "Ppv hy")

    n_line = n_line + 1
    sh1.write(n_line, 0, "CHP_w demand")
    n_line = n_line + 1
    sh1.write(n_line, 0, "CHP_w supply")
    n_line = n_line + 1
    sh1.write(n_line, 0, "CHP_w up")
    n_line = n_line + 1
    sh1.write(n_line, 0, "CHP_w down")
    n_line = n_line + 1
    sh1.write(n_line, 0, "CHP_g")
    n_line = n_line + 1
    sh1.write(n_line, 0, "CHP_hy")

    n_line = n_line + 1
    sh1.write(n_line, 0, "P_P2G_E demand")
    n_line = n_line + 1
    sh1.write(n_line, 0, "P_P2G_E supply")
    n_line = n_line + 1
    sh1.write(n_line, 0, "U_P2G_E")
    n_line = n_line + 1
    sh1.write(n_line, 0, "D_P2G_E")
    n_line = n_line + 1
    sh1.write(n_line, 0, "P_P2G_g")
    n_line = n_line + 1
    sh1.write(n_line, 0, "P_P2G_hy")

    n_line = n_line + 1
    sh1.write(n_line, 0, "P_FC_E demand")
    n_line = n_line + 1
    sh1.write(n_line, 0, "P_FC_E supply")
    n_line = n_line + 1
    sh1.write(n_line, 0, "U_FC_E")
    n_line = n_line + 1
    sh1.write(n_line, 0, "D_FC_E")
    n_line = n_line + 1
    sh1.write(n_line, 0, "P_FC_E gas")
    n_line = n_line + 1
    sh1.write(n_line, 0, "P_FC_E hy")

    t = 0
    n_line = 0

    sh1.write(n_line, t + 1, sum(sum(m.P_sto_ch[n, t].value for n in range(0, len(buildings))) for t in range(0, h))/1000)
    n_line = n_line + 1
    sh1.write(n_line, t + 1, sum(sum(- m.P_sto_dis[n, t].value for n in range(0, len(buildings))) for t in range(0, h))/1000)
    n_line = n_line + 1
    sh1.write(n_line, t + 1, sum(sum(m.P_sto_up[n, t].value for n in range(0, len(buildings))) for t in range(0, h))/1000)
    n_line = n_line + 1
    sh1.write(n_line, t + 1, sum(sum(m.P_sto_down[n, t].value for n in range(0, len(buildings))) for t in range(0, h))/1000)
    n_line = n_line + 1
    sh1.write(n_line, t + 1, 0)
    n_line = n_line + 1
    sh1.write(n_line, t + 1, 0)

    n_line = n_line + 1
    sh1.write(n_line, t + 1, sum(sum(m.P_hp[n, t].value for n in range(0, len(buildings))) for t in range(0, h))/1000)
    n_line = n_line + 1
    sh1.write(n_line, t + 1, 0)
    n_line = n_line + 1
    sh1.write(n_line, t + 1, sum(sum(m.P_hp_up[n, t].value for n in range(0, len(buildings))) for t in range(0, h))/1000)
    n_line = n_line + 1
    sh1.write(n_line, t + 1, sum(sum(m.P_hp_down[n, t].value for n in range(0, len(buildings))) for t in range(0, h))/1000)
    n_line = n_line + 1
    sh1.write(n_line, t + 1, 0)
    n_line = n_line + 1
    sh1.write(n_line, t + 1, 0)


    n_line = n_line + 1
    sh1.write(n_line, t + 1, 0)
    n_line = n_line + 1
    sh1.write(n_line, t + 1, sum(sum(-m.P_PV[n, t].value for n in range(0, len(buildings))) for t in range(0, h))/1000)
    n_line = n_line + 1
    sh1.write(n_line, t + 1, sum(sum(m.P_PV_up[n, t].value for n in range(0, len(buildings))) for t in range(0, h))/1000)
    n_line = n_line + 1
    sh1.write(n_line, t + 1, sum(sum(m.P_PV_down[n, t].value for n in range(0, len(buildings))) for t in range(0, h))/1000)
    n_line = n_line + 1
    sh1.write(n_line, t + 1, 0)
    n_line = n_line + 1
    sh1.write(n_line, t + 1, 0)




    n_line = n_line + 1
    sh1.write(n_line, t + 1, 0)
    n_line = n_line + 1
    sh1.write(n_line, t + 1, sum(sum(-m.P_chp_w[n, t].value for n in range(26, 28)) for t in range(0, h))/1000)
    n_line = n_line + 1
    sh1.write(n_line, t + 1, sum(sum(m.P_chp_w_up[n, t].value for n in range(26, 28)) for t in range(0, h))/1000)
    n_line = n_line + 1
    sh1.write(n_line, t + 1, sum(sum(m.P_chp_w_down[n, t].value for n in range(26, 28)) for t in range(0, h))/1000)
    n_line = n_line + 1
    sh1.write(n_line, t + 1, sum(sum(m.P_chp_g[n, t].value for n in range(26, 28)) for t in range(0, h))/1000)
    n_line = n_line + 1
    sh1.write(n_line, t + 1, 0)


    n_line = n_line + 1
    sh1.write(n_line, t + 1, sum(sum(m.P_P2G_E[n, t].value for n in range(0, len(load_in_bus_w))) for t in range(0, h))/1000)
    n_line = n_line + 1
    sh1.write(n_line, t + 1, 0)
    n_line = n_line + 1
    sh1.write(n_line, t + 1, sum(sum(m.U_P2G_E[n, t].value for n in range(0, len(load_in_bus_w))) for t in range(0, h))/1000)
    n_line = n_line + 1
    sh1.write(n_line, t + 1, sum(sum(m.D_P2G_E[n, t].value for n in range(0, len(load_in_bus_w))) for t in range(0, h))/1000)
    n_line = n_line + 1
    sh1.write(n_line, t + 1, 0)
    n_line = n_line + 1
    sh1.write(n_line, t + 1, sum(sum(m.P_P2G_net_hy[n, t].value for n in range(0, len(load_in_bus_w))) for t in range(0, h))/1000)


    n_line = n_line + 1
    sh1.write(n_line, t + 1, 0)
    n_line = n_line + 1
    sh1.write(n_line, t + 1, sum(sum(-m.P_FC_E[n, t].value for n in range(0, len(load_in_bus_w))) for t in range(0, h))/1000)
    n_line = n_line + 1
    sh1.write(n_line, t + 1, sum(sum(m.U_FC_E[n, t].value for n in range(0, len(load_in_bus_w))) for t in range(0, h))/1000)
    n_line = n_line + 1
    sh1.write(n_line, t + 1, sum(sum(m.D_FC_E[n, t].value for n in range(0, len(load_in_bus_w))) for t in range(0, h))/1000)
    n_line = n_line + 1
    sh1.write(n_line, t + 1, 0)
    n_line = n_line + 1
    sh1.write(n_line, t + 1, 0)

    sh1 = book.add_sheet("DMERs")

    sh1.write(1, 0, "Php")
    sh1.write(2, 0, "Php up")
    sh1.write(3, 0, "Php down")
    sh1.write(4, 0, "Ppv")
    sh1.write(5, 0, "Ppv up")
    sh1.write(6, 0, "Ppv down")
    sh1.write(7, 0, "Psto")
    sh1.write(8, 0, "Psto up")
    sh1.write(9, 0, "Psto down")
    sh1.write(10, 0, "CHP_g")
    sh1.write(11, 0, "CHP_w")
    sh1.write(12, 0, "CHP_w up")
    sh1.write(13, 0, "CHP_w down")
    sh1.write(14, 0, "P_P2G_E")
    sh1.write(15, 0, "U_P2G_E")
    sh1.write(16, 0, "D_P2G_E")
    sh1.write(17, 0, "P_P2G_Hy")
    sh1.write(18, 0, "P_FC_E")
    sh1.write(19, 0, "U_FC_E")
    sh1.write(20, 0, "D_FC_E")
    for t in range(0, h):
        sh1.write(0, t + 1, t)
        sh1.write(1, t + 1, sum(m.P_hp[n, t].value for n in range(0, len(buildings)))/1000)
        sh1.write(2, t + 1, sum(m.P_hp_up[n, t].value for n in range(0, len(buildings)))/1000)
        sh1.write(3, t + 1, sum(m.P_hp_down[n, t].value for n in range(0, len(buildings)))/1000)
        sh1.write(4, t + 1, sum(-m.P_PV[n, t].value for n in range(0, len(buildings)))/1000)
        sh1.write(5, t + 1, sum(m.P_PV_up[n, t].value for n in range(0, len(buildings)))/1000)
        sh1.write(6, t + 1, sum(m.P_PV_down[n, t].value for n in range(0, len(buildings)))/1000)
        sh1.write(7, t + 1, sum(m.P_sto_ch[n, t].value - m.P_sto_dis[n, t].value for n in range(0, len(buildings)))/1000)
        sh1.write(8, t + 1, sum(m.P_sto_up[n, t].value for n in range(0, len(buildings)))/1000)
        sh1.write(9, t + 1, sum(m.P_sto_down[n, t].value for n in range(0, len(buildings)))/1000)
        sh1.write(10, t + 1, sum(m.P_chp_g[n, t].value for n in range(26, 28))/1000)
        sh1.write(11, t + 1, sum(-m.P_chp_w[n, t].value for n in range(26, 28))/1000)
        sh1.write(12, t + 1, sum(m.P_chp_w_up[n, t].value for n in range(26, 28))/1000)
        sh1.write(13, t + 1, sum(m.P_chp_w_down[n, t].value for n in range(26, 28))/1000)
        sh1.write(14, t + 1, sum(m.P_P2G_E[n, t].value for n in range(0, len(load_in_bus_w)))/1000)
        sh1.write(15, t + 1, sum(m.U_P2G_E[n, t].value for n in range(0, len(load_in_bus_w)))/1000)
        sh1.write(16, t + 1, sum(m.D_P2G_E[n, t].value for n in range(0, len(load_in_bus_w)))/1000)
        sh1.write(17, t + 1, sum(-m.P_P2G_net_hy[n, t].value for n in range(0, len(load_in_bus_w)))/1000)
        sh1.write(18, t + 1, sum(-m.P_FC_E[n, t].value for n in range(0, len(load_in_bus_w)))/1000)
        sh1.write(19, t + 1, sum(m.U_FC_E[n, t].value for n in range(0, len(load_in_bus_w)))/1000)
        sh1.write(20, t + 1, sum(m.D_FC_E[n, t].value for n in range(0, len(load_in_bus_w)))/1000)

    # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    #       Costs
    # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    sh1 = book.add_sheet("Costs")

    sh1.write(0, 0, "Electricity energy")
    sh1.write(1, 0, "Electricity band")
    sh1.write(2, 0, "Gas")
    sh1.write(3, 0, "Hydrogen")
    sh1.write(4, 0, "Water")
    sh1.write(5, 0, "Oxygen")
    sh1.write(6, 0, "Guarantees of origin")
    sh1.write(7, 0, "Carbon")

    sh1.write(0, 1, sum(m.P_aggr_w[t].value * prices_w[t] for t in range(0, h))/1000)
    sh1.write(1, 1, sum(sum(-prices_w_sec[t] * (m.P_aggr_w_up[i, t].value + m.P_aggr_w_down[i, t].value) for i in range(0, len(load_in_bus_w))) for t in range(0, h))/1000)
    sh1.write(2, 1, sum(m.P_aggr_g[t].value * prices_g[t] for t in range(0, h))/1000)
    sh1.write(3, 1, (sum(-m.P_aggr_hy[t].value * prices_hy[t]  for t in range(0, h)))/1000)
    sh1.write(4, 1, sum(sum(prices_H2O[t] * m.P_P2G_E[j, t].value * c_H2O for j in range(0, len(load_in_bus_w))) for t in range(0, h))/1000)
    sh1.write(5, 1, sum(sum(- prices_O2[t] * m.P_P2G_E[j, t].value * c_O2 for j in range(0, len(load_in_bus_w))) for t in range(0, h))/1000)
    sh1.write(6, 1, sum(m.P_renew[t].value * 0.5 for t in range(0, h))/1000)
    sh1.write(7, 1, sum(sum(m.P_chp_w[n, t].value + m.P_chp_w_up[n, t].value * ratio_up[t] - m.P_chp_w_down[n, t].value * ratio_down[t]
                            for n in range(0, len(load_in_bus_h))) for t in range(0, h)) / 1000 * 0.2 * price_CO2
                        +
                        (sum(sum(m.P_chp_h[n, t].value + m.P_chp_h_up[n, t].value * ratio_up[t] - m.P_chp_h_down[n, t].value * ratio_down[t]
                                 for n in range(0, len(load_in_bus_h))) for t in range(0, h)) / 1000 * 0.2 - CO2_license) * price_CO2)


    # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    #       Carbon allowances
    # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    sh1 = book.add_sheet("Carbon")

    sh1.write(0, 0, "Electricity")
    sh1.write(1, 0, "Heat")
    sh1.write(2, 0, "Free")
    sh1.write(3, 0, "Total")

    c_w = sum(sum(m.P_chp_w[n, t].value + m.P_chp_w_up[n, t].value * ratio_up[t] - m.P_chp_w_down[n, t].value * ratio_down[t]
                            for n in range(0, len(load_in_bus_h))) for t in range(0, h)) / 1000 * 0.2
    c_h = sum(sum(m.P_chp_h[n, t].value + m.P_chp_h_up[n, t].value * ratio_up[t] - m.P_chp_h_down[n, t].value * ratio_down[t]
                                 for n in range(0, len(load_in_bus_h))) for t in range(0, h)) / 1000 * 0.2
    sh1.write(0, 1, c_w)
    sh1.write(1, 1, c_h)
    sh1.write(2, 1, 2.8)
    sh1.write(3, 1, c_w + c_h - 2.8)

    # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    #      Voltage
    # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

    sh1 = book.add_sheet("Voltage")

    sh1.write(0, 0, "Electricity energy")
    for t in range(0, h):
        sh1.write(0, t + 1, t)
    for k in range(0, 3):
        for j in range(0, len(branch)):
            sh1.write(1 + j + k * len(branch), 0, j)

    for k in range(0, 3):
        for j in range(0, len(branch)):
            for t in range(0, h):
                print(k, j + k * len(branch), t)
                sh1.write(1 + j + k * len(branch), t + 1, sqrt(results_m1[k][t].V[0, j, 0].value))

    sh1 = book.add_sheet("V tot")

    sh1.write(0, 1, "Energy")
    sh1.write(0, 2, "Upward")
    sh1.write(0, 3, "Downward")
    sh1.write(1, 0, "Number")
    sh1.write(2, 0, "Max")
    sh1.write(3, 0, "Min")

    results_m1_new = []
    for i in range(0, len(results_m1)):
        results_new = []
        for j in range(0, len(results_m1[0])):
            for k in range(0, len(branch)):
                results_new.append(sqrt(results_m1[i][j].V[0, k, 0].value))
        results_m1_new.append(results_new)

    sh1.write(1, 1, sum((i > 1.05 or i < 0.95) for i in results_m1_new[0]))
    sh1.write(1, 2, sum((i > 1.05 or i < 0.95) for i in results_m1_new[1]))
    sh1.write(1, 3, sum((i > 1.05 or i < 0.95) for i in results_m1_new[2]))
    sh1.write(2, 1, max(results_m1_new[0]))
    sh1.write(2, 2, max(results_m1_new[1]))
    sh1.write(2, 3, max(results_m1_new[2]))
    sh1.write(3, 1, min(results_m1_new[0]))
    sh1.write(3, 2, min(results_m1_new[1]))
    sh1.write(3, 3, min(results_m1_new[2]))


    # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    #      Gas
    # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

    if flag_gas:
        sh1 = book.add_sheet("Gas")

        sh1.write(0, 0, "Gas energy")
        for t in range(0, h):
            sh1.write(0, t + 1, t)
        for k in range(0, 3):
                sh1.write(1 + 0 + k * 2, 0, "HHV")
                sh1.write(1 + 1 + k * 2, 0, "WI")

        for k in range(0, 3):
                for t in range(0, h):
                    sh1.write(1 + 0 + k * 2, t + 1, results_m2[k][t].HHV_mix[0, 0].value)
                    WI = sqrt(results_m2[k][t].HHV_mix[0, 0].value * results_m2[k][t].HHV_mix[0, 0].value / results_m2[k][t].S_mix[0, 0].value)
                    sh1.write(1 + 1 + k * 2, t + 1, WI)

        sh1 = book.add_sheet("G tot")

        sh1.write(0, 1, "Energy HHV")
        sh1.write(0, 2, "Upward")
        sh1.write(0, 3, "Downward")
        sh1.write(1, 0, "Number")
        sh1.write(2, 0, "Max")
        sh1.write(3, 0, "Min")

        results_m2_new = []
        for i in range(0, len(results_m2)):
            results_new = []
            for j in range(0, len(results_m2[0])):

                    results_new.append((results_m2[i][j].HHV_mix[0, 0].value))
            results_m2_new.append(results_new)

        sh1.write(1, 1, sum((i > 47.8 or i < 35.5) for i in results_m2_new[0]))
        sh1.write(1, 2, sum((i > 47.8 or i < 35.5) for i in results_m2_new[1]))
        sh1.write(1, 3, sum((i > 47.8 or i < 35.5) for i in results_m2_new[2]))
        sh1.write(2, 1, max(results_m2_new[0]))
        sh1.write(2, 2, max(results_m2_new[1]))
        sh1.write(2, 3, max(results_m2_new[2]))
        sh1.write(3, 1, min(results_m2_new[0]))
        sh1.write(3, 2, min(results_m2_new[1]))
        sh1.write(3, 3, min(results_m2_new[2]))


        sh1.write(7, 1, "Energy HHV")
        sh1.write(7, 2, "Upward")
        sh1.write(7, 3, "Downward")
        sh1.write(8, 0, "Number")
        sh1.write(9, 0, "Max")
        sh1.write(10, 0, "Min")

        results_m2_new = []
        for i in range(0, len(results_m2)):
            results_new = []
            for j in range(0, len(results_m2[0])):
                WI = sqrt(results_m2[i][j].HHV_mix[0, 0].value * results_m2[i][j].HHV_mix[0, 0].value / results_m2[i][j].S_mix[0, 0].value)
                results_new.append(WI)

            results_m2_new.append(results_new)

        sh1.write(8, 1, sum((i > 55.9 or i < 45.7) for i in results_m2_new[0]))
        sh1.write(8, 2, sum((i > 55.9 or i < 45.7) for i in results_m2_new[1]))
        sh1.write(8, 3, sum((i > 55.9 or i < 45.7) for i in results_m2_new[2]))
        sh1.write(9, 1, max(results_m2_new[0]))
        sh1.write(9, 2, max(results_m2_new[1]))
        sh1.write(9, 3, max(results_m2_new[2]))
        sh1.write(10, 1, min(results_m2_new[0]))
        sh1.write(10, 2, min(results_m2_new[1]))
        sh1.write(10, 3, min(results_m2_new[2]))


    # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    #      Heat
    # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

    if flag_heat:
        sh1 = book.add_sheet("Heat")

        sh1.write(0, 0, "Heat energy")
        for t in range(0, h):
            sh1.write(0, t + 1, t)
        for k in range(0, 3):
            for j in range(0, len(branch_h)):
                sh1.write(1 + j + k * len(branch_h), 0, j)

        for k in range(0, 3):
            for j in range(0, len(branch_h)):
                for t in range(0, h):
                    print(k, j + k * len(branch_h), t)
                    sh1.write(1 + j + k * len(branch_h), t + 1, results_m3[k][t].m[j, 0].value)

        sh1 = book.add_sheet("Heat tot")

        results_m3_new = []
        for i in range(0, len(results_m3)):
            results_new = []
            for j in range(0, len(results_m3[0])):
                for k in range(0, len(branch_h)):
                    results_new.append(results_m3[i][j].m[k, 0].value)
            results_m3_new.append(results_new)

        sh1.write(0, 1, "Energy")
        sh1.write(0, 2, "Upward")
        sh1.write(0, 3, "Downward")
        sh1.write(1, 0, "Number")
        sh1.write(2, 0, "Max")

        sh1.write(1, 1, sum(i > 40 for i in results_m3_new[0]))
        sh1.write(1, 2, sum(i > 40 for i in results_m3_new[1]))
        sh1.write(1, 3, sum(i > 40 for i in results_m3_new[0]))
        sh1.write(2, 1, max(results_m3_new[0]))
        sh1.write(2, 2, max(results_m3_new[1]))
        sh1.write(2, 3, max(results_m3_new[0]))


    # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

    def not_in_use(filename):
            try:
                os.rename(filename, filename)
                return True
            except:
                return False
    excelFilePath = "Results\\Aggregator_bids_organized_secure.xls"
    if not_in_use(excelFilePath):
        book.save(excelFilePath)
    else:
        import winsound
        duration = 1000  # milliseconds
        freq = 440  # Hz
        winsound.Beep(freq, duration)
        sleep(20)
        book.save(excelFilePath)



    #book.save("Results\\Aggregator_bids_organized.xls")




def save_results_aggr3(m, load_initial, h, load_in_bus_w, load_in_bus_g, load_in_bus_h, buildings, prices, EVs):

    instance = m
    var_names = [v.name for v in instance.component_objects(ctype=Var, active=True, descend_into=True)]

    # 've_in_es', 've_in_eb', 've_in_hp', 've_in_p2g', 've_in_fc', 've_in_ac', 've_in_ef', 've_in_eh', 've_in_ephac', 've_pv_es', 've_pv_eb', 've_pv_hp', 've_pv_p2g', 've_pv_ac', 've_pv_ef', 've_pv_eh', 've_pv_ephac', 've_out_chp_es', 've_out_chp_eb', 've_out_chp_hp', 've_out_chp_p2g', 've_out_chp_fc', 've_out_chp_ac', 've_out_chp_ef', 've_out_chp_eh', 've_out_chp_ephac', 've_out_cchp_es', 've_out_cchp_eb', 've_out_cchp_hp', 've_out_cchp_p2g', 've_out_cchp_fc', 've_out_cchp_ac', 've_out_cchp_ef', 've_out_cchp_eh', 've_out_cchp_ephac', 've_out_es_eb', 've_out_es_hp', 've_out_es_p2g', 've_out_es_ac', 've_out_es_ef', 've_out_es_eh', 've_out_es_ephac', 've_out_es', 've_out_chp', 've_out_cchp', 've_pv_out', 've_out', 've_out_es_net', 've_out_chp_net', 've_out_cchp_net', 've_pv_out_net', 'vh_in_c', 'vh_in_hs', 'vh_out_eb_c', 'vh_out_eb_hs', 'vh_out_hp_c', 'vh_out_hp_hs', 'vh_out_gb_c', 'vh_out_gb_hs', 'vh_out_chp_c', 'vh_out_chp_hs', 'vh_out_cchp_c', 'vh_out_cchp_hs', 'vh_out_ef_c', 'vh_out_ef_hs', 'vh_out_eh_c', 'vh_out_eh_hs', 'vh_out_ephac_c', 'vh_out_ephac_hs', 'vh_out_gf_c', 'vh_out_gf_hs', 'vh_out_gfi_c', 'vh_out_gfi_hs', 'vh_out_gphac_c', 'vh_out_gphac_hs', 'vh_out_hs_c', 'vh_out_eb', 'vh_out_hp', 'vh_out_hs', 'vh_out_chp', 'vh_out_cchp', 'vh_out_ef', 'vh_out_eh', 'vh_out_ephac', 'vh_out_gb', 'vh_out_gf', 'vh_out_gfi', 'vh_out_gphac', 'vh_out', 'vc_in_cs', 'vc_out_c_cs', 'vc_out_ac_cs', 'vc_out_cchp_cs', 'vc_out_ac', 'vc_out_c', 'vc_out_cs', 'vc_out', 'vg_in_chp', 'vg_in_cchp', 'vg_in_gb', 'vg_in_gf', 'vg_in_gfi', 'vg_in_gphac', 'vg_out', 'vhy_in_hys', 'vhy_out_p2g_hys', 'vhy_out_hys', 'vhy_out_p2g', 'vhy_out', 'state_es', 'state_hs', 'state_cs', 'state_hys',


    # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    #       Energy hub
    # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||


    other = ['P_aggr_w', 'P_aggr_g', 'P_aggr_hy', 'P_renew']
    other_ud = ['P_aggr_w_up', 'P_aggr_w_down']

    aggr_bids = ['P_w', 'P_g', 'P_h']
    aggr_bids_up = ['P_w_up', 'P_g_up', 'P_h_up']
    aggr_bids_down = ['P_w_down', 'P_g_down', 'P_h_down']
    aggr_bids = ['P_w']
    aggr_bids_up = ['P_w_up']
    aggr_bids_down = ['P_w_down']
    power_hp = ["hp", "hp_up", "hp_down"]
    power_hp_up = ["hp_up"]
    power_hp_down = ["hp_down"]

    eh_hp = ['P_hp']
    eh_hp_up = ['P_hp_up']
    eh_hp_down = ['P_hp_down']

    eh_dh = ['P_dh']
    eh_dh_up = ['P_dh_up']
    eh_dh_down = ['P_dh_down']


    DF_other = pd.DataFrame()
    DF_other_ud = pd.DataFrame()
    DF_aggr_bids = pd.DataFrame()
    DF_aggr_bids_up = pd.DataFrame()
    DF_aggr_bids_down = pd.DataFrame()
    DF_power_hp = pd.DataFrame()
    DF_power_hp_up = pd.DataFrame()
    DF_power_hp_down = pd.DataFrame()
    DF_eh_hp = pd.DataFrame()
    DF_eh_hp_up = pd.DataFrame()
    DF_eh_hp_down = pd.DataFrame()
    DF_eh_dh = pd.DataFrame()
    DF_eh_dh_up = pd.DataFrame()
    DF_eh_dh_down = pd.DataFrame()

    ratio_up = prices['ratio_up']
    ratio_down = prices['ratio_down']

    for v in m.component_objects(Var, active=True):
        if v.name not in other:
            for index in v:  # index[0, j, i, t]; v = ve_in_hp,...
                # print(v, index)
                if index[1] < h:
                    if v.name in other_ud:
                        DF_other_ud.at[index[1], v.name] = value(v[index])

                    if v.name in aggr_bids:
                        DF_aggr_bids.at[index[1], index[0]] = value(v[index])
                    if v.name in aggr_bids_up:
                        DF_aggr_bids_up.at[index[1], index[0]] = value(v[index])
                    if v.name in aggr_bids_down:
                        DF_aggr_bids_down.at[index[1], index[0]] = value(v[index])

                    if v.name in power_hp:
                        DF_power_hp.at[index[1], index[0]] = value(v[index])
                    if v.name in power_hp_up:
                        DF_power_hp_up.at[index[1], index[0]] = value(v[index])
                    if v.name in power_hp_down:
                        DF_power_hp_down.at[index[1], index[0]] = value(v[index])

                    if v.name in eh_hp:
                        DF_eh_hp.at[index[1], index[0]] = value(v[index])
                    if v.name in eh_hp_up:
                        DF_eh_hp_up.at[index[1], index[0]] = value(v[index])
                    if v.name in eh_hp_down:
                        DF_eh_hp_down.at[index[1], index[0]] = value(v[index])

                    if v.name in eh_dh:
                        DF_eh_dh.at[index[1], index[0]] = value(v[index])
                    if v.name in eh_dh_up:
                        DF_eh_dh_up.at[index[1], index[0]] = value(v[index])
                    if v.name in eh_dh_down:
                        DF_eh_dh_down.at[index[1], index[0]] = value(v[index])

        else:
            if v.name in other:
                for index in v:
                    if index < h:
                        print(index)
                        DF_other.at[v.name] = value(v[index])

    print("...Save data aggregator's bids...")
    with pd.ExcelWriter("Results/Aggregator_bids_all.xls") as writer:
            DF_other.to_excel(writer, sheet_name='other')
            DF_other_ud.to_excel(writer, sheet_name='other_ud')

            DF_aggr_bids.to_excel(writer, sheet_name='aggr_bids')
            DF_aggr_bids_up.to_excel(writer, sheet_name='aggr_bids_up')
            DF_aggr_bids_down.to_excel(writer, sheet_name='aggr_bids_down')

            DF_power_hp.to_excel(writer, sheet_name='power_hp')
            DF_power_hp_up.to_excel(writer, sheet_name='power_hp_up')
            DF_power_hp_down.to_excel(writer, sheet_name='power_hp_down')

            DF_eh_hp.to_excel(writer, sheet_name='power_eh_hp')
            DF_eh_hp_up.to_excel(writer, sheet_name='power_eh_hp_up')
            DF_eh_hp_down.to_excel(writer, sheet_name='power_eh_hp_down')

            DF_eh_dh.to_excel(writer, sheet_name='power_eh_dh')
            DF_eh_dh_up.to_excel(writer, sheet_name='power_eh_dh_up')
            DF_eh_dh_down.to_excel(writer, sheet_name='power_eh_dh_down')


    book = xlwt.Workbook()
    sh1 = book.add_sheet("All Bids for RT")
    sh1.write(1, 0, "Energy bid")
    sh1.write(2, 0, "Upward bid")
    sh1.write(3, 0, "Downward bid")
    sh1.write(4, 0, "Pg")
    sh1.write(5, 0, "Phy")
    sh1.write(6, 0, "P_P2G_E")
    sh1.write(7, 0, "P_chp_w")
    sh1.write(8, 0, "P_P2G_net_hy")
    sh1.write(9, 0, "P_GO")
    for t in range(0, h):
        sh1.write(0, t + 1, t)
        sh1.write(1, t + 1, m.P_aggr_w[t].value)
        sh1.write(2, t + 1, sum(m.P_aggr_w_up[i, t].value for i in range(0, len(load_in_bus_w))))
        sh1.write(3, t + 1, sum(m.P_aggr_w_down[i, t].value for i in range(0, len(load_in_bus_w))))
        sh1.write(4, t + 1, sum(m.P_g[n, t].value for n in range(0, len(load_in_bus_g))))
        sh1.write(5, t + 1, sum(m.P_hy[n, t].value for n in range(0, len(load_in_bus_g))))
        sh1.write(6, t + 1, sum(m.P_P2G_E[n, t].value for n in range(0, len(load_in_bus_w))))
        sh1.write(7, t + 1, sum(m.P_chp_w[n, t].value for n in range(26, 28)))
        sh1.write(8, t + 1, sum(m.P_P2G_net_hy[n, t].value for n in range(0, len(load_in_bus_w))))
        sh1.write(9, t + 1, (m.P_renew[t].value))

    PV = []
    PV_down = []
    PV_up = []
    P2G = []
    P2G_down = []
    P2G_up = []
    for t in range(0, h):
        PV.append(sum(m.P_PV[n, t].value for n in range(0, len(buildings))))
        PV_down.append(sum(m.P_PV_down[n, t].value for n in range(0, len(buildings))))
        PV_up.append(sum(m.P_PV_up[n, t].value for n in range(0, len(buildings))))
        P2G.append(sum(m.P_P2G_E[n, t].value for n in range(0, len(load_in_bus_w))))
        P2G_down.append(sum(m.D_P2G_E[n, t].value for n in range(0, len(load_in_bus_w))))
        P2G_up.append(sum(m.U_P2G_E[n, t].value for n in range(0, len(load_in_bus_w))))

        sh1.write(10, t + 1, (P2G[t] - PV[t] + (PV_down[t] + P2G_down[t]) * ratio_down[t] - (PV_up[t] + P2G_up[t]) * ratio_up[t])/1000)



    sh1 = book.add_sheet("Bids")
    sh1.write(1, 0, "Energy bid")
    sh1.write(2, 0, "Upward bid")
    sh1.write(3, 0, "Downward bid")
    for t in range(0, h):
        sh1.write(0, t + 1, t)
        sh1.write(1, t + 1, m.P_aggr_w[t].value)
        sh1.write(2, t + 1, sum(m.P_aggr_w_up[i, t].value for i in range(0, len(load_in_bus_w))))
        sh1.write(3, t + 1, sum(m.P_aggr_w_down[i, t].value for i in range(0, len(load_in_bus_w))))

    sh1 = book.add_sheet("Vectors")
    sh1.write(1, 0, "Pw")
    sh1.write(2, 0, "Pg")
    sh1.write(3, 0, "Ph load")
    sh1.write(4, 0, "Ph gen")
    sh1.write(5, 0, "Phy")
    for t in range(0, h):
        sh1.write(0, t + 1, t)
        sh1.write(1, t + 1, sum(m.P_w[n, t].value for n in range(0, len(load_in_bus_w))))
        sh1.write(2, t + 1, sum(m.P_g[n, t].value for n in range(0, len(load_in_bus_g))))
        sh1.write(3, t + 1, sum(m.P_h[n, t].value for n in range(0, 26)))
        sh1.write(4, t + 1, sum(m.P_h[n, t].value for n in range(26, 28)))
        sh1.write(5, t + 1, sum(m.P_hy[n, t].value for n in range(0, len(load_in_bus_g))))

    sh1 = book.add_sheet("Vectors up")
    sh1.write(1, 0, "Pw")
    sh1.write(2, 0, "Pg")
    sh1.write(3, 0, "Ph load")
    sh1.write(4, 0, "Ph gen")
    sh1.write(5, 0, "Phy")
    for t in range(0, h):
        sh1.write(0, t + 1, t)
        sh1.write(1, t + 1, sum(m.P_w_up[n, t].value for n in range(0, len(load_in_bus_w))))
        sh1.write(2, t + 1, sum(m.P_g_up[n, t].value for n in range(0, len(load_in_bus_g))))
        sh1.write(3, t + 1, sum(m.P_h_up[n, t].value for n in range(0, 26)))
        sh1.write(4, t + 1, sum(m.P_h_up[n, t].value for n in range(26, 28)))
        sh1.write(5, t + 1, sum(m.P_hy_up[n, t].value for n in range(0, len(load_in_bus_g))))

    sh1 = book.add_sheet("Vectors down")
    sh1.write(1, 0, "Pw")
    sh1.write(2, 0, "Pg")
    sh1.write(3, 0, "Ph load")
    sh1.write(4, 0, "Ph gen")
    sh1.write(5, 0, "Phy")
    for t in range(0, h):
        sh1.write(0, t + 1, t)
        sh1.write(1, t + 1, sum(m.P_w_down[n, t].value for n in range(0, len(load_in_bus_w))))
        sh1.write(2, t + 1, sum(m.P_g_down[n, t].value for n in range(0, len(load_in_bus_g))))
        sh1.write(3, t + 1, sum(m.P_h_down[n, t].value for n in range(0, 26)))
        sh1.write(4, t + 1, sum(m.P_h_down[n, t].value for n in range(26, 28)))
        sh1.write(5, t + 1, sum(m.P_hy_down[n, t].value for n in range(0, len(load_in_bus_g))))

    sh1 = book.add_sheet("Sto")
    sh1.write(1, 0, "Psto")
    sh1.write(2, 0, "Psto up")
    sh1.write(3, 0, "Psto down")
    sh1.write(4, 0, "Psto up ch")
    sh1.write(5, 0, "Psto up dis")
    sh1.write(6, 0, "Psto down ch")
    sh1.write(7, 0, "Psto down dis")
    sh1.write(8, 0, "Psto soc")
    sh1.write(9, 0, "Psto ch")
    sh1.write(10, 0, "Psto dis")
    sh1.write(11, 0, "Psto ch space")
    sh1.write(12, 0, "Psto dis space")
    sh1.write(13, 0, "P_soc_up")
    sh1.write(14, 0, "P_soc_down")
    sh1.write(15, 0, "b_sto_ch")
    sh1.write(16, 0, "P_soc_res")
    for t in range(0, h):
        sh1.write(0, t + 1, t)
        sh1.write(1, t + 1, sum(m.P_sto_ch[n, t].value - m.P_sto_dis[n, t].value for n in range(0, len(buildings))))
        sh1.write(2, t + 1, sum(m.P_sto_up[n, t].value for n in range(0, len(buildings))))
        sh1.write(3, t + 1, sum(m.P_sto_down[n, t].value for n in range(0, len(buildings))))
        sh1.write(4, t + 1, sum(m.P_sto_up_ch[n, t].value for n in range(0, len(buildings))))
        sh1.write(5, t + 1, sum(m.P_sto_up_dis[n, t].value for n in range(0, len(buildings))))
        sh1.write(6, t + 1, sum(m.P_sto_down_ch[n, t].value for n in range(0, len(buildings))))
        sh1.write(7, t + 1, sum(m.P_sto_down_dis[n, t].value for n in range(0, len(buildings))))
        sh1.write(8, t + 1, sum(m.P_soc[n, t].value for n in range(0, len(buildings))))
        sh1.write(9, t + 1, sum(m.P_sto_ch[n, t].value for n in range(0, len(buildings))))
        sh1.write(10, t + 1, sum(m.P_sto_dis[n, t].value for n in range(0, len(buildings))))
        sh1.write(11, t + 1, sum(m.P_sto_ch_space[n, t].value for n in range(0, len(buildings))))
        sh1.write(12, t + 1, sum(m.P_sto_dis_space[n, t].value for n in range(0, len(buildings))))
        sh1.write(13, t + 1, sum(m.P_soc_up[n, t].value for n in range(0, len(buildings))))
        sh1.write(14, t + 1, sum(m.P_soc_down[n, t].value for n in range(0, len(buildings))))
        sh1.write(15, t + 1, sum(m.b_sto_ch[n, t].value for n in range(0, len(buildings))))
        #sh1.write(16, t + 1, sum(m.P_soc_res[n, t].value for n in range(0, len(buildings))))

    sh1 = book.add_sheet("HP eh")
    sh1.write(1, 0, "Php")
    sh1.write(2, 0, "Php up")
    sh1.write(3, 0, "Php down")
    for t in range(0, h):
        sh1.write(0, t + 1, t)
        sh1.write(1, t + 1, sum(m.P_hp[n, t].value for n in range(0, len(buildings))))
        sh1.write(2, t + 1, sum(m.P_hp_up[n, t].value for n in range(0, len(buildings))))
        sh1.write(3, t + 1, sum(m.P_hp_down[n, t].value for n in range(0, len(buildings))))

    sh1 = book.add_sheet("PV")
    sh1.write(1, 0, "Ppv")
    sh1.write(2, 0, "Ppv up")
    sh1.write(3, 0, "Ppv down")
    for t in range(0, h):
        sh1.write(0, t + 1, t)
        sh1.write(1, t + 1, sum(m.P_PV[n, t].value for n in range(0, len(buildings))))
        sh1.write(2, t + 1, sum(m.P_PV_up[n, t].value for n in range(0, len(buildings))))
        sh1.write(3, t + 1, sum(m.P_PV_down[n, t].value for n in range(0, len(buildings))))

    if 1:
        sh1 = book.add_sheet("CHP net")
        sh1.write(1, 0, "CHP")
        sh1.write(2, 0, "CHP up")
        sh1.write(3, 0, "CHP down")
        for t in range(0, h):
            sh1.write(0, t + 1, t)
            sh1.write(1, t + 1, sum(m.P_chp_w[n, t].value for n in range(26, 28)))
            sh1.write(2, t + 1, sum(m.P_chp_w_up[n, t].value for n in range(26, 28)))
            sh1.write(3, t + 1, sum(m.P_chp_w_down[n, t].value for n in range(26, 28)))

    sh1 = book.add_sheet("Fuel station")
    sh1.write(1, 0, "P_P2G_E")
    sh1.write(2, 0, "U_P2G_E")
    sh1.write(3, 0, "D_P2G_E")

    sh1.write(5, 0, "P_P2G_hy")
    sh1.write(6, 0, "U_P2G_hy")
    sh1.write(7, 0, "D_P2G_hy")
    sh1.write(8, 0, "P_P2G_net_hy")
    sh1.write(9, 0, "U_P2G_net_hy")
    sh1.write(10, 0, "D_P2G_net_hy")
    sh1.write(11, 0, "P_P2G_sto_hy")
    sh1.write(12, 0, "D_P2G_sto_hy")
    sh1.write(13, 0, "U_P2G_sto_hy")
    sh1.write(14, 0, "P_P2G_HV_hy")

    sh1.write(15, 0, "y_soc_sto_hy")
    sh1.write(16, 0, "y_sto_hy_ch")
    sh1.write(17, 0, "y_sto_hy_dis")
    sh1.write(18, 0, "y_sto_hy_HV")
    sh1.write(19, 0, "P_sto_net_hy")
    sh1.write(20, 0, "P_sto_FC_hy")
    sh1.write(21, 0, "U_sto_FC_hy")
    sh1.write(22, 0, "D_sto_FC_hy")

    sh1.write(23, 0, "P_FC_E")
    sh1.write(24, 0, "U_FC_E")
    sh1.write(25, 0, "D_FC_E")

    for t in range(0, h):
        sh1.write(0, t + 1, t)
        sh1.write(1, t + 1, sum(m.P_P2G_E[n, t].value for n in range(0, len(load_in_bus_w))))
        sh1.write(2, t + 1, sum(m.U_P2G_E[n, t].value for n in range(0, len(load_in_bus_w))))
        sh1.write(3, t + 1, sum(m.D_P2G_E[n, t].value for n in range(0, len(load_in_bus_w))))

        sh1.write(5, t + 1, sum(m.P_P2G_hy[n, t].value for n in range(0, len(load_in_bus_w))))
        sh1.write(6, t + 1, sum(m.U_P2G_hy[n, t].value for n in range(0, len(load_in_bus_w))))
        sh1.write(7, t + 1, sum(m.D_P2G_hy[n, t].value for n in range(0, len(load_in_bus_w))))
        sh1.write(8, t + 1, sum(m.P_P2G_net_hy[n, t].value for n in range(0, len(load_in_bus_w))))
        sh1.write(9, t + 1, sum(m.U_P2G_net_hy[n, t].value for n in range(0, len(load_in_bus_w))))
        sh1.write(10, t + 1, sum(m.D_P2G_net_hy[n, t].value for n in range(0, len(load_in_bus_w))))
        sh1.write(11, t + 1, sum(m.P_P2G_sto_hy[n, t].value for n in range(0, len(load_in_bus_w))))
        sh1.write(12, t + 1, sum(m.D_P2G_sto_hy[n, t].value for n in range(0, len(load_in_bus_w))))
        sh1.write(13, t + 1, sum(m.U_P2G_sto_hy[n, t].value for n in range(0, len(load_in_bus_w))))
        sh1.write(14, t + 1, sum(m.P_P2G_HV_hy[n, t].value for n in range(0, len(load_in_bus_w))))

        sh1.write(15, t + 1, sum(m.y_soc_sto_hy[n, t].value for n in range(0, len(load_in_bus_g))))
        sh1.write(16, t + 1, sum(m.y_sto_hy_ch[n, t].value for n in range(0, len(load_in_bus_g))))
        sh1.write(17, t + 1, sum(m.y_sto_hy_dis[n, t].value for n in range(0, len(load_in_bus_g))))
        sh1.write(18, t + 1, sum(m.y_sto_HV_hy[n, t].value for n in range(0, len(load_in_bus_g))))
        sh1.write(19, t + 1, sum(m.P_sto_net_hy[n, t].value for n in range(0, len(load_in_bus_g))))
        sh1.write(20, t + 1, sum(m.P_sto_FC_hy[n, t].value for n in range(0, len(load_in_bus_g))))
        sh1.write(21, t + 1, sum(m.U_sto_FC_hy[n, t].value for n in range(0, len(load_in_bus_g))))
        sh1.write(22, t + 1, sum(m.D_sto_FC_hy[n, t].value for n in range(0, len(load_in_bus_g))))

        sh1.write(23, t + 1, sum(m.P_FC_E[n, t].value for n in range(0, len(load_in_bus_w))))
        sh1.write(24, t + 1, sum(m.U_FC_E[n, t].value for n in range(0, len(load_in_bus_w))))
        sh1.write(25, t + 1, sum(m.D_FC_E[n, t].value for n in range(0, len(load_in_bus_w))))


    if 1:
        sh1 = book.add_sheet("EVs")
        sh1.write(1, 0, "P_EV_ch")
        sh1.write(2, 0, "P_EV_dis")
        sh1.write(3, 0, "P_EV_soc")
        sh1.write(4, 0, "P_EV_up")
        sh1.write(5, 0, "P_EV_up_ch")
        sh1.write(6, 0, "P_EV_up_dis")
        sh1.write(7, 0, "P_EV_down")
        sh1.write(8, 0, "P_EV_down_ch")
        sh1.write(9, 0, "P_EV_down_dis")
        sh1.write(10, 0, "P_EV_ch_space")
        sh1.write(11, 0, "P_EV_dis_space")
        for t in range(0, h):
            sh1.write(0, t + 1, t)
            sh1.write(1, t + 1, sum(m.P_EV_ch[k, t].value for k in range(0, 20)))
            sh1.write(2, t + 1, sum(m.P_EV_dis[k, t].value for k in range(0, 20)))
            sh1.write(3, t + 1, sum(m.P_EV_soc[k, t].value for k in range(0, 20)))
            sh1.write(4, t + 1, sum(m.P_EV_up[k, t].value for k in range(0, 20)))
            #sh1.write(5, t + 1, sum(m.P_EV_up_ch[k, t].value for k in range(0, 20)))
            #sh1.write(6, t + 1, sum(m.P_EV_up_dis[k, t].value for k in range(0, 20)))
            sh1.write(7, t + 1, sum(m.P_EV_down[k, t].value for k in range(0, 20)))
            #sh1.write(8, t + 1, sum(m.P_EV_down_ch[k, t].value for k in range(0, 20)))
            #sh1.write(9, t + 1, sum(m.P_EV_down_dis[k, t].value for k in range(0, 20)))
            #sh1.write(10, t + 1, sum(m.P_EV_ch_space[k, t].value for k in range(0, 20)))
            #sh1.write(11, t + 1, sum(m.P_EV_dis_space[k, t].value for k in range(0, 20)))

    if 1:
        k = 0
        soc_max = EVs['soc_max'][k]
        soc_min = min(0.1 * EVs['soc_max'][k], EVs['soc_arrival'][k])
        soc_arrival = EVs['soc_arrival'][k]
        soc_departure = EVs['soc_departure'][k]
        time_arrival = EVs['time_arrival'][k]
        time_departure = EVs['time_departure'][k]

        sh1 = book.add_sheet("EVs ind")
        for k in range(0, 20):
            sh1.write(1 + k, 0, "Time arrival" + str(k))
            for t in range(0, 1):
                sh1.write(1 + k, t + 1, EVs['time_arrival'][k])
        n_tot = 20 + 2
        for k in range(0, 20):
            sh1.write(n_tot + k, 0, "Time departure" + str(k))
            for t in range(0, 1):
                sh1.write(n_tot + k, t + 1, EVs['time_departure'][k])
        n_tot = 20 + n_tot + 2
        for k in range(0, 20):
            sh1.write(n_tot + k, 0, "P_EV_soc_" + str(k))
            for t in range(0, h):
                if k == 0:
                    sh1.write(n_tot - 1, t + 1, t)
                sh1.write(n_tot + k, t + 1, m.P_EV_soc[k, t].value)
        n_tot = 20 + n_tot + 2
        for k in range(0, 20):
            sh1.write(n_tot + k, 0, "P_EV_ch_" + str(k))
            for t in range(0, h):
                if k == 0:
                    sh1.write(n_tot - 1, t + 1, t)
                sh1.write(n_tot + k, t + 1, m.P_EV_ch[k, t].value)
        n_tot = 20 + n_tot + 2
        for k in range(0, 20):
            sh1.write(n_tot + k, 0, "P_EV_dis" + str(k))
            for t in range(0, h):
                if k == 0:
                    sh1.write(n_tot - 1, t + 1, t)
                sh1.write(n_tot + k, t + 1, m.P_EV_dis[k, t].value)

    book.save("Results\\DA_Aggregator_bids_free.xls")





    return 0



def save_results_dso_w(m1_h, h, s):

    k = 0

    # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    #       Power flow
    # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    branch = ['PF']
    voltage = ['V']
    load = ['net_load']
    current = ['I']
    Pres = ['Pres']
    Qres = ['Qres']
    var_obj = ['var_obj']

    DF_branch = pd.DataFrame()
    DF_voltage = pd.DataFrame()
    DF_load = pd.DataFrame()
    DF_current = pd.DataFrame()
    DF_Pres = pd.DataFrame()
    DF_Qres = pd.DataFrame()
    for t in range(0, h):
        m1 = m1_h[t]
        for v in m1.component_objects(Var, active=True):
            for index in v:  # index[0, j, i, t]; v = ve_in_hp,...
                if v.name not in var_obj:
                    if index[0] == k and index[2] == 0:
                        if v.name in branch:
                            #print(v, index)
                            DF_branch.at[t, index[1]] = value(v[index])
                        if v.name in voltage:
                            DF_voltage.at[t, index[1]] = sqrt(value(v[index]))
                        if v.name in load:
                            DF_load.at[t, index[1]] = value(v[index])
                        if v.name in Pres:
                            DF_Pres.at[t, index[1]] = value(v[index])
                        if v.name in Qres:
                            DF_Qres.at[t, index[1]] = value(v[index])
                        if v.name in current:
                            if value(v[index]) < 0:
                                DF_current.at[t, index[1]] = 0
                            else:
                                DF_current.at[t, index[1]] = sqrt(value(v[index]))

    print("...Save power flow...")
    with pd.ExcelWriter("Results/Power_flow_elec_" + str(s) + " - NC.xls") as writer:
        DF_branch.to_excel(writer, sheet_name='branch')
        DF_voltage.to_excel(writer, sheet_name='voltage')
        DF_load.to_excel(writer, sheet_name='load')
        DF_current.to_excel(writer, sheet_name='I')
        DF_Pres.to_excel(writer, sheet_name='Pres')
        DF_Qres.to_excel(writer, sheet_name='Qres')


    return 0



def save_results_dso_g(m2_h, h, s):


    # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    #       Gas flow
    # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    branch = ['PF_gas']
    branch_in = ['PF_in']
    branch_out = ['PF_out']
    load = ['g']
    pressure = ['p2']
    Pgen = ['Pgen_gas']

    DF_branch = pd.DataFrame()
    DF_branch_in = pd.DataFrame()
    DF_branch_out = pd.DataFrame()
    DF_load = pd.DataFrame()
    DF_pressure = pd.DataFrame()
    DF_Pgen = pd.DataFrame()
    DF_WI = pd.DataFrame()
    DF_HHV = pd.DataFrame()
    for t in range(0, h):
        m2 = m2_h[t]
        for v in m2.component_objects(Var, active=True):
            for index in v:  # index[0, j, i, t]; v = ve_in_hp,...
                        if v.name in branch:
                            # print(v, index)
                            DF_branch.at[t, index[0]] = value(v[index])
                        if v.name in branch_in:
                            DF_branch_in.at[t, index[0]] = value(v[index])
                        if v.name in branch_out:
                            DF_branch_out.at[t, index[0]] = value(v[index])
                        if v.name in load:
                            DF_load.at[t, index[0]] = value(v[index])
                        if v.name in pressure:
                            DF_pressure.at[t, index[0]] = value(v[index])
                        if v.name in Pgen:
                            DF_Pgen.at[t, index[0]] = value(v[index])
                        if v.name in ['S_mix']:
                            DF_WI.at[t, index[0]] = value(v[index])
                        if v.name in ['HHV_mix']:
                            DF_HHV.at[t, index[0]] = value(v[index])

    print("...Save gas flow...")
    with pd.ExcelWriter("Results/Power_flow_gas" + str(s) + " - NC.xls") as writer:
        DF_branch.to_excel(writer, sheet_name='branch')
        DF_branch_in.to_excel(writer, sheet_name='branch_in')
        DF_branch_out.to_excel(writer, sheet_name='branch_out')
        DF_load.to_excel(writer, sheet_name='load')
        DF_pressure.to_excel(writer, sheet_name='pressure')
        DF_Pgen.to_excel(writer, sheet_name='Pgen')
        DF_WI.to_excel(writer, sheet_name='WI')
        DF_HHV.to_excel(writer, sheet_name='HHV')



    return 0



def save_results_dso_h(m3_h, h, s):

    # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    #       Heat flow
    # ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    generation = ['theta']
    mass = ['m']
    mass_return = []
    #mass_return = ['m_return']
    mass_load = ['mq_load']
    mass_gen = ['mq_gen']
    pressure = []
    #pressure = ['p']
    temperature_buses = ['Ts']
    temperature_buses_return = ['Ts_return']

    DF_mass = pd.DataFrame()
    DF_mass_return = pd.DataFrame()
    DF_mass_load = pd.DataFrame()
    DF_mass_gen = pd.DataFrame()
    DF_pressure = pd.DataFrame()
    DF_temperature_buses = pd.DataFrame()
    DF_temperature_buses_return = pd.DataFrame()

    for t in range(0, h):
        m3 = m3_h[t]
        for v in m3.component_objects(Var, active=True):
            for index in v:  # index[0, j, i, t]; v = ve_in_hp,...
                    if v.name in mass:
                        DF_mass.at[t, index[0]] = value(v[index])
                    if v.name in mass_return:
                        DF_mass_return.at[t, index[0]] = value(v[index])
                    if v.name in mass_load:
                        DF_mass_load.at[t, index[0]] = value(v[index])
                    if v.name in mass_gen:
                        DF_mass_gen.at[t, index[0]] = value(v[index])
                    if v.name in pressure:
                        DF_pressure.at[t, index[0]] = value(v[index])
                    if v.name in temperature_buses:
                        DF_temperature_buses.at[t, index[0]] = value(v[index])
                    if v.name in temperature_buses_return:
                        DF_temperature_buses_return.at[t, index[0]] = value(v[index])

    print("...Save power flow...")
    with pd.ExcelWriter("Results/Power_flow_heat_" + str(s) + " - NC.xls") as writer:
        DF_mass.to_excel(writer, sheet_name='mass')
        DF_mass_return.to_excel(writer, sheet_name='mass_return')
        DF_mass_load.to_excel(writer, sheet_name='mass_load')
        DF_mass_gen.to_excel(writer, sheet_name='mass_gen')
        DF_pressure.to_excel(writer, sheet_name='pressure')
        DF_temperature_buses.to_excel(writer, sheet_name='temperature_buses')
        DF_temperature_buses_return.to_excel(writer, sheet_name='temperature_buses_return')


    return 0



def save_time(time_h):
    aggregator = pd.DataFrame(time_h['aggregator'])
    dso_w = pd.DataFrame(time_h['dso_w'])
    dso_g = pd.DataFrame(time_h['dso_g'])
    dso_h = pd.DataFrame(time_h['dso_h'])

    print("...Save time...")
    with pd.ExcelWriter("Results/time_values.xls") as writer:
        aggregator.to_excel(writer, sheet_name='aggregator')
        dso_w.to_excel(writer, sheet_name='dso_w')
        dso_g.to_excel(writer, sheet_name='dso_g')
        dso_h.to_excel(writer, sheet_name='dso_h')


    return 0






