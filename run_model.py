from xlrd import *
from numpy import *
from pyomo.environ import *
from math import *
import pandas as pd



def run_model_buildings(m, h, number_buildings, profile_solar, temp_outside, resources_agr,
                        buses_w, buses_g, buses_h, gen_dh, other_h, fuel_station0, EVs, prices):

    buildings_with_pv = []
    buildings_with_sto = []
    buildings_with_hp_thermal = []
    buildings_with_dh = []
    buildings_with_dh_thermal = []
    for i in range(0, len(resources_agr)):
        if resources_agr[i]['installed']['PV']:
            buildings_with_pv.append(i + 1)
        if resources_agr[i]['installed']['sto']:
            buildings_with_sto.append(i + 1)
        if resources_agr[i]['installed']['dh'] == 0 and resources_agr[i]['thermal']['R'] > 0:
            buildings_with_hp_thermal.append(i + 1)
        if resources_agr[i]['installed']['dh'] == 1:
            buildings_with_dh.append(i + 1)
            if resources_agr[i]['thermal']['R'] > 0:
                buildings_with_dh_thermal.append(i + 1)


    buses_with_w_gen_dh = []
    buses_with_g_gen_dh = []
    buses_with_gen_dh = []
    buses_with_w_chp_gen_dh = []
    buses_with_g_chp_gen_dh = []
    for i in range(0, len(gen_dh)):
        if gen_dh[i]['type'] == 'w':
            buses_with_w_gen_dh.append(gen_dh[i]['bus_elec'])
        if gen_dh[i]['type'] == 'g':
            buses_with_g_gen_dh.append(gen_dh[i]['bus_gas'])
        if gen_dh[i]['type'] == 'chp':
            buses_with_w_chp_gen_dh.append(gen_dh[i]['bus_elec'])
            buses_with_g_chp_gen_dh.append(gen_dh[i]['bus_gas'])
        buses_with_gen_dh.append(gen_dh[i]['bus_dh'])

    building = resources_agr

    buses_with_fuel_station_hy = []
    buses_with_fuel_station_w = []
    fuel_station = []
    for k in range(0, 1):
        buses_with_fuel_station_w.append(fuel_station0['bus_elec'])
        buses_with_fuel_station_hy.append(fuel_station0['bus_gas'])
        fuel_station.append(fuel_station0)
    for t in range(0, h):

        ###############################################################################################################
        ###############################################################################################################
        # Delivery scenarios constraints
        ###############################################################################################################
        ###############################################################################################################

        # ____________ P electricity ____________________
        for n in range(0, len(buses_w)):
            prov_w = 0
            for j in range(0, number_buildings):
                m.c1.add(m.P_inf_load_w[j, t] == float(building[j]["consumption"]['elec'][t]))
                if j + 1 in buses_w[n]:
                    prov_w = prov_w + float(building[j]["consumption"]['elec'][t]) + m.P_hp[j, t] - m.P_PV[j, t] + \
                             m.P_sto_ch[j, t] - m.P_sto_dis[j, t]
                    for k in range(0, len(EVs['building_number'])):
                        if j + 1 == EVs['building_number'][k]:
                            prov_w = prov_w + m.P_EV_ch[k, t] - m.P_EV_dis[k, t]

            if n in buses_with_fuel_station_w:
                for k in range(0, len(fuel_station)):
                    if fuel_station[k]['bus_elec'] == n:
                        prov_w = prov_w + m.P_P2G_E[n, t] - m.P_FC_E[n, t]

            if n in buses_with_w_gen_dh:
                for k in range(0, len(gen_dh)):
                    if gen_dh[k]['bus_elec'] == n:
                        m.c1.add(m.P_w[n, t] == prov_w + m.gen_dh_w[gen_dh[k]['bus_dh'], t])
            elif n in buses_with_w_chp_gen_dh:
                for k in range(0, len(gen_dh)):
                    if gen_dh[k]['bus_elec'] == n:
                        m.c1.add(m.P_w[n, t] == prov_w - m.P_chp_w[gen_dh[k]['bus_dh'], t])
            else:
                m.c1.add(m.P_w[n, t] == prov_w)

        # ____________ P gas ____________________
        for n in range(0, len(buses_g)):
            prov_g = 0
            for j in range(0, number_buildings):
                if j + 1 in buses_g[n]:
                    prov_g = prov_g + float(building[j]["consumption"]['gas'][t]) + m.P_gb[j, t]

            if n in buses_with_g_gen_dh:
                for k in range(0, len(gen_dh)):
                    if gen_dh[k]['bus_gas'] == n:
                        m.c1.add(m.P_g[n, t] == prov_g + m.gen_dh_g[gen_dh[k]['bus_dh'], t])
            elif n in buses_with_g_chp_gen_dh:
                for k in range(0, len(gen_dh)):
                    if gen_dh[k]['bus_gas'] == n:
                        m.c1.add(m.P_g[n, t] == prov_g + m.P_chp_g[gen_dh[k]['bus_dh'], t])
            else:
                m.c1.add(m.P_g[n, t] == prov_g)

        # ____________ P heat ____________________
        if other_h['b_network']:
            for n in range(0, len(buses_h)):
                prov_dh = 0
                for j in range(0, number_buildings):
                    if j + 1 in buses_h[n]:
                        prov_dh = prov_dh + m.P_dh[j, t]

                if n in buses_with_gen_dh:
                    for k in range(0, len(gen_dh)):
                        if gen_dh[k]['bus_dh'] == n and gen_dh[k]['type'] == 'w':
                            m.c1.add(m.P_h[n, t] == prov_dh + m.gen_dh_w[n, t] * gen_dh[k]['rend'])
                        if gen_dh[k]['bus_dh'] == n and gen_dh[k]['type'] == 'g':
                            m.c1.add(m.P_h[n, t] == prov_dh + m.gen_dh_g[n, t] * gen_dh[k]['rend'])
                        if gen_dh[k]['bus_dh'] == n and gen_dh[k]['type'] == 'chp':
                            m.c1.add(m.P_h[n, t] == prov_dh + m.P_chp_h[n, t])
                else:
                    m.c1.add(m.P_h[n, t] == prov_dh)

        # ____________ P hydrogen ____________________
        for n in range(0, len(buses_g)):

            if n in buses_with_fuel_station_hy:
                for k in range(len(fuel_station)):
                    if fuel_station[k]['bus_gas'] == n:
                        m.c1.add(m.P_hy[n, t] == m.P_P2G_net_hy[fuel_station[k]['bus_elec'], t] + m.P_sto_net_hy[n, t])
            else:
                m.c1.add(m.P_hy[n, t] == 0)

        ###############################################################################################################
        # Reserves
        ###############################################################################################################
        # ____________ P electricity  up ____________________
        for n in range(0, len(buses_w)):
            prov_w = 0
            for j in range(0, number_buildings):
                if j + 1 in buses_w[n]:
                    prov_w = prov_w + m.P_sto_up[j, t] + m.P_PV_up[j, t] + m.P_hp_up[j, t]
                    for k in range(0, len(EVs['building_number'])):
                        if j + 1 == EVs['building_number'][k]:
                            prov_w = prov_w + m.P_EV_up[k, t]

            if n in buses_with_fuel_station_w:
                for k in range(0, len(fuel_station)):
                    if fuel_station[k]['bus_elec'] == n:
                        prov_w = prov_w + m.U_P2G_E[n, t] + m.U_FC_E[n, t]

            if n in buses_with_w_gen_dh:
                for k in range(0, len(gen_dh)):
                    if gen_dh[k]['bus_elec'] == n:
                        m.c1.add(m.P_w_up[n, t] == m.P_w[n, t] - prov_w - m.gen_dh_w_up[gen_dh[k]['bus_dh'], t])
                        m.c1.add(m.P_aggr_w_up[n, t] == prov_w + m.gen_dh_w_up[gen_dh[k]['bus_dh'], t])
            elif n in buses_with_w_chp_gen_dh:
                for k in range(0, len(gen_dh)):
                    if gen_dh[k]['bus_elec'] == n:
                        m.c1.add(m.P_w_up[n, t] == m.P_w[n, t] - prov_w - m.P_chp_w_up[gen_dh[k]['bus_dh'], t])
                        m.c1.add(m.P_aggr_w_up[n, t] == prov_w + m.P_chp_w_up[gen_dh[k]['bus_dh'], t])
            else:
                m.c1.add(m.P_w_up[n, t] == m.P_w[n, t] - prov_w)
                m.c1.add(m.P_aggr_w_up[n, t] == prov_w)

        # ____________ P electricity  down ____________________
        for n in range(0, len(buses_w)):
            prov_w = 0
            for j in range(0, number_buildings):
                if j + 1 in buses_w[n]:
                    prov_w = prov_w + m.P_sto_down[j, t] + m.P_PV_down[j, t] + m.P_hp_down[j, t]

            if n in buses_with_fuel_station_w:
                for k in range(0, len(fuel_station)):
                    if fuel_station[k]['bus_elec'] == n:
                        prov_w = prov_w + m.D_P2G_E[n, t] + m.D_FC_E[n, t]

            if n in buses_with_w_gen_dh:
                for k in range(0, len(gen_dh)):
                    if gen_dh[k]['bus_elec'] == n:
                        m.c1.add(m.P_w_down[n, t] == m.P_w[n, t] + prov_w + m.gen_dh_w_down[gen_dh[k]['bus_dh'], t])
                        m.c1.add(m.P_aggr_w_down[n, t] == prov_w + m.gen_dh_w_down[gen_dh[k]['bus_dh'], t])
            elif n in buses_with_w_chp_gen_dh:
                for k in range(0, len(gen_dh)):
                    if gen_dh[k]['bus_elec'] == n:
                        m.c1.add(m.P_w_down[n, t] == m.P_w[n, t] + prov_w + m.P_chp_w_down[gen_dh[k]['bus_dh'], t])
                        m.c1.add(m.P_aggr_w_down[n, t] == prov_w + m.P_chp_w_down[gen_dh[k]['bus_dh'], t])
            else:
                m.c1.add(m.P_w_down[n, t] == m.P_w[n, t] + prov_w)
                m.c1.add(m.P_aggr_w_down[n, t] == prov_w)

        m.c1.add(sum(m.P_aggr_w_up[n, t] for n in range(0, len(buses_w))) ==
                 2 * sum(m.P_aggr_w_down[n, t] for n in range(0, len(buses_w))))


        if len(buses_with_w_gen_dh) + len(buses_with_w_chp_gen_dh) > 0:
            # ____________ P heat up ____________________
            for n in range(0, len(buses_h)):
                prov_dh = 0
                for j in range(0, number_buildings):
                    if j + 1 in buses_h[n]:
                        prov_dh = prov_dh + m.P_dh_up[j, t]


                if n in buses_with_gen_dh:
                    for k in range(0, len(gen_dh)):
                        if gen_dh[k]['bus_dh'] == n and gen_dh[k]['type'] == 'w':
                            m.c1.add(m.P_h_up[n, t] == m.P_h[n, t] - prov_dh - m.gen_dh_w_up[n, t] * gen_dh[k]['rend'])
                        if gen_dh[k]['bus_dh'] == n and gen_dh[k]['type'] == 'chp':
                            m.c1.add(m.P_h_up[n, t] == m.P_h[n, t] + prov_dh + m.P_chp_h_up[n, t])
                else:
                    m.c1.add(m.P_h_up[n, t] == m.P_h[n, t] + prov_dh)

            # ____________ P heat down ____________________
            for n in range(0, len(buses_h)):
                prov_dh = 0
                for j in range(0, number_buildings):
                    if j + 1 in buses_h[n]:
                        prov_dh = prov_dh + m.P_dh_down[j, t]

                if n in buses_with_gen_dh:
                    for k in range(0, len(gen_dh)):
                        if gen_dh[k]['bus_dh'] == n and gen_dh[k]['type'] == 'w':
                            m.c1.add(
                                m.P_h_down[n, t] == m.P_h[n, t] + prov_dh + m.gen_dh_w_down[n, t] * gen_dh[k]['rend'])
                        if gen_dh[k]['bus_dh'] == n and gen_dh[k]['type'] == 'chp':
                            m.c1.add(m.P_h_down[n, t] == m.P_h[n, t] - prov_dh - m.P_chp_h_down[n, t])
                else:
                    m.c1.add(m.P_h_down[n, t] == m.P_h[n, t] - prov_dh)

        if len(buses_with_g_chp_gen_dh) > 0:
            # ____________ P gas up ____________________
            for n in range(0, len(buses_g)):
                if n in buses_with_g_chp_gen_dh:
                    for k in range(0, len(gen_dh)):
                        if gen_dh[k]['bus_gas'] == n and gen_dh[k]['type'] == 'chp':
                            m.c1.add(m.P_g_up[n, t] == m.P_g[n, t] + m.P_chp_g_up[gen_dh[k]['bus_dh'], t])
                            m.c1.add(m.P_aggr_g_up[n, t] == m.P_chp_g_up[gen_dh[k]['bus_dh'], t])
                else:
                    m.c1.add(m.P_g_up[n, t] == m.P_g[n, t])
                    m.c1.add(m.P_aggr_g_up[n, t] == 0)

            # ____________ P gas down ____________________
            for n in range(0, len(buses_g)):
                if n in buses_with_g_chp_gen_dh:
                    for k in range(0, len(gen_dh)):
                        if gen_dh[k]['bus_gas'] == n and gen_dh[k]['type'] == 'chp':
                            m.c1.add(m.P_g_down[n, t] == m.P_g[n, t] - m.P_chp_g_down[gen_dh[k]['bus_dh'], t])
                            m.c1.add(m.P_aggr_g_down[n, t] == m.P_chp_g_down[gen_dh[k]['bus_dh'], t])
                else:
                    m.c1.add(m.P_g_down[n, t] == m.P_g[n, t])
                    m.c1.add(m.P_aggr_g_down[n, t] == 0)


        if len(buses_with_fuel_station_hy) > 0:
            # ____________ P hydrogen up ____________________
            for n in range(0, len(buses_g)):
                if n in buses_with_fuel_station_hy:
                    for k in range(0, len(fuel_station)):
                        if fuel_station[k]['bus_gas'] == n:
                                m.c1.add(m.P_hy_up[n, t] == m.P_hy[n, t] - m.U_P2G_net_hy[fuel_station[k]['bus_elec'], t])
                                m.c1.add(m.P_aggr_hy_up[n, t] == m.U_P2G_net_hy[fuel_station[k]['bus_elec'], t])
                else:
                    m.c1.add(m.P_hy_up[n, t] == m.P_hy[n, t])
                    m.c1.add(m.P_aggr_hy_up[n, t] == 0)
            # ____________ P hydrogen down ____________________
            for n in range(0, len(buses_g)):
                if n in buses_with_fuel_station_hy:
                    for k in range(0, len(fuel_station)):
                        if fuel_station[k]['bus_gas'] == n:
                                m.c1.add(m.P_hy_down[n, t] == m.P_hy[n, t] + m.D_P2G_net_hy[fuel_station[k]['bus_elec'], t])
                                m.c1.add(m.P_aggr_hy_down[n, t] == m.D_P2G_net_hy[fuel_station[k]['bus_elec'], t])
                else:
                    m.c1.add(m.P_hy_down[n, t] == m.P_hy[n, t])
                    m.c1.add(m.P_aggr_hy_down[n, t] == 0)


        ###############################################################################################################
        ###############################################################################################################
        # Delivery scenarios constraints
        ###############################################################################################################
        ###############################################################################################################
        # ____________ P aggregator ____________________
        m.c1.add(m.P_aggr_w[t] == sum(m.P_w[n, t] for n in range(0, len(buses_w))))
        m.c1.add(m.P_aggr_g[t] == sum(m.P_g[n, t] for n in range(0, len(buses_g))))
        m.c1.add(m.P_aggr_hy[t] == sum(m.P_hy[n, t] for n in range(0, len(buses_g))))

        prov_dh_load = 0
        prov_dh_load_up = 0
        prov_dh_load_down = 0
        prov_dh_gen = 0
        prov_dh_gen_up = 0
        prov_dh_gen_down = 0
        for n in range(0, len(buses_h)):
            if n in buses_with_gen_dh:
                prov_dh_gen = prov_dh_gen + m.P_h[n, t]
                prov_dh_gen_up = prov_dh_gen_up + m.P_h_up[n, t]
                prov_dh_gen_down = prov_dh_gen_down + m.P_h_down[n, t]
            else:
                prov_dh_load = prov_dh_load + m.P_h[n, t]
                prov_dh_load_up = prov_dh_load_up + m.P_h_up[n, t]
                prov_dh_load_down = prov_dh_load_down + m.P_h_down[n, t]

        m.c1.add(prov_dh_load <= prov_dh_gen)
        #m.c1.add(prov_dh_load_up == prov_dh_gen_up)
        #m.c1.add(prov_dh_load_down == prov_dh_gen_down)

        ###############################################################################################################
        ###############################################################################################################
        # Resources constraints
        ###############################################################################################################
        ###############################################################################################################
        # ____________ hp = load ____________________
        for j in range(0, number_buildings):
            if building[j]['installed']['hp'] == 1:
                rend = building[j]['rend']['hp']
            elif building[j]['installed']['gb'] == 1:
                rend = building[j]['rend']['gb']

            if j + 1 not in buildings_with_dh:
                if j + 1 in buildings_with_hp_thermal:
                    building_r = building[j]['thermal']['R']
                    alpha = exp(-1/(building[j]['thermal']['R'] * building[j]['thermal']['C']))
                    T_init = building[j]['thermal']['T_init']

                    if building[j]['installed']['hp'] == 1:
                        m.c1.add(m.P_hp[j, t] <= building[j]['limits']['hp'])
                        m.c1.add(m.P_hp[j, t] >= 0)
                        m.c1.add(m.P_hp_down[j, t] <= building[j]['limits']['hp'] - m.P_hp[j, t])
                        m.c1.add(m.P_hp_up[j, t] <= m.P_hp[j, t])
                        m.c1.add(m.P_gb[j, t] == 0)
                    elif building[j]['installed']['gb'] == 1:
                        m.c1.add(m.P_gb[j, t] <= building[j]['limits']['hp'])
                        m.c1.add(m.P_gb[j, t] >= 0)
                        m.c1.add(m.P_hp[j, t] == 0)
                        m.c1.add(m.P_hp_up[j, t] == 0)
                        m.c1.add(m.P_hp_down[j, t] == 0)

                    m.c1.add(m.temp_building[j, t] >= building[j]['thermal']['T_min'][t])
                    m.c1.add(m.temp_building[j, t] <= building[j]['thermal']['T_max'][t])
                    m.c1.add(m.temp_building_up[j, t] >= building[j]['thermal']['T_min'][t])
                    m.c1.add(m.temp_building_up[j, t] <= building[j]['thermal']['T_max'][t])
                    m.c1.add(m.temp_building_down[j, t] >= building[j]['thermal']['T_min'][t])
                    m.c1.add(m.temp_building_down[j, t] <= building[j]['thermal']['T_max'][t])


                    if t == 0:
                        m.c1.add(m.temp_building[j, t] == alpha * T_init + (1 - alpha) *
                                 (temp_outside[t] + building_r * rend * (m.P_hp[j, t] + m.P_gb[j, t])))

                        if building[j]['installed']['hp'] == 1:
                            m.c1.add(m.temp_building_up[j, t] == alpha * T_init + (1 - alpha) *
                                     (temp_outside[t] + building_r * rend * (m.P_hp[j, t] - m.P_hp_up[j, t])))
                            m.c1.add(m.temp_building_down[j, t] == alpha * T_init + (1 - alpha) *
                                     (temp_outside[t] + building_r * rend * (m.P_hp[j, t] + m.P_hp_down[j, t])))

                    else:
                        m.c1.add(m.temp_building[j, t] == alpha * m.temp_building[j, t - 1] + (1 - alpha) *
                                 (temp_outside[t] + building_r * rend * (m.P_hp[j, t] + m.P_gb[j, t])))

                        if building[j]['installed']['hp'] == 1:
                            m.c1.add(m.temp_building_up[j, t] == alpha * m.temp_building_up[j, t - 1] + (1 - alpha) *
                                     (temp_outside[t] + building_r * rend * (m.P_hp[j, t] - m.P_hp_up[j, t])))
                            m.c1.add(m.temp_building_down[j, t] == alpha * m.temp_building_down[j, t - 1] + (1 - alpha) *
                                     (temp_outside[t] + building_r * rend * (m.P_hp[j, t] + m.P_hp_down[j, t])))

                else:
                    if building[j]['installed']['hp'] == 1:
                        m.c1.add(m.P_hp[j, t] == building[j]["consumption"]['heat'][t] / rend)
                        m.c1.add(m.P_gb[j, t] == 0)
                    elif building[j]['installed']['gb'] == 1:
                        m.c1.add(m.P_gb[j, t] == building[j]["consumption"]['heat'][t] / rend)
                        m.c1.add(m.P_hp[j, t] == 0)

                    m.c1.add(m.P_hp_up[j, t] == 0)
                    m.c1.add(m.P_hp_down[j, t] == 0)
                    m.c1.add(m.temp_building[j, t] == 0)
                    m.c1.add(m.temp_building_up[j, t] == 0)
                    m.c1.add(m.temp_building_down[j, t] == 0)
            else:
                m.c1.add(m.P_hp[j, t] == 0)
                m.c1.add(m.P_hp_up[j, t] == 0)
                m.c1.add(m.P_hp_down[j, t] == 0)
                m.c1.add(m.P_gb[j, t] == 0)




        # ____________ DH = load ____________________
        for j in range(0, number_buildings):
            if j + 1 in buildings_with_dh:
                if j + 1 in buildings_with_dh_thermal:
                    building_r = building[j]['thermal']['R']
                    alpha = exp(-1/(building[j]['thermal']['R']*building[j]['thermal']['C']))
                    T_init = building[j]['thermal']['T_init']
                    building[j]['limits']['dh'] = 500
                    m.c1.add(m.P_dh[j, t] <= building[j]['limits']['dh'])
                    m.c1.add(m.P_dh[j, t] >= 0)
                    m.c1.add(m.P_dh[j, t] >= 10)
                    #m.c1.add(m.P_dh_up[j, t] >= 0.1)
                    if len(buses_with_w_gen_dh) or len(buses_with_w_chp_gen_dh) > 0:
                        m.c1.add(m.P_dh_up[j, t] <= building[j]['limits']['dh'] - m.P_dh[j, t])
                        m.c1.add(m.P_dh_down[j, t] <= m.P_dh[j, t])
                    else:
                        m.c1.add(m.P_dh_up[j, t] == 0)
                        m.c1.add(m.P_dh_down[j, t] == 0)

                    if t >= 7 and t <= 20:
                        m.c1.add(m.temp_building[j, t] >= 19)
                        m.c1.add(m.temp_building[j, t] <= 23)
                        m.c1.add(m.temp_building_up[j, t] >= 19)
                        m.c1.add(m.temp_building_up[j, t] <= 23)
                        m.c1.add(m.temp_building_down[j, t] >= 19)
                        m.c1.add(m.temp_building_down[j, t] <= 23)
                    else:
                        m.c1.add(m.temp_building[j, t] >= 16)
                        m.c1.add(m.temp_building[j, t] <= 26)
                        m.c1.add(m.temp_building_up[j, t] >= 16)
                        m.c1.add(m.temp_building_up[j, t] <= 26)
                        m.c1.add(m.temp_building_down[j, t] >= 16)
                        m.c1.add(m.temp_building_down[j, t] <= 26)

                    if t == 0:
                        m.c1.add(m.temp_building[j, t] == alpha * T_init + (1 - alpha) *
                                 (temp_outside[t] + building_r * (m.P_dh[j, t])))

                        if len(buses_with_w_gen_dh) or len(buses_with_w_chp_gen_dh) > 0:
                            m.c1.add(m.temp_building_up[j, t] == alpha * T_init + (1 - alpha) *
                                     (temp_outside[t] + building_r * (m.P_dh[j, t] + m.P_dh_up[j, t])))


                            m.c1.add(m.temp_building_down[j, t] == alpha * T_init + (1 - alpha) *
                                     (temp_outside[t] + building_r * (m.P_dh[j, t] - m.P_dh_down[j, t])))


                    else:
                        m.c1.add(m.temp_building[j, t] == alpha * m.temp_building[j, t - 1] + (1 - alpha) *
                                 (temp_outside[t] + building_r * (m.P_dh[j, t])))

                        if len(buses_with_w_gen_dh) or len(buses_with_w_chp_gen_dh) > 0:
                            m.c1.add(m.temp_building_up[j, t] == alpha * m.temp_building_up[j, t - 1] + (1 - alpha) *
                                     (temp_outside[t] + building_r * (m.P_dh[j, t] + m.P_dh_up[j, t])))


                            m.c1.add(m.temp_building_down[j, t] == alpha * m.temp_building_down[j, t - 1] + (1 - alpha) *
                                     (temp_outside[t] + building_r * (m.P_dh[j, t] - m.P_dh_down[j, t])))


                else:
                    m.c1.add(m.P_dh[j, t] == building[j]["consumption"]['heat'][t])
                    m.c1.add(m.P_dh_up[j, t] == 0)
                    m.c1.add(m.P_dh_down[j, t] == 0)
                    m.c1.add(m.temp_building[j, t] == 0)
                    m.c1.add(m.temp_building_up[j, t] == 0)
                    m.c1.add(m.temp_building_down[j, t] == 0)
            else:
                m.c1.add(m.P_dh[j, t] == 0)
                m.c1.add(m.P_dh_up[j, t] == 0)
                m.c1.add(m.P_dh_down[j, t] == 0)





        # ____________ PV limits____________________
        for j in range(0, number_buildings):
            if j + 1 in buildings_with_pv:
                m.c1.add(m.P_PVr[j, t] == building[j]['limits']['PV'] * profile_solar[t])

                m.c1.add(m.P_PV_cu[j, t] <= m.P_PVr[j, t])
                m.c1.add(m.P_PV[j, t] == m.P_PVr[j, t] - m.P_PV_cu[j, t])

                m.c1.add(m.P_PV_up[j, t] <= m.P_PV_cu[j, t])
                m.c1.add(m.P_PV_down[j, t] <= m.P_PVr[j, t] - m.P_PV_cu[j, t])
            else:
                m.c1.add(m.P_PVr[j, t] == 0)
                m.c1.add(m.P_PV_cu[j, t] == 0)
                m.c1.add(m.P_PV[j, t] == 0)
                m.c1.add(m.P_PV_up[j, t] == 0)
                m.c1.add(m.P_PV_down[j, t] == 0)

        # ____________ Storage limits____________________
        # Strategy - hour to hour
        if 1:
            for j in range(0, number_buildings):
                if j + 1 in buildings_with_sto:
                    soc_max = building[j]['limits']['sto_soc_max']
                    soc_min = building[j]['limits']['sto_soc_min']
                    P_max = building[j]['limits']['sto_P']
                    rend = building[j]['rend']['sto']
                    soc_init = building[j]['limits']['sto_soc_init']

                    m.c1.add(m.P_sto_dis[j, t] + m.P_sto_dis_space[j, t] <= P_max - P_max * m.b_sto_ch[j, t])
                    m.c1.add(m.P_sto_ch[j, t] + m.P_sto_ch_space[j, t] <= P_max * m.b_sto_ch[j, t])

                    if t == 0:
                        m.c1.add(m.P_soc[j, t + 1] == soc_init + (
                                    m.P_sto_ch[j, t] * rend - m.P_sto_dis[j, t] / rend))
                    if t > 0:
                        m.c1.add(m.P_soc[j, t + 1] == m.P_soc[j, t] + (
                                    m.P_sto_ch[j, t] * rend - m.P_sto_dis[j, t] / rend))
                    if t == 23:
                        m.c1.add(m.P_soc[j, t] == soc_init)
                        m.c1.add(m.P_sto_up_dis[j, t] == 0)
                        m.c1.add(m.P_sto_up_ch[j, t] == 0)
                        m.c1.add(m.P_sto_down_dis[j, t] == 0)
                        m.c1.add(m.P_sto_down_ch[j, t] == 0)

                    m.c1.add(m.P_soc[j, t] <= soc_max)
                    m.c1.add(m.P_soc[j, t] >= soc_min)

                    m.c1.add(m.P_sto_up_dis[j, t] <= P_max - m.P_sto_dis[j, t])
                    m.c1.add(m.P_sto_up_ch[j, t] <= m.P_sto_ch[j, t])

                    m.c1.add(m.P_sto_down_ch[j, t] <= P_max - m.P_sto_ch[j, t])
                    m.c1.add(m.P_sto_down_dis[j, t] <= m.P_sto_dis[j, t])

                    m.c1.add(
                        m.P_sto_up_dis[j, t] / rend + m.P_sto_up_ch[j, t] * rend <= m.P_soc[j, t + 1] - soc_min)
                    m.c1.add(m.P_sto_down_dis[j, t] / rend + m.P_sto_down_ch[j, t] * rend <= soc_max - m.P_soc[
                        j, t + 1])

                    m.c1.add(m.P_sto_up[j, t] == m.P_sto_up_ch[j, t] + m.P_sto_up_dis[j, t])
                    m.c1.add(m.P_sto_down[j, t] == m.P_sto_down_ch[j, t] + m.P_sto_down_dis[j, t])

                    m.c1.add(
                        m.P_sto_up_dis[j, t] + m.P_sto_up_ch[j, t] + m.P_sto_down_dis[j, t] + m.P_sto_down_ch[
                            j, t] <=
                        m.P_sto_ch_space[j, t + 1] + m.P_sto_dis_space[j, t + 1])

                    m.c1.add(
                        m.P_sto_up_dis[j, t] + m.P_sto_up_ch[j, t] + m.P_sto_down_dis[j, t] + m.P_sto_down_ch[
                            j, t] <=
                        m.b_sto_space[j, t] * 10000000)

                    m.c1.add(m.P_sto_ch_space[j, t] + m.P_sto_dis_space[j, t] <= (
                                1 - m.b_sto_space[j, t]) * 10000000)

                    m.c1.add(m.P_soc_up[j, t] == 0)
                    m.c1.add(m.P_soc_down[j, t] == 0)
                else:
                    m.c1.add(m.P_sto_ch[j, t] == 0)
                    m.c1.add(m.P_sto_dis[j, t] == 0)
                    m.c1.add(m.P_soc[j, t] == 0)
                    m.c1.add(m.b_sto_ch[j, t] == 0)

                    m.c1.add(m.P_sto_up[j, t] == 0)
                    m.c1.add(m.P_sto_down[j, t] == 0)
                    m.c1.add(m.P_sto_up_ch[j, t] == 0)
                    m.c1.add(m.P_sto_down_ch[j, t] == 0)
                    m.c1.add(m.P_sto_up_dis[j, t] == 0)
                    m.c1.add(m.P_sto_down_dis[j, t] == 0)

                    m.c1.add(m.P_sto_ch_space[j, t] == 0)
                    m.c1.add(m.P_sto_dis_space[j, t] == 0)

                    m.c1.add(m.P_soc_up[j, t] == 0)
                    m.c1.add(m.P_soc_down[j, t] == 0)



        if 1:
            for k in range(0, len(EVs['building_number'])):
                '''EVs = {"time_arrival": time_arrival, "time_departure": time_departure,
                       "soc_arrival": soc_arrival,
                       "soc_departure": soc_departure, "soc_max": soc_max,
                       "building_number": building_number}'''
                soc_max = EVs['soc_max'][k] + 1
                soc_min = min(0.1 * EVs['soc_max'][k], EVs['soc_arrival'][k])
                soc_arrival = EVs['soc_arrival'][k]
                soc_departure = EVs['soc_departure'][k]
                time_arrival = EVs['time_arrival'][k]
                time_departure = EVs['time_departure'][k]
                rend = 0.95
                P_max = 7
                m.c1.add(m.P_EV_ch[k, t] <= 7)
                m.c1.add(m.P_EV_dis[k, t] <= 7)
                m.c1.add(m.P_EV_up[k, t] <= 7)
                m.c1.add(m.P_EV_down[k, t] <= 7)

                if t < time_arrival:
                    #print("EVs time", 1, "t_main", t_main, "t", t)
                    m.c1.add(m.P_EV_soc[k, t] == 0)
                    m.c1.add(m.P_EV_soc[k, t + 1] == 0)
                    m.c1.add(m.P_EV_ch[k, t] == 0)
                    m.c1.add(m.P_EV_dis[k, t] == 0)
                    m.c1.add(m.P_EV_up[k, t] == 0)
                    m.c1.add(m.P_EV_down[k, t] == 0)

                elif t == time_arrival:
                    #print("EVs time", 2, "t_main", t_main, "t", t)
                    m.c1.add(m.P_EV_soc[k, t] == 0)
                    m.c1.add(m.P_EV_soc[k, t + 1] == soc_arrival)
                    m.c1.add(m.P_EV_ch[k, t] == 0)
                    m.c1.add(m.P_EV_dis[k, t] == 0)
                    m.c1.add(m.P_EV_up[k, t] == 0)
                    m.c1.add(m.P_EV_down[k, t] == 0)


                elif t > time_arrival and t < time_departure:
                    m.c1.add(
                        m.P_EV_soc[k, t + 1] == m.P_EV_soc[k, t] + (m.P_EV_ch[k, t] * rend - m.P_EV_dis[k, t] / rend))
                    m.c1.add(m.P_EV_soc[k, t + 1] >= soc_min)
                    m.c1.add(m.P_EV_soc[k, t] <= soc_max)


                elif t == time_departure:

                    # print("EVs time", 6, "t_main", t_main, "t", t)
                    m.c1.add(m.P_EV_soc[k, t] >= soc_departure)
                    m.c1.add(m.P_EV_soc[k, t] <= soc_max)
                    m.c1.add(m.P_EV_soc[k, t + 1] <= soc_max)
                    m.c1.add(m.P_EV_soc[k, t + 1] == 0)
                    m.c1.add(m.P_EV_ch[k, t] == 0)
                    m.c1.add(m.P_EV_dis[k, t] == 0)
                    m.c1.add(m.P_EV_up[k, t] == 0)
                    m.c1.add(m.P_EV_down[k, t] == 0)


                else:

                    # print("EVs time", 7, "t_main", t_main, "t", t)
                    m.c1.add(m.P_EV_soc[k, t] == 0)
                    m.c1.add(m.P_EV_soc[k, t + 1] == 0)
                    m.c1.add(m.P_EV_ch[k, t] == 0)
                    m.c1.add(m.P_EV_dis[k, t] == 0)
                    m.c1.add(m.P_EV_up[k, t] == 0)
                    m.c1.add(m.P_EV_down[k, t] == 0)


                if t > time_arrival and t < time_departure:
                    m.c1.add(m.P_EV_soc[k, t] <= soc_max)

                    # print("soc_departure", soc_departure, "soc_max", soc_max)
                    m.c1.add(m.P_EV_dis[k, t] <= P_max - P_max * m.b_EV_ch[k, t])
                    m.c1.add(m.P_EV_ch[k, t] <= P_max * m.b_EV_ch[k, t])




                    m.c1.add(m.P_EV_up[k, t] <= (soc_max - m.P_EV_soc[k, t + 1]) / rend)
                    m.c1.add(m.P_EV_down[k, t] <= (soc_max - m.P_EV_soc[k, t + 1]) / rend)
                    m.c1.add(m.P_EV_up[k, t] <= (m.P_EV_soc[k, t + 1] - soc_min) * rend)
                    m.c1.add(m.P_EV_down[k, t] <= (m.P_EV_soc[k, t + 1] - soc_min) * rend)
                    m.c1.add(sum(m.P_EV_up[k, t] + m.P_EV_down[k, t] for t in range(t, h)) <= sum((P_max - m.P_EV_ch[k, t] - m.P_EV_dis[k, t])/2 for t in range(t, h)))


        # ____________ Electrolyzer limits____________________


        for n in range(0, len(buses_w)):
            if n in buses_with_fuel_station_w:
                for k in range(0, len(fuel_station)):
                    if fuel_station[k]['bus_elec'] == n:
                        m.c1.add(m.P_P2G_hy[n, t] == fuel_station[k]['P2G_rend'] * m.P_P2G_E[n, t])
                        m.c1.add(m.P_P2G_hy[n, t] == m.P_P2G_HV_hy[n, t] + m.P_P2G_net_hy[n, t] + m.P_P2G_sto_hy[n, t])
                        m.c1.add(m.P_P2G_E[n, t] <= fuel_station[k]['P2G_max_P'])
                        m.c1.add(m.U_P2G_E[n, t] <= m.P_P2G_E[n, t])
                        m.c1.add(m.D_P2G_E[n, t] <= fuel_station[k]['P2G_max_P'] - m.P_P2G_E[n, t])
                        m.c1.add(m.U_P2G_hy[n, t] == fuel_station[k]['P2G_rend'] * m.U_P2G_E[n, t])
                        m.c1.add(m.D_P2G_hy[n, t] == fuel_station[k]['P2G_rend'] * m.D_P2G_E[n, t])
                        m.c1.add(m.U_P2G_hy[n, t] == m.U_P2G_net_hy[n, t] + m.U_P2G_sto_hy[n, t])
                        m.c1.add(m.D_P2G_hy[n, t] == m.D_P2G_net_hy[n, t] + m.D_P2G_sto_hy[n, t])

                        #m.c1.add(m.P_P2G_net_hy[n, t] <= fuel_station[k]['P2G_max_P'] * fuel_station[k]['P2G_rend'])
                        m.c1.add(m.U_P2G_net_hy[n, t] <= m.P_P2G_net_hy[n, t])
                        m.c1.add(m.U_P2G_sto_hy[n, t] <= m.P_P2G_sto_hy[n, t])
                        #m.c1.add(m.D_P2G_net_hy[n, t] <= fuel_station[k]['P2G_max_P'] * fuel_station[k]['P2G_rend'] - m.P_P2G_net_hy[n, t])
            else:
                m.c1.add(m.P_P2G_E[n, t] == 0)
                m.c1.add(m.U_P2G_E[n, t] == 0)
                m.c1.add(m.D_P2G_E[n, t] == 0)
                m.c1.add(m.P_P2G_hy[n, t] == 0)
                m.c1.add(m.U_P2G_hy[n, t] == 0)
                m.c1.add(m.D_P2G_hy[n, t] == 0)
                m.c1.add(m.P_P2G_HV_hy[n, t] == 0)
                m.c1.add(m.P_P2G_net_hy[n, t] == 0)
                m.c1.add(m.U_P2G_net_hy[n, t] == 0)
                m.c1.add(m.D_P2G_net_hy[n, t] == 0)
                m.c1.add(m.P_P2G_sto_hy[n, t] == 0)
                m.c1.add(m.U_P2G_sto_hy[n, t] == 0)
                m.c1.add(m.D_P2G_sto_hy[n, t] == 0)




        # ____________ Storage hydrogen limits____________________
        for n in range(0, len(buses_g)):
            if n in buses_with_fuel_station_hy:
                for k in range(0, len(fuel_station)):
                    if fuel_station[k]['bus_gas'] == n:
                        soc_max = 80
                        soc_min = fuel_station[k]['sto_hy_soc_min']
                        P_max = fuel_station[k]['sto_hy_max_P']
                        soc_init = 30
                        LHV_hy = fuel_station[k]['hy_LHV']

                        if t == 0:
                            print('soc_init', soc_init)
                            m.c1.add(m.y_soc_sto_hy[n, t] == soc_init)
                            m.c1.add(
                                m.y_soc_sto_hy[n, t + 1] == soc_init + (m.y_sto_hy_ch[n, t] - m.y_sto_hy_dis[n, t]))
                        if t > 0:
                            m.c1.add(m.y_soc_sto_hy[n, t + 1] == m.y_soc_sto_hy[n, t] + (
                                    m.y_sto_hy_ch[n, t] - m.y_sto_hy_dis[n, t]))
                        if t == h - 1:
                            m.c1.add(m.y_soc_sto_hy[n, t + 1] == soc_init)
                            m.c1.add(m.D_sto_FC_hy[n, t] + m.U_sto_FC_hy[n, t] == 0)

                        m.c1.add(m.y_soc_sto_hy[n, t] <= soc_max)
                        m.c1.add(m.y_soc_sto_hy[n, t] >= soc_min)

                        m.c1.add(m.y_sto_hy_ch[n, t] == m.P_P2G_sto_hy[fuel_station[k]['bus_elec'], t] / LHV_hy)
                        m.c1.add(m.y_sto_hy_dis[n, t] == m.y_sto_HV_hy[n, t] + m.P_sto_net_hy[n, t] / LHV_hy
                                 + m.P_sto_FC_hy[n, t] / LHV_hy)

                        # m.c1.add(m.P_sto_hy_net[n, t] == 0)

                        m.c1.add(m.y_sto_hy_ch[n, t] <= P_max * m.b_sto_hy_ch[j, t])
                        m.c1.add(m.y_sto_hy_dis[n, t] <= P_max * m.b_sto_hy_dis[j, t])
                        m.c1.add(m.b_sto_hy_ch[j, t] + m.b_sto_hy_dis[j, t] <= 1)

                        m.c1.add(
                            m.D_P2G_sto_hy[fuel_station[k]['bus_elec'], t] <= (P_max - m.y_sto_hy_ch[n, t]) * LHV_hy)
                        m.c1.add(m.U_P2G_sto_hy[fuel_station[k]['bus_elec'], t] <= (m.y_sto_hy_ch[n, t]) * LHV_hy)
                        m.c1.add(m.D_sto_FC_hy[n, t] <= m.P_sto_FC_hy[n, t])
                        m.c1.add(m.U_sto_FC_hy[n, t] <= (P_max - m.y_sto_hy_dis[n, t]) * LHV_hy)

                        m.c1.add(m.D_P2G_sto_hy[fuel_station[k]['bus_elec'], t] + m.D_sto_FC_hy[n, t] <= (
                                soc_max - m.y_soc_sto_hy[n, t]) * LHV_hy)
                        m.c1.add(m.U_P2G_sto_hy[fuel_station[k]['bus_elec'], t] + m.U_sto_FC_hy[n, t] <= (
                                m.y_soc_sto_hy[n, t] - soc_min) * LHV_hy)

                        print(t)
                        m.c1.add(sum(m.U_P2G_sto_hy[fuel_station[k]['bus_elec'], t] + m.D_P2G_sto_hy[
                            fuel_station[k]['bus_elec'], t]
                                     + m.U_sto_FC_hy[n, t] + m.D_sto_FC_hy[n, t] for t in range(t, h)) <=
                                 sum((P_max - m.y_sto_hy_ch[n, t] - m.y_sto_hy_dis[n, t]) / 2 for t in range(t, h)))
                        print(t)
                        """                        m.c1.add(sum(m.U_P2G_sto_hy[fuel_station[k]['bus_elec'], t1] + m.D_P2G_sto_hy[fuel_station[k]['bus_elec'], t1]
                                     + m.U_sto_FC_hy [n, t1] + m.D_sto_FC_hy [n, t1] for t1 in range(t, h)) <=
                                 sum((P_max - m.y_sto_hy_ch[n, t1] - m.y_sto_hy_dis[n, t1])/2 for t1 in range(t, h)))"""

            else:
                m.c1.add(m.y_soc_sto_hy[n, t] == 0)
                m.c1.add(m.y_sto_hy_ch[n, t] == 0)
                m.c1.add(m.y_sto_hy_dis[n, t] == 0)
                m.c1.add(m.y_sto_HV_hy[n, t] == 0)
                m.c1.add(m.P_sto_net_hy[n, t] == 0)
                m.c1.add(m.P_sto_FC_hy[n, t] == 0)
                m.c1.add(m.D_sto_FC_hy[n, t] == 0)
                m.c1.add(m.U_sto_FC_hy[n, t] == 0)

        # ____________ Fuel cell limits____________________
        for n in range(0, len(buses_w)):
            if n in buses_with_fuel_station_w:
                for k in range(0, len(fuel_station)):
                    if fuel_station[k]['bus_elec'] == n:
                        #FC_rend = fuel_cell[k]["FC_rend"]
                        #FC_max = fuel_Cell[k]["FC_max_P"]
                        FC_rend = 0.6
                        FC_max = 250

                        m.c1.add(m.P_FC_E[n, t] == FC_rend * m.P_sto_FC_hy[fuel_station[k]['bus_gas'], t])

                        m.c1.add(m.P_sto_FC_hy[fuel_station[k]['bus_gas'], t] <= FC_max * m.b_sto_hy_dis[j, t])
                        m.c1.add(m.U_sto_FC_hy[fuel_station[k]['bus_gas'], t] <= FC_max * m.b_sto_hy_dis[j, t])
                        m.c1.add(m.D_sto_FC_hy[fuel_station[k]['bus_gas'], t] <= FC_max * m.b_sto_hy_dis[j, t])

                        m.c1.add(m.U_sto_FC_hy[fuel_station[k]['bus_gas'], t] <= FC_max - m.P_sto_FC_hy[fuel_station[k]['bus_gas'], t])
                        m.c1.add(m.D_sto_FC_hy[fuel_station[k]['bus_gas'], t] <= m.P_sto_FC_hy[fuel_station[k]['bus_gas'], t])

                        m.c1.add(m.U_FC_E[n, t] == FC_rend * m.U_sto_FC_hy[fuel_station[k]['bus_gas'], t])
                        m.c1.add(m.D_FC_E[n, t] == FC_rend * m.D_sto_FC_hy[fuel_station[k]['bus_gas'], t])

            else:
                m.c1.add(m.P_FC_E[n, t] == 0)
                m.c1.add(m.U_FC_E[n, t] == 0)
                m.c1.add(m.D_FC_E[n, t] == 0)

        # ____________ Hydrogen vehicles limits____________________
        for n in range(0, len(buses_g)):
            if n in buses_with_fuel_station_hy:
                for k in range(0, len(fuel_station)):
                    if fuel_station[k]['bus_gas'] == n:
                        m.c1.add(m.P_P2G_HV_hy[fuel_station[k]['bus_elec'], t]/LHV_hy + m.y_sto_HV_hy[n, t] == fuel_station[k]['load'][t])


        # ____________ Electrical district heating generators limits____________________
        if other_h['b_network']:
            for n in range(0, len(buses_h)):

                if n in buses_with_gen_dh:
                    for j in range(0, len(gen_dh)):
                        if n == gen_dh[j]['bus_dh']:
                            if gen_dh[j]['type'] == 'w':
                                m.c1.add(m.gen_dh_w[n, t] <= gen_dh[j]['P_max'])
                                m.c1.add(m.gen_dh_w_up[n, t] <= gen_dh[j]['P_max'])
                                m.c1.add(m.gen_dh_w_down[n, t] <= gen_dh[j]['P_max'])
                                m.c1.add(m.gen_dh_w_down[n, t] <= gen_dh[j]['P_max'] - m.gen_dh_w[n, t])
                                m.c1.add(m.gen_dh_w_up[n, t] <= m.gen_dh_w[n, t])
                            else:
                                m.c1.add(m.gen_dh_w[n, t] == 0)
                                m.c1.add(m.gen_dh_w_up[n, t] == 0)
                                m.c1.add(m.gen_dh_w_down[n, t] == 0)

                            if gen_dh[j]['type'] == 'g':
                                m.c1.add(m.gen_dh_g[n, t] <= gen_dh[j]['P_max'])
                            else:
                                m.c1.add(m.gen_dh_g[n, t] == 0)

                            if gen_dh[j]['type'] == 'chp':
                                m.c1.add(m.P_chp_g[n, t] <= gen_dh[j]['P_max'])
                                m.c1.add(m.P_chp_g_up[n, t] <= (2/6) * gen_dh[j]['P_max'])
                                m.c1.add(m.P_chp_g_down[n, t] <= (2/6) * gen_dh[j]['P_max'])

                                m.c1.add(m.P_chp_g_up[n, t] <= gen_dh[j]['P_max'] - m.P_chp_g[n, t])
                                m.c1.add(m.P_chp_g_down[n, t] <= m.P_chp_g[n, t])

                                m.c1.add(m.P_chp_w[n, t] == m.P_chp_g[n, t] * gen_dh[j]['rend_elec'])
                                m.c1.add(m.P_chp_w_up[n, t] == m.P_chp_g_up[n, t] * gen_dh[j]['rend_elec'])
                                m.c1.add(m.P_chp_w_down[n, t] == m.P_chp_g_down[n, t] * gen_dh[j]['rend_elec'])
                                m.c1.add(m.P_chp_h[n, t] == m.P_chp_g[n, t] * gen_dh[j]['rend_heat'])
                                m.c1.add(m.P_chp_h_up[n, t] == m.P_chp_g_up[n, t] * gen_dh[j]['rend_heat'])
                                m.c1.add(m.P_chp_h_down[n, t] == m.P_chp_g_down[n, t] * gen_dh[j]['rend_heat'])
                if n not in buses_with_gen_dh:
                    m.c1.add(m.P_chp_h[n, t] == 0)
                    m.c1.add(m.P_chp_h_up[n, t] == 0)
                    m.c1.add(m.P_chp_h_down[n, t] == 0)
                    m.c1.add(m.P_chp_g[n, t] == 0)
                    m.c1.add(m.P_chp_g_up[n, t] == 0)
                    m.c1.add(m.P_chp_g_down[n, t] == 0)
                    m.c1.add(m.P_chp_w[n, t] == 0)
                    m.c1.add(m.P_chp_w_up[n, t] == 0)
                    m.c1.add(m.P_chp_w_down[n, t] == 0)


    return m





