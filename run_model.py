from xlrd import *
from numpy import *
from pyomo.environ import *
from math import *
import pandas as pd

from run_resources_model import *



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
        m = run_resources_model(m, t, h, number_buildings, profile_solar, temp_outside, resources_agr,
                        buses_w, buses_g, buses_h, gen_dh, other_h, fuel_station0, EVs)


    return m




