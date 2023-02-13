from numpy import *
from math import *
from copy import *


def power_flow_heat(m, m0, t0, s, branch, load, buses, resources_agr, heat_network, other_h):
    t = 0
    # ____________________________________Power Flow___________________________________________________
    Cp = other_h['Cp']
    Ta = other_h['Ta']
    m_max = other_h['m_max']
    m_min = other_h['m_min']
    Ts_min = other_h['Ts_min']
    Ts_max = other_h['Ts_max']
    Tr_min = other_h['Tr_min']
    Tr_max = other_h['Tr_max']
    friction = other_h['friction']
    T_load = other_h['T_load']
    T_gen = other_h['T_gen']
    p_min = other_h['p_min']
    p_max = other_h['p_max']

    buses_gen = heat_network['buses_gen']
    load_in_bus_h = heat_network['load_in_bus_h']


    # Define buses with thermal loads
    buses_load = []
    buses_with_dh_thermal = []
    buildings_with_dh = []
    for i in range(0, len(load_in_bus_h)):
        if len(load_in_bus_h[i]) > 0:
            buses_load.append(i)
        for j in range(0, len(load_in_bus_h[i])):
            num_building = load_in_bus_h[i][j] - 1
            if resources_agr[num_building]['thermal']['R'] > 0:
                buses_with_dh_thermal.append(i)
            if resources_agr[num_building]['installed']['dh'] == 1:
                buildings_with_dh.append(num_building + 1)

    # Define buses load at extreme points and middle points
    buses_load_extremes = []
    buses_middle = []
    for i in range(0, len(buses)):
        num_in = []
        num_out = []
        for j in range(0, len(branch)):
            if branch[j][0] == i:  # out of node
                num_out.append([branch[j][0], branch[j][1]])
            if branch[j][1] == i:  # into the node
                num_in.append([branch[j][0], branch[j][1]])
        if len(num_out) == 0 and len(num_in) > 0 and i in buses_load:
            buses_load_extremes.append(i)
        elif len(num_out) > 0 and len(num_in) > 0 and i in buses_load:
            buses_middle.append(i)



    # ____________________________ Load and DR__________________________________________
    # Define load and generation at different type of buses
    for k in range(0, len(buses)):
        if k in buses_load and k in buses_gen:
            m.c1.add(m.mq_load[k, t] >= 0)
            m.c1.add(m.mq_gen[k, t] >= 0)

        elif k in buses_load and k not in buses_gen:
            m.c1.add(m.mq_load[k, t] >= 0)
            m.c1.add(m.mq_gen[k, t] == 0)

        elif k in buses_gen and k not in buses_load:
            m.c1.add(m.mq_load[k, t] == 0)
            m.c1.add(m.mq_gen[k, t] >= 0)

        else:
            m.c1.add(m.mq_load[k, t] == 0)
            m.c1.add(m.mq_gen[k, t] == 0)

        if k not in buses_middle:
            m.c1.add(m.To[k, t] == T_load)

    for i in range(0, len(load)):
        if i + 1 not in buildings_with_dh:
            m.c1.add(m.h[i, t] == 0)


    ###################################################################################################################
    ###################################################################################################################
    #  Supply Model
    ###################################################################################################################
    ###################################################################################################################

    # ____________________________(1) Continuity of flow__________________________________________
    for i in range(0, len(buses)):
        m_out1 = 0
        m_in1 = 0
        for j in range(0, len(branch)):
            if branch[j][0] == i:
                m_out1 = m_out1 + m.m[j, t]
            if branch[j][1] == i:
                m_in1 = m_in1 + m.m[j, t]

        mqload = 0
        mqgen = 0
        if i in buses_load:
            mqload = m.mq_load[i, t]
        if i in buses_gen:
            mqgen = m.mq_gen[i, t]

        m.c1.add(m_in1 - m_out1 + mqgen - mqload == 0)

    # ____________________________(2) Pressure loss equation__________________________________________
    for i in range(0, len(branch)):
        a = friction * branch[i][2] / (2 * branch[i][3] / 1000 * 1000 * pi * (branch[i][3] / 1000 / 2) ** 2)
        m.c1.add(m.p[branch[i][0], t] - m.p[branch[i][1], t] == a * m.m[i, t] ** 2)

    # ____________________________(3) Heat power equations__________________________________________

    if s == 0: # Scenario energy
        for i in range(0, len(buses)):
            if i in buses_load:
                if i in buses_middle:
                    m.c1.add(m.P_dso_heat[i, t] * 1000 == Cp * m.Ts[i, t] * m.mq_load[i, t] - Cp * m.mq_load[i, t] * T_load )
                    if i not in buses_with_dh_thermal:
                        m.c1.add(m.P_dso_heat[i, t] == m0.P_h[i, t0].value)
                    m.c1.add(m.P_dso_heat[i, t] <= 20000)
                    m.c1.add(m.P_dso_heat[i, t] >= 0)
                    m.c1.add(m.To[i, t] == T_load)

                else:
                    m.c1.add(m.P_dso_heat[i, t] * 1000 == Cp * m.Ts[i, t] * m.mq_load[i, t] - Cp * m.mq_load[i, t] * T_load)
                    if i not in buses_with_dh_thermal:
                        m.c1.add(m.P_dso_heat[i, t] == m0.P_h[i, t0].value)
                    m.c1.add(m.P_dso_heat[i, t] <= 20000)
                    m.c1.add(m.P_dso_heat[i, t] >= 0)
                    m.c1.add(m.Ts_return[i, t] == T_load)

            elif i in buses_gen:
                m.c1.add(m.P_dso_heat[i, t] * 1000 == Cp * m.mq_gen[i, t] * T_gen - Cp * m.Ts_return[i, t] * m.mq_gen[i, t])
                m.c1.add(m.P_dso_heat[i, t] <= 20000)
                m.c1.add(m.P_dso_heat[i, t] >= 0)
                m.c1.add(m.Ts[i, t] == T_gen)
            else:
                m.c1.add(m.P_dso_heat[i, t] == 0)

    elif s == 1: # Scenario upward
        for i in range(0, len(buses)):
            if i in buses_load:
                if i in buses_middle:
                    m.c1.add(m.P_dso_heat_up[i, t] * 1000 == Cp * m.Ts[i, t] * m.mq_load[i, t] - Cp * m.mq_load[i, t] * T_load )
                    if i not in buses_with_dh_thermal:
                        m.c1.add(m.P_dso_heat_up[i, t] == m0.P_h_up[i, t0].value)
                    m.c1.add(m.P_dso_heat[i, t] <= 20000)
                    m.c1.add(m.P_dso_heat[i, t] >= 0)
                    m.c1.add(m.P_dso_heat_up[i, t] <= 20000)
                    m.c1.add(m.P_dso_heat_up[i, t] >= 0)
                    m.c1.add(m.To[i, t] == T_load)

                else:
                    m.c1.add(m.P_dso_heat_up[i, t] * 1000 == Cp * m.Ts[i, t] * m.mq_load[i, t] - Cp * m.mq_load[i, t] * T_load)
                    if i not in buses_with_dh_thermal:
                        m.c1.add(m.P_dso_heat_up[i, t] == m0.P_h_up[i, t0].value)
                    m.c1.add(m.P_dso_heat[i, t] <= 20000)
                    m.c1.add(m.P_dso_heat[i, t] >= 0)
                    m.c1.add(m.P_dso_heat_up[i, t] <= 20000)
                    m.c1.add(m.P_dso_heat_up[i, t] >= 0)
                    m.c1.add(m.Ts_return[i, t] == T_load)

            elif i in buses_gen:
                m.c1.add(m.P_dso_heat_up[i, t] * 1000 == Cp * m.mq_gen[i, t] * T_gen - Cp * m.Ts_return[i, t] * m.mq_gen[i, t])
                m.c1.add(m.P_dso_heat_up[i, t] <= 20000)
                m.c1.add(m.Ts[i, t] == T_gen)
            else:
                m.c1.add(m.P_dso_heat_up[i, t] == 0)

    elif s == 2: # Scenario downward
        for i in range(0, len(buses)):
            if i in buses_load:
                if i in buses_middle:
                    m.c1.add(m.P_dso_heat_down[i, t] * 1000 == Cp * m.Ts[i, t] * m.mq_load[i, t] - Cp * m.mq_load[i, t] * T_load )
                    if i not in buses_with_dh_thermal:
                        m.c1.add(m.P_dso_heat_down[i, t] == m0.P_h_down[i, t0].value)
                    m.c1.add(m.P_dso_heat[i, t] <= 20000)
                    m.c1.add(m.P_dso_heat[i, t] >= 0)
                    m.c1.add(m.P_dso_heat_down[i, t] <= 20000)
                    m.c1.add(m.P_dso_heat_down[i, t] >= 0)
                    m.c1.add(m.To[i, t] == T_load)

                else:
                    m.c1.add(m.P_dso_heat_down[i, t] * 1000 == Cp * m.Ts[i, t] * m.mq_load[i, t] - Cp * m.mq_load[i, t] * T_load)
                    if i not in buses_with_dh_thermal:
                        m.c1.add(m.P_dso_heat_down[i, t] == m0.P_h_down[i, t0].value)
                    m.c1.add(m.P_dso_heat[i, t] <= 20000)
                    m.c1.add(m.P_dso_heat[i, t] >= 0)
                    m.c1.add(m.P_dso_heat_down[i, t] <= 20000)
                    m.c1.add(m.P_dso_heat_down[i, t] >= 0)
                    m.c1.add(m.Ts_return[i, t] == T_load)

            elif i in buses_gen:
                m.c1.add(m.P_dso_heat_down[i, t] * 1000 == Cp * m.mq_gen[i, t] * T_gen - Cp * m.Ts_return[i, t] * m.mq_gen[i, t])
                m.c1.add(m.P_dso_heat_down[i, t] <= 20000)
                m.c1.add(m.Ts[i, t] == T_gen)
            else:
                m.c1.add(m.P_dso_heat_down[i, t] == 0)

    # ____________________________(4) Supply - Pipeline temperature__________________________________________
    for i in range(0, len(branch)):
        m.c1.add( m.T_end[i, t] == (m.T_start[i, t] - Ta) * (1 - ((branch[i][4] * branch[i][2]) / (Cp * m.m[i, t])))
                  + Ta)

    # ____________________________(5) Supply - Node temperature mixing__________________________________________

    for i in range(0, len(buses)):
        m_out = 0
        m_in = 0
        num_in = []
        num_out = []
        for j in range(0, len(branch)):
            if branch[j][0] == i:  # out of node
                m_out = m_out + m.T_start[j, t] * m.m[j, t]

                num_out.append([branch[j][0], branch[j][1]])
            if branch[j][1] == i:  # into the node
                m_in = m_in + m.T_end[j, t] * m.m[j, t]

                num_in.append([branch[j][0], branch[j][1]])

        if len(num_in) > 1:  # 2 in, 1 out
            m.c1.add(m_out == m_in)
            for j in range(0, len(branch)):
                if branch[j][0] == i:
                    m.c1.add(m.Ts[i, t] == m.T_start[j, t])

        elif len(num_in) == 0:  # 0 in, 1 out
            for j in range(0, len(branch)):
                if branch[j][0] == i:
                    m.c1.add(m.Ts[i, t] == m.T_start[j, t])

        elif len(num_out) == 0:  # 1 in, 0 out
            for j in range(0, len(branch)):
                if branch[j][1] == i:
                    m.c1.add(m.Ts[i, t] == m.T_end[j, t])

        elif len(num_in) == 1 and len(num_out) >= 1:  # 1 in, 1 out
            for j in range(0, len(branch)):
                if branch[j][1] == i:
                    m.c1.add(m.Ts[i, t] == m.T_end[j, t])
                if branch[j][0] == i:
                    m.c1.add(m.Ts[i, t] == m.T_start[j, t])

    ###################################################################################################################
    ###################################################################################################################
    #  Return model
    ###################################################################################################################
    ###################################################################################################################
    branch1 = deepcopy(branch)
    for i in range(0, len(branch1)):
        in2 = branch1[i][0]
        out2 = branch[i][1]
        branch1[i][0] = out2
        branch1[i][1] = in2

    # ____________________________(1) Return - Continuity of flow__________________________________________
    for i in range(0, len(buses)):
        m_out1 = 0
        m_in1 = 0
        num = []
        num1 = []
        for j in range(0, len(branch1)):
            if branch1[j][0] == i:
                m_out1 = m_out1 + m.m_return[j, t]
                m.c1.add(m.m_return[j, t] == m.m[j, t])
                num.append([branch1[j][0], branch1[j][1]])
            if branch1[j][1] == i:
                m_in1 = m_in1 + m.m_return[j, t]
                num1.append([branch1[j][0], branch1[j][1]])
                m.c1.add(m.m_return[j, t] == m.m[j, t])

        mqload = 0
        mqgen = 0
        if i in buses_load_extremes:
            mqload = m.mq_load[i, t]
        elif i in buses_middle:
            mqload = m.mq_load[i, t]

        if i in buses_gen:
            mqgen = m.mq_gen[i, t]

        m.c1.add(m_in1 - m_out1 - mqgen + mqload == 0)

    # ____________________________(2) Pressure loss equation__________________________________________
    for i in range(0, len(branch1)):
        a = friction * branch1[i][2] / (2 * branch1[i][3] / 1000 * 1000 * pi * (branch1[i][3] / 1000 / 2) ** 2)
        m.c1.add(m.p_return[branch1[i][0], t] - m.p_return[branch1[i][1], t] == a * m.m_return[i, t] ** 2)

    # ____________________________(4) Return - Pipeline temperature__________________________________________

    for i in range(0, len(branch1)):
        m.c1.add(m.T_end_return[i, t] == (m.T_start_return[i, t] - Ta) *
                 (1 - ((branch1[i][4] * branch1[i][2]) / (Cp * m.m_return[i, t]))) + Ta)


    # ____________________________(5) Return - Node temperature mixing__________________________________________
    for i in range(0, len(buses)):
        num_in = []
        num_out = []

        m_out_return = 0
        m_in_return = 0
        for j in range(0, len(branch1)):
            if branch1[j][0] == i:  # out of node
                m_out_return = m_out_return + m.T_start_return[j, t] * m.m_return[j, t]


                num_out.append([branch1[j][0], branch1[j][1]])

            if branch1[j][1] == i:  # into the node
                m_in_return = m_in_return + m.T_end_return[j, t] * m.m_return[j, t]

                num_in.append([branch1[j][0], branch1[j][1]])

        if i in buses_middle:
            m_in_return = m_in_return + m.mq_load[i, t] * T_load
            num_in.append([i])
            m.c1.add(m.To[i, t] == T_load)

        if len(num_in) > 1:  # 2 in, 1 out
            m.c1.add(m_out_return == m_in_return)
            for j in range(0, len(branch1)):
                if branch1[j][0] == i:
                    m.c1.add(m.Ts_return[i, t] == m.T_start_return[j, t])

        if len(num_in) == 0 and len(num_out) == 1:  # 0 in, 1 out
            for j in range(0, len(branch1)):
                if branch1[j][0] == i:
                    m.c1.add(m.Ts_return[i, t] == m.T_start_return[j, t])

        elif len(num_in) == 1 and len(num_out) == 0:  # 1 in, 0 out
            for j in range(0, len(branch1)):
                if branch1[j][1] == i:
                    m.c1.add(m.Ts_return[i, t] == m.T_end_return[j, t])

        if len(num_in) == 1 and len(num_out) >= 1:  # 1 in, 1 out
            for j in range(0, len(branch1)):
                if branch1[j][1] == i:
                    m.c1.add(m.Ts_return[i, t] == m.T_end_return[j, t])
                if branch1[j][0] == i:
                    m.c1.add(m.Ts_return[i, t] == m.T_start_return[j, t])


    ###################################################################################################################
    #  Variables limits
    ###################################################################################################################

    for i in range(0, len(buses)):
        m.c1.add(m.Ts[i, t] >= Ts_min)
        m.c1.add(m.Ts[i, t] <= Ts_max)
        m.c1.add(m.To[i, t] <= Tr_max)
        m.c1.add(m.To[i, t] >= Tr_min)
        m.c1.add(m.Ts_return[i, t] >= Tr_min)
        m.c1.add(m.Ts_return[i, t] <= Tr_max)

        m.c1.add(m.P_dso_heat[i, t] <= 20000)
        m.c1.add(m.P_dso_heat_up[i, t] <= 20000)
        m.c1.add(m.P_dso_heat_down[i, t] <= 20000)
        m.c1.add(m.P_dso_heat[i, t] >= 0)
        m.c1.add(m.P_dso_heat_up[i, t] >= 0)
        m.c1.add(m.P_dso_heat_down[i, t] >= 0)

        1
    for i in range(0, len(branch)):
        m.c1.add(m.T_end[i, t] <= Ts_max)
        m.c1.add(m.T_end[i, t] >= Ts_min)
        m.c1.add(m.T_start[i, t] <= Ts_max)
        m.c1.add(m.T_start[i, t] >= Ts_min)

        m.c1.add(m.T_end_return[i, t] <= Tr_max)
        m.c1.add(m.T_end_return[i, t] >= Tr_min)
        m.c1.add(m.T_start_return[i, t] <= Tr_max)
        m.c1.add(m.T_start_return[i, t] >= Tr_min)

    for i in range(0, len(buses)):
        m.c1.add(m.p[i, t] <= p_max * 100000)
        m.c1.add(m.p[i, t] >= p_min * 100000)
        m.c1.add(m.p_return[i, t] <= p_max * 100000)
        m.c1.add(m.p_return[i, t] >= p_min * 100000)
        m.c1.add(m.mq_gen[i, t] <= m_max)
        m.c1.add(m.mq_load[i, t] <= m_max)

    for i in range(0, len(branch)):
        m.c1.add(m.m[i, t] <= m_max)
        m.c1.add(m.m[i, t] >= m_min)
        m.c1.add(m.m_return[i, t] <= m_max)
        m.c1.add(m.m_return[i, t] >= m_min)


    return m




