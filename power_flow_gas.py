from xlrd import *
from numpy import *
from pyomo.environ import *

import pandas as pd

def power_flow_gas(m, s, branch, load_in_bus, other_g):
    pressure_max = other_g['p_max']
    ref_gas_bus = other_g['ref_gas']

    # Limits of pressure and gas well production
    m = limits_of_pressure_and_gas_production(m, load_in_bus, pressure_max, ref_gas_bus)

    # Gas flow constraints
    m = gas_flow_constraints(m, branch)

    for i in range(0, len(load_in_bus)):
        m, m_hy = gas_balance_per_node(m, s, i, branch)
        m = HHV_and_S_definition(m, i, m_hy)

    m = maximum_limits_of_variables(m, load_in_bus, branch)

    return m


def limits_of_pressure_and_gas_production(m, load_in_bus, pressure_max, ref_gas_bus):
    t = 0
    for i in range(0, len(load_in_bus)):
        m.c1.add(m.p2[i, t] <= pressure_max)
        if i != ref_gas_bus:
            m.c1.add(m.Pgen_gas[i, t] == 0)
        else:
            m.c1.add(m.Pgen_gas[i, t] <= 1000 * 1000)
            m.c1.add(m.Pgen_gas[i, t] >= 0)
    return m

def gas_flow_constraints(m, branch):
    t = 0
    for i in range(0, len(branch)):
        # 𝐾_𝐺 = ((𝑝_𝐺,𝑆𝑡)2 . (𝑆_𝑀𝑖𝑥)0.848 . 𝜃_𝐺)    /    (57.3×10−8 . (𝜃𝐺,𝑆𝑡)2 . 143.52)    .    ( 𝐿 /    ((𝜂)2 . (𝑑)4.848))
        K2 = (43.348 * pow(m.S_mix[branch[i][0], t], 0.848) * branch[i][2]) / ((0.9 ** 2) * branch[i][3] ** 4.848)

        # 𝐾_𝐺 ∙ 𝑞 . |(𝑞)0.848| = (𝑝𝑚_𝐺)2 − (𝑝𝑛_𝐺)2
        m.c1.add(pow(m.q_gas[i, t], 1.848) * K2 == (m.p2[branch[i][0], t] - m.p2[branch[i][1], t]))

        m.c1.add(m.q_gas[i, t] == m.q_in[i, t])
        m.c1.add(m.q_gas[i, t] == m.q_out[i, t])

    return m

def gas_balance_per_node(m, s, i, branch):
    t = 0
    soma_q_in = 0
    soma_q_out = 0
    bus1 = [i]
    # Definition of Σ𝑞𝑛 and Σ𝑞𝑚
    for k in range(0, len(branch)):
        if branch[k][1] == bus1[0]:
            soma_q_in = soma_q_in + m.q_out[k, t]
        if branch[k][0] == bus1[0]:
            soma_q_out = soma_q_out + m.q_in[k, t]

    # Gas balance in each node per scenario
    # 𝑞𝑁𝐺 + 𝑞̂𝐻2 − 𝑞̂𝐺 + Σ𝑞𝑛 − Σ𝑞𝑚 = 0
    m_hy = 0
    if s == 0:  # Energy scenario
        if i == 0:
            # This is where hydrogen is injected and there is a heat generator consuming gas (the HHVmix affects its generation)
            m.c1.add(m.Pgen_gas[i, t] + m.P_dso_hy[i, t] / 3.54 - m.P_dso_gas[i, t] * 41.04 / (
                        m.HHV_mix[i, t] * m.HHV_mix[i, t] / 3.6) + soma_q_in - soma_q_out == 0)
            # 𝑤_𝐻2 = 𝑞̂_𝐻2 / (𝑞_𝑁𝐺 + 𝑞̂_𝐻2)
            m_hy = (m.P_dso_hy[i, t] / 3.54) / (m.P_dso_hy[i, t] / 3.54 + m.Pgen_gas[i, t])

        elif i == 14:
            # In this bus there is a heat generator consuming gas (the HHVmix affects its generation)
            m.c1.add(m.Pgen_gas[i, t] - m.P_dso_gas[i, t] * 41.04 / (
                        m.HHV_mix[0, t] * m.HHV_mix[0, t] / 3.6) + soma_q_in - soma_q_out == 0)

        else:
            # In these buses there are gas loads (the HHVmix does not affect it)
            m.c1.add(m.Pgen_gas[i, t] - m.P_dso_gas[i, t] / 11.4 + soma_q_in - soma_q_out == 0)

    elif s == 1:  # Upward scenario
        if i == 0:
            # This is where hydrogen is injected and there is a heat generator consuming gas (the HHVmix affects its generation)
            m.c1.add(m.Pgen_gas[i, t] + m.P_dso_hy_up[i, t] / 3.54 - m.P_dso_gas_up[i, t] * 41.04 / (
                        m.HHV_mix[i, t] * m.HHV_mix[i, t] / 3.6) + soma_q_in - soma_q_out == 0)
            # 𝑤_𝐻2 = 𝑞̂_𝐻2 / (𝑞_𝑁𝐺 + 𝑞̂_𝐻2)
            m_hy = (m.P_dso_hy_up[i, t] / 3.54) / (m.P_dso_hy_up[i, t] / 3.54 + m.Pgen_gas[i, t])

        elif i == 14:
            # In this bus there is a heat generator consuming gas (the HHVmix affects its generation)
            m.c1.add(m.Pgen_gas[i, t] - m.P_dso_gas_up[i, t] * 41.04 / (
                        m.HHV_mix[0, t] * m.HHV_mix[0, t] / 3.6) + soma_q_in - soma_q_out == 0)
        else:
            # In these buses there are gas loads (the HHVmix does not affect it)
            m.c1.add(m.Pgen_gas[i, t] - m.P_dso_gas_up[i, t] / 11.4 + soma_q_in - soma_q_out == 0)

    elif s == 2:  # Downward scenario
        if i == 0:
            # This is where hydrogen is injected and there is a heat generator consuming gas (the HHVmix affects its generation)
            m.c1.add(m.Pgen_gas[i, t] + m.P_dso_hy_down[i, t] / 3.54 - m.P_dso_gas_down[i, t] * 41.04 / (
                        m.HHV_mix[i, t] * m.HHV_mix[i, t] / 3.6) + soma_q_in - soma_q_out == 0)
            # 𝑤_𝐻2 = 𝑞̂_𝐻2 / (𝑞_𝑁𝐺 + 𝑞̂_𝐻2)
            m_hy = (m.P_dso_hy_down[i, t] / 3.54) / (m.P_dso_hy_down[i, t] / 3.54 + m.Pgen_gas[i, t])
        elif i == 0:
            # In this bus there is a heat generator consuming gas (the HHVmix affects its generation)
            m.c1.add(m.Pgen_gas[i, t] - m.P_dso_gas_down[i, t] * 41.04 / (
                        m.HHV_mix[0, t] * m.HHV_mix[0, t] / 3.6) + soma_q_in - soma_q_out == 0)
        else:
            # In these buses there are gas loads (the HHVmix does not affect it)
            m.c1.add(m.Pgen_gas[i, t] - m.P_dso_gas_down[i, t] / 11.4 + soma_q_in - soma_q_out == 0)

    return m, m_hy

def HHV_and_S_definition(m, i, m_hy):
    t = 0
    if i == 0:
        # 𝑊𝐼 = 𝐻𝐻𝑉_𝑚𝑖𝑥 / √𝑆_𝑚𝑖𝑥
        m.c1.add(m.WI[i, t] * m.WI[i, t] * m.S_mix[i, t] == m.HHV_mix[i, t] * m.HHV_mix[i, t])
        # 𝐻𝐻𝑉_𝑚𝑖𝑥 = 𝑤_𝐻2 ∙ 𝐻𝐻𝑉_𝐻2 + (1 − 𝑤_𝐻2) . 𝐻𝐻𝑉_𝐺
        m.c1.add(m.HHV_mix[i, t] == m_hy * 12.75 + (1 - m_hy) * 41.04)
        # 𝑆_𝑚𝑖𝑥 = 𝑤_𝐻2 ∙ 𝑆_𝐻2 + (1 − 𝑤_𝐻2) . 𝑆𝐺
        m.c1.add(m.S_mix[i, t] == m_hy * 0.0696 + (1 - m_hy) * 0.6049)

        m.c1.add(m.WI[i, t] >= 45.7)
        m.c1.add(m.WI[i, t] <= 55.9)
        m.c1.add(m.HHV_mix[i, t] >= 35.5)
        m.c1.add(m.HHV_mix[i, t] <= 47.8)
    else:
        m.c1.add(m.WI[i, t] == 0)
        m.c1.add(m.HHV_mix[i, t] == m.HHV_mix[0, t])
        m.c1.add(m.S_mix[i, t] == m.S_mix[0, t])
        m.c1.add(m.P_dso_hy_down[i, t] == 0)
        m.c1.add(m.P_dso_hy_up[i, t] == 0)
        m.c1.add(m.P_dso_hy[i, t] == 0)
    return m

def maximum_limits_of_variables(m, load_in_bus, branch):
    t = 0
    for i in range(0, len(branch)):
        m.c1.add(m.q_gas[i, t] <= 1000 * 1000)
        m.c1.add(m.q_in[i, t] <= 1000 * 1000)
        m.c1.add(m.q_out[i, t] <= 1000 * 1000)
        m.c1.add(m.q_gas[i, t] >= 0)
        m.c1.add(m.q_in[i, t] >= 0)
        m.c1.add(m.q_out[i, t] >= 0)

    for i in range(0, len(load_in_bus)):
        m.c1.add(m.Pgen_gas[i, t] <= 1000 * 1000)
        m.c1.add(m.Pgen_gas[i, t] >= 0)

        m.c1.add(m.P_dso_gas[i, t] <= 1000 * 1000)
        m.c1.add(m.P_dso_gas_up[i, t] <= 1000 * 1000)
        m.c1.add(m.P_dso_gas_down[i, t] <= 1000 * 1000)

        m.c1.add(m.p2[i, t] <= 1000 * 1000)

        m.c1.add(m.P_dso_hy[i, t] <= 1000 * 1000)
        m.c1.add(m.P_dso_hy_up[i, t] <= 1000 * 1000)
        m.c1.add(m.P_dso_hy_down[i, t] <= 1000 * 1000)

        m.c1.add(m.S_mix[i, t] <= 1)

    return m











