from xlrd import *
from numpy import *
from pyomo.environ import *

import pandas as pd




def power_flow_gas(m, m0, t0, s, branch, load_in_bus, other_g):
    # ____________________________________Power Flow___________________________________________________
    p_max = other_g['p_max']
    ref_gas = other_g['ref_gas']
    h = 0

    for t in range(h, h + 1):
        #__________________________________________________________________________________________________________________
        # Limits of pressure and gas well production
        for i in range(0, len(load_in_bus)):
            m.c1.add(m.p2[i, t] <= p_max)
            if i != ref_gas:
                m.c1.add(m.Pgen_gas[i, t] == 0)


        #__________________________________________________________________________________________________________________
        # Gas flow constraint
        for i in range(0, len(branch)):

            #K2 = ((0.9**2)*branch[i][3]**4.848)/(27.24 * branch[i][2])
            K2 = (43.348 * pow(m.S_mix[branch[i][0], t], 0.848) * branch[i][2])/((0.9 ** 2) * branch[i][3] ** 4.848)

            #m.c1.add(m.PF_gas[i, t] ** 2 == K2 * (m.p2[branch[i][0], t] - m.p2[branch[i][1], t]))
            m.c1.add(pow(m.PF_gas[i, t], 1.848) * K2 == (m.p2[branch[i][0], t] - m.p2[branch[i][1], t]))

            m.c1.add(m.PF_gas[i, t] == m.PF_in[i, t])
            m.c1.add(m.PF_gas[i, t] == m.PF_out[i, t])

        for i in range(0, len(load_in_bus)):
            #if m0.P_dso_gas[0, i, t0] == 0:
            #    m.c1.add(m.P_dso_gas[i, t] == 0)

            somaP_in = 0
            somaP_out = 0
            bus1 = [i]
            for k in range(0, len(branch)):
                if branch[k][1] == bus1[0]:
                    somaP_in = somaP_in + m.PF_out[k, t]
                if branch[k][0] == bus1[0]:
                    somaP_out = somaP_out + m.PF_in[k, t]

            if s == 0:
                if i == 0:
                    m.c1.add(m.Pgen_gas[i, t] + m.P_dso_hy[i, t]/3.54 - m.P_dso_gas[i, t] * 41.04/(m.HHV_mix[i, t] * m.HHV_mix[i, t]/3.6) + somaP_in - somaP_out == 0)
                    m_hy = (m.P_dso_hy[i, t] / 3.54) / (m.P_dso_hy[i, t] / 3.54 + m.Pgen_gas[i, t])
                elif i == 14:
                    m.c1.add(m.Pgen_gas[i, t] - m.P_dso_gas[i, t] * 41.04/(m.HHV_mix[0, t] * m.HHV_mix[0, t]/3.6) + somaP_in - somaP_out == 0)
                else:
                    m.c1.add(m.Pgen_gas[i, t] - m.P_dso_gas[i, t] /11.4 + somaP_in - somaP_out == 0)

            elif s == 1:
                if i == 0:
                    m.c1.add(m.Pgen_gas[i, t] + m.P_dso_hy_up[i, t]/3.54 - m.P_dso_gas_up[i, t] * 41.04/(m.HHV_mix[i, t] * m.HHV_mix[i, t]/3.6) + somaP_in - somaP_out == 0)
                    m_hy = (m.P_dso_hy_up[i, t]/3.54)/(m.P_dso_hy_up[i, t]/3.54 + m.Pgen_gas[i, t])
                elif i == 14:
                    m.c1.add(m.Pgen_gas[i, t] - m.P_dso_gas_up[i, t] * 41.04/(m.HHV_mix[0, t] * m.HHV_mix[0, t]/3.6) + somaP_in - somaP_out == 0)
                else:
                    m.c1.add(m.Pgen_gas[i, t] - m.P_dso_gas_up[ i, t] / 11.4 + somaP_in - somaP_out == 0)

            elif s == 2:
                if i == 0:
                    m.c1.add(m.Pgen_gas[i, t] + m.P_dso_hy_down[i, t]/3.54 - m.P_dso_gas_down[i, t] * 41.04/(m.HHV_mix[i, t] * m.HHV_mix[i, t]/3.6) + somaP_in - somaP_out == 0)
                    m_hy = (m.P_dso_hy_down[i, t]/3.54)/(m.P_dso_hy_down[i, t]/3.54 + m.Pgen_gas[i, t])
                elif i == 0:
                    m.c1.add(m.Pgen_gas[i, t] - m.P_dso_gas_down[i, t] * 41.04/(m.HHV_mix[0, t] * m.HHV_mix[0, t]/3.6) + somaP_in - somaP_out == 0)
                else:
                    m.c1.add(m.Pgen_gas[i, t] - m.P_dso_gas_down[i, t] /11.4 + somaP_in - somaP_out == 0)


            if i == 0:
                m.c1.add(m.WI[i, t] * m.WI[i, t] * m.S_mix[i, t] == m.HHV_mix[i, t] * m.HHV_mix[i, t])
                m.c1.add(m.HHV_mix[i, t] == m_hy * 12.75 + (1 - m_hy) * 41.04)
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
                #m.c1.add( == 0)


        for i in range(0, len(branch)):
            m.c1.add(m.PF_gas[i, t] <= 1000 * 1000)
            m.c1.add(m.PF_in[i, t] <= 1000 * 1000)
            m.c1.add(m.PF_out[i, t] <= 1000 * 1000)
            m.c1.add(m.PF_gas[i, t] >= 0)
            m.c1.add(m.PF_in[i, t] >= 0)
            m.c1.add(m.PF_out[i, t] >= 0)

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









