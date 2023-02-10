from xlrd import *
from numpy import *
from pyomo.environ import *
from math import *
import pandas as pd





def power_flow_elec(m, s, branch, load_in_bus, other_w):
    h = 0
    scenario = 0
    # ____________________________________Power Flow___________________________________________________
    branch_div = other_w['MVA_base'] / 1.6 #20
    rx_div = 1000
    ref_bus = other_w['ref_bus']
    v_min = other_w['v_min']
    v_max = other_w['v_max']
    v_ref = other_w['v_ref']
    I_max = other_w['I_max']

    for t in range(h, h + 1):
        for i in range(0, len(branch)):
            somaP = 0
            somaQ = 0
            bus1 = [branch[i][1]]

            for k in range(0,len(branch)):
                if branch[k][0] == bus1[0]:
                    somaP = somaP + m.PF[scenario, k, t]
                    somaQ = somaQ + m.QF[scenario, k, t]


            if s == 0:
                # _________ (1) PF = load - Pres + somaP + r * I^2 ________________________________________________________
                m.c1.add(m.net_load[scenario, branch[i][1].astype(int), t] == m.P_dso[scenario, branch[i][1].astype(int), t])

                m.c1.add(m.PF[scenario, i, t] == m.P_dso[scenario, branch[i][1].astype(int), t]/branch_div -
                         m.Pres[scenario, branch[i][1].astype(int), t] / 1 + somaP + branch[i][2] * (m.I[scenario, i, t]) / rx_div)

                # _________ (2) QF = load - Qres + somaQ + r * I^2 ________________________________________________________
                m.c1.add(m.QF[scenario, i, t] == m.P_dso[scenario, branch[i][1].astype(int), t] * 0.2/branch_div -
                         m.Qres[scenario, branch[i][1].astype(int), t] / 1 + somaQ + branch[i][3] * (m.I[scenario, i, t]) / rx_div)

            elif s == 1:
                # _________ (1) PF = load - Pres + somaP + r * I^2 ________________________________________________________
                m.c1.add(m.net_load[scenario, branch[i][1].astype(int), t] == m.P_dso_up[scenario, branch[i][1].astype(int), t])

                m.c1.add(m.PF[scenario, i, t] == m.P_dso_up[scenario, branch[i][1].astype(int), t]/branch_div -
                         m.Pres[ scenario, branch[i][1].astype(int), t] / 1 + somaP + branch[i][2] * (m.I[scenario, i, t]) / rx_div)

                # _________ (2) QF = load - Qres + somaQ + r * I^2 ________________________________________________________
                m.c1.add(m.QF[scenario, i, t] == m.P_dso_up[scenario, branch[i][1].astype(int), t] * 0.2/branch_div -
                         m.Qres[scenario, branch[i][1].astype(int), t] / 1 + somaQ + branch[i][3] * (m.I[scenario, i, t]) / rx_div)

            elif s == 2:
                # _________ (1) PF = load - Pres + somaP + r * I^2 ________________________________________________________
                m.c1.add(m.net_load[scenario, branch[i][1].astype(int), t] == m.P_dso_down[scenario, branch[i][1].astype(int), t])

                m.c1.add(m.PF[scenario, i, t] == m.P_dso_down[scenario, branch[i][1].astype(int), t]/branch_div -
                         m.Pres[scenario, branch[i][1].astype(int), t] / 1 + somaP + branch[i][2] * (m.I[scenario, i, t]) / rx_div)

                # _________ (2) QF = load - Qres + somaQ + r * I^2 ________________________________________________________
                m.c1.add(m.QF[scenario, i, t] == m.P_dso_down[scenario, branch[i][1].astype(int), t] * 0.2/branch_div -
                         m.Qres[scenario, branch[i][1].astype(int), t] / 1 + somaQ + branch[i][3] * (m.I[scenario, i, t]) / rx_div)



            # _________ (3) Vm^2 - 2(r x PF + x x QF) + (r^2 + x^2). I^2 = Vn ^2 ______________________________________
            m.c1.add(m.V[scenario, branch[i][0].astype(int), t] - 2 * (branch[i][2] * m.PF[scenario, i, t] + branch[i][3] * m.QF[scenario, i, t])/rx_div +
                     ((branch[i][2]/rx_div) ** 2 + (branch[i][3]/rx_div) ** 2) * m.I[scenario, i, t] == m.V[scenario, branch[i][1].astype(int), t])

            # _________ (4) V^2 x I^2 = PF^2 + QF^2 ___________________________________________________________________
            m.c1.add((m.V[scenario, branch[i][0].astype(int), t]) * m.I[scenario, i, t] == (m.PF[scenario, i, t] ** 2) + (m.QF[scenario, i, t] ** 2))

        # _________ Generator limits _________________________________________________________________________________
        for j in range(0, len(load_in_bus)):
            if j == ref_bus:
                m.c1.add(m.Pres[scenario, j, t] == m.PF[scenario, 0, t])
                m.c1.add(m.Qres[scenario, j, t] == 0.2 * m.Pres[scenario, j, t])
            else:
                m.c1.add(m.Pres[scenario, j, t] == 0)
                m.c1.add(m.Qres[scenario, j, t] == 0)

        # _________ Voltage limits _________________________________________________________________________________
        for j in range(0, len(load_in_bus)):
            if j != ref_bus:
                m.c1.add(m.V[scenario, j, t] >= 0.95 ** 2)
                m.c1.add(m.V[scenario, j, t] <= 1.05 ** 2)

        m.c1.add(m.V[scenario, ref_bus, t] == v_ref)

        # ____________________________________________________________________________________________________________
        for j in range(0, len(branch)):
            m.c1.add(m.PF[scenario, j, t] <= 1000 * 1000)
            m.c1.add(m.QF[scenario, j, t] <= 0.2 * 1000 * 1000)

            m.c1.add(m.I[scenario, j, t] <= I_max ** 2)
            m.c1.add(m.I[scenario, j, t] >= -I_max ** 2)


        for i in range(0, len(load_in_bus)):
            m.c1.add(m.P_dso[0, i, t] <= 1000 * 1000)
            m.c1.add(m.P_dso[0, i, t] >= -1000 * 1000)
            m.c1.add(m.P_dso_up[0, i, t] <= 1000 * 1000)
            m.c1.add(m.P_dso_up[0, i, t] >= -1000 * 1000)
            m.c1.add(m.P_dso_down[0, i, t] <= 1000 * 1000)
            m.c1.add(m.P_dso_down[0, i, t] >= -1000 * 1000)

    return m
