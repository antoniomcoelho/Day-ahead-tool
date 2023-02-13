from copy import *

from create_variables import *
from optimization import *
from power_flow_gas import *


def run_gas_DSO_model(m, m2_h_up, h, branch_g, load_in_bus_g, other_g, pi, ro, time_h, iter):
    # Initialization of variables
    P_dso_g = []
    P_dso_g_up = []
    P_dso_g_down = []

    P_dso_hy = []
    P_dso_hy_up = []
    P_dso_hy_down = []

    m2_old = deepcopy(m2_h_up)
    m2_h = []
    m2_h_up = []
    m2_h_down = []
    results_m2 = []

    for s in range(1, 3):  # 0 - energy scenario / 1 - upward scenario / 2 - downward scenario
        print_scenario_details(s)

        for t in range(0, h):
            # Model initialization
            m2 = model_initialization(branch_g, load_in_bus_g, t, iter, m2_old)

            # Optimal gas flow restrictions
            m2 = power_flow_gas(m2, s, branch_g, load_in_bus_g, other_g)

            # Solve optimal gas flow
            if s == 0:
                m2_h, P_dso_g, P_dso_hy = run_scenario_energy(m2, m, t, pi, ro, load_in_bus_g, time_h, m2_h, P_dso_g, P_dso_hy)
            elif s == 1:
                m2_h_up, P_dso_g_up, P_dso_hy_up = run_scenario_upward(m2, m, t, pi, ro, load_in_bus_g, time_h, m2_h_up, P_dso_g_up, P_dso_hy_up)
            elif s == 2:
                m2_h_down, P_dso_g_down, P_dso_hy_down = run_scenario_downward(m2, m, t, pi, ro, load_in_bus_g, time_h, m2_h_down, P_dso_g_down, P_dso_hy_down)

    # Append results to be used later
    results_m2.append(m2_h_down)
    results_m2.append(m2_h_up)
    results_m2.append(m2_h_down)

    return results_m2, m2_h_up, P_dso_g_up, P_dso_hy_up, m2_h_down, P_dso_g_down, P_dso_hy_down


def print_scenario_details(s):
    print("")
    if s == 0:
        print("______ Scenario Energy ______")
    elif s == 1:
        print("______ Scenario Up ______")
    elif s == 2:
        print("______ Scenario Down ______")

def model_initialization( branch_g, load_in_bus_g, t, iter, m2_old):
    m2 = ConcreteModel()
    m2.c1 = ConstraintList()
    m2 = create_variables_power_flow_gas(m2, branch_g, load_in_bus_g, t, iter, m2_old)

    return m2

def run_scenario_energy(m2, m, t, pi, ro, load_in_bus_g, time_h, m2_h, P_dso_g, P_dso_hy):
    # Optimize model
    m2, time_h = optimization_dso_gas(m2, m, t, pi, ro, load_in_bus_g, time_h)

    # Append results
    m2_h.append(m2)
    P_dso_prov_g = []
    P_dso_prov_hy = []
    for i in range(0, len(load_in_bus_g)):
        P_dso_prov_g.append(m2.P_dso_gas[i, 0].value)
        P_dso_prov_hy.append(m2.P_dso_hy[i, 0].value)
    P_dso_g.append(P_dso_prov_g)
    P_dso_hy.append(P_dso_prov_hy)

    return m2_h, P_dso_g, P_dso_hy


def run_scenario_upward(m2, m, t, pi, ro, load_in_bus_g, time_h, m2_h_up, P_dso_g_up, P_dso_hy_up):
    # Optimize model
    m2, time_h = optimization_dso_gas_up(m2, m, t, pi, ro, load_in_bus_g, time_h)

    # Append results
    m2_h_up.append(m2)
    P_dso_prov_g_up = []
    P_dso_prov_hy_up = []
    for i in range(0, len(load_in_bus_g)):
        P_dso_prov_g_up.append(m2.P_dso_gas_up[i, 0].value)
        P_dso_prov_hy_up.append(m2.P_dso_hy_up[i, 0].value)
    P_dso_g_up.append(P_dso_prov_g_up)
    P_dso_hy_up.append(P_dso_prov_hy_up)

    return m2_h_up, P_dso_g_up, P_dso_hy_up

def run_scenario_downward(m2, m, t, pi, ro, load_in_bus_g, time_h, m2_h_down, P_dso_g_down, P_dso_hy_down):
    # Optimize model
    m2, time_h = optimization_dso_gas_down(m2, m, t, pi, ro, load_in_bus_g, time_h)

    # Append results
    m2_h_down.append(m2)
    P_dso_prov_g_down = []
    P_dso_prov_hy_down = []
    for i in range(0, len(load_in_bus_g)):
        P_dso_prov_g_down.append(m2.P_dso_gas_down[i, 0].value)
        P_dso_prov_hy_down.append(m2.P_dso_hy_down[i, 0].value)
    P_dso_g_down.append(P_dso_prov_g_down)
    P_dso_hy_down.append(P_dso_prov_hy_down)

    return m2_h_down, P_dso_g_down, P_dso_hy_down


