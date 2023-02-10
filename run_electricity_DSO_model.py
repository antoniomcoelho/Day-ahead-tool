from copy import *

from create_variables import *
from optimization import *
from power_flow import *


def print_scenario_details(s):
    print("")
    if s == 0:
        print("______ Scenario Energy ______")
    elif s == 1:
        print("______ Scenario Up ______")
    elif s == 2:
        print("______ Scenario Down ______")


def run_scenario_upward(m1, m, t, pi, ro, load_in_bus_w, time_h, m1_h_up, P_dso_w_up):
    # Optimize model
    m1, time_h = optimization_dso_up(m1, m, t, pi, ro, load_in_bus_w, time_h)

    # Append results
    m1_h_up.append(deepcopy(m1))
    P_dso_prov_w_up = []
    for i in range(0, len(load_in_bus_w)):
        P_dso_prov_w_up.append(m1.P_dso_up[0, i, 0].value)
    P_dso_w_up.append(P_dso_prov_w_up)

    return m1_h_up, P_dso_w_up


def run_scenario_downward(m1, m, t, pi, ro, load_in_bus_w, time_h, m1_h_down, P_dso_w_down):
    # Optimize model
    m1, time_h = optimization_dso_down(m1, m, t, pi, ro, load_in_bus_w, time_h)

    # Append results
    m1_h_down.append(deepcopy(m1))
    P_dso_prov_w_down = []
    for i in range(0, len(load_in_bus_w)):
        P_dso_prov_w_down.append(deepcopy(m1.P_dso_down[0, i, 0].value))
    P_dso_w_down.append(P_dso_prov_w_down)

    return m1_h_down, P_dso_w_down





def run_electricity_DSO_model(m, h, branch_w, load_in_bus_w, other_w, pi, ro, time_h):
    results_m1 = []

    for s in range(1, 3): # 1 - upward scenario / 2 - downward scenario
        print_scenario_details(s)

        P_dso_w_up = []
        m1_h_up = []
        P_dso_w_down = []
        m1_h_down = []
        for t in range(0, h):

            # Model initialization
            m1 = ConcreteModel()
            m1.c1 = ConstraintList()
            m1 = create_variables_power_flow(m1, h, branch_w, load_in_bus_w)

            # Optimal power flow restrictions
            m1 = power_flow_elec(m1, s, branch_w, load_in_bus_w, other_w)

            # Solve optimal power flow
            if s == 1:
                m1_h_up, P_dso_w_up = run_scenario_upward(m1, m, t, pi, ro, load_in_bus_w, time_h, m1_h_up, P_dso_w_up)
            elif s == 2:
                m1_h_down, P_dso_w_down = run_scenario_downward(m1, m, t, pi, ro, load_in_bus_w, time_h, m1_h_down, P_dso_w_down)

    # Append results to be used later
    results_m1.append(m1_h_down)
    results_m1.append(m1_h_up)
    results_m1.append(m1_h_down)

    return results_m1, P_dso_w_up, P_dso_w_down, m1_h_up, m1_h_down