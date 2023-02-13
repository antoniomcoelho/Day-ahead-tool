from copy import *

from create_variables import *
from optimization import *
from power_flow_heat import *


def run_heat_DSO_model(m, m3_h, h, branch_h, load_in_bus_h, other_h, pi, ro, time_h, iter, number_buildings, load_h, resources_agr, heat_network):
    # Initialization of variables
    P_dso_h = []
    P_dso_h_up = []
    P_dso_h_down = []

    m3_old = deepcopy(m3_h)
    m3_h = []
    m3_h_up = []
    m3_h_down = []

    for s in range(0, 3): # 0 - energy scenario / 1 - upward scenario / 2 - downward scenario
        print_scenario_details(s)

        for t in range(0, h):
            flag_error = 1
            while flag_error:

                # Model Initialization
                m3 = model_initialization(m3_old, t, iter, branch_h, load_in_bus_h, number_buildings)

                # Optimal heat flow restrictions
                m3 = power_flow_heat(m3, m, t, s, branch_h, load_h, load_in_bus_h, resources_agr, heat_network, other_h)

                # Solve optimal gas flow
                if s == 0:
                    m3_h, P_dso_h, flag_error, time_h  = run_scenario_energy(m3, m, t, pi, ro, load_in_bus_h, time_h, m3_h, P_dso_h)

                elif s == 1:
                    m3_h_up, P_dso_h_up, flag_error, time_h = run_scenario_upward(m3, m, t, pi, ro, load_in_bus_h, time_h, m3_h_up, P_dso_h_up)

                elif s == 2:
                    m3_h_down, P_dso_h_down, flag_error, time_h = run_scenario_downward(m3, m, t, pi, ro, load_in_bus_h, time_h, m3_h_down, P_dso_h_down)

    results_m3 = [m3_h, m3_h_up, m3_h_down]

    return results_m3, m3_h, P_dso_h, m3_h_up, P_dso_h_up, m3_h_down, P_dso_h_down, time_h



def print_scenario_details(s):
    print("")
    if s == 0:
        print("______ Scenario Energy ______")
    elif s == 1:
        print("______ Scenario Up ______")
    elif s == 2:
        print("______ Scenario Down ______")

def model_initialization(m3_old, t, iter, branch_h, load_in_bus_h, number_buildings):
    m3 = ConcreteModel()
    m3.c1 = ConstraintList()

    m3 = create_variables_power_flow_heat(m3, m3_old, 1, t, iter, branch_h, load_in_bus_h, number_buildings)

    return m3



def run_scenario_energy(m3, m, t, pi, ro, load_in_bus_h, time_h, m3_h, P_dso_h):
    m3, flag_error, time_h = optimization_dso_heat(m3, m, t, pi, ro, load_in_bus_h, time_h)

    if flag_error == 0:
        m3_h.append(m3)

        P_dso_prov_gen = []
        for i in range(0, len(load_in_bus_h)):
            P_dso_prov_gen.append(m3.P_dso_heat[i, 0].value)
        P_dso_h.append(P_dso_prov_gen)

    return m3_h, P_dso_h, flag_error, time_h

def run_scenario_upward(m3, m, t, pi, ro, load_in_bus_h, time_h, m3_h_up, P_dso_h_up):
    m3, flag_error, time_h = optimization_dso_heat_up(m3, m, t, pi, ro, load_in_bus_h, time_h)

    if flag_error == 0:
        m3_h_up.append(m3)

        P_dso_prov_gen_up = []
        for i in range(0, len(load_in_bus_h)):
            P_dso_prov_gen_up.append(m3.P_dso_heat_up[i, 0].value)
        P_dso_h_up.append(P_dso_prov_gen_up)

    return m3_h_up, P_dso_h_up, flag_error, time_h

def run_scenario_downward(m3, m, t, pi, ro, load_in_bus_h, time_h, m3_h_down, P_dso_h_down):
    m3, flag_error, time_h = optimization_dso_heat_down(m3, m, t, pi, ro, load_in_bus_h, time_h)

    if flag_error == 0:
        m3_h_down.append(m3)

        P_dso_prov_gen_down = []
        for i in range(0, len(load_in_bus_h)):
            P_dso_prov_gen_down.append(m3.P_dso_heat_down[i, 0].value)
        P_dso_h_down.append(P_dso_prov_gen_down)

    return m3_h_down, P_dso_h_down, flag_error, time_h
