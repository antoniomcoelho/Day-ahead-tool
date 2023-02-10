from xlrd import *
from numpy import *
from pyomo.environ import *
from math import *
import pandas as pd



def print_results_aggregator(m, h, load_in_bus_w, load_in_bus_g, load_in_bus_h):
    print("_____ Scenario Energy _____")
    for t in range(0, h):
        print("")
        print("Hour", t)
        print("Electricity consumption:", round(sum(m.P_w[n, t].value for n in range(0, len(load_in_bus_w))), 2), "kW")
        print("Gas consumption:", round(sum(m.P_g[n, t].value for n in range(0, len(load_in_bus_g))), 2), "kW")
        print("Districh heating consumption:", round(sum(m.P_h[n, t].value for n in range(0, len(load_in_bus_h))), 2), "kW")
        print("Districh heating load:", round(sum(m.P_h[n, t].value for n in range(0, 26)), 2), "kW")
        print("Districh heating gen:", round(sum(m.P_h[n, t].value for n in range(26, 28)), 2), "kW")
    for t in range(11, 14):
            print("")
            for i in range(0, len(load_in_bus_h)):
                print(t, i, m.P_h[i, t].value, m.P_dso_heat[0, i, t].value)


    print("")
    print("_____ Scenario Up _____")
    for t in range(0, h):
        print("")
        print("Hour", t)
        print("Electricity consumption:", round(sum(m.P_w_up[n, t].value for n in range(0, len(load_in_bus_w))), 2), "kW")

    print("")
    print("_____ Scenario Down _____")
    for t in range(0, h):
        print("")
        print("Hour", t)
        print("Electricity consumption:", round(sum(m.P_w_down[n, t].value for n in range(0, len(load_in_bus_w))), 2), "kW")

    return 0



