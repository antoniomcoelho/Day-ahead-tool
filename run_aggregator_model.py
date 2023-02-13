from math import *
import pandas as pd

from create_variables import *
from run_model import *
from optimization import *


def run_aggregator_model(h, number_buildings, load_in_bus_w, load_in_bus_g, load_in_bus_h, number_EVs, profile_solar, temperature_outside, resources_agr,
                                gen_dh, other_g, other_h, fuel_station, EVs, prices, iter, pi, ro, Pa, time_h, costs_all):

    # Model initialization
    m = model_initialization(h, number_buildings, load_in_bus_w, load_in_bus_g, load_in_bus_h, number_EVs)

    # Add aggregator constraints to model
    m = run_model_buildings(m, h, number_buildings, profile_solar, temperature_outside, resources_agr,
                            load_in_bus_w, load_in_bus_g, load_in_bus_h, gen_dh, other_h, fuel_station, EVs, prices)

    m, time_h, cost = optimization_aggregator(m, h, prices, 0, iter, pi, ro, Pa, load_in_bus_w, load_in_bus_g,
                                              load_in_bus_h,
                                              other_g, other_h, 1, time_h, fuel_station, number_buildings)

    costs_all.append(cost)
    pd.DataFrame(costs_all).to_csv("costs.csv")

    return m, time_h


def model_initialization(h, number_buildings, load_in_bus_w, load_in_bus_g, load_in_bus_h, number_EVs):
    m = ConcreteModel()
    m.c1 = ConstraintList()
    m = create_variables(m, h, number_buildings, load_in_bus_w, load_in_bus_g, load_in_bus_h, number_EVs)

    return m
