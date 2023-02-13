run_aggregator_model()
m = ConcreteModel()
m.c1 = ConstraintList()

m = create_variables(m, h, number_buildings, load_in_bus_w, load_in_bus_g, load_in_bus_h, number_EVs)

m = run_model_buildings(m, h, number_buildings, profile_solar, temperature_outside, resources_agr,
                        load_in_bus_w, load_in_bus_g, load_in_bus_h, gen_dh, other_h, fuel_station, EVs, prices)

m, time_h, cost = optimization_aggregator(m, h, prices, 0, iter, pi, ro, Pa, load_in_bus_w, load_in_bus_g,
                                          load_in_bus_h,
                                          other_g, other_h, b_prints, time_h, fuel_station, number_buildings)

costs_all.append(cost)
pd.DataFrame(costs_all).to_csv("costs.csv")