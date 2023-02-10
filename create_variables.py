from xlrd import *
from numpy import *
from pyomo.environ import *



def create_variables(m, h, number_buildings, load_in_bus_w, load_in_bus_g, load_in_bus_h, number_EVs):
    n_type_buildings = 1

    m.P_dso = Var(arange(n_type_buildings), arange(len(load_in_bus_w)), arange(h), domain=Reals)
    m.P_dso_gas = Var(arange(n_type_buildings), arange(len(load_in_bus_g)), arange(h), domain=Reals)
    m.P_dso_heat = Var(arange(n_type_buildings), arange(len(load_in_bus_h)), arange(h), domain=NonNegativeReals)
    m.P_dso_up = Var(arange(n_type_buildings), arange(len(load_in_bus_w)), arange(h), domain=Reals)
    m.P_dso_down = Var(arange(n_type_buildings), arange(len(load_in_bus_w)), arange(h), domain=Reals)
    m.P_dso_gas_up = Var(arange(n_type_buildings), arange(len(load_in_bus_g)), arange(h), domain=Reals)
    m.P_dso_gas_down = Var(arange(n_type_buildings), arange(len(load_in_bus_g)), arange(h), domain=Reals)
    m.P_dso_heat_up = Var(arange(n_type_buildings), arange(len(load_in_bus_h)), arange(h), domain=NonNegativeReals)
    m.P_dso_heat_down = Var(arange(n_type_buildings), arange(len(load_in_bus_h)), arange(h), domain=NonNegativeReals)


    m.P_aggr_w = Var(arange(h), domain=Reals)
    m.P_aggr_g = Var(arange(h), domain=Reals)
    m.P_aggr_w_up = Var(arange(len(load_in_bus_w)), arange(h), domain=Reals)
    m.P_aggr_w_down = Var(arange(len(load_in_bus_w)), arange(h), domain=Reals)
    m.P_aggr_g_up = Var(arange(len(load_in_bus_g)), arange(h), domain=Reals)
    m.P_aggr_g_down = Var(arange(len(load_in_bus_g)), arange(h), domain=Reals)


    m.P_w = Var(arange(len(load_in_bus_w)), arange(h), domain=Reals)
    m.P_g = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)
    m.P_h = Var(arange(len(load_in_bus_h)), arange(h), domain=NonNegativeReals)
    m.P_hp = Var(arange(number_buildings), arange(h), domain=NonNegativeReals)
    m.P_gb = Var(arange(number_buildings), arange(h), domain=NonNegativeReals)
    m.P_dh = Var(arange(number_buildings), arange(h), domain=NonNegativeReals)
    m.P_PV = Var(arange(number_buildings), arange(h), domain=NonNegativeReals)
    m.P_PVr = Var(arange(number_buildings), arange(h), domain=NonNegativeReals)
    m.P_PV_cu = Var(arange(number_buildings), arange(h), domain=NonNegativeReals)
    m.P_sto_ch = Var(arange(number_buildings), arange(h + 1), domain=NonNegativeReals)
    m.P_sto_dis = Var(arange(number_buildings), arange(h + 1), domain=NonNegativeReals)
    m.b_sto_ch = Var(arange(number_buildings), arange(h + 1), domain=Binary)
    m.P_soc = Var(arange(number_buildings), arange(h + 1), domain=NonNegativeReals)
    m.P_soc_res = Var(arange(number_buildings), arange(h + 1), domain=NonNegativeReals)
    m.P_EV_ch = Var(arange(number_EVs), arange(h + 1), domain=NonNegativeReals)
    m.P_EV_dis = Var(arange(number_EVs), arange(h + 1), domain=NonNegativeReals)
    m.b_EV_ch = Var(arange(number_EVs), arange(h + 1), domain=Binary)
    m.P_EV_soc = Var(arange(number_EVs), arange(h + 1), domain=NonNegativeReals)
    m.P_inf_load_w = Var(arange(number_buildings), arange(h), domain=Reals)
    m.P_renew = Var(arange(h), domain=Reals)


    m.P_hp_up = Var(arange(number_buildings), arange(h), domain=NonNegativeReals)
    m.P_dh_up = Var(arange(number_buildings), arange(h), domain=NonNegativeReals)
    m.P_PV_up = Var(arange(number_buildings), arange(h), domain=NonNegativeReals)
    m.P_sto_up = Var(arange(number_buildings), arange(h + 1), domain=NonNegativeReals)
    m.P_sto_up_ch = Var(arange(number_buildings), arange(h + 1), domain=NonNegativeReals)
    m.P_sto_up_dis = Var(arange(number_buildings), arange(h + 1), domain=NonNegativeReals)
    m.P_EV_up = Var(arange(number_EVs), arange(h + 1), domain=NonNegativeReals)

    m.P_w_up = Var(arange(len(load_in_bus_w)), arange(h), domain=Reals)
    m.P_g_up = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)
    m.P_h_up = Var(arange(len(load_in_bus_h)), arange(h), domain=NonNegativeReals)


    m.P_hp_down = Var(arange(number_buildings), arange(h), domain=NonNegativeReals)
    m.P_dh_down = Var(arange(number_buildings), arange(h), domain=NonNegativeReals)
    m.P_PV_down = Var(arange(number_buildings), arange(h), domain=NonNegativeReals)
    m.P_sto_down = Var(arange(number_buildings), arange(h + 1), domain=NonNegativeReals)
    m.P_sto_down_ch = Var(arange(number_buildings), arange(h + 1), domain=NonNegativeReals)
    m.P_sto_down_dis = Var(arange(number_buildings), arange(h + 1), domain=NonNegativeReals)
    m.P_EV_down = Var(arange(number_EVs), arange(h + 1), domain=NonNegativeReals)

    m.P_w_down = Var(arange(len(load_in_bus_w)), arange(h), domain=Reals)
    m.P_g_down = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)
    m.P_h_down = Var(arange(len(load_in_bus_h)), arange(h), domain=NonNegativeReals)


    m.temp_building = Var(arange(number_buildings), arange(h + 1), domain=NonNegativeReals)
    m.temp_building_up = Var(arange(number_buildings), arange(h + 1), domain=NonNegativeReals)
    m.temp_building_down = Var(arange(number_buildings), arange(h + 1), domain=NonNegativeReals)


    m.gen_dh_w = Var(arange(len(load_in_bus_h)), arange(h), domain=NonNegativeReals)
    m.gen_dh_w_up = Var(arange(len(load_in_bus_h)), arange(h), domain=NonNegativeReals)
    m.gen_dh_w_down = Var(arange(len(load_in_bus_h)), arange(h), domain=NonNegativeReals)
    m.gen_dh_g = Var(arange(len(load_in_bus_h)), arange(h), domain=NonNegativeReals)


    m.P_chp_w = Var(arange(len(load_in_bus_h)), arange(h), domain=NonNegativeReals)
    m.P_chp_w_up = Var(arange(len(load_in_bus_h)), arange(h), domain=NonNegativeReals)
    m.P_chp_w_down = Var(arange(len(load_in_bus_h)), arange(h), domain=NonNegativeReals)
    m.P_chp_g = Var(arange(len(load_in_bus_h)), arange(h), domain=NonNegativeReals)
    m.P_chp_g_up = Var(arange(len(load_in_bus_h)), arange(h), domain=NonNegativeReals)
    m.P_chp_g_down = Var(arange(len(load_in_bus_h)), arange(h), domain=NonNegativeReals)
    m.P_chp_h = Var(arange(len(load_in_bus_h)), arange(h), domain=NonNegativeReals)
    m.P_chp_h_up = Var(arange(len(load_in_bus_h)), arange(h), domain=NonNegativeReals)
    m.P_chp_h_down = Var(arange(len(load_in_bus_h)), arange(h), domain=NonNegativeReals)


    m.P_aggr_hy = Var(arange(h), domain=Reals)
    m.P_aggr_hy_up = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)
    m.P_aggr_hy_down = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)

    m.P_hy = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)
    m.P_hy_up = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)
    m.P_hy_down = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)

    m.P_P2G_E = Var(arange(len(load_in_bus_w)), arange(h), domain=NonNegativeReals)
    m.U_P2G_E = Var(arange(len(load_in_bus_w)), arange(h), domain=NonNegativeReals)
    m.D_P2G_E = Var(arange(len(load_in_bus_w)), arange(h), domain=NonNegativeReals)
    m.P_P2G_hy = Var(arange(len(load_in_bus_w)), arange(h), domain=NonNegativeReals)
    m.U_P2G_hy = Var(arange(len(load_in_bus_w)), arange(h), domain=NonNegativeReals)
    m.D_P2G_hy = Var(arange(len(load_in_bus_w)), arange(h), domain=NonNegativeReals)
    m.P_P2G_HV_hy = Var(arange(len(load_in_bus_w)), arange(h), domain=NonNegativeReals)
    m.P_P2G_net_hy = Var(arange(len(load_in_bus_w)), arange(h), domain=NonNegativeReals)
    m.U_P2G_net_hy = Var(arange(len(load_in_bus_w)), arange(h), domain=NonNegativeReals)
    m.D_P2G_net_hy = Var(arange(len(load_in_bus_w)), arange(h), domain=NonNegativeReals)
    m.P_P2G_sto_hy = Var(arange(len(load_in_bus_w)), arange(h), domain=NonNegativeReals)
    m.U_P2G_sto_hy = Var(arange(len(load_in_bus_w)), arange(h), domain=NonNegativeReals)
    m.D_P2G_sto_hy = Var(arange(len(load_in_bus_w)), arange(h), domain=NonNegativeReals)
    m.b_sto_hy_ch = Var(arange(number_buildings), arange(h + 1), domain=Binary)
    m.b_sto_hy_dis = Var(arange(number_buildings), arange(h + 1), domain=Binary)

    m.y_soc_sto_hy = Var(arange(len(load_in_bus_g)), arange(h + 1), domain=NonNegativeReals)
    m.y_sto_hy_ch = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)
    m.y_sto_hy_dis = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)
    m.y_sto_HV_hy = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)
    m.P_sto_net_hy = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)
    m.P_sto_FC_hy = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)
    m.D_sto_FC_hy = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)
    m.U_sto_FC_hy = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)

    m.P_dso_hy = Var(arange(n_type_buildings), arange(len(load_in_bus_g)), arange(h), domain=Reals)
    m.P_dso_hy_up = Var(arange(n_type_buildings), arange(len(load_in_bus_g)), arange(h), domain=Reals)
    m.P_dso_hy_down = Var(arange(n_type_buildings), arange(len(load_in_bus_g)), arange(h), domain=Reals)

    m.P_FC_E = Var(arange(len(load_in_bus_w)), arange(h), domain=NonNegativeReals)
    m.D_FC_E = Var(arange(len(load_in_bus_w)), arange(h), domain=NonNegativeReals)
    m.U_FC_E = Var(arange(len(load_in_bus_w)), arange(h), domain=NonNegativeReals)

    m.P_sto_ch_space = Var(arange(number_buildings), arange(h + 1), domain=NonNegativeReals)
    m.P_sto_dis_space = Var(arange(number_buildings), arange(h + 1), domain=NonNegativeReals)
    m.b_sto_space = Var(arange(number_buildings), arange(h + 1), domain=Binary)
    m.P_soc_up = Var(arange(number_buildings), arange(h + 1), domain=NonNegativeReals)
    m.P_soc_down = Var(arange(number_buildings), arange(h + 1), domain=NonNegativeReals)


    return m

def create_variables_EVs(m, h, number_buildings, load_in_bus_w, load_in_bus_g, load_in_bus_h, number_EVs):
    n_type_buildings = 1

    m.P_dso = Var(arange(n_type_buildings), arange(len(load_in_bus_w)), arange(h), domain=Reals)
    m.P_dso_gas = Var(arange(n_type_buildings), arange(len(load_in_bus_g)), arange(h), domain=Reals)
    m.P_dso_heat = Var(arange(n_type_buildings), arange(len(load_in_bus_h)), arange(h), domain=NonNegativeReals)
    m.P_dso_up = Var(arange(n_type_buildings), arange(len(load_in_bus_w)), arange(h), domain=Reals)
    m.P_dso_down = Var(arange(n_type_buildings), arange(len(load_in_bus_w)), arange(h), domain=Reals)
    m.P_dso_gas_up = Var(arange(n_type_buildings), arange(len(load_in_bus_g)), arange(h), domain=Reals)
    m.P_dso_gas_down = Var(arange(n_type_buildings), arange(len(load_in_bus_g)), arange(h), domain=Reals)
    m.P_dso_heat_up = Var(arange(n_type_buildings), arange(len(load_in_bus_h)), arange(h), domain=NonNegativeReals)
    m.P_dso_heat_down = Var(arange(n_type_buildings), arange(len(load_in_bus_h)), arange(h), domain=NonNegativeReals)


    m.P_aggr_w = Var(arange(h), domain=Reals)
    m.P_aggr_g = Var(arange(h), domain=Reals)
    m.P_aggr_w_up = Var(arange(len(load_in_bus_w)), arange(h), domain=Reals)
    m.P_aggr_w_down = Var(arange(len(load_in_bus_w)), arange(h), domain=Reals)
    m.P_aggr_g_up = Var(arange(len(load_in_bus_g)), arange(h), domain=Reals)
    m.P_aggr_g_down = Var(arange(len(load_in_bus_g)), arange(h), domain=Reals)


    m.P_w = Var(arange(len(load_in_bus_w)), arange(h), domain=Reals)
    m.P_g = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)
    m.P_h = Var(arange(len(load_in_bus_h)), arange(h), domain=NonNegativeReals)
    m.P_hp = Var(arange(number_buildings), arange(h), domain=NonNegativeReals)
    m.P_gb = Var(arange(number_buildings), arange(h), domain=NonNegativeReals)
    m.P_dh = Var(arange(number_buildings), arange(h), domain=NonNegativeReals)
    m.P_PV = Var(arange(number_buildings), arange(h), domain=NonNegativeReals)
    m.P_PVr = Var(arange(number_buildings), arange(h), domain=NonNegativeReals)
    m.P_PV_cu = Var(arange(number_buildings), arange(h), domain=NonNegativeReals)
    m.P_sto_ch = Var(arange(number_buildings), arange(h + 1), domain=NonNegativeReals)
    m.P_sto_dis = Var(arange(number_buildings), arange(h + 1), domain=NonNegativeReals)
    m.b_sto_ch = Var(arange(number_buildings), arange(h + 1), domain=Binary)
    m.P_soc = Var(arange(number_buildings), arange(h + 1), domain=NonNegativeReals)
    m.P_soc_res = Var(arange(number_buildings), arange(h + 1), domain=NonNegativeReals)
    m.P_EV_ch = Var(arange(number_EVs), arange(h + 1), domain=NonNegativeReals)
    m.P_EV_dis = Var(arange(number_EVs), arange(h + 1), domain=NonNegativeReals)
    m.b_EV_ch = Var(arange(number_EVs), arange(h + 1), domain=Binary)
    m.P_EV_soc = Var(arange(number_EVs), arange(h + 1), domain=NonNegativeReals)
    m.P_inf_load_w = Var(arange(number_buildings), arange(h), domain=Reals)
    m.P_renew = Var(arange(h), domain=Reals)


    m.P_hp_up = Var(arange(number_buildings), arange(h), domain=NonNegativeReals)
    m.P_dh_up = Var(arange(number_buildings), arange(h), domain=NonNegativeReals)
    m.P_PV_up = Var(arange(number_buildings), arange(h), domain=NonNegativeReals)
    m.P_sto_up = Var(arange(number_buildings), arange(h + 1), domain=NonNegativeReals)
    m.P_sto_up_ch = Var(arange(number_buildings), arange(h + 1), domain=NonNegativeReals)
    m.P_sto_up_dis = Var(arange(number_buildings), arange(h + 1), domain=NonNegativeReals)
    m.P_EV_up = Var(arange(number_EVs), arange(h + 1), domain=NonNegativeReals)
    m.P_EV_up_ch = Var(arange(number_EVs), arange(h + 1), domain=NonNegativeReals)
    m.P_EV_up_dis = Var(arange(number_EVs), arange(h + 1), domain=NonNegativeReals)
    m.P_w_up = Var(arange(len(load_in_bus_w)), arange(h), domain=Reals)
    m.P_g_up = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)
    m.P_h_up = Var(arange(len(load_in_bus_h)), arange(h), domain=NonNegativeReals)


    m.P_hp_down = Var(arange(number_buildings), arange(h), domain=NonNegativeReals)
    m.P_dh_down = Var(arange(number_buildings), arange(h), domain=NonNegativeReals)
    m.P_PV_down = Var(arange(number_buildings), arange(h), domain=NonNegativeReals)
    m.P_sto_down = Var(arange(number_buildings), arange(h + 1), domain=NonNegativeReals)
    m.P_sto_down_ch = Var(arange(number_buildings), arange(h + 1), domain=NonNegativeReals)
    m.P_sto_down_dis = Var(arange(number_buildings), arange(h + 1), domain=NonNegativeReals)
    m.P_EV_down = Var(arange(number_EVs), arange(h + 1), domain=NonNegativeReals)
    m.P_EV_down_ch = Var(arange(number_EVs), arange(h + 1), domain=NonNegativeReals)
    m.P_EV_down_dis = Var(arange(number_EVs), arange(h + 1), domain=NonNegativeReals)
    m.P_w_down = Var(arange(len(load_in_bus_w)), arange(h), domain=Reals)
    m.P_g_down = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)
    m.P_h_down = Var(arange(len(load_in_bus_h)), arange(h), domain=NonNegativeReals)


    m.temp_building = Var(arange(number_buildings), arange(h + 1), domain=NonNegativeReals)
    m.temp_building_up = Var(arange(number_buildings), arange(h + 1), domain=NonNegativeReals)
    m.temp_building_down = Var(arange(number_buildings), arange(h + 1), domain=NonNegativeReals)


    m.gen_dh_w = Var(arange(len(load_in_bus_h)), arange(h), domain=NonNegativeReals)
    m.gen_dh_w_up = Var(arange(len(load_in_bus_h)), arange(h), domain=NonNegativeReals)
    m.gen_dh_w_down = Var(arange(len(load_in_bus_h)), arange(h), domain=NonNegativeReals)
    m.gen_dh_g = Var(arange(len(load_in_bus_h)), arange(h), domain=NonNegativeReals)


    m.P_chp_w = Var(arange(len(load_in_bus_h)), arange(h), domain=NonNegativeReals)
    m.P_chp_w_up = Var(arange(len(load_in_bus_h)), arange(h), domain=NonNegativeReals)
    m.P_chp_w_down = Var(arange(len(load_in_bus_h)), arange(h), domain=NonNegativeReals)
    m.P_chp_g = Var(arange(len(load_in_bus_h)), arange(h), domain=NonNegativeReals)
    m.P_chp_g_up = Var(arange(len(load_in_bus_h)), arange(h), domain=NonNegativeReals)
    m.P_chp_g_down = Var(arange(len(load_in_bus_h)), arange(h), domain=NonNegativeReals)
    m.P_chp_h = Var(arange(len(load_in_bus_h)), arange(h), domain=NonNegativeReals)
    m.P_chp_h_up = Var(arange(len(load_in_bus_h)), arange(h), domain=NonNegativeReals)
    m.P_chp_h_down = Var(arange(len(load_in_bus_h)), arange(h), domain=NonNegativeReals)


    m.P_aggr_hy = Var(arange(h), domain=Reals)
    m.P_aggr_hy_up = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)
    m.P_aggr_hy_down = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)

    m.P_hy = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)
    m.P_hy_up = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)
    m.P_hy_down = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)

    m.P_P2G_E = Var(arange(len(load_in_bus_w)), arange(h), domain=NonNegativeReals)
    m.U_P2G_E = Var(arange(len(load_in_bus_w)), arange(h), domain=NonNegativeReals)
    m.D_P2G_E = Var(arange(len(load_in_bus_w)), arange(h), domain=NonNegativeReals)
    m.P_P2G_hy = Var(arange(len(load_in_bus_w)), arange(h), domain=NonNegativeReals)
    m.U_P2G_hy = Var(arange(len(load_in_bus_w)), arange(h), domain=NonNegativeReals)
    m.D_P2G_hy = Var(arange(len(load_in_bus_w)), arange(h), domain=NonNegativeReals)
    m.P_P2G_HV_hy = Var(arange(len(load_in_bus_w)), arange(h), domain=NonNegativeReals)
    m.P_P2G_net_hy = Var(arange(len(load_in_bus_w)), arange(h), domain=NonNegativeReals)
    m.U_P2G_net_hy = Var(arange(len(load_in_bus_w)), arange(h), domain=NonNegativeReals)
    m.D_P2G_net_hy = Var(arange(len(load_in_bus_w)), arange(h), domain=NonNegativeReals)
    m.P_P2G_sto_hy = Var(arange(len(load_in_bus_w)), arange(h), domain=NonNegativeReals)
    m.U_P2G_sto_hy = Var(arange(len(load_in_bus_w)), arange(h), domain=NonNegativeReals)
    m.D_P2G_sto_hy = Var(arange(len(load_in_bus_w)), arange(h), domain=NonNegativeReals)
    m.b_sto_hy_ch = Var(arange(number_buildings), arange(h + 1), domain=Binary)
    m.b_sto_hy_dis = Var(arange(number_buildings), arange(h + 1), domain=Binary)

    m.y_soc_sto_hy = Var(arange(len(load_in_bus_g)), arange(h + 1), domain=NonNegativeReals)
    m.y_sto_hy_ch = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)
    m.y_sto_hy_dis = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)
    m.y_sto_HV_hy = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)
    m.P_sto_net_hy = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)
    m.P_sto_FC_hy = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)
    m.D_sto_FC_hy = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)
    m.U_sto_FC_hy = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)

    m.P_dso_hy = Var(arange(n_type_buildings), arange(len(load_in_bus_g)), arange(h), domain=Reals)
    m.P_dso_hy_up = Var(arange(n_type_buildings), arange(len(load_in_bus_g)), arange(h), domain=Reals)
    m.P_dso_hy_down = Var(arange(n_type_buildings), arange(len(load_in_bus_g)), arange(h), domain=Reals)

    m.P_FC_E = Var(arange(len(load_in_bus_w)), arange(h), domain=NonNegativeReals)
    m.D_FC_E = Var(arange(len(load_in_bus_w)), arange(h), domain=NonNegativeReals)
    m.U_FC_E = Var(arange(len(load_in_bus_w)), arange(h), domain=NonNegativeReals)

    m.P_sto_ch_space = Var(arange(number_buildings), arange(h + 1), domain=NonNegativeReals)
    m.P_sto_dis_space = Var(arange(number_buildings), arange(h + 1), domain=NonNegativeReals)
    m.b_sto_space = Var(arange(number_buildings), arange(h + 1), domain=Binary)
    m.P_soc_up = Var(arange(number_buildings), arange(h + 1), domain=NonNegativeReals)
    m.P_soc_down = Var(arange(number_buildings), arange(h + 1), domain=NonNegativeReals)

    m.P_EV_ch_space = Var(arange(number_EVs), arange(h + 1), domain=NonNegativeReals)
    m.P_EV_dis_space = Var(arange(number_EVs), arange(h + 1), domain=NonNegativeReals)
    m.b_EV_space = Var(arange(number_EVs), arange(h + 1), domain=Binary)
    m.P_EV_soc_up = Var(arange(number_EVs), arange(h + 1), domain=NonNegativeReals)
    m.P_EV_soc_down = Var(arange(number_EVs), arange(h + 1), domain=NonNegativeReals)

    return m


def create_variables_power_flow(m, h, branch, load_in_bus_w):
    number_scenarios = 1
    m.PF = Var(arange(number_scenarios), arange(len(branch)), arange(h), domain=Reals)  # purchased power
    m.QF = Var(arange(number_scenarios), arange(len(branch)), arange(h), domain=Reals)  # purchased power
    m.V = Var(arange(number_scenarios), arange(len(load_in_bus_w)), arange(h), domain=Reals)
    m.U = Var(arange(number_scenarios), arange(len(load_in_bus_w)), arange(h), domain=NonNegativeReals)
    m.I = Var(arange(number_scenarios), arange(len(branch)), arange(h), domain=Reals)
    m.Pres = Var(arange(number_scenarios), arange(len(load_in_bus_w)), arange(h), domain=Reals)
    m.Qres = Var(arange(number_scenarios), arange(len(load_in_bus_w)), arange(h), domain=Reals)
    m.V_viol = Var(arange(number_scenarios), arange(len(load_in_bus_w)), arange(h), domain=NonNegativeReals)
    m.net_load = Var(arange(number_scenarios), arange(len(load_in_bus_w)-1), arange(h), domain = Reals)
    m.chp_gen_w = Var(arange(number_scenarios), arange(2), arange(h), domain = NonNegativeReals)



    m.P_dso = Var(arange(number_scenarios), arange(len(load_in_bus_w)), arange(h), domain=Reals)
    m.P_dso_up = Var(arange(number_scenarios), arange(len(load_in_bus_w)), arange(h), domain=Reals)
    m.P_dso_down = Var(arange(number_scenarios), arange(len(load_in_bus_w)), arange(h), domain=Reals)


    return m


def create_variables_power_flow_gas(m, branch, load_in_bus_g, t, iter, m_old):
    h = 1
    m.PF_gas = Var(arange(len(branch)), arange(h), domain=NonNegativeReals, bounds=(1e-20,None), initialize=1e-20)  # purchased power
    if iter > 0:
        for i in range(0, len(branch)):
            m.PF_gas[i, 0] = m_old[t].PF_gas[i, 0].value
    m.PF_in = Var(arange(len(branch)), arange(h), domain=NonNegativeReals)  # purchased power
    if iter > 0:
        for i in range(0, len(branch)):
            m.PF_in[i, 0] = m_old[t].PF_in[i, 0].value
    m.PF_out = Var(arange(len(branch)), arange(h), domain=NonNegativeReals)  # purchased power
    if iter > 0:
        for i in range(0, len(branch)):
            m.PF_out[i, 0] = m_old[t].PF_out[i, 0].value
    m.Pgen_gas = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)
    if iter > 0:
        for i in range(0, len(load_in_bus_g)):
            m.Pgen_gas[i, 0] = m_old[t].Pgen_gas[i, 0].value
    m.p = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)
    if iter > 0:
        for i in range(0, len(load_in_bus_g)):
            m.p[i, 0] = m_old[t].p[i, 0].value
    m.p2 = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)
    if iter > 0:
        for i in range(0, len(load_in_bus_g)):
            m.p2[i, 0] = m_old[t].p2[i, 0].value

    m.P_dso_gas = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)
    if iter > 0:
        for i in range(0, len(load_in_bus_g)):
            m.P_dso_gas[i, 0] = m_old[t].P_dso_gas[i, 0].value
    m.P_dso_gas_up = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)
    if iter > 0:
        for i in range(0, len(load_in_bus_g)):
            m.P_dso_gas_up[i, 0] = m_old[t].P_dso_gas_up[i, 0].value
    m.P_dso_gas_down = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)
    if iter > 0:
        for i in range(0, len(load_in_bus_g)):
            m.P_dso_gas_down[i, 0] = m_old[t].P_dso_gas_down[i, 0].value

    m.P_dso_hy = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)
    if iter > 0:
        for i in range(0, len(load_in_bus_g)):
            m.P_dso_hy[i, 0] = m_old[t].P_dso_hy[i, 0].value
    m.P_dso_hy_up = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)
    if iter > 0:
        for i in range(0, len(load_in_bus_g)):
            m.P_dso_hy_up[i, 0] = m_old[t].P_dso_hy_up[i, 0].value
    m.P_dso_hy_down = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)
    if iter > 0:
        for i in range(0, len(load_in_bus_g)):
            m.P_dso_hy_down[i, 0] = m_old[t].P_dso_hy_down[i, 0].value

    m.WI = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)
    if iter > 0:
        for i in range(0, len(load_in_bus_g)):
            m.WI[i, 0] = m_old[t].WI[i, 0].value
    m.HHV_mix = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)
    if iter > 0:
        for i in range(0, len(load_in_bus_g)):
            m.HHV_mix[i, 0] = m_old[t].HHV_mix[i, 0].value
    m.S_mix = Var(arange(len(load_in_bus_g)), arange(h), domain=NonNegativeReals)
    if iter > 0:
        for i in range(0, len(load_in_bus_g)):
            m.S_mix[i, 0] = m_old[t].S_mix[i, 0].value

    return m



def create_variables_power_flow_heat(m, m_old, h, t, iter, branch, buses, number_buildings):

    m.P_dso_heat = Var(arange(number_buildings), arange(h), domain=Reals)
    if iter > 0:
        for i in range(0, number_buildings):
            m.P_dso_heat[i, 0] = m_old[t].P_dso_heat[i, 0].value
    m.P_dso_heat_up = Var(arange(number_buildings), arange(h), domain=Reals)
    if iter > 0:
        for i in range(0, number_buildings):
            m.P_dso_heat_up[i, 0] = m_old[t].P_dso_heat_up[i, 0].value
    m.P_dso_heat_down = Var(arange(number_buildings), arange(h), domain=Reals)
    if iter > 0:
        for i in range(0, number_buildings):
            m.P_dso_heat_down[i, 0] = m_old[t].P_dso_heat_down[i, 0].value
    m.h = Var(arange(number_buildings), arange(h), domain=NonNegativeReals)



    m.m = Var(arange(len(branch)), arange(h), domain = NonNegativeReals)
    if iter > 0:
        for i in range(0, len(branch)):
            m.m[i, 0] = m_old[t].m[i, 0].value
    m.mq_gen = Var(arange(len(buses)), arange(h), domain = NonNegativeReals)
    if iter > 0:
        for i in range(0, len(buses)):
            m.mq_gen[i, 0] = m_old[t].mq_gen[i, 0].value
    m.mq_load = Var(arange(len(buses)), arange(h), domain = NonNegativeReals)
    if iter > 0:
        for i in range(0, len(buses)):
            m.mq_load[i, 0] = m_old[t].mq_load[i, 0].value
    m.Ts = Var(arange(len(buses)), arange(h), domain = NonNegativeReals)
    if iter > 0:
        for i in range(0, len(buses)):
            m.Ts[i, 0] = m_old[t].Ts[i, 0].value
    m.To = Var(arange(len(buses)), arange(h), domain = NonNegativeReals)
    if iter > 0:
        for i in range(0, len(buses)):
            m.To[i, 0] = m_old[t].To[i, 0].value
    m.Ts_o = Var(arange(len(buses)), arange(h), domain = NonNegativeReals)
    if iter > 0:
        for i in range(0, len(buses)):
            m.Ts_o[i, 0] = m_old[t].Ts_o[i, 0].value
    m.T_start = Var(arange(len(branch)), arange(h), domain = NonNegativeReals)
    if iter > 0:
        for i in range(0, len(branch)):
            m.T_start[i, 0] = m_old[t].T_start[i, 0].value
    m.T_end = Var(arange(len(branch)), arange(h), domain = NonNegativeReals)
    if iter > 0:
        for i in range(0, len(branch)):
            m.T_end[i, 0] = m_old[t].T_end[i, 0].value
    m.p = Var(arange(len(buses)), arange(h), domain = NonNegativeReals)
    if iter > 0:
        for i in range(0, len(buses)):
            m.p[i, 0] = m_old[t].p[i, 0].value

    m.m_return = Var(arange(len(branch)), arange(h), domain = NonNegativeReals)
    if iter > 0:
        for i in range(0, len(branch)):
            m.m_return[i, 0] = m_old[t].m_return[i, 0].value
    m.p_return = Var(arange(len(buses)), arange(h), domain=NonNegativeReals)
    if iter > 0:
        for i in range(0, len(buses)):
            m.p_return[i, 0] = m_old[t].p_return[i, 0].value
    m.Ts_return = Var(arange(len(buses)), arange(h), domain=NonNegativeReals)
    if iter > 0:
        for i in range(0, len(buses)):
            m.Ts_return[i, 0] = m_old[t].Ts_return[i, 0].value
    m.T_start_return = Var(arange(len(branch)), arange(h), domain = NonNegativeReals)
    if iter > 0:
        for i in range(0, len(branch)):
            m.T_start_return[i, 0] = m_old[t].T_start_return[i, 0].value
    m.T_end_return = Var(arange(len(branch)), arange(h), domain = NonNegativeReals)
    if iter > 0:
        for i in range(0, len(branch)):
            m.T_end_return[i, 0] = m_old[t].T_end_return[i, 0].value
    m.mq_gen_return = Var(arange(len(buses)), arange(h), domain = NonNegativeReals)
    if iter > 0:
        for i in range(0, len(buses)):
            m.mq_gen_return[i, 0] = m_old[t].mq_gen_return[i, 0].value
    m.mq_load_return = Var(arange(len(buses)), arange(h), domain = NonNegativeReals)
    if iter > 0:
        for i in range(0, len(buses)):
            m.mq_load_return[i, 0] = m_old[t].mq_load_return[i, 0].value


    return m



