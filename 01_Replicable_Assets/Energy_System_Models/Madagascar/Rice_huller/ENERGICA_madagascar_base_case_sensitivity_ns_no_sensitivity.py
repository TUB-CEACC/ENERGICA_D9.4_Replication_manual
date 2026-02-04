#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Jul 15 16:49:43 2021

@author: nikolas

model for ENERGICA T1.2.2 (energy system models)
Demonstrator: madagascar

Case: Base Case Rice huller

Comments:
    + PV size: variable
    + Battery Capacity: variable
    + DC Rice huller Power
    + Rice Huller operation: fixed time series as supply
    + rice revenue is an external calculation
   
@Edouard updates: 01/12/2022  
    + modifications name investment functions
    + sensitivity analysis corrected with Sensitivity Analyzer
    + components added: rice huller, #diesel huller, demand 1 to 5 (change the input excel data)
    + fixed cost taken into account in oemof optimisation
    + Evaluation Criterias defined and returned by the script
    
    
 


"""

###############################################################################
# imports

import pandas as pd
import xlsxwriter
import numpy as np
import os
import logging
import pprint as pp

from oemof.tools import logger
from oemof import solph
import pyomo.environ as po
from oemof.tools import economics
from oemof.solph import (Sink, Transformer, Source, Bus, Flow, NonConvex,
                               Model, EnergySystem, components, custom, Investment, views)
from oemof.solph import constraints

from sensitivity import SensitivityAnalyzer
try:
    import matplotlib.pyplot as plt
    plt.style.use('seaborn')
except ImportError:
    plt = None

        
# custom function 
from madagascar_economic_fncs_base_case import (number_of_investement,
                                   residual_value,
                                   Capex,
                                   CRF_fnc,
                                   annuity,
                                   cost_component_optimized,
                                   system_opt_result,
                                   fuel_price_change,
                                   list_price,
                                   Capex_replacement, #not used
                                   Capex_reinvest   #not used
                                   )

#to be considered: To use from flexible case? --> try flexible case
#from system_results_fncs_base_case_v1        
#changed from system_results_fncs_base_case_v1 to Madagascar_system_results_fncs_status_quo_v1 
from system_results_fncs_flexible_case_v1 import (sizing_fnc,
                                         results_postprocessing,
                                         pv_optimized,
                                         battery_optimized,
                                         controller_optimized,
                                         rice_huller_optimized,                   
                                         legend_fnc,
                                         get_energy_controller,
                                         )
                             
###############################################################################         
#SAVE IN EXCEL: Definition of the output file name and the file path

File_name = 'Results_rice_huller_high_wntl.xlsx'
path_out = "C:/Energica/WP2/"



# path and filename to save the hourly energyflow data: operation behaviour
operation_behaviour_analysis= 'on' 
file_name_operation_behaviour=  'Results_rice_huller_high_wntl_opb.xlsx'  
sheet_name_OPB='OPB'


date_from= "2019-11-01 00:00:00"#"01/01/2019  00:00:00"
date_to= "2019-11-01 00:00:00"# "01/01/2019  23:00:00" 
figname_OPB_plot='OBP_plot'


# select on if you want to produce a operation behaviour plot (quick and dirty) a better 
# plot style can be done in a seperate analysis 
operation_behaviour_analysis_plot= 'off'
# choose th period of time you want to plot for the operation behaviour
#2021-11-01 00:00:00

###############################################################################
  #Sensitivity analysis 

#sensitivity: fill the parameters name and values in the sensitivity_dict and add the parameter in the sensitivity_model function
#sensitivity_dict = {
#    'scenario': [1,2,3,4,5],   # scenario index
     
#}

#def sensitivity_model(scenario=1):

###############################################################################
# System configuration 
result_dict={}

###############################################################################
# Add a solver 
# solver = "cbc"

solver = 'cbc'


###############################################################################
# Load input Data 
# Read data file

input_filename = os.path.join(os.getcwd(), "Madagascar_input_Base_case.xlsx")
feedin = pd.read_excel(input_filename,sheet_name='INPUT' ) 

#########
feedin["Demand_el_0"]          #cluster 0
feedin["Demand_el_1"]          #cluster 1
feedin["Demand_el_2"]          #cluster 2
#########
demand_1 = feedin["Demand_el_0"]          #cluster 0
demand_normalized_1 = feedin["Demand_el_normalized_0"]
demand_2 = feedin["Demand_el_0"]          #cluster 0
demand_normalized_2 = feedin["Demand_el_normalized_0"]
demand_3 = feedin["Demand_el_0"]          #cluster 0
demand_normalized_3 = feedin["Demand_el_normalized_0"]
demand_4 = feedin["Demand_el_0"]          #cluster 0
demand_normalized_4 = feedin["Demand_el_normalized_0"]
demand_5 = feedin["Demand_el_2"]          #cluster 2
demand_normalized_5 = feedin["Demand_el_normalized_2"]

"""    
    
    if scenario == 1:
        
         
        demand_1 = feedin["Demand_el_0"]          #cluster 0
        demand_normalized_1 = feedin["Demand_el_normalized_0"]

        demand_2 = feedin["Demand_el_1"]          #cluster 1
        demand_normalized_2 = feedin["Demand_el_normalized_1"]
        demand_3 = feedin["Demand_el_1"]          #cluster 1
        demand_normalized_3 = feedin["Demand_el_normalized_1"]
        demand_4 = feedin["Demand_el_1"]          #cluster 1
        demand_normalized_4 = feedin["Demand_el_normalized_1"]

        demand_5 = feedin["Demand_el_2"]          #cluster 2
        demand_normalized_5 = feedin["Demand_el_normalized_2"]

    ######################################################
    #scenario 2 low demand consumer composition [0,5,0]
    if scenario == 2:
    
        demand_1=feedin["Demand_el_1"] 
        demand_normalized_1 = feedin["Demand_el_normalized_1"]
        demand_2 = feedin["Demand_el_1"]          #cluster 1
        demand_normalized_2 = feedin["Demand_el_normalized_1"]
        demand_3 = feedin["Demand_el_1"]          #cluster 1
        demand_normalized_3 = feedin["Demand_el_normalized_1"]
        demand_4 = feedin["Demand_el_1"]          #cluster 1
        demand_normalized_4 = feedin["Demand_el_normalized_1"]
        demand_5 = feedin["Demand_el_1"]          #cluster 1
        demand_normalized_5 = feedin["Demand_el_normalized_1"]
    
    ######################################################
    #scenario 3 high demand consumer composition [5,0,0]
    if scenario == 3:
        demand_1 = feedin["Demand_el_0"]          #cluster 0
        demand_normalized_1 = feedin["Demand_el_normalized_0"]
        demand_2 = feedin["Demand_el_0"]          #cluster 0
        demand_normalized_2 = feedin["Demand_el_normalized_0"]
        demand_3 = feedin["Demand_el_0"]          #cluster 0
        demand_normalized_3 = feedin["Demand_el_normalized_0"]
        demand_4 = feedin["Demand_el_0"]          #cluster 0
        demand_normalized_4 = feedin["Demand_el_normalized_0"]
        demand_5 = feedin["Demand_el_0"]          #cluster 0
        demand_normalized_5 = feedin["Demand_el_normalized_0"]
    
        ######################################################
    #scenario 4 low demand with public lighting [0,4,1]
    if scenario == 4:
        demand_1=feedin["Demand_el_1"] 
        demand_normalized_1 = feedin["Demand_el_normalized_1"]
        demand_2 = feedin["Demand_el_1"]          #cluster 1
        demand_normalized_2 = feedin["Demand_el_normalized_1"]
        demand_3 = feedin["Demand_el_1"]          #cluster 1
        demand_normalized_3 = feedin["Demand_el_normalized_1"]
        demand_4 = feedin["Demand_el_1"]          #cluster 1
        demand_normalized_4 = feedin["Demand_el_normalized_1"]
        demand_5 = feedin["Demand_el_2"]          #cluster 2
        demand_normalized_5 = feedin["Demand_el_normalized_2"]
    
        ######################################################
    #scenario 5 high demand with public lighting [4,0,1]
    if scenario == 5:
        demand_1 = feedin["Demand_el_0"]          #cluster 0
        demand_normalized_1 = feedin["Demand_el_normalized_0"]
        demand_2 = feedin["Demand_el_0"]          #cluster 0
        demand_normalized_2 = feedin["Demand_el_normalized_0"]
        demand_3 = feedin["Demand_el_0"]          #cluster 0
        demand_normalized_3 = feedin["Demand_el_normalized_0"]
        demand_4 = feedin["Demand_el_0"]          #cluster 0
        demand_normalized_4 = feedin["Demand_el_normalized_0"]
        demand_5 = feedin["Demand_el_2"]          #cluster 2
        demand_normalized_5 = feedin["Demand_el_normalized_2"]
    
"""    
    ###############################################################################
#Input: Project's parameter 

project_param= {
                     'project_lifetime' : 10,     # Project duration 
                     'wacc'             : 0.10,   # Weighted average cost of capital (0<wacc<1) 
                     'currency'         : 'EUR'
               }


###############################################################################
#Input: component's parameter 

#/!\ order of components needs to be kept in the following dicts & lists 

Comp_param = {

       'pv'      :{   

                 'pv_peak'                   : 1,                          # Peak generation in kWp
                 'fix'                       : 101,                        # Fix investment costs in EUR (independent of the size_comp)
                 'fuel_price'                : 0,                          # EUR/l
                 'fuel_escalation_rate'      : 0,                          # fuel_escalation_rate in percent 
                 'var'                       : 0,                         # Utilisation costs in currency/kWh/a
                 'c_inv'                     : 540,    # Investment cost currency/kWp     SENSITIVITY      -> not per a ?
                 'c_inv_kWh'                 : 0,                          # Investment cost currency/kWh  ->  not per a ?
                 'o&m'                       : 14,     # OPEX in currency/kWp/a       SENSITIVITY
                 'life_time'                 : 10,                          # Component's life time  

                     }, 


       'battery':{
                  'technology'            :'LA',     # LI: Li-ion , LA: leadacid
                  'status'                :'on',
                  'LHV'                   : 0,       # Low Heat Value in kWh/l
                  'fix'                   : 26,       # Fix investment costs in EUR  (independent of the size_comp) 
                  'fuel_price'            : 0,       # EUR/l
                  'fuel_escalation_rate'  : 0,       # fuel_escalation_rate in percent 
                  'var'                   : 0,   # Utilisation costs in EUR/kWh 00
                  'c_inv'                 : 0,      # Investment cost EUR/kW    SENSITIVITY
                  'c_inv_kWh'             : 246.45,     # Investment cost EUR/kWh     SENSITIVITY
                  'o&m'                   : 14,    # OPEX in EUR/kW/a
                  'o&m_kWh'               : 0,     # OPEX in EUR/kWh/a        SENSITIVITY
                  'life_time'             : 3.5,    # Component's life time  
                     },   


    'controller':{  
                  'LHV'                   : 0,        # Low Heat Value in kWh/l
                  'fix'                   : 306,        # Fix investment costs in EUR  (independent of the size_comp)
                  'fuel_price'            : 0,        # EUR/l
                  'fuel_escalation_rate'  : 0,        # fuel_escalation_rate in percent 
                  'var'                   : 0,       # Utilisation costs in EUR/kWh 
                  'c_inv'                 : 0,      # Investment cost EUR/kW
                  'c_inv_kWh'             : 0,        # Investment cost EUR/kWh
                  'o&m'                   : 0.03*306,      # OPEX in EUR/kW/a
                  'life_time'             : 10,       # Component's life time 
                  'efficiency'            : 1, 
                      },

    'rice_huller':{
                  'status'                :'on',
                  'LHV'                   : 0,        # Low Heat Value in kWh/l
                  'fix'                   : 0,        # Fix investment costs in EUR  (independent of the size_comp)
                  'fuel_price'            : 0,        # EUR/l
                  'fuel_escalation_rate'  : 0,        # fuel_escalation_rate in percent 
                  'var'                   : 0,        # Utilisation costs in EUR/kWh 
                  'c_inv'                 : 607,      # Investment cost EUR/kW
                  'c_inv_kWh'             : 0,        # Investment cost EUR/kWh
                  'o&m'                   : 28,       # OPEX in EUR/kW/a  
                  'life_time'             : 5,       # Component's life time 
                  'efficiency'            : 70,        # kg/kWh (rice)
                       }, 

   
                    }   


###############################################################################
#Add components in component_List, wich should be analyzed

#Component 
component_List=['pv',
                'battery',
                'controller', 
                'rice_huller']

#Add component_List with sizing optimized based on component_List based on 
#

#(battery, dc_electricity): battery_out [kWh]
#(battery, nan)           : battery_cap [kWh]
#(dc_electricity, battery): battery_in  [kWh]


component_sizing_opt= ['pv',                    
                       'battery_out',
                       'battery',
                       'battery_in',
                       'controller_in',
                       'rice_huller_in'] 


# Elec flows only?

component_flows_list= ['pv',                    
                       'battery_out',
                       'battery',
                       'battery_in',
                       'controller_in',
                       'rice_huller_in']  

###############################################################################
#Cost_preprocessing

#output:
#replacement_number (project_lifetime,component_lifetime)
#residual_value (c_inv,project_lifetime,component_lifetime,Replacements,wacc)
#Capex(c_invest,c_residual,project_lifetime,component_lifetime,Replacements,wacc)
#CRF (discount_rate= WACC,project_lifetime)
#annuity(CAPEX,OPEX,CRF) 


res_eco ={}


lifetime_component         =[]
Investments_Number         =[]
Residual_value_per_kW      =[]
Residual_value_per_kWh     =[]
Replacements               =[] 
Investment_per_kW          =[] 
Investment_per_kWh         =[]
Fuel_price                 =[] 
Utilisation_cost_per_unit  =[] 
Capex_per_kW               =[] 
Capex_per_kWh              =[] 
Opex_per_kW                =[] 
Opex_per_kWh               =[] 
annuity_per_kW             =[]
annuity_per_kWh            =[]

Residual_value_fix         =[]
Capex_fix                  =[]
annuity_fix                =[]

CRF= CRF_fnc(project_param['wacc'], project_param['project_lifetime'])   ## WACC is a discount rate used for CRF calculation
f= lambda x : x-1 if x > 0 else 0 



for i in range(len(component_List)):



    lifetime_component.append(Comp_param[component_List[i]]['life_time'])

    Investments_Number.append(number_of_investement(project_param['project_lifetime'],
                                          Comp_param[component_List[i]]['life_time'],
                                                 ))

    Replacements.append(f(Investments_Number[i]))

    Investment_per_kW.append(Comp_param[component_List[i]]['c_inv'])

    Investment_per_kWh.append(Comp_param[component_List[i]]['c_inv_kWh'])

    Fuel_price.append(fuel_price_change(Comp_param[component_List[i]]['fuel_price'],Comp_param[component_List[i]]['fuel_escalation_rate'],project_param['project_lifetime'],project_param['wacc'],CRF)) 

    Utilisation_cost_per_unit.append(Comp_param[component_List[i]]['var'])


    Residual_value_per_kW.append(residual_value(Comp_param[component_List[i]]['c_inv'],
                                                 project_param['project_lifetime'],
                                                 Comp_param[component_List[i]]['life_time'],
                                                 Investments_Number[i],
                                                 project_param['wacc']))

    Residual_value_per_kWh.append(residual_value(Comp_param[component_List[i]]['c_inv_kWh'],
                                                 project_param['project_lifetime'],
                                                 Comp_param[component_List[i]]['life_time'],
                                                 Investments_Number[i],
                                                 project_param['wacc']))



    Capex_per_kW.append(Capex(Comp_param[component_List[i]]['c_inv'],
                                             Residual_value_per_kW[i],
                                             project_param['project_lifetime'],
                                             Comp_param[component_List[i]]['life_time'],
                                             Investments_Number[i],
                                             project_param['wacc']))



    Capex_per_kWh.append(Capex(Comp_param[component_List[i]]['c_inv_kWh'],
                                                 Residual_value_per_kWh[i],
                                                 project_param['project_lifetime'],
                                                 Comp_param[component_List[i]]['life_time'],
                                                 Investments_Number[i],
                                                 project_param['wacc']))

    Opex_per_kW.append(Comp_param[component_List[i]]['o&m']
                                                 )




    annuity_per_kW.append(annuity(Capex_per_kW[i],
                                   Comp_param[component_List[i]]['o&m'],
                                                 CRF))

    if (component_List[i] == 'battery'):

        annuity_per_kWh.append(annuity(Capex_per_kWh[i],
                                   Comp_param[component_List[i]]['o&m_kWh'],
                                                 CRF))
    else:
        annuity_per_kWh.append(0)

        
        
    Residual_value_fix.append(residual_value(Comp_param[component_List[i]]['fix'],
                                             project_param['project_lifetime'],
                                             Comp_param[component_List[i]]['life_time'],
                                             Investments_Number[i],
                                             project_param['wacc']))
    
    Capex_fix.append(Capex(Comp_param[component_List[i]]['fix'],
                           Residual_value_fix[i],
                           project_param['project_lifetime'],
                           Comp_param[component_List[i]]['life_time'],
                           Investments_Number[i],
                           project_param['wacc']))
    
    annuity_fix.append(annuity(Capex_fix[i],
                                   0,  #OPEX values already accounted for in annuity per kW
                                 CRF))
    
    
    res_eco['Components lifetime']=lifetime_component
    res_eco['Number of Investment']=Investments_Number
    res_eco['Replacements']=Replacements
    res_eco['Investment_'+project_param['currency']+'/kW']=Investment_per_kW
    res_eco['Investment_'+project_param['currency']+'/kWh']=Investment_per_kWh
    res_eco['Residual_value_'+project_param['currency']+'/kW']=Residual_value_per_kW
    res_eco['Residual_value_'+project_param['currency']+'/kWh']=Residual_value_per_kWh
    res_eco['Fuel_price_'+project_param['currency']+'/liter']= Fuel_price
    res_eco['Utilisation_cost_(Var)'+project_param['currency']+'/kWh']=Utilisation_cost_per_unit
    res_eco['Capex_per_'+project_param['currency']+'/kW']=Capex_per_kW
    res_eco['Capex_per_'+project_param['currency']+'/kWh']=Capex_per_kWh
    res_eco['Opex_(O&M)_'+project_param['currency']+'/kW']= Opex_per_kW
    res_eco['Annuity_'+project_param['currency']+'/kW']=annuity_per_kW
    res_eco['Annuity_'+project_param['currency']+'/kWh']=annuity_per_kWh
    
    res_eco['Residual_value_fix_'+project_param['currency']]=Residual_value_fix
    res_eco['Capex_fix_'+project_param['currency']]=Capex_fix
    res_eco['Annuity_fix_'+project_param['currency']]=annuity_fix


DF_Project_param= pd.DataFrame.from_dict(project_param,orient='index',columns=['Project_param'])


cost_param= pd.DataFrame.from_dict(res_eco,orient='index', columns=component_List)

###############################################################################
# input data calculated from the loaded data 



###############################################################################
# Initialize energy system

my_index = pd.date_range('1/1/2019', periods=8760, freq='H') #periods = 
ES_base_case = solph.EnergySystem(timeindex=my_index)


###############################################################################
###busses

dc_b1 = solph.Bus(label = 'dc_bus_1_label')
dc_b2 = solph.Bus(label = 'dc_bus_2_label')

    #rice huller busses
r_b1 = solph.Bus(label = 'rice_bus_1_label')
r_b2 = solph.Bus(label = 'rice_bus_2_label')
r_b3 = solph.Bus(label = 'rice_bus_3_label')


ES_base_case.add(dc_b1, dc_b2, r_b1, r_b2, r_b3) 


##############sources

#PV source
pv = solph.Source(label='pv', outputs={dc_b1: solph.Flow(
                              fix=feedin["PV"],
                              investment=solph.Investment(ep_costs=cost_param.loc['Annuity_'+project_param['currency']+'/kW','pv']/Comp_param['pv']['pv_peak'], 
                              existing = 0,
                              variable_costs=Comp_param['pv']['var']/Comp_param['pv']['pv_peak'],
                              ))})

# rice source

rice_source = solph.Source(label = 'rice_source', 
                           outputs={r_b1: solph.Flow(
                               fix = feedin['rice_source_normalized'],         # rice supply in kg/h
                               nominal_value = feedin['rice_source'].max()     
                           )}) 



ES_base_case.add( pv, rice_source)


###############transformers

#controller

controller = solph.Transformer(
    label="controller",
    inputs={dc_b1: solph.Flow(investment=solph.Investment(
        ep_costs=cost_param.loc['Annuity_'+project_param['currency']+'/kW','controller'],
        ))},
    outputs={dc_b2: solph.Flow()},
    conversion_factors={dc_b2: Comp_param['controller']['efficiency']}) 

#rice_huller

rice_huller = solph.Transformer(
    label="rice_huller",
    inputs={dc_b2:solph.Flow(investment=solph.Investment(
        ep_cost=cost_param.loc['Annuity_'+project_param['currency']+'/kW','rice_huller'])),
            r_b1:solph.Flow()},
    outputs={r_b2: solph.Flow(), r_b3: solph.Flow()},
    conversion_factors={r_b1:1, dc_b2: 1/Comp_param['rice_huller']['efficiency'], r_b2: 0.55, r_b3: 0.45})   #1/70 kWh/kg




ES_base_case.add( controller, rice_huller)        



##################storage
#battery

if Comp_param['battery']['status'] == 'off':


     battery_status=0

else:

    battery_status= float("+inf")


if  Comp_param['battery']['technology'] == 'LA':

    battery = solph.GenericStorage(
        label='battery',

        inputs={dc_b1: solph.Flow(investment=solph.Investment(
                                   ep_costs=cost_param.loc['Annuity_'+project_param['currency']+'/kW','battery']),
                                   variable_costs=Comp_param['battery']['var'],
                                   )},

        outputs={dc_b1: solph.Flow()},
        loss_rate=0, 
        inflow_conversion_factor=1,  
        outflow_conversion_factor=0.8,#total efficiency 80%
        initial_storage_level = 0.5,
        min_storage_level=0.3, #SOC-min
        max_storage_level=1,
        balanced = True, 
        nominal_storage_capacity = None,
        invest_relation_input_capacity  =  0.1 , # C-rate charge
        invest_relation_output_capacity =  0.1 , # C-rate discharge
        investment = solph.Investment(ep_costs=cost_param.loc['Annuity_'+project_param['currency']+'/kWh','battery'],
                                      maximum=battery_status) 
        ) 


ES_base_case.add(battery) 



#########sinks

#elec demand:
demand_el_1 = solph.Sink(label = "el_demand_1", inputs = {dc_b2: solph.Flow(fix=demand_normalized_1, nominal_value= demand_1.max()/1000)},)     # Check unit demand in W with rest of model

demand_el_2 = solph.Sink(label = "el_demand_2", inputs = {dc_b2: solph.Flow(fix=demand_normalized_2, nominal_value= demand_2.max()/1000)},)     # Check unit demand in W with rest of model

demand_el_3 = solph.Sink(label = "el_demand_3", inputs = {dc_b2: solph.Flow(fix=demand_normalized_3, nominal_value= demand_3.max()/1000)},)     # Check unit demand in W with rest of model

demand_el_4 = solph.Sink(label = "el_demand_4", inputs = {dc_b2: solph.Flow(fix=demand_normalized_4, nominal_value= demand_4.max()/1000)},)     # Check unit demand in W with rest of model

demand_el_5 = solph.Sink(label = "el_demand_5", inputs = {dc_b2: solph.Flow(fix=demand_normalized_5, nominal_value= demand_5.max()/1000)},)     # Check unit demand in W with rest of model



excess_el = solph.Sink(label="excess_sink", inputs={dc_b1: solph.Flow()})


    #rice demand with costs only. We can add a minimum or maximum production as well
rice_demand = solph.Sink(label = "rice_demand", inputs = {r_b2: solph.Flow()}) # if  1€/kg white rice revenue -> Flow(variable_costs = -1)
husks_sink = solph.Sink(label ="husks_sink", inputs ={r_b3: solph.Flow()}) # simple sink, no revenue for husks?




ES_base_case.add(demand_el_1, demand_el_2, demand_el_3, demand_el_4, demand_el_5, excess_el, rice_demand, husks_sink)






###############################################################################
#Optimization, constraints and Results 


logging.info("Optimise the energy system")
#set up a simple least cost optimisation
model = solph.Model(ES_base_case)


# solve the energy model using the GLPK solver
model.solve(solver='cbc', solve_kwargs={'tee': True})

logging.info("Store the energy system with the results.")


# get information on the meta data, e.g. objective function etc
ES_base_case.results["meta"] = solph.processing.meta_results(model)
# get main results from optimization
ES_base_case.results["main"] = solph.processing.results(model)
res=solph.processing.results(model)

####Now have function gives back the results as a python dictionary holding pandas Series for scalar values
# and pandas DataFrames for all nodes and flows between them. 

df_results= pd.DataFrame(data = ES_base_case.results["main"])

dc_b1 = solph.views.node(df_results, "dc_bus_1_label")
dc_b2 = solph.views.node(df_results, "dc_bus_2_label")

r_b1 = solph.views.node(df_results, "rice_bus_1_label")
r_b2 = solph.views.node(df_results, "rice_bus_2_label")
r_b3 = solph.views.node(df_results, "rice_bus_3_label")


print("********* Meta results *********")
# pp.pprint(ES_base_case.results["main"])
print("")

print("********* Main results *********")

print(dc_b1["sequences"].sum(axis=0))
print(dc_b2["sequences"].sum(axis=0))
print(r_b1["sequences"].sum(axis=0))
print(r_b2["sequences"].sum(axis=0))
print(r_b3["sequences"].sum(axis=0))

    
print("********* Main results *********")



###################Results_processing





###############################################################################
#plots for some cross-checking

#component's data as dataframe ,

data_pv = solph.views.node(df_results, "pv")
data_controller= solph.views.node(df_results, "controller")
data_battery= solph.views.node(df_results, "battery")
data_excess= solph.views.node(df_results, "excess_sink")
data_rice_huller = solph.views.node(df_results, "rice_huller")

print('data rice huller',data_rice_huller["sequences"].sum(axis=0))             

df_pv=pd.DataFrame(data_pv["sequences"].sum(axis=0))
df_controller=pd.DataFrame(data_controller["sequences"].sum(axis=0))      
df_battery=pd.DataFrame(data_battery["sequences"].sum(axis=0)) 
df_excess= pd.DataFrame(data_excess["sequences"].sum(axis=0))
df_rice_huller= pd.DataFrame(data_rice_huller["sequences"].sum(axis=0))

demand_el_sum = ( demand_normalized_1.sum()* demand_1.max()
                 +demand_normalized_2.sum()* demand_2.max()
                 +demand_normalized_3.sum()* demand_3.max()
                 +demand_normalized_4.sum()* demand_4.max()
                 +demand_normalized_5.sum()* demand_5.max())/1000    #coeffictient 1000: W to kW


Technology= ['pv'
             ,'controller'
             ,'battery'
             ,'excess_sink' 
             ,'rice_huller'
             ,'el_demand']

power_kWh  = [df_pv.iloc[0,0],
            df_controller.iloc[0,0],
            df_battery.iloc[2],
            df_excess.iloc[0,0],
            df_rice_huller.iloc[0,0],
            demand_el_sum]

"""
    fig, ax = plt.subplots(figsize=(10, 5))
    
    plt.xlabel('Technology')
    plt.ylabel('Power kWh')
    ypos=np.arange(len(Technology))
    plt.xticks(ypos,Technology)
    ax.bar(ypos,power_kWh)
    pathname = path + '\Power_per_tech_'+str(scenario)+'.jpeg'
    fig.savefig(pathname,dpi=400)

    if plt is not None:
        fig, ax = plt.subplots(figsize=(10, 5))
        data_battery["sequences"][
            "2019-01-01 01:00:00":"2019-01-01 23:00:00"
        ].plot(
            ax=ax, kind="line", drawstyle="steps-post"
        )
        plt.legend(
            loc="upper center",
            prop={"size": 8},
            bbox_to_anchor=(0.5, 1.25),
            ncol=2,
        )
        fig.subplots_adjust(top=0.8)
        plt.show()
        pathname = path + '\Battery_'+str(scenario)+'.jpeg'
        fig.savefig(pathname)

        fig, ax = plt.subplots(figsize=(10, 5))
        dc_b1["sequences"][
            "2019-01-01 01:00:00":"2019-01-01 23:00:00"
        ].plot(
            ax=ax, kind="line", drawstyle="steps-post"
        )
        plt.legend(
            loc="upper center", prop={"size": 8}, bbox_to_anchor=(0.5, 1.3), ncol=2
        )
        fig.subplots_adjust(top=0.8)
        plt.show()
        pathname=path + '\Elec_generation_'+str(scenario)+'.jpeg'
        fig.savefig(pathname)

        
        fig, ax = plt.subplots(figsize=(10, 5))
        dc_b2["sequences"][
            "2019-01-01 01:00:00":"2019-01-01 23:00:00"
        ].plot(
            ax=ax, kind="line", drawstyle="steps-post"
        )
        plt.legend(
            loc="upper center", prop={"size": 8}, bbox_to_anchor=(0.5, 1.3), ncol=2
        )
        fig.subplots_adjust(top=0.8)
        plt.show()

        fig.subplots_adjust(top=0.8)
        plt.show()
        pathname = path + '\Bus2_elec_'+str(scenario)+'.jpeg'
        fig.savefig(pathname)
"""
    ############# Sizing, cost processing


## Cost_Preproccessing

print("****************** Project_parameter ********************")
print()
print(DF_Project_param)
print()
print("****************** Cost_Preproccessing ******************")
print()
print(cost_param.round(2))
print()



###############################################################################
# optimized sizing and flow of each component 

#pv  
pv_opt =pv_optimized ('pv',df_results,model) 

#battery 
dc_bus_1_label = 'dc_bus_1_label'
battery_opt =battery_optimized ('battery',df_results,model,dc_bus_1_label)

#controller
dc_bus_2_label = 'dc_bus_2_label'
controller_opt=controller_optimized ('controller',df_results,model,dc_bus_1_label,dc_bus_2_label) 

#rice_huller
rice_bus_1_label= 'rice_bus_1_label'
rice_bus_2_label= 'rice_bus_2_label'

rice_huller_opt = rice_huller_optimized ( 'rice_huller',df_results,model,dc_bus_2_label, rice_bus_1_label, rice_bus_2_label)

#excess 
excess_generation= dc_b1['sequences'][(('dc_bus_1_label', "excess_sink"),'flow')].sum()

#total_generation
total_generation = dc_b1['sequences'][(("pv",'dc_bus_1_label' ),'flow')].sum()

#rice generation
rice_generation = r_b2['sequences'][(('rice_bus_2_label', "rice_demand"),'flow')].sum()
rice_husks_generation = r_b3['sequences'][(('rice_bus_3_label', "husks_sink"),'flow')].sum()



###############################################################################
#Sizing after solving the optimization, DataFrame of all the components


sizing=sizing_fnc(df_results,model, component_List ,['Capacity_optimized'])    # taking into account the investment costs

sizing['component']= component_sizing_opt     # Where is it defined?

list_cap_kWh=['battery']

sizing.set_index('component',inplace=True)


# create DateFrame to consider Capacity_optimized_kWh and Capacity_optimized_kW seperatly
sizing_kW_kWh= pd.DataFrame()


sizing_kW_kWh['component']= component_sizing_opt  

sizing_kW_kWh.set_index('component',inplace=True)

sizing_kW_kWh['Capacity_optimized_kWh']=[sizing.Capacity_optimized[i] if i in list_cap_kWh  else 0 for i in sizing.index]

sizing_kW_kWh['Capacity_optimized_kW']=[sizing.Capacity_optimized[i] if i not in list_cap_kWh  else 0 for i in sizing.index]

storage_kW= ['battery_in']  

sizing_kW_kWh.loc['battery','Capacity_optimized_kW']= sizing_kW_kWh.loc['battery_in','Capacity_optimized_kW']


sizing_kW_kWh.drop(['battery_out','battery_in'],inplace=True)




print('Sizing\n', sizing_kW_kWh)

####    sizing_kW_kWh contains the optimized component power and capacities of the optimized (invest_mode) components 

#get the Energysflows of each component after solving the optimization, DataFrame of all the components

#e_flow_controller = get_energy_controller(controller_label, df_results, model, dc_bus_2_label)


e_flows_main=[pv_opt.loc['Energyflow_optimized_kWh','pv'],
         battery_opt.loc['Energyflow_optimized_in_kWh','battery'],
         controller_opt.loc['Energyflow_optimized_in_kWh','controller'],
         rice_huller_opt.loc['Energyflow_optimized_in_kWh','rice_huller']]

sizing_kW_kWh['Energyflow_optimized_kWh']= e_flows_main

#In this case no fuel consumption considered
#sizing_kW_kWh['Fuel_consumption']= [fuel_consumption_diesel.loc['Fuel_consumption_liter','diesel'] if x == 'diesel' else 0 
#                                 for x in sizing_kW_kWh.index ]



#Sizing + energy flow after solving the optimization 
System_opt = sizing

#Energysflows after solving the optimization, DataFrame of all the components

e_flows_all=[pv_opt.loc['Energyflow_optimized_kWh','pv'],
         battery_opt.loc['Energyflow_optimized_out_kWh','battery'],
         battery_opt.loc['Energyflow_optimized_content_kWh','battery'],
         battery_opt.loc['Energyflow_optimized_in_kWh','battery'],
         controller_opt.loc['Energyflow_optimized_in_kWh','controller'],
         rice_huller_opt.loc['Energyflow_optimized_in_kWh','rice_huller']]

e_flows_all_series = pd.Series(e_flows_all,index=component_sizing_opt)



#Energysflows after solving the optimization, DataFrame of all the components
System_opt['Energyflow_optimized']= e_flows_all_series


#rice consumption after optimization, rice, unless 0
#Fuel consumption list
# Fuel_consumption_list = [0,0,0,0,0,huller_opt.loc['Riceflow_optimized_in_kg','huller']]
# Fuel_consumption_series = pd.Series(Fuel_consumption_list, index = component_sizing_opt)
# System_opt['Fuel_consumption'] = Fuel_consumption_series



print()
print("********* Optimized sizing and energy flow of each component *************")
print()
print(System_opt.round(2))
print()
print('*storage_out ist the optimized power in kW ') 
print('*storage optimized capacity is in kWh') 
print('the Fuel_consumption is given in liter') 




if  operation_behaviour_analysis== 'on':

    Result_concat= results_postprocessing(df_results,component_List)           
    Result_concat['demand_el_1']=dc_b2['sequences'][(('dc_bus_2_label', "el_demand_1"),'flow')]         
    Result_concat['demand_el_2']=dc_b2['sequences'][(('dc_bus_2_label', "el_demand_2"),'flow')]
    Result_concat['demand_el_3']=dc_b2['sequences'][(('dc_bus_2_label', "el_demand_3"),'flow')]
    Result_concat['demand_el_4']=dc_b2['sequences'][(('dc_bus_2_label', "el_demand_4"),'flow')]
    Result_concat['demand_el_5']=dc_b2['sequences'][(('dc_bus_2_label', "el_demand_5"),'flow')]
    
    Result_concat['rice_huller']=dc_b2['sequences'][(('dc_bus_2_label', "rice_huller"),'flow')]
    
    #demand_el_1 = solph.Sink(label = "el_demand_1", inputs = {dc_b2: solph.Flow(fix=demand_normalized_1, nominal_value= demand_1.max()/1000)},)     # Check unit demand in W with rest of model
   
    
    
    Result_concat['SOC_battery']= data_battery['sequences'][(('battery','nan'),'storage_content')].values / battery_opt.loc['Capacity_optimized_kWh','battery']
    Result_concat['excess']=dc_b1['sequences'][(('dc_bus_1_label', "excess_sink"),'flow')]

    
    writer = pd.ExcelWriter(path_out+file_name_operation_behaviour, engine='xlsxwriter')

   
    Result_concat.to_excel(writer, sheet_name= sheet_name_OPB)
    
 
    
    writer.save() 
    


###############################################################################
#economic Results of each component
'''

Parameter
---------
Annuity_per_unit_list
Capex_per_unit_list
Utilisation_cost_list
fuel_price_list
Fuel_comsumption_expend
Capacity_opt_list
Energy_opt_list
component_List
currency

Results
--------- 

Annuity_Optimised_Sizing    
Capex_Optimised_Sizing      
Utili_Cost_Optimised_Energy 
Fuel_expenditures           
Total_Annuity  

'''




Annuity_per_kW_list= pd.Series.tolist(cost_param.loc['Annuity_'+project_param['currency']+'/kW'])

Annuity_per_kWh_list=pd.Series.tolist(cost_param.loc['Annuity_'+project_param['currency']+'/kWh'])

Capex_per_kW_list = pd.Series.tolist(cost_param.loc['Capex_per_'+project_param['currency']+'/kW'])

Capex_per_kWh_list= pd.Series.tolist(cost_param.loc['Capex_per_'+project_param['currency']+'/kWh'])

Utilisation_cost_list= pd.Series.tolist(cost_param.loc['Utilisation_cost_(Var)'+project_param['currency']+'/kWh'])


Annuity_fix_list = pd.Series.tolist(cost_param.loc['Annuity_fix_'+project_param['currency']])

Capex_fix_list = pd.Series.tolist(cost_param.loc['Capex_fix_'+project_param['currency']])


# Fuel_comsumption_list= pd.Series.tolist(System_opt['Fuel_consumption'].drop(['battery_out','battery_in']))

#In this case no fuel is considered
#        Fuel_comsumption_list= pd.Series.tolist(sizing_kW_kWh['Fuel_consumption'])

fuel_price_list= pd.Series.tolist(cost_param.loc['Fuel_price_'+project_param['currency']+'/liter'])

#Capacity_opt_kW_list= [sizing.loc[component_List[i],'Capacity_optimized'] if component_List[i] != 'battery' else battery_opt.loc['Power_optimized_kW','battery'] for i in range(len(component_List))]


Capacity_opt_kW_list= sizing_kW_kWh['Capacity_optimized_kW']

#Capacity_opt_kW_storage=[battery_opt.loc['Power_optimized_kW','battery'], h2storage_opt.loc['Power_optimized_kW','h2storage']]

Capacity_opt_kWh_list= sizing_kW_kWh['Capacity_optimized_kWh']



#Capacity_opt_kWh_list=[0 if component_List[i] != 'battery' else battery_opt.loc['Capacity_optimized_kWh','battery'] for i in range(len(component_List))]

#Capacity_opt_kWh_list=[battery_opt.loc['Capacity_optimized_kWh','battery'],h2storage_opt.loc['Capacity_optimized_kWh','h2storage']]


#Energy_opt_list= pd.Series.tolist(System_opt['Energyflow_optimized'].drop(['battery_out','battery','h2storage','h2storage_out']))
Energy_opt_list= pd.Series.tolist(sizing_kW_kWh['Energyflow_optimized_kWh'])
comma=3


Cost_system_opt=cost_component_optimized(Annuity_per_kW_list,
                Annuity_per_kWh_list,
                Annuity_fix_list,
                Capex_fix_list,                         
                Capex_per_kW_list,
                Capex_per_kWh_list,
                Utilisation_cost_list,
                fuel_price_list,

                Capacity_opt_kW_list,
                Capacity_opt_kWh_list,
                Energy_opt_list,
                component_List,
                project_param['currency']
                )





print()
print("********* Main economic Parameter of each component *************")
print()
print(Cost_system_opt.round(2))

###############################################################################


    
    #######################################    Evaluation of PUEs in the Community
    

###### Technical Criterias:

total_energy_production = pv_opt.loc['Energyflow_optimized_kWh','pv']


EXCESS_GENERATION_SHARE= excess_generation/total_energy_production*100

#Hours excess per year
EXCESS_HOURS = dc_b1['sequences'][(('dc_bus_1_label', "excess_sink"),'flow')].loc[dc_b1['sequences'][(('dc_bus_1_label', "excess_sink"),'flow')]>0].count()


#Frequency of excess during the day
ds = dc_b1['sequences'][(('dc_bus_1_label', "excess_sink"),'flow')]
events = pd.DataFrame({'timestamp':ds.index, "Flow":ds.values})
events['hour'] = events['timestamp'].dt.hour
hourly_events = events.loc[events['Flow']>0].groupby('hour').size().copy()#group event by hour & exclude 0 values (no excess el)
hourly_events = hourly_events.reindex(events.groupby('hour').size().index, fill_value=0)
distribution = hourly_events/hourly_events.sum() *100

"""
    fig, ax= plt.subplots(1,1, figsize=(18,10))
    
    ax.set_xticks(np.arange(0,23))
    distribution.plot.bar(ax=ax, legend=False,rot=0)
    ax.set_xlabel('hour')
    ax.set_ylabel('Frequency [%]')
    ax.set_title('Distribution of excess electricity hours')
    pathname = path + '\Excess_elec_'+str(scenario)+'.jpeg'
    fig.savefig(pathname)
    
"""    
    #######  Economic criterias:

#total annualized cost
total_cost_system= Cost_system_opt.loc['TOTAL_ANNUITY_'+project_param['currency']].sum() #total annualized cost
#

#total energy supply
dc_b2_out = 0
for flow in dc_b2["sequences"]:
        if 'dc_bus_2_label' == flow[0][0]:
            dc_b2_out +=dc_b2["sequences"][flow].sum(axis = 0)
            
total_energy_supply= dc_b2_out # Takes into account residential el_demand + PUE if existing, this is only for calculating LCOS as the definition of total system costs/total energy supply
#total_energy_supply= (dieselGen_opt.loc['Energyflow_optimized_kWh','diesel']+inverter_opt.loc['Energyflow_optimized_out_kWh','inverter'])
#

#LCOE:
LCOS = total_cost_system/total_energy_supply

res_flow = 0
service_flow = 0
for flow in dc_b2["sequences"]:
        if ('dc_bus_2_label' == flow[0][0]):
            if ('el_demand_' == flow[0][1][:-1]):
                res_flow += dc_b2["sequences"][flow].sum(axis = 0)
            else:
                service_flow += dc_b2["sequences"][flow].sum(axis = 0)

if service_flow + res_flow != total_energy_supply:
    raise NameError("Res & Service flow don't match supply")
    
ratio_residential = res_flow/total_energy_supply
ratio_service = service_flow/total_energy_supply
cost_PUE = Cost_system_opt.loc['ANNUITY_EUR','rice_huller'] # annualized cost of service
Residential_LCOE = (total_cost_system-cost_PUE)*ratio_residential/res_flow
Service_LCOE = (cost_PUE + (total_cost_system-cost_PUE)*ratio_service)/service_flow

print('cost_PUE',cost_PUE,'Residential_LCOE',Residential_LCOE,'Service_LCOE', Service_LCOE, 'LCOS',LCOS)
#



#CO2_Emission     =  Diesel/Gasoline_opt_flow unit?  # no emission in this system

currency =project_param['currency']
#FUEL_CONSUMPTION= sizing.loc['diesel','Fuel_consumption'] # no emission in this system
Residential_DEMAND= res_flow 
Service_DEMAND = service_flow







System_opt_result= system_opt_result(total_cost_system,
                   total_energy_supply,
                   total_energy_production,
                   EXCESS_GENERATION_SHARE,
                   EXCESS_HOURS,
                   LCOS,
                   Residential_LCOE,
                   Service_LCOE,
                   Residential_DEMAND,
                   Service_DEMAND,
                   currency
                   )    
#FUEL_CONSUMPTION
#CO2_Emission,

print() 
print("*********** Main resulted Parameter of the system *****************")
print()
print(System_opt_result.apply(lambda x: '%.2f' % x, axis=1)) 

result_dict['System_opt_result'] = System_opt_result
result_dict['Capacity_optimized'] = System_opt['Capacity_optimized']
result_dict['Energyflow_optimized'] = System_opt['Energyflow_optimized']


###############################################################################################
###########
####### SAVE IN EXCEL
df_list = [DF_Project_param, cost_param.round(2), System_opt.round(2),Cost_system_opt.round(2), System_opt_result, distribution]
def multiple_dfs(df_list, sheets, file_name, spaces):
    writer = pd.ExcelWriter(File_name, engine="openpyxl",
                        mode='a', if_sheet_exists = 'overlay')# as writer:  
    row = 0
    for dataframe in df_list :
        dataframe.to_excel(writer,sheet_name=sheets,startrow=row , startcol=1)   
        row = row + len(dataframe.index) + 1

    print('Saved in Excel file')
    writer.save()


multiple_dfs(df_list,'Scenario_', File_name, 2)
#################################################################################################


"""

    return result_dict['System_opt_result'],

sa = SensitivityAnalyzer(sensitivity_dict, sensitivity_model)

sa.df




##############################################################################
####Sensitivity Analysis Display Function

# Only for 1 sensi-param analysis
# sensi_var should be same as in sensitivity dict
# output_var should be same as in System_opt_result dataframe

def sensitivity_display(sensi_var, ref_value, output_var):
    x_list = []
    y_list = []
    for value in sensitivity_dict[sensi_var]:
        x_list += [(value-ref_value)/ref_value*100]
        y_list += [sa.df.Result.loc[list(sa.df[sa.df[sensi_var]==ref_value].index)[0]].loc[output_var,'System_parameter']]
    return (x_list, y_list)

    
'''
System_opt_result.to_excel('syst_main_res.xlsx', sheet_name= 'sizing',
                           startrow= 1,
                           startcol= list_1[j],
                           index=[list_2[j]])
'''

#Excel export        
###############################################################################
# File_name = 'sensitivity_results.xlsx'
# def save_sensi():
#     return 'Saved'

# File_name = 'Nom_Workbook.xlsx'
# with pd.ExcelWriter(File_name, engine="openpyxl",
#                     mode='a', if_sheet_exists = 'replace') as writer:  
#     Load_data.to_excel(writer, sheet_name=sheet)
# print('Saved in Excel file')
"""