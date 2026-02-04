#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Aug 11 16:32:03 2021

@author: nikolas

ENERGICA T1.2.2 (energy system model)

system result function for case_1_dc_rice_huller



"""
from oemof import solph
import pandas as pd 
from oemof.solph import (Sink,
                         Transformer,
                         Source, Bus,
                         Flow,
                         NonConvex,
                         Model, 
                         EnergySystem, 
                         components, 
                         custom, 
                         Investment, 
                         views)




def sizing_fnc(results,m,sizing_list,columnname) :
    
    res={}
    
    for i in range(len(sizing_list)): #in component_list
        node=m.es.groups[sizing_list[i]]
        node_data=solph.views.node(results,node)
        for nodes, flow_name in node_data['sequences']:
            if 'scalars' in node_data and node_data['scalars'].get((nodes, 'invest')) is not None :
                res[nodes] = node_data['scalars'].get((nodes, 'invest'))
                
        DF= pd.DataFrame.from_dict(res,orient='index', columns=columnname)
        
        
    return DF



def results_postprocessing(model, component_list):
    generator_list = []

    for i in range( len( component_list ) ):
        d1 = views.node( model, component_list[i] )['sequences']
        generator_list.append( d1 )

    res = pd.concat( generator_list, axis=1 )
    
    '''
    Result_concat.rename(columns={'(('pv', 'dc_electricity'), 'flow')': 'pv-dc_electricity',
                              '(('dieselGen', 'ac_electricity'), 'flow')': 'pv-dc_electricity'
                     }
                     )

    '''

    return res



def excel_export (xls_file,dataframe,sheetname):
    
    writer = pd.ExcelWriter(xls_file, engine='xlsxwriter')
    dataframe.to_excel(writer, sheet_name= sheetname) 
    writer.save() 
    
    

def pv_optimized (pv_label,results,model) :
    
    #capacity
    
    res1={}
    node=model.es.groups[pv_label]
    node_data= views.node(results,node)
    for nodes, flow_name in node_data['sequences']:
        res1['Capacity_optimized_kWp']= node_data['scalars'].get((nodes,'invest'))
      
    #energyflow
           
    res1['Energyflow_optimized_kWh']= node_data['sequences'][(nodes,'flow')].sum()
    
                     
    DF= pd.DataFrame.from_dict(res1,orient='index', columns=[pv_label])
    
    
    return DF



def battery_optimized (battery_label,results,model,dc_bus_1_label) :
    
    res1={}
    
    #Capacity
    
    battery = solph.views.node(results, battery_label)
    
    battery_capacity = battery['scalars'][((battery_label,'nan'),'invest')]
    
    dc_b1 = solph.views.node(results,dc_bus_1_label)
    
    battery_power = dc_b1['scalars'][((battery_label,dc_bus_1_label),'invest')]
    
    res1['Capacity_optimized_kWh']= battery_capacity
    res1['Power_optimized_kW']= battery_power
    
    
    #Energyflow
    
    battery_flow_out= battery['sequences'][((battery_label,dc_bus_1_label),'flow')].sum()
    battery_flow_content= battery['sequences'][((battery_label,'nan'),'storage_content')].sum()
    battery_flow_in= battery['sequences'][((dc_bus_1_label,battery_label),'flow')].sum()
    
    res1['Energyflow_optimized_out_kWh']    = battery_flow_out
    res1['Energyflow_optimized_content_kWh']= battery_flow_content
    res1['Energyflow_optimized_in_kWh']= battery_flow_in
    
   
   
    DF= pd.DataFrame.from_dict(res1,orient='index', columns=[battery_label])
        
    return DF



def controller_optimized (controller_label,results,model,dc_bus_1_label,dc_bus_2_label) :
    
    res1={}
    
    #Capacity
    
    dc_b1= solph.views.node(results, dc_bus_1_label)
    Capacity_optimized_kW= dc_b1['scalars'][((dc_bus_1_label,controller_label),'invest')]
    
    
    #Energyflow
    dc_b2= solph.views.node(results, dc_bus_2_label)
    
    Energyflow_in_kW =  dc_b1['sequences'][((dc_bus_1_label, controller_label),'flow')].sum()
    
    Energyflow_out_kW=  dc_b2['sequences'][((controller_label,dc_bus_2_label),'flow')].sum()
    
    res1['Capacity_optimized_kW']         =  Capacity_optimized_kW
    res1['Energyflow_optimized_in_kWh']   =  Energyflow_in_kW
    res1['Energyflow_optimized_out_kWh']  =  Energyflow_out_kW
    
   
    DF= pd.DataFrame.from_dict(res1,orient='index', columns=[controller_label]) 
    
    
    return DF


# def rice_huller_optimized(rice_huller_label, results, model, dc_bus_2_label, rice_bus_1_label, rice_bus_2_label) :
    
#     res1={}
    
#     dc_b2= solph.views.node(results, dc_bus_2_label)
#     r_b2= solph.views.node(results, rice_bus_2_label)
#     r_b1= solph.views.node(results, rice_bus_1_label)
    
#     #Capacity
    
#     Capacity_optimized_kW= dc_b2['scalars'][((dc_bus_2_label,rice_huller_label),'invest')]
    
#     #Energyflow
    
#     Energyflow_in_kW =  dc_b2['sequences'][((dc_bus_2_label,rice_huller_label),'flow')].sum()
    
#     Riceflow_out_kg=  r_b2['sequences'][((rice_huller_label,rice_bus_2_label),'flow')].sum()
    
    
#     Riceflow_in_kg=  r_b1['sequences'][((rice_bus_1_label, rice_huller_label),'flow')].sum()
    
#     res1['Capacity_optimized_kW']         =  Capacity_optimized_kW
#     res1['Energyflow_optimized_in_kWh']   =  Energyflow_in_kW
#     res1['Riceflow_optimized_out_kg']  =  Riceflow_out_kg
#     res1['Riceflow_optimized_in_kg']  =  Riceflow_in_kg
    
   
    # DF= pd.DataFrame.from_dict(res1,orient='index', columns=[rice_huller_label]) 
    
    
#     return DF
                                        
    
def diesel_huller_optimized(rice_huller_label, results, model, fuel_bus_1_label, rice_bus_1_label, rice_bus_2_label) :
    
    res1={}
    
    f_b1= solph.views.node(results, fuel_bus_1_label)
    r_b2= solph.views.node(results, rice_bus_2_label)
    r_b1= solph.views.node(results, rice_bus_1_label)
    
    #Capacity
    
    Capacity_optimized_kW= r_b1['scalars'][((rice_bus_1_label,rice_huller_label),'invest')]
    
    #Energyflow
    
    Energyflow_in_kW =  f_b1['sequences'][((fuel_bus_1_label,rice_huller_label),'flow')].sum()
    
    Riceflow_out_kg=  r_b2['sequences'][((rice_huller_label,rice_bus_2_label),'flow')].sum()
    Riceflow_in_kg=  r_b1['sequences'][((rice_bus_1_label, rice_huller_label),'flow')].sum()
    
    res1['Capacity_optimized_kW']         =  Capacity_optimized_kW     #Capacity actually in Production rate: kg_paddy/h
    res1['Energyflow_optimized_in_kWh']   =  Energyflow_in_kW          #Riceflow_input in kg_paddy
    
    res1['Riceflow_optimized_out_kg']  =  Riceflow_out_kg
    res1['Riceflow_optimized_in_kg']  =  Riceflow_in_kg
    
   
    DF= pd.DataFrame.from_dict(res1,orient='index', columns=[rice_huller_label]) 
    
    
    return DF
                  
                           
                           
                           
                           
                           

def dc_water_pump_optimized (dc_water_pump_label,results,model,dc_bus_2_label,water_bus_1_label, water_bus_2_label) :
    
    res1={}
    
    #Capacity
    
    dc_b2= solph.views.node(results, dc_bus_2_label)
    Capacity_optimized_kW= dc_b2['scalars'][((dc_bus_2_label,dc_water_pump_label),'invest')]
    
    
    #waterflow
    w_b2= solph.views.node(results, water_bus_2_label)
    
    
    Energyflow_in_kW=  dc_b2['sequences'][((dc_water_pump_label,dc_bus_2_label),'flow')].sum()
    
    Waterflow_out_l =  w_b2['sequences'][((water_bus_2_label, dc_water_pump_label),'flow')].sum()
    

    
    res1['Capacity_optimized_kW']         =  Capacity_optimized_kW
    res1['Energyflow_optimized_in_kWh']   =  Energyflow_in_kW
    res1['Waterflow_optimized_out_l']  =  Waterflow_out_l
    
   
    DF= pd.DataFrame.from_dict(res1,orient='index', columns=[dc_water_pump_label]) 
    
    
    return DF

def dc_water_treatment_optimized (dc_water_treatment_label,results,model,dc_bus_2_label,water_bus_2_label, water_bus_3_label) :
    
    res1={}
    
    #Capacity
    
    dc_b2= solph.views.node(results, dc_bus_2_label)
    Capacity_optimized_kW= dc_b2['scalars'][((dc_bus_2_label,dc_water_treatment_label),'invest')]
    
    
    #waterflow
    w_b3= solph.views.node(results, water_bus_3_label)
    
    
    Energyflow_in_kW=  dc_b2['sequences'][((dc_water_treatment_label,dc_bus_2_label),'flow')].sum()
    
    Waterflow_out_l =  w_b3['sequences'][((water_bus_3_label, dc_water_treatment_label),'flow')].sum()
    

    
    res1['Capacity_optimized_kW']         =  Capacity_optimized_kW
    res1['Energyflow_optimized_in_kWh']   =  Energyflow_in_kW
    res1['Waterflow_optimized_out_l']  =  Waterflow_out_l
    
   
    DF= pd.DataFrame.from_dict(res1,orient='index', columns=[dc_water_treatment_label]) 
    
    
    return DF


def get_energy_controller(controller_label, results, model, dc_bus_2_label):
    
    res1={}
    
    dc_b2 = solph.views.node(results,dc_bus_2_label)  
 
    
    #Energyflow
    
    controller_flow= dc_b2['sequences'][((controller_label,dc_bus_2_label),'flow')].sum()
    res1['Energyflow_optimized_out_kWh']    = controller_flow
   
    DF= pd.DataFrame.from_dict(res1,orient='index', columns=[controller_label])
        
    return DF



def excess_generation (dc_bus_1_label,excess_sink_label,results):
    
    dc_b1 = solph.views.node(results,excess_sink_label)
    excess_generation= dc_b1['sequences'][((dc_bus_1_label,excess_sink_label),'flow')].sum()
    
    
    return excess_generation





def legend_fnc(currency):
    
    res1={}
    
    res1['Residual_value_per_kW']           =  [currency+'/kW', 'residual value of the component']
    res1['Investement_per_kW']               =  [currency+'/kW', 'in case of pv '+ currency+ '/kWp else '+ currency+'/kW' ]
    res1['Investement_per_kWh']              =  [currency+'/kWh', 'usually used for battery']
    res1['Residual_value_per_kWh ']          =  [currency+'/kWh','Capex of the component']
    res1['Utilisation_cost_per_kWh (Var)']   =  [currency+'/kWh','Utilisations costs (variable)']
    res1['Capex_per_kW']                     =  [currency+'/kW','capex= invest+ replacem - residual']
    res1['Capex_per_kWh']                    =  [currency+'/kWh','capex= invest+ replacem - residual']
    res1['Opex_per_kW (O&M)']                =  [currency+'/kW','Operation & Maintenance Cost']
    res1['Annuity_per_kW']                   =  [currency+'/kW','Annuity= capex_kW* CRF + opex_kW']
    res1['Annuity_per_kWh']                  =  [currency+'/kWh','Annuity= capex* CRF + opex']
    res1['LCOE']                             =  [currency+'/kWh','Levelized Cost of Electricity']
    res1['ANNUITY_'+currency]                =  [currency,'=annuity x capacity optimized']
    res1['CAPEX_'+currency]                  =  [currency]
    res1['UTI_COSTS_'+currency]              =  [currency,'Utilisations costs']
    res1['TOTAL_ANNUITY_'+currency]          =  [currency,'= ANNUITY+ UTI_COSTS+ FUEL_EXP']
    res1['battery']                          =  ['kWh','optimizd capacity of the battry']
    res1['battery_in']                       =  ['kW','optimized input of the battery']
    res1['battery_out']                      =  ['kW','optimized power of the battery']
    res1['raw_water_storage']                          =  ['l','optimizd capacity of the raw_water_storage']
    res1['raw_water_storage_in']                       =  ['l','optimized input of the raw_water_storage']
    res1['raw_water_storage_out']                      =  ['l','optimized power of the raw_water_storage']
    res1['clean_water_storage']                          =  ['l','optimizd capacity of the clean_water_storage']
    res1['clean_water_storage_in']                       =  ['l','optimized input of the clean_water_storage']
    res1['clean_water_storage_in']                      =  ['l','optimized power of the clean_water_storage']
    
    
    legend= pd.DataFrame.from_dict(res1,orient='index', columns=['UNIT','MEANING']) 
    
    return legend.sort_index()













"""
Not used here, but for other cases:

def inverter_optimized (inverter_label,results,model,dc_bus_label,ac_bus_label) :
    
    res1={}
    
    #Capacity
    
    bel_dc= solph.views.node(results, dc_bus_label)
    Capacity_optimized_kW= bel_dc['scalars'][((dc_bus_label,inverter_label),'invest')]
    
    
    #Energyflow
    bel_ac= solph.views.node(results, ac_bus_label)
    
    Energyflow_in_kW =  bel_dc['sequences'][((dc_bus_label,inverter_label),'flow')].sum()
    
    Energyflow_out_kW=  bel_ac['sequences'][((inverter_label,ac_bus_label),'flow')].sum()
    
    res1['Capacity_optimized_kW']         =  Capacity_optimized_kW
    res1['Energyflow_optimized_in_kWh']   =  Energyflow_in_kW
    res1['Energyflow_optimized_out_kWh']  =  Energyflow_out_kW
    
   
    DF= pd.DataFrame.from_dict(res1,orient='index', columns=[inverter_label]) 
    
    
    return DF



def fuel_consumption(fuel_consumption_name,results,fuel_bus_label,fuel_source_label,combustion_value_fuel,fuel_price,curreny):
    
    res1={}
    
    fuel_bus= solph.views.node(results,fuel_bus_label)
    Fuel_consumption_kWh   = fuel_bus['sequences'][((fuel_source_label,fuel_bus_label),'flow')].sum()
    Fuel_consumption_Liter = Fuel_consumption_kWh/combustion_value_fuel
    Fuel_consumption_expen = Fuel_consumption_Liter * fuel_price
    
    res1['Fuel_consumption_kWh']       =  Fuel_consumption_kWh
    res1['Fuel_consumption_liter']     =  Fuel_consumption_Liter
    res1['Fuel_expenditures_'+curreny] =  Fuel_consumption_expen
    
    DF= pd.DataFrame.from_dict(res1,orient='index', columns=[fuel_consumption_name]) 

     
    return DF

def inverter_optimized (inverter_label,results,model,dc_bus_label,ac_bus_label) :
    
    res1={}
    
    #Capacity
    
    bel_dc= solph.views.node(results, dc_bus_label)
    Capacity_optimized_kW= bel_dc['scalars'][((dc_bus_label,inverter_label),'invest')]
    
    
    #Energyflow
    bel_ac= solph.views.node(results, ac_bus_label)
    
    Energyflow_in_kW =  bel_dc['sequences'][((dc_bus_label,inverter_label),'flow')].sum()
    
    Energyflow_out_kW=  bel_ac['sequences'][((inverter_label,ac_bus_label),'flow')].sum()
    
    res1['Capacity_optimized_kW']         =  Capacity_optimized_kW
    res1['Energyflow_optimized_in_kWh']   =  Energyflow_in_kW
    res1['Energyflow_optimized_out_kWh']  =  Energyflow_out_kW
    
   
    DF= pd.DataFrame.from_dict(res1,orient='index', columns=[inverter_label]) 
    
    
    return DF



"""










'''

legend_df= legend(currency)

class legend_class :  
    def __init__(self, unit, meaning):
        self.unit     = unit  
        self.meaning  = meaning
        
        
battery_in= legend_class(legend_df.loc['battery_out','UNIT'],legend_df.loc['battery_out','MEANING'])


for i in legend_df.index:
    legend_df.index[i] = legend_class(legend_df.loc[legend_df.index[i],'UNIT']
                                      ,legend_df.loc[legend_df.index[i],'MEANING'])


class Legend():
    def __init__(self,unit,meaning):
        self.unit= unit
        self.meaning= meaning
    
    
df = pd.DataFrame(
    data = {
        'unit':['m/s','Pa'],
        'meaning':['distance moved divided by the time','force divided by the area'], 
    },
    index=['velocity','pressure'],
)


for i,r in df.iterrows():
    exec('{} = Legend("{}","{}")'.format(i,r['unit'],r['meaning']))
    
'''