#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Jul 27 13:41:57 2021

@author: nikol
"""

import pandas as pd

## costs_preprocessing
#Oemof (in the investement method) needs the annuity of NPV of each component 


#input 
# project duration:  project_lifetime
# component's (i) life time t
# wacc
# investement per unit: c_invest
# i: discounting factor 
# OPEX(fix): Operation & Mangement costs  [currency/unit]



#output
# 1- replacement cost 
# 2- residual costs
# 3- capex calculation
# 4- annuity = capex + opex (O&M) 


###############################################################################
# How much replacement of component are needed ?

def number_of_investement (project_lifetime,component_lifetime): 
    
     #Parameters:
     #project_lifetime
     #component_lifetime
    
     #Result:
     #umber_of_investement
    
    
    if project_lifetime == component_lifetime :
        
        n=1
   
    else:    
        n= round(project_lifetime/component_lifetime+0.5)
        
        '''
        why do we add 0.5 in the round function ? 
        this a exampke for n= 2.3 
        
        case 1: 
        round(2.3)= 2, this is no true because investement should be done more than 2 times (2.3), should be 3 times 
        
        case 2: 
        round(2.3+ 0.5)= 3, this true! we need 3 investement
        
        '''
        
        
    return n


###############################################################################
# After the end the project the component has a residual value if t>T

def residual_value (c_invest,project_lifetime,component_lifetime,Investement_num,wacc):
      
     #Parameters:
     #Investement_num: number_of_investement needed for each component
     #life time of the component 
     #project duration 
     #c_invest: first investement 
     #c_last: discounted investment cost of the last replacement
     #wacc
     
     #Result:
     #residual value of the component based on linear depreciation c_last/component_lifetime
      
     
    if Investement_num* component_lifetime > project_lifetime: 
        
         c_last     = c_invest/(1+wacc)**((Investement_num-1)*component_lifetime)
         
         c_residual =c_last/component_lifetime*(Investement_num*component_lifetime-project_lifetime)
         
    else: 
        
        c_residual=0
        
    return c_residual


###############################################################################
# The CAPEX of the per-unit costs for each component i
# CAPEX of a component takes into consideration the residual value 

def Capex (c_invest,c_residual,project_lifetime,component_lifetime,Investement_num,wacc):
    
    #Parameters:
    #c_invest: specific investment costs [currency/kW]
    #project_lifetime: project duration 
    #n: number of investement 
    #wacc
    
    
    #Result:
    #Capex_per_unit:
    #Capex_per_kW: if the c_invest in [currency/kW]
    #Capex_per_kWh: if the c_invest in [currency/kWh] for instance for the storage 
    
    
    capex=0
    
    if component_lifetime == project_lifetime:
        capex= c_invest
    
    
    
    if  Investement_num*component_lifetime != project_lifetime:
        
          
            Replacement_i=0
            
            while Replacement_i < Investement_num:
                
                 capex= capex + c_invest/(1+wacc)**(Replacement_i*component_lifetime)
                 
                 Replacement_i= Replacement_i+ 1
            
         
            #another method using the for loop 
            '''
            for Replacement_i in range(0,Investement_num):
                capex= capex + c_invest/(1+wacc)**(Replacement_i*component_lifetime)
                '''
           
    
    if Investement_num*component_lifetime> project_lifetime:
        
        capex= capex  - (c_residual/(1+wacc)**project_lifetime)
        

    return capex


###############################################################################
# # The CAPEX based on replacement cost

def Capex_replacement (c_invest,
                       c_residual,
                       project_lifetime,
                       component_lifetime,
                       Investement_num,
                       wacc):
    
    #Parameters:
    #c_invest: specific investment costs [currency/kW]
    #project_lifetime: project duration 
    #n: number of investement 
    #wacc
    
    
    #Result:
    #Capex_per_unit:
    #Capex_per_kW: if the c_invest in [currency/kW]
    #Capex_per_kWh: if the c_invest in [currency/kWh] for instance for the storage 
    
    
    capex=0
    
    if component_lifetime == project_lifetime:
        capex= 0
    
    
    
    if  Investement_num*component_lifetime != project_lifetime:
        
          
            Replacement_i=0
            
            while Replacement_i < Investement_num:
                
                 capex= capex + c_invest/(1+wacc)**(Replacement_i*component_lifetime)
                 
                 Replacement_i= Replacement_i+ 1
            
         
    
    if Investement_num*component_lifetime> project_lifetime:
        
        capex= capex- (c_residual/(1+wacc)**project_lifetime)
        

    return capex

###############################################################################
#Capex_reinvest: used to calculate the capex of  FC and EL system and stacks 
#seperatly consedering the replacements costs   

def Capex_reinvest (c_invest,
                       c_residual,
                       project_lifetime,
                       component_lifetime,
                       Investement_num,
                       wacc,
                       c_reinvest=None):
    
    #Parameters:
    #c_invest: specific investment costs [currency/kW]
    #project_lifetime: project duration 
    #n: number of investement 
    #wacc
    #c_reinvest: for example 50% of c_invest (specific investment costs) have to be reinvest
    
    
    #Result:
    #Capex_per_unit:
    #Capex_per_kW: if the c_invest in [currency/kW]
    #Capex_per_kWh: if the c_invest in [currency/kWh] for instance for the storage 
    
    
    capex=0
    
    if component_lifetime == project_lifetime:
        capex= 0
    
    
    
    if  Investement_num*component_lifetime != project_lifetime:
        
        if c_reinvest is not None:
            
            
             
            capex= c_invest
            Replacement_i=1
            
            while Replacement_i < Investement_num:
                
                 capex= capex + c_reinvest/(1+wacc)**(Replacement_i*component_lifetime)
                 
                 Replacement_i= Replacement_i+ 1
            
         
                        
       
        else:    
            Replacement_i=1

            while Replacement_i < Investement_num:

                     capex= capex + c_invest/(1+wacc)**(Replacement_i*component_lifetime)

                     Replacement_i= Replacement_i+ 1

        
          
           
            
         
    
    if Investement_num*component_lifetime> project_lifetime:
        
        capex= capex- (c_residual/(1+wacc)**project_lifetime)
        

    return capex
###############################################################################
# Capital Recovery Factor: CRF to calculate the present value of an annuity 

def CRF_fnc (i,project_lifetime):
    
    #Parameters:
    #project_lifetime: project duration 
    #i: discounting factor 
    
    #Result:
    #CRF
 
    CRF= (i*(1+i)**project_lifetime)/((1+i)**project_lifetime-1)

    return CRF



###############################################################################
# Annual costs of each component: annuity wich oemof in the investement method will be required  
# Annuity takes into account CAPEX using CRF and O&M in form of OEPX per unit
# The Annuity are not taking account the Utilisation_cost_per_kWh_per_year (Var) 
#CRF ist used calculate the present value of an annuity 


def annuity(capex,Opex,CRF):
    
    #Parameters:
    #capex: [currency/kW] or #CAPEX: [currency/kWh] (e.g in case of battery)
    #Opex(O&M): Operation & Mangement costs [currency/kW] or Opex(O&M): [currency/kWh] (e.g in case of battery)
    #CRF: capital recovery factor 
    
    
    #Result:
    #annuity_per_kW  if capex in [currency/kW]  and Opex(O&M)  in [currency/kW]
    #annuity_per_kWh if capex in [currency/kWh] and Opex(O&M)  in [currency/kWh] (e.g in case of battery)
  
    annuity= capex * CRF + Opex
    
    return annuity 

###############################################################################
# the total cost of the component including the utilization_cost [currency/kWh]


def total_cost_component(annuity,cost_var,cap_opt,E_opt):
    
    #annuity: CAPEX + OPEX(=O&M) [currency/kW]
    #cost_var: utilization_cost  [currency/kWh]
    #cap_opt: optimized capacity after the optimization problem is solved in kW
    #E_opt:   optimized energy Flow the optimization problem is solved in kWh
    #total_cost_comp: currency for example EURO or USD
  
    total_cost_comp = cap_opt * annuity + E_opt * cost_var
    
    return total_cost_comp 


###############################################################################
# After getting the optimised operations of each component
# we calculate the component-specific Annuity (Total_Annuity) 
# including the utilization costs per kWh and Fuel_expenditures 


def cost_component_optimized (Annuity_per_kW_list
                              ,Annuity_per_kWh_list
                              ,Capex_per_kW_list
                              ,Capex_per_kWh_list
                              ,Utilisation_cost_list
                              ,fuel_price_list
                              ,Fuel_comsumption_list
                              ,Capacity_opt_kW_list
                              ,Capacity_opt_kWh_list                              
                              ,Energy_opt_list
                              ,component_List
                              ,currency):
    
    #Parameters:
    #Annuity_per_kW_list : list of component-specific annuities of the component in [currency/kW]
    #Annuity_per_kWh_list: list of component-specific annuities of the component in [currency/kWh]
    #Capex_per_kW_list     [currency/kW]
    #Capex_per_kWh_list    [currency/kWh]
    #Utilisation_cost_list [currency/kWh]
    #fuel_price_list in    [currency/liter]
    #Fuel_comsumption_list [liter]
    #Capacity_opt_kW_list  [kW]
    #Capacity_opt_kWh_list [kWh]
    #component_List
    #curreny               
   
    
    #Result:
    #Annuity_Optimised_Sizing:
    #Utili_Cost_Optimised_Energy
    #Fuel_expenditures 
    #Total_Annuity= Annuity_Optimised_Sizing + Utili_Cost_Optimised_Energy + Fuel_expenditures
    
    
  
    
    res_eco ={}
   
    Annuity_Optimised_Sizing    =[]
    Capex_Optimised_Sizing      =[]
    Utili_Cost_Optimised_Energy =[]
    Fuel_expenditures           =[]
    Total_Annuity               =[]
    
   
    
    for i in range(len(component_List)): 
   
        
        Annuity_Optimised_Sizing.append(Capacity_opt_kW_list[i] * Annuity_per_kW_list[i] + Capacity_opt_kWh_list[i] * Annuity_per_kWh_list[i])
        
        '''
        
        for the battery 
        Annuity_Optimised_Sizing = Annuity_per_kW [currency/kW] * Optimised_Energieflow [kW] + Annuity_per_kWh [currency/kWh] * Optimised_Sizing [kWh]
        
        '''
        
        Capex_Optimised_Sizing.append(Capacity_opt_kW_list[i] * Capex_per_kW_list[i] +  Capacity_opt_kWh_list[i] * Capex_per_kWh_list[i])
        
        '''
        
        for the battery 
        Capex_Optimised_Sizing  = Capex_per_kW [currency/kW] * Optimised_Sizing [kW] + Capex_per_kWh [currency/kWh] * Optimised_Energieflow [kWh]
        
        '''
        
        
        Utili_Cost_Optimised_Energy.append(Energy_opt_list[i] * Utilisation_cost_list[i])
        
#        Fuel_expenditures.append(fuel_price_list[i] * Fuel_comsumption_list[i])
        
        Total_Annuity.append(Annuity_Optimised_Sizing[i] + Utili_Cost_Optimised_Energy[i])#+ Fuel_expenditures[i]
        
        
        res_eco['ANNUITY_'+currency]           = Annuity_Optimised_Sizing
        res_eco['CAPEX_'+currency]             = Capex_Optimised_Sizing
        res_eco['UTI_COSTS_'+currency]         = Utili_Cost_Optimised_Energy
#        res_eco['FUEL_EXP_'+currency]          = Fuel_expenditures
        res_eco['TOTAL_ANNUITY_'+currency]     = Total_Annuity

    
    DF= pd.DataFrame.from_dict(res_eco,orient='index', columns=component_List)
    
    
    return DF


###############################################################################
# system_opt_result function gives a resume about the main parameter of the system
#FUEL_CONSUMPTION

def system_opt_result (total_cost_system,total_energy_supply,total_energy_production,
                       RE_energyflow
                      
                       ,EXCESS_GENERATION
                       ,TOTAL_DEMAND,
                       TOTAL_DEMAND_PUSE
                       ,FUEL_CONSUMPTION


                       ,currency) :
 
    
    res_eco ={}
    
    TOTAL_SYS_COST    = total_cost_system
    TOTAL_GENERATION  = total_energy_production
    TOTAL_DEMAND      = TOTAL_DEMAND
    TOTAL_DEMAND_PUSE = TOTAL_DEMAND_PUSE
    
    TOTAL_DEMAND_SUP  = TOTAL_DEMAND
    LCOE              = total_cost_system/total_energy_supply
    FUEL_CONSUMPTION  = FUEL_CONSUMPTION

    EXCESS_GENERATION =  EXCESS_GENERATION
    RE_Share          =  RE_energyflow/total_energy_production *100
    Supply_Demand_Ratio = RE_energyflow/ TOTAL_DEMAND *100

    res_eco['TOTAL_SYS_COST_'+currency]         = TOTAL_SYS_COST 
    res_eco['TOTAL_GENERATION_kWh']             = TOTAL_GENERATION
    res_eco['TOTAL_DEMAND_kWh']                 = TOTAL_DEMAND

    res_eco['TOTAL_DEMAND_PUSE_kWh']             = TOTAL_DEMAND_PUSE
    res_eco['EXCESS_GENERATION_kWh']            = EXCESS_GENERATION
    res_eco['FUEL_CONSUMPTION_kg']           = FUEL_CONSUMPTION 

    res_eco['LCOE_'+currency+'/kWh']            = LCOE 
    res_eco['RE_SHARE_%']                       = RE_Share
    res_eco['SUPPLY_DEMAND_RATIO_%']            = Supply_Demand_Ratio  
    
        
    DF= pd.DataFrame.from_dict(res_eco,orient='index',columns=['System_parameter'])
    
    return DF 

def fuel_price_change(fuel_price_t0,fuel_escalation_rate,project_lifetime,wacc,CRF ):
    
    if fuel_escalation_rate== 0:
        
        fuel_new_price= fuel_price_t0
    
    else:
        
        fuel_price_CF = 0
        
        fuel_price_i = fuel_price_t0
        
        for i in range(0, project_lifetime):
            
                    fuel_price_CF += fuel_price_i / (1 + wacc) ** (i)
                
                    fuel_price_i = fuel_price_i * (1 + fuel_escalation_rate)
            
        fuel_new_price = fuel_price_CF * CRF   
    
    
    
    
    return fuel_new_price


def list_price(start_price,LENGTH,kind=None,step=None,price_percent=None,price_referenz=None,step_referenz=None):
    
    price_list=[start_price]
    
    
    if price_percent is None and price_referenz is None and step_referenz is None :
        
        
        for i in range(LENGTH-1):
            EL_price_new= price_list[i] + step
            price_list.append(EL_price_new)
            
    if step is None and price_referenz is None and step_referenz is None :
        
        for i in range(LENGTH-1):
            
            EL_price_new= price_list[i] + (price_list[i] *price_percent)
            price_list.append(EL_price_new)
            
    if price_referenz is not None and step_referenz is not None and step is None and price_percent is None :
        
        price_ref=[price_referenz]
        step_ref = step_referenz
        
        factor=price_referenz/start_price
        
        for i in range(LENGTH-1):
            
            EL_price_ref_new= (price_ref[i] + step_ref) 
            
            price_ref.append(EL_price_ref_new)
            
            if kind== 'multi':
                EL_price_new=price_ref[i+1] * factor
            
                price_list.append(EL_price_new)
                
            if kind== 'div':
                EL_price_new=price_ref[i+1] / factor
            
                price_list.append(EL_price_new)
            
            
         
    return [round(i) for i in price_list ]
    
  
    
price_referenz= 4000      
            
    