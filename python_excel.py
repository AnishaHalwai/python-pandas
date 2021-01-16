# -*- coding: utf-8 -*-
"""
Created on Mon Oct 12 16:46:03 2020

@author: Anisha
"""

# Writing to an excel  
# sheet using Python 
import pandas as pd
import os
import matplotlib.pyplot as plt

input_dir = r"C:\Users\Anisha\Dropbox\URPduration\duration-2020\input data"
output_dir = r"C:\Users\Anisha\Dropbox\URPduration\duration-2020\output data"
#------------------------------------------------------------------------------
# functions to find max STA and FTG
#------------------------------------------------------------------------------

# getting from FTG files from directory either sector wise of total:
# max of : FTG , lenght FTG, occupancy
def formulate_ftg(file, dictionary):
 
    dictionary["City"]=[]
    dictionary["PH FTG"]=[]
    dictionary["Max # of Parked Vehicles FTG"]=[]
    dictionary["Total length FTG"]=[]
    dictionary["Occupancy FTG %"]=[]
    for city in file:
        FTG_allsectors(city,dictionary)           #calling FTG function for values of that city

    return

def formulate_sta(file, dictionary):
   
    dictionary["PH STA"]=[]
    dictionary["Max # of Parked Vehicles STA"]=[]
    dictionary["Total length STA"]=[]
    dictionary["Occupancy STA %"]=[]
    for city in file:
        STA_allsectors(city,dictionary)
    return 

def formulate_sector(file, dictionary):
    numsect=2
    for sect in range(1,numsect+1):
        main_key = "Sector "+str(sect)
        sub_dict=dict()
        sub_dict["City"]=[]
        sub_dict["PH FTG"]=[]
        sub_dict["PH STA"]=[]
        sub_dict["Max # of Parked Vehicles"]=[]
        sub_dict["Max Road Length"]=[]
        dictionary[main_key]=sub_dict
        
        for city in file:
            sector_wise(city,dictionary,sect)
                
    return

def sector_wise(city,dictionary,sect):
    file = r"TotalsAllocatedbySector.csv"
    path = r"%s\%s\%s" % (output_dir,city,file)
    
    #reading excel columns
    #dictionary of max values of specified columns
    ftg = "FTG"     
    sta = "STA"     
    max_veh = "Total Vehicles Occupied"
    max_len = "Total Road Capacity Occupied"
    columns = [ftg, sta, max_veh, max_len, "Industry Sector"]
    
    max_values=pd.read_csv(path, usecols=columns )
    
    #finding range for this sector, thus finding max in that sector
    low_high = find_range(max_values, sect)
    ph_ftg = max_values.loc[low_high[0]:low_high[1], ftg].max()
    ph_sta = max_values.loc[low_high[0]:low_high[1], sta].max()
    veh = max_values.loc[low_high[0]:low_high[1], max_veh].max()
    length = max_values.loc[low_high[0]:low_high[1], max_len].max()
    
    key = "Sector " + str(sect)
    # structure :
    # dictionary = { Sector 1 : {"City":value, "PH FTG":value, "Max veh":value, "FTG len":value},
    #                Sector 2: {"City":value, "PH FTG":value, "Max veh":value, "FTG len":value}
    #               }
    dictionary[key]["City"].append(city)
    dictionary[key]["PH FTG"].append(ph_ftg)
    dictionary[key]["PH STA"].append(ph_sta)
    dictionary[key]["Max # of Parked Vehicles"].append(veh)
    dictionary[key]["Max Road Length"].append(length)
    
    return

def FTG_allsectors(city, dictionary):
    
    ftg_file = r"FTGTotalsAllocated.csv"
    ftg_path = r"%s\%s\%s" % (output_dir,city,ftg_file)

    #reading excel columns
    #dictionary of max values of specified columns
    #ph_ftg = "% FTG"
    veh_ftg = "Total Vehicles Occupied"          
    ftg_len = "Road Capacity Occupied"
    ftg_occ = "Occupancy Ratio"
    columns = [veh_ftg , ftg_len, ftg_occ]

    max_values=pd.read_csv(ftg_path, usecols=columns )
    
    low_high = (144,192)        #day 4
    fvef = max_values.loc[low_high[0]:low_high[1], veh_ftg].max()
    leng = max_values.loc[low_high[0]:low_high[1], ftg_len].max()
    focc = max_values.loc[low_high[0]:low_high[1], ftg_occ].max()
    
    # structure :
    # city : city name
    # Total vehicles occupied : max veh
    # Road capacity occupied : max road lenght
    # Occupancy Ratio: max occupancy
    dictionary["City"].append(city)
    #dictionary["PH FTG"].append(max_values[ph_ftg])
    dictionary["Max # of Parked Vehicles FTG"].append(fvef)
    dictionary["Total length FTG"].append(leng)
    dictionary["Occupancy FTG %"].append(focc*100)
    
    return 

def STA_allsectors(city, dictionary):
  
    sta_file = r"STATotalsAllocated.csv"
    sta_path = r"%s\%s\%s" % (output_dir,city,sta_file)
    
    #reading excel columns
    #dictionary of max values of specified columns

    #ph_sta = "% STA"
    veh_sta = "Total Vehicles Occupied"          
    sta_len = "Road Capacity Occupied"
    sta_occ = "Occupancy Ratio"
    columns = [veh_sta, sta_len, sta_occ]
    
    max_values=pd.read_csv(sta_path, usecols=columns )
    
    low_high = (144,192)        #day 4
    svef = max_values.loc[low_high[0]:low_high[1], veh_sta].max()
    sleng = max_values.loc[low_high[0]:low_high[1], sta_len].max()
    socc = max_values.loc[low_high[0]:low_high[1], sta_occ].max()
    
    # structure :
    # city : city name
    # Total vehicles occupied : max veh
    # Road capacity occupied : max road lenght
    # Occupancy Ratio: max occupancy
    #dictionary["PH STA"].append(max_values[ph_sta])
    dictionary["Max # of Parked Vehicles STA"].append(svef)
    dictionary["Total length STA"].append(sleng)
    dictionary["Occupancy STA %"].append(socc*100)
        
    return 

def find_range(dataframe, num):

    #results=dataframe[dataframe["Industry Sector"]==int(num)]
    #low = results.index[0]
    #high = results.index[-1]
    
    #day 4
    low=num*144
    high=low+48
    return (low,high)

def combine(dictionary,final_column, column1, column2):
    #adding individual lenghts and occupancies for total
    dictionary[final_column]=[x + y for x, y in 
              zip(column1, column2)]
    
    return

def sort_dict(dictionary):
    occ=dictionary["Occupancy %"]
    sort=sorted(range(len(occ)), key=lambda k: occ[k])
    new_dict=dict()
    
    for k in dictionary.keys():
        new_dict[k]=[]
        
    for x in reversed(range(len(sort))):
        for k in dictionary.keys():
            new_dict[k].append(dictionary[k][sort[x]])
            
    return new_dict

def plot_graph(graph, file):
    for city in file:
        file_name = r"TotalsAllocated.csv"
        path = r"%s\%s\%s" % (output_dir,city,file_name)
        values=pd.read_csv(path, usecols=["Road Capacity Occupied"])
        v=values["Road Capacity Occupied"].loc[144:192].tolist()
        
        if (city[:4]=="Soho"): graph[city]=v
        else: graph[city[:-3]]=v
    return


if __name__ == "__main__":
    
    #--------------------------------------------------------------------------
    # counting files
    #
    # getting the names and number of cities from path of output folder for 
    # URP durations
    #--------------------------------------------------------------------------
    
    files = os.listdir(input_dir) # dir is your directory path
    file_count = len(files)
    files.remove('Soho-Manhattan-S1')
    files.remove('Soho-Manhattan-S2')
    
    #-S1 and S2 cities----------------
    s1=[]
    s2=[]
    for city in files:
        if city[-2:]=="S1":
            s1.append(city)
        else:
            s2.append(city)
            
            
    #--------------------------
  
    soho = []
    soho.extend(['Soho-Manhattan-S1','Soho-Manhattan-S2'])
    print(soho)
    print(s1)
    
    #--------------------------------------------------------------------------
    # industry sectors considered
    #--------------------------------------------------------------------------
    
    values_s1=dict()
    values_s2=dict()
    formulate_sector(s1, values_s1)
    formulate_sector(s2, values_s2)
    
    
    writer = pd.ExcelWriter("Tables.xlsx", engine='xlsxwriter')
    
    list_both=[values_s1,values_s2]
    
    index=1
    for val in list_both:
        dataframe=[]
        frames=[]
        sectors=2
        
        for s in range(1,sectors+1):
            name = "Sector " + str(s)
            dframe = pd.DataFrame(val[name])
            dframe = dframe.round(decimals=2)
            dframe.set_index("City", inplace=True)
            dataframe.append(dframe)
        
        for s in range(1,sectors+1):
            name = "Sector " + str(s)
            df = pd.concat({name: pd.DataFrame(dataframe[s-1])}, axis=1)
            frames.append(df)
        
        final = pd.concat(frames, axis=1, sort=False)
      
        final.to_excel(writer, sheet_name='Sheet S'+str(index))
        index+=1
    
    #---------------------------------
    #formatting (incomplete)----------
    
    workbook  = writer.book
    #worksheet = writer.sheets['Sheet1']
    
    
    # Add a header format.
    header_format = workbook.add_format({
        'text_wrap': True})
    
    f=workbook.add_format()
    f.set_text_wrap()
    
    #for f in dataframe:
     #   for col in f.columns.get_indexer(f.columns.values):
      #      print(value)
       #     worksheet.write(1, col_num+1 , value, header_format)
         
    
    #for f in dataframe:
     #   print(f.columns)
      #  for col_num, value in enumerate(f.columns):
       #             worksheet.write(1,col_num + 1, value, header_format)
    #worksheet.set_column('A:N', 10, columns_format)
    #worksheet.set_row(1, 10, columns_format)
    
    #--------------------------------------------------------------------------
    # total values of all sectors and cities
    #--------------------------------------------------------------------------
    all_values_s1=dict()
    all_values_s2=dict()
    
    formulate_ftg(s1,all_values_s1)
    formulate_sta(s1,all_values_s1)
    
    formulate_ftg(s2,all_values_s2)
    formulate_sta(s2,all_values_s2)
    
    all_values_both=[all_values_s1,all_values_s2]
    
    index=1
    for all_val in all_values_both:
        tmp=dict()
        
        if(all_val==all_values_s1): tmp=values_s1
        else: tmp=values_s2
        
        
        combine(all_val, "Total Length", all_val["Total length STA"], all_val["Total length FTG"])
        combine(all_val, "Occupancy %", all_val["Occupancy STA %"], all_val["Occupancy FTG %"])
        combine(all_val, "PH FTG", tmp["Sector 1"]["PH FTG"], tmp["Sector 2"]["PH FTG"])
        combine(all_val, "PH STA", tmp["Sector 1"]["PH STA"], tmp["Sector 2"]["PH STA"])
        
        all_val=sort_dict(all_val)
        
        #manually adding occupancy after 50% amd 25% linear caps
        all_val["Occupancy % (50%)"]=[x*2 for x in all_val["Occupancy %"] ]
        all_val["Occupancy % (25%)"]=[x*4 for x in all_val["Occupancy %"] ]
        
        order=["PH FTG", "PH STA","Max # of Parked Vehicles FTG",
               "Max # of Parked Vehicles STA", "Total length FTG", "Total length STA", 
               "Total Length", "Occupancy FTG %", "Occupancy STA %", "Occupancy %",
               "Occupancy % (50%)","Occupancy % (25%)"]
        
        df = pd.DataFrame(all_val)
        df = df.round(decimals=2)
        df = df.set_index("City",drop=False)
        df.to_excel(writer, sheet_name='Sheet S'+str(index), columns=order,
                    startrow=final.shape[0] + 5, startcol=0)
        index+=1
        #worksheet.set_column('A:N', None, header_format)
        #worksheet.set_row(14, None, header_format)
    
    #--------------------------------------------------------------------------
    
    graph_s1=dict()
    graph_s2=dict()
    soho_graph=dict() #dictionary graph for soho results
    
    index=["12:00AM","12:30AM","1:00AM","1:30AM","2:00AM","2:30AM","3:00AM","3:30AM",
           "4:00AM","4:30AM","5:00AM","5:30AM", "6:00AM","6:30AM","7:00AM","7:30AM",
           "8:00AM","8:30AM","9:00AM","9:30AM","10:00AM","10:30AM","11:00AM","11:30AM",
           "12:00PM","12:30PM","1:00PM","1:30PM","2:00PM","2:30PM","3:00PM","3:30PM",
           "4:00PM","4:30PM","5:00PM","5:30PM", "6:00PM","6:30PM","7:00PM","7:30PM",
           "8:00PM","8:30PM","9:00PM","9:30PM","10:00PM","10:30PM","11:00PM","11:30PM",
           "12:00AM"]
    
    graphs=[soho_graph, graph_s1, graph_s2]
    sheet_names=["soho results", "S1 results","S2 results" ]
    file_names=[soho, s1, s2]
    
    for x in range(len(graphs)):
        plot_graph(graphs[x], file_names[x])    #dataframe for soho results
        df = pd.DataFrame(graphs[x], index=index)
        df.to_excel(writer, sheet_name=sheet_names[x] )
        
        workbook = writer.book
        worksheet = writer.sheets[sheet_names[x]]
        chart = workbook.add_chart({'type': 'line'})
        
        marker_types=['automatic', 'none', 'square', 'diamond', 'triangle', 'x', 'star', 
                      'short_dash', 'plus', 'circle']
        
        for i in range(len(file_names[x])):
            col = i + 1
            mark = marker_types[i]
            chart.add_series({
                'name':       [sheet_names[x], 0, col],
                'categories': [sheet_names[x], 1, 0,   49, 0],
                'values':     [sheet_names[x], 1, col, 49, col],
                'marker': {'type': mark },
            })
        # Configure the chart axes.
        chart.set_x_axis({'name': 'Time Interval','interval_unit': 2})
        chart.set_y_axis({'name': 'FSA Parking Road Length(m)'})
      
        worksheet.insert_chart('D2', chart)
    
    worksheet.set_row(1, None, header_format)
    writer.save()
#------------------------------------------------------------------------------
    
    # to do:
    # STA sector wise
    # compare max vehicles and change data
    # add sector headers
    # search for zip

    #Table1
    #add FTG STA vehicles in table 1
    #change PH STA, FTG from vehicles to STA. FTG
