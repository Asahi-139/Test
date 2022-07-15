import pandas as pd
import numpy as np
from datetime import datetime
import openpyxl

from dateutil.parser import parse
pd.options.mode.chained_assignment = None 
import warnings
warnings.simplefilter(action='ignore', category=UserWarning)

def reading_data(x):
    main_file = pd.read_excel(x)
    main_file = main_file.replace(np.nan,"",regex=True)
 
    main_file = main_file[main_file["TestStatus"].str.contains("In-Process")==False]
    unique_part_no = list(main_file["PartNumber"].unique())
    Unique_serial = list(main_file["SerialNumber"].unique())
    main_file = main_file.reset_index(drop=True)
    data = main_file["SerialNumber"] +"|"+ main_file["Station"]
    main_file.insert(0, 'Combine_Serial', data)
    
    # converting timestamp to real time format
    for i in range (len(main_file)):
        timestamp = main_file.iloc[i][3]
        timestamp = timestamp.split(" ")
        date = timestamp[0]
        date = date.split("/")
        date_d = date[1]
        date_m = date[0]
        date_y = date[2]
        date = date_d+"/"+date_m+"/"+date_y
        timestamp[0] =  date
        timestamp = " ".join(timestamp)
        # date = parse(timestamp)
        date =  datetime.strptime(timestamp, '%d/%m/%Y %H:%M:%S %p')
        main_file.at[i,["InDate"]]=date
    main_file = main_file.dropna(subset = ["TestStatus"])
    
    # Now we are gonna start making final sheet
    
    df = pd.DataFrame(Unique_serial, columns=["Serial No."])
    unique_station = list(main_file["Station"].unique())
    station_df_list = []
    # for station in unique_station:
    #     station= main_file[main_file["Station"].str.contains(station)]
    #     station_df_list.append(station)
        
    #for serial in Unique_serial:
        
    
    
    
    
    main_file.to_excel("output1.xlsx")
    return main_file ,df, station_df_list



x = reading_data("SMR 2kW.xlsx")