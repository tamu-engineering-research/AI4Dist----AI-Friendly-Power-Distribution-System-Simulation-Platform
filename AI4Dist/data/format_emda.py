import numpy as np
import pandas as pd
from datetime import datetime


iso = 'spp'
city = 'south'

load_path = rf"C:\Users\Dongqi Wu\OneDrive\Work\Finished\LAADS\EMDA\COVID-EMDA\data_release\{iso}\{iso}_{city}_load.csv"
genmix_path = rf"C:\Users\Dongqi Wu\OneDrive\Work\Finished\LAADS\EMDA\COVID-EMDA\data_release\{iso}\{iso}_rto_genmix.csv"

date_range = pd.date_range(datetime(2017, 1, 1), datetime(2020, 12, 31))
    

df_load = pd.read_csv(load_path, index_col='date', parse_dates=True)
df_genmix = pd.read_csv(genmix_path, index_col='date', parse_dates=True)

df_genmix = pd.read_csv(genmix_path, index_col='date', parse_dates=True)
df_genmix = df_genmix.loc[df_genmix.index.isin(date_range)]
df_genmix = df_genmix.set_index([df_genmix.index, 'fuel'])
df_genmix = pd.DataFrame({'value': df_genmix.stack()}).reset_index(1)
df_genmix.index = pd.to_datetime(
        df_genmix.index.get_level_values(0).strftime('%Y-%m-%d ') + df_genmix.index.get_level_values(1))
df_genmix.index.name = 'date_time'


df_load = df_load.loc[df_load.index.isin(date_range)]
df_load = pd.DataFrame({'load': df_load.stack()})
df_load.index = pd.to_datetime(
df_load.index.get_level_values(0).strftime('%Y-%m-%d ') + df_load.index.get_level_values(1))
df_load.index.name = 'date_time'
df_load = df_load.rename(columns={'load':'value'})

df_pv = df_genmix.loc[df_genmix['fuel']=='solar']
df_pv = df_pv.drop(columns='fuel')
df_wind = df_genmix.loc[df_genmix['fuel']=='wind']
df_wind = df_wind.drop(columns='fuel')


df_load = (df_load-df_load.min())/(df_load.max()-df_load.min())
df_pv = (df_pv-df_pv.min())/(df_pv.max()-df_pv.min())
df_wind = (df_wind-df_wind.min())/(df_wind.max()-df_wind.min())


df_load.to_csv(rf'load_profile\{iso}_{city}_load.csv')
df_pv.to_csv(rf'pv_profile\{iso}_{city}_pv.csv')
df_wind.to_csv(rf'wind_profile\{iso}_{city}_wind.csv')
