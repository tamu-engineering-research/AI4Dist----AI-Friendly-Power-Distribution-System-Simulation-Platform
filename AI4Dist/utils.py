import numpy as np
import pandas as pd
from datetime import datetime


# convert Cartesian to Polar for complex numbers
def cart_to_pol(arr):
    # the format of input array is:
    # [Re1, Im1, Re2, Im2, ...]
    dim = int(len(arr) / 2)
    mag = np.zeros(dim)
    angle = np.zeros(dim)
    for i in range(dim):
        re = arr[i*2]
        im = arr[i*2 + 1]
        cplx = re + 1j*im
        mag[i] = np.abs(cplx)
        angle[i] = np.angle(cplx)

    return mag, angle

## Get the number of a string, ignoring other characters
def getNum(string):
    length = len(string)
    i = 0
    while i < length:
        if (ord(string[i]) == 46) or (ord(string[i]) >= 48 and ord(string[i]) <= 57):
            pass
        else:
            string = string.replace(string[i],'')
            length = len(string)
            i = i - 1
        i = i + 1
    if length > 2:
        if ord(string[length-1]) == 46:
            string = string[0:(length-1)]
    try:
        num = float(string)
        return num
    # return None if conversion fails
    except Exception:
        return None

# normalize a dataframe
def normalize(df):
    result = df.copy()
    min_val = df.min()
    max_val = df.max()
    result = (df - min_val) / (max_val - min_val)

    return result


# load pv and load profile
def parse_profile(pv_path, load_path):
    date_range = pd.date_range(datetime(2019, 1, 1), datetime(2019, 12, 31))
    
    df_pv = pd.read_csv(pv_path, index_col='date', parse_dates=True)
    df_pv = df_pv.loc[df_pv.index.isin(date_range)]
    df_pv = df_pv.set_index([df_pv.index, 'fuel'])
    df_pv = pd.DataFrame({'value': df_pv.stack()}).reset_index(1)
    df_pv.index = pd.to_datetime(
        df_pv.index.get_level_values(0).strftime('%Y-%m-%d ') + df_pv.index.get_level_values(1))
    df_pv.index.name = 'date_time'

    df_load = pd.read_csv(load_path, index_col='date', parse_dates=True)
    df_load = df_load.loc[df_load.index.isin(date_range)]
    df_load = pd.DataFrame({'load': df_load.stack()})
    df_load.index = pd.to_datetime(
        df_load.index.get_level_values(0).strftime('%Y-%m-%d ') + df_load.index.get_level_values(1))
    df_load.index.name = 'date_time'

    pv_profile = df_pv[df_pv['fuel']=='solar']['value']
    load_profile = df_load['load']

    pv_profile = normalize(pv_profile)
    load_profile = normalize(load_profile)

    return pv_profile, load_profile

