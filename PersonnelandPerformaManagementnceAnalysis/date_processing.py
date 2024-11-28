import os
import openpyxl
import pandas as pd
import numpy as np


def date_processing():
    all_date = pd.DataFrame({'all_date': pd.date_range('2023-01-01', '2023-12-31')})
    # dayofweek: 0-Mon ---> 6-Sun
    weekend_cond_ = (all_date['all_date'].dt.dayofweek != 5) & (all_date['all_date'].dt.dayofweek != 6)
    all_date = all_date[weekend_cond_]

    all_date['Month'] = all_date['all_date'].dt.month
    all_date['Week'] = all_date['all_date'].dt.isocalendar().week

    ref_date = pd.DataFrame()
    for month in range(1, 13):
        cond_ = all_date['Month'] == month
        each_month = all_date[cond_].copy()
        each_month['Week number'] = each_month['Week'] - (each_month['Week'].min() - 1)
        ref_date = pd.concat([ref_date, each_month])

    return ref_date