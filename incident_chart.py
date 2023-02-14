import pandas as pd
import numpy as np
import warnings
import xlsxwriter
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import matplotlib.pyplot as plt
from datetime import datetime
from openpyxl.formatting.rule import DataBarRule
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import colors
from pandas import ExcelWriter
from datetime import timedelta
import sys

# Set pandas display options
pd.set_option("display.precision", 4)
pd.set_option("display.expand_frame_repr", False)
pd.set_option("display.max_rows", None)

# turn off warnings for set on a copy of a slice from a df etc
warnings.filterwarnings("ignore")

# set up excel writer
writer = pd.ExcelWriter('t2.xlsx', engine='xlsxwriter', datetime_format='yyyy-mm-dd')

# Load data exported from AssetPlus
data_raw = pd.read_excel("incident_simulation1.xls")

# rename columns for easy labelling
data_raw = data_raw.rename(columns={
    'Incident Date': 'inc',
    'Received Date': 'rec',
    'Added to Assetplus Date': 'add',
    'Action Date': 'act',
    'Due Date': 'due',
    'Closed Date': 'clo',
    'Elapsed Workdays': 'elap_wd',
    'Workdays to Due Date': 'wd_due'
    })

# convert all datetimes
data_raw['inc'] = pd.to_datetime(data_raw['inc'], format="%Y%m%d", utc=False)
data_raw['rec'] = pd.to_datetime(data_raw['rec'], format="%Y%m%d", utc=False)
data_raw['add'] = pd.to_datetime(data_raw['add'], format="%Y%m%d", utc=False)
data_raw['act'] = pd.to_datetime(data_raw['act'], format="%Y%m%d", utc=False)
data_raw['due'] = pd.to_datetime(data_raw['due'], format="%Y%m%d", utc=False)
data_raw['clo'] = pd.to_datetime(data_raw['clo'], format="%Y%m%d", utc=False)

# convert wd_due from float to int64
data_raw['wd_due'] = data_raw['wd_due'].astype('Int64')


# Note: the perf column is not included in the A+ report, hence we should calculate it
# Best method to date is a bit of a work around:-
# split the df to eliminate the NaNs
not_na = data_raw[pd.notna(data_raw.clo)]
is_na = data_raw[pd.isna(data_raw.clo)]

# get the due to closed days
not_na['perf'] = np.busday_count(not_na['due'].values.astype('datetime64[D]'),
                                 not_na['clo'].values.astype('datetime64[D]'))

# join dfs to get the full data set again
data = pd.concat([is_na, not_na], axis=0)

# perf is set as floats so convert to integers
data['perf'] = data['perf'].astype('Int64')

# get the year-month for grouping and graph labels
data['month_inc'] = pd.to_datetime(data['inc']).dt.to_period('M')   # used for KPI calcs
# don't need but there for possible future use
data['day_of_week'] = data['due'].apply(lambda x: x.weekday())
data['working_day'] = (data['day_of_week'] >= 0) & (data['day_of_week'] <= 4)

# add to first excel sheet
data.to_excel(writer, sheet_name='DATA')

# DO KPI STATS
# check KPI for days to add - label each True (OK), False (Not OK)
data['kpi_ok'] = data['elap_wd'].apply(lambda x: x < 6)
month_ok_count = data.groupby(['month_inc', 'kpi_ok'])['elap_wd'].count()
elap_wd_pct = month_ok_count.groupby(level=0).apply(lambda x: 100 * x / x.sum())

# now the main stats regarding the KPI
kpi_1_stats = data.groupby('month_inc').agg({"elap_wd": [max, min, "mean", "count"]}).round(2)

# add to second sheet
kpi_1_stats.to_excel(writer, sheet_name='PLOTS')

# DO CLOSING STATS
# Use the subset acquired earlier that just includes the rows with Closed dates
print("Now do closing stats ...")
not_na['month_due'] = pd.to_datetime(not_na['due']).dt.to_period('M')

perf_stats = not_na.groupby('month_due').agg({"perf": [max, min, "mean", "count"]})
print(perf_stats)

# add to PLOTS sheet
perf_stats.to_excel(writer, sheet_name='PLOTS', startrow=20)

# now the perf percentages
not_na['on_time'] = not_na['perf'].apply(lambda x: x <= 0)
month_on_time_count = not_na.groupby(['month_due', 'on_time'])['perf'].count()
on_time_pct = month_on_time_count.groupby(level=0).apply(lambda x: 100 * x / x.sum())
print(on_time_pct)

# add to third sheet
on_time_pct.to_excel(writer, sheet_name='PLOTS', startrow=40)

# save excel
print("saving ...")
writer.save()








