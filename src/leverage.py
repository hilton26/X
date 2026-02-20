#!/usr/bin/env python
# coding: utf-8

# # Calculate leverage in PIM portofolios for BNP

# In[266]:


# libraries, libraries!
import time
from datetime import datetime
from pathlib import Path
import pandas as pd
import numpy as np
import re # to extract dates
from tqdm import tqdm
from constants import pthCmp
from utilities import timediff


# In[267]:


# list the fund holdings files

start_time = time.time()
start_time_lvrg = time.time()
pthL = pthCmp + r"\BNP Leverage DD\Annual Review Data"
files = [
    f.name
    for f in Path(pthL).iterdir()
    if f.is_file()
    and f.name.startswith("Portfolio Analytics Report")
    and f.name.endswith(".xlsx")
]

# for file in files:
#     print(file)
print(f"{len(files)} files")

print(timediff(start_time, time.time()))


# In[268]:


# list the funds

pthF = pthCmp + r"\BNP Leverage DD\Prescient additional information.xlsx"
df2 = pd.read_excel(pthF, usecols = "N")
funds = df2.iloc[:,0].unique()
print(len(funds), funds)


# In[269]:


# load each holdings file and then create a single dataframe from them
df_all = []
for file in tqdm(files):
    df_f = pd.read_excel(pthL + '\\' + file)
    df_all.append(df_f)

df = pd.concat(df_all, ignore_index = True)
# df


# In[270]:


# normalise fund holding percentages
df['fund_mv_total'] = df.groupby(['i Position Effective Date', 'Entity ID'])['Sum of Market Value Income'].transform('sum')
df['% of Total Market Value'] = df['Sum of Market Value Income'] / df['fund_mv_total'] * 100

df['fund_ce_total'] = df.groupby(['i Position Effective Date', 'Entity ID'])['Current Exposure'].transform('sum')
df['Current Exposure %'] = df['Current Exposure'] / df['fund_ce_total'] * 100

# # check percentages
# df.groupby(['i Position Effective Date', 'Entity ID'])['% of Total Market Value'].sum()
# df.groupby(['i Position Effective Date', 'Entity ID'])['Current Exposure %'].sum()


# In[271]:


# remove lagging 00:00:00 from the date column
df['i Position Effective Date'] = df['i Position Effective Date'].dt.date


# In[272]:


# list the holdings dates
dates = df['i Position Effective Date'].unique()


# In[273]:


# get leverage values

start_time = time.time()
print(f"\nCalculating {len(files) * len(funds)} exposures \
for {len(funds)} funds over {len(files)} periods ...")

# define exclusions, being cash and cash equivalents and synthetic cash, per AIFMD
excl = ['CASH', 'MONEY MARKET', 'UNKNOWN', 'SYTH']

vals = []
for dt in tqdm(dates):

    # loop through each fund for that date
    for fund in funds:
        #dataframe the fund holdings
        df_x = df[(df['i Position Effective Date'] == dt) & (df['Entity ID'] == fund)]

        # calculate market value and effective exposure
        mv = df_x.loc[df_x["Entity ID"] == fund, "Sum of Market Value Income"].sum()
        ce = df_x.loc[df_x["Entity ID"] == fund, "Current Exposure"          ].sum()

        # sum exposure
        gross      = df_x.loc[(df_x["Entity ID"] == fund) & (~df_x["Valuation First Level"].isin(excl)), "Current Exposure"].abs().sum()
        commitment = abs(df_x.loc[(df_x["Entity ID"] == fund) & (~df_x["Valuation First Level"].isin(excl)), "Current Exposure"].sum())

        # calculate leverage
        leverage_gross      = gross / mv
        leverage_commitment = commitment / mv
     
        # populate the summary dataframe
        new_data = [dt, fund, mv, gross, commitment, leverage_gross, leverage_commitment, mv - ce]
        vals.append(new_data)

# dataframe the calculation results
column_names = ['Date', 'Fund', 'NAV', 'Exposure (Gross)', 'Exposure (Commitment)', 'Leverage (Gross)', 'Leverage (Commitment)', 'MV - CE']
summary_df = pd.DataFrame(vals)
summary_df.columns = column_names
summary_df = summary_df.sort_values(by = 'Date', ascending = False)
summary_df = summary_df.reset_index(drop = True)

print(f" {timediff(start_time, time.time())} calculating \
{len(files) * len(funds)} exposures for \
{len(funds)} funds over {len(files)} periods ...\n")


# In[274]:


# calculate averages and standard deviations over the entire period

start_time = time.time()
print(f"\nCalculating averages and standard deviations for the {len(funds)} funds")

# scale = np.sqrt(len(files))
scale = 1
rptDate = summary_df['Date'].max()
averages = []

for fund in funds:
    df_a = summary_df[summary_df['Fund'] == fund]
    average_gross = df_a['Leverage (Gross)'].mean()
    average_commitment = df_a['Leverage (Commitment)'].mean()
    stddev_gross = df_a['Leverage (Gross)'].std() * scale
    stddev_commitment = df_a['Leverage (Commitment)'].std() * scale
    new_data = [rptDate, fund, average_gross, stddev_gross, average_commitment, stddev_commitment]
    averages.append(new_data)

# dataframe the calculation results
col_names = [f'{len(files)} months ended {rptDate.strftime("%d%b%Y")}', 
             'Fund', 'Average Leverage (Gross)', 'Std Dev Leverage (Gross)', 
            'Average Leverage (Commitment)', 'Std Dev Leverage (Commitment)']
leverages = pd.DataFrame(averages)
leverages.columns = col_names

# leverages

print(f" {timediff(start_time, time.time())} calculating averages and standard deviations for the {len(funds)} funds\n")


# In[275]:


# Proportion (% of NAV) of illiquid assets that the fund structurally holds 
# (or is expected to hold). [Hedge Funds, Funds of Hedge Funds, 
# Physical Real Estate, Real Estate Funds, Infrastructure fund, 
# Private Equity, Private Equity Funds, Precious Metals, 
# Loans, and Non-agency ABS]

rptDate = pd.to_datetime('2025-11-30').date()


# In[276]:


# check what kinds of assets are held in the funds
all_assets = df.loc[(~df["Valuation First Level"].isin(excl)), "Valuation First Level"].unique()
all_assets


# In[277]:


# check 'PRIVATE COMPANY'
print(df[(df['Valuation First Level'] == "PRIVATE COMPANY")]["PrimaryAssetID"].unique())


# In[278]:


# check out the prefs held in the funds
print(df[(df['Valuation First Level'] == "PREFERENCE SHARES")]["PrimaryAssetID"].unique())


# In[279]:


# list the 'illiquid' assets
illiquid_assets = ['PIMEVOA']


# In[280]:


# show the funds holding 'illiquid' assets
illiquid_holdings = df[(df['i Position Effective Date'] == rptDate) & (df['PrimaryAssetID'].isin(illiquid_assets))][['Entity ID', 'Current Exposure %']]
illiquid_holdings


# In[281]:


# If % of NAV for high-risk asset >25%, please provide the split of illiquid assets 
# [Hedge Funds, Funds of Hedge Funds, Physical Real Estate, Real Estate Funds, 
# Private Equity, Private Equity Funds, Infrastructure fund, Precious Metals, 
# Loans, and Non-agency ABS]


# In[282]:


# add a new column naming each holding by code and description
df['Asset Description'] = df['Issue Description'] + \
' (' + df['PrimaryAssetID'].astype(str) + ') ' + \
round(df['% of Total Market Value'],1).astype(str) + '%'
df['Asset Description']


# In[283]:


# make a subset of holdings exclusding 'CASH' and 'UNKNOWN'
df_ex_CASH = df[~df['Valuation First Level'].isin(['CASH', 'UNKNOWN'])]


# In[284]:


# get top ten holdings for each fund and for each date

# Claude prompt 2 Feb 2026: given a dataframe of serval funds and their 
# respective holdings overa several dates, how to get the 
# top ten holdings of each fund over each of the dates
top10 = (df_ex_CASH.groupby(['i Position Effective Date', 'Entity ID'])
           .apply(lambda x: x.nlargest(10, '% of Total Market Value'), include_groups=False)
           .reset_index(level=[0, 1])
           .reset_index(drop=True))
# reset_index(level=[0, 1]) pulls # the date and fund 
# groupby keys back in as columns, then
# reset_index(drop=True) cleans up the leftover numeric index


# In[285]:


top10_holdings = top10[['i Position Effective Date', 'Entity ID', 'Asset Description', '% of Total Market Value']]
top10_holdings


# In[286]:


# create a new column concatenating the top ten holdings

# Geini prompt 2 Feb 2026: "for each date and each fund,
# create a new column combining the securityu names 
# making up the top ten holdings into a comm-separated string"
top10_summary = (top10.groupby(['i Position Effective Date', 'Entity ID'])['Asset Description']
                      .apply(lambda x: ', '.join(x))
                      .reset_index())
top10_summary
# top10_summary.columns = ['i Position Effective Date', 'Entity ID', 'Asset Description']


# In[287]:


# write the summary and leverage dataframe to Excel

start_time  = time.time()
filename    = pthL + '\\' + f'Fund Leverages {rptDate.strftime("%d%b%Y")}_b.xlsx'
with pd.ExcelWriter(filename, engine  = 'xlsxwriter') as writer:
    leverages.to_excel(        writer, index = False, sheet_name = 'averages' )
    summary_df.to_excel(       writer, index = False, sheet_name = 'exposures')
    illiquid_holdings.to_excel(writer, index = False, sheet_name = 'illiquids')
    top10_summary.to_excel(    writer, index = False, sheet_name = 'top10'    )
    df.to_excel(               writer, index = False, sheet_name = 'holdings' )
    
writer.close()

print(filename)

print(f'\n{timediff(start_time, time.time())} writing the dataframes to a file\n')

print(f"\n\n{timediff(start_time_lvrg, time.time())} total time\n")


# In[289]:


# !jupyter nbconvert --to script leverage_2.ipynb # convert from .ipynb to .py


# In[ ]:




