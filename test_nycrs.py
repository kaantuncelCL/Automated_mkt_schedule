#!/usr/bin/env python3
"""
Test run of the RockSling notebook with NYCRS / 2713
"""

import os
import requests
import openpyxl as opxl
import pandas as pd
import numpy as np
import glob
from datetime import datetime
from dateutil.relativedelta import relativedelta
from pandas.tseries.offsets import QuarterEnd

print("="*80)
print("ROCKSLING AUTOMATION TEST - NYCRS / 2713")
print("="*80)

# Configuration
PREQIN_API_USERNAME = os.environ.get('PREQIN_API_USERNAME', 'preqinapiuser@campbell-lutyens.com')
PREQIN_API_KEY = os.environ.get('PREQIN_API_KEY', '6a5a2291b5c84cab968009e51ba0aaf9')
DEFAULT_COMMITMENT_MN = 20

print("\n📋 Configuration:")
print(f"  API Username: {PREQIN_API_USERNAME}")
print(f"  Default Commitment: ${DEFAULT_COMMITMENT_MN}M")

# Test parameters
preqin_firm_id_input = '2713'
preqin_firm_name = 'NYCRS'
strategy_lookup = 'pe'
portfolio_name_rocksling = 'NYCRS test'
underwriting_input = 'moic-based'

print(f"\n🎯 Test Parameters:")
print(f"  Firm ID: {preqin_firm_id_input}")
print(f"  Firm Name: {preqin_firm_name}")
print(f"  Strategy: {strategy_lookup}")
print(f"  Portfolio Name: {portfolio_name_rocksling}")
print(f"  Underwriting: {underwriting_input}")

# Initialize Preqin API
print("\n🔐 Authenticating with Preqin API...")
url = "https://api.preqin.com/connect/token"
headers = {"Content-Type": "application/x-www-form-urlencoded"}
data = f"username={PREQIN_API_USERNAME}&apikey={PREQIN_API_KEY}"

try:
    response = requests.post(url, headers=headers, data=data, timeout=30)

    if response.status_code == 200:
        access_token = response.json()["access_token"]
        print("✓ Access token successfully retrieved")
    else:
        print(f"✗ API request failed with status code: {response.status_code}")
        print(f"Response: {response.text}")
        exit(1)
except Exception as e:
    print(f"✗ Error connecting to API: {str(e)}")
    exit(1)

# Find and load Preqin file
print("\n📁 Loading Preqin fund performance data...")
pattern = "Preqin_Fundperformance_export-*.xlsx"
matching_files = glob.glob(pattern)

if not matching_files:
    print(f"✗ No file matching pattern '{pattern}' found")
    exit(1)

preqin_file = max(matching_files, key=os.path.getmtime)
print(f"  Using: {preqin_file}")

try:
    fundperformance_df = pd.read_excel(preqin_file)
    print(f"✓ Loaded {len(fundperformance_df)} fund performance records")
except Exception as e:
    print(f"✗ Error loading file: {str(e)}")
    exit(1)

# Load rocksling returns dictionary
print("\n📊 Loading strategy mapping...")
rocksling_returns = {'Rocksling strategy': {0: 'Balanced',
  1: 'Buyout',
  2: 'Co-investment',
  3: 'Co-Investment Multi-Manager',
  4: 'Direct Lending',
  5: 'Direct Lending',
  6: 'Direct Lending',
  7: 'Direct Lending',
  8: 'Direct Lending',
  9: 'Direct Secondaries',
  10: 'Distressed Debt',
  11: 'Early Stage',
  12: 'Early Stage: Seed',
  13: 'Early Stage: Start-up',
  14: 'Expansion / Late Stage',
  15: 'Fund of Funds',
  16: 'Growth',
  17: 'Infrastructure',
  18: 'Infrastructure',
  19: 'Real Estate Debt',
  20: 'Infrastructure Fund of Funds',
  21: 'Infrastructure Opportunistic',
  22: 'Infrastructure Secondaries',
  23: 'Infrastructure',
  24: 'Mezzanine',
  25: 'Natural Resources',
  26: 'Private Debt Fund of Funds',
  27: 'Real Estate Co-Investment',
  28: 'Real Estate Core',
  29: 'Real Estate Core-Plus',
  30: 'Real Estate Debt',
  31: 'Real Estate Distressed',
  32: 'Real Estate Fund of Funds',
  33: 'Real Estate Opportunistic',
  34: 'Real Estate Secondaries',
  35: 'Real Estate Value Added',
  36: 'Secondaries',
  37: 'Special Situations',
  38: 'Timber',
  39: 'Turnaround',
  40: 'VC',
  41: 'Venture Debt'},
 'Preqin strategy': {0: 'Balanced',
  1: 'Buyout',
  2: 'Co-Investment',
  3: 'Co-Investment Multi-Manager',
  4: 'Direct Lending',
  5: 'Direct Lending - Blended / Opportunistic Debt',
  6: 'Direct Lending - Junior / Subordinated Debt',
  7: 'Direct Lending - Senior Debt',
  8: 'Direct Lending - Unitranche Debt',
  9: 'Direct Secondaries',
  10: 'Distressed Debt',
  11: 'Early Stage',
  12: 'Early Stage: Seed',
  13: 'Early Stage: Start-up',
  14: 'Expansion / Late Stage',
  15: 'Fund of Funds',
  16: 'Growth',
  17: 'Infrastructure Core',
  18: 'Infrastructure Core Plus',
  19: 'Infrastructure Debt',
  20: 'Infrastructure Fund of Funds',
  21: 'Infrastructure Opportunistic',
  22: 'Infrastructure Secondaries',
  23: 'Infrastructure Value Added',
  24: 'Mezzanine',
  25: 'Natural Resources',
  26: 'Private Debt Fund of Funds',
  27: 'Real Estate Co-Investment',
  28: 'Real Estate Core',
  29: 'Real Estate Core-Plus',
  30: 'Real Estate Debt',
  31: 'Real Estate Distressed',
  32: 'Real Estate Fund of Funds',
  33: 'Real Estate Opportunistic',
  34: 'Real Estate Secondaries',
  35: 'Real Estate Value Added',
  36: 'Secondaries',
  37: 'Special Situations',
  38: 'Timber',
  39: 'Turnaround',
  40: 'Venture (General)',
  41: 'Venture Debt'},
 'Target Net IRR': {0: 0.17193133421332657,
  1: 0.17193133421332657,
  2: 0.1920941688392953,
  3: 0.1920941688392953,
  4: 0.13814902889156994,
  5: 0.13814902889156994,
  6: 0.13814902889156994,
  7: 0.13814902889156994,
  8: 0.13814902889156994,
  9: 0.1613242592088194,
  10: 0.17193133421332657,
  11: 0.20246920158004958,
  12: 0.20246920158004958,
  13: 0.20246920158004958,
  14: 0.20246920158004958,
  15: 0.17193133421332657,
  16: 0.19391517070503603,
  17: 0.13784644227229806,
  18: 0.13784644227229806,
  19: 0.15870079486124036,
  20: 0.13784644227229806,
  21: 0.13784644227229806,
  22: 0.1613242592088194,
  23: 0.13784644227229806,
  24: 0.13814902889156994,
  25: 0.219462850007873,
  26: 0.13814902889156994,
  27: 0.15870079486124036,
  28: 0.15870079486124036,
  29: 0.15870079486124036,
  30: 0.15870079486124036,
  31: 0.15870079486124036,
  32: 0.15870079486124036,
  33: 0.15870079486124036,
  34: 0.1613242592088194,
  35: 0.15870079486124036,
  36: 0.1613242592088194,
  37: 0.17193133421332657,
  38: 0.13784644227229806,
  39: 0.17193133421332657,
  40: 0.20246920158004958,
  41: 0.20246920158004958},
 'Target Net MoIC': {0: 1.6087231197318523,
  1: 1.6087231197318523,
  2: 1.9075525598173824,
  3: 1.9075525598173824,
  4: 1.435901802681836,
  5: 1.435901802681836,
  6: 1.435901802681836,
  7: 1.435901802681836,
  8: 1.435901802681836,
  9: 1.5560609304646806,
  10: 1.6087231197318523,
  11: 1.9507169962824704,
  12: 1.9507169962824704,
  13: 1.9507169962824704,
  14: 1.9507169962824704,
  15: 1.6087231197318523,
  16: 1.8418198272546982,
  17: 1.4997913776471483,
  18: 1.4997913776471483,
  19: 1.5534083018706242,
  20: 1.4997913776471483,
  21: 1.4997913776471483,
  22: 1.5560609304646806,
  23: 1.4997913776471483,
  24: 1.435901802681836,
  25: 2.018400147618794,
  26: 1.435901802681836,
  27: 1.5534083018706242,
  28: 1.5534083018706242,
  29: 1.5534083018706242,
  30: 1.5534083018706242,
  31: 1.5534083018706242,
  32: 1.5534083018706242,
  33: 1.5534083018706242,
  34: 1.5560609304646806,
  35: 1.5534083018706242,
  36: 1.5560609304646806,
  37: 1.6087231197318523,
  38: 1.4997913776471483,
  39: 1.6087231197318523,
  40: 1.9507169962824704,
  41: 1.9507169962824704}}

print(f"✓ Loaded {len(rocksling_returns['Preqin strategy'])} strategy mappings")

# Helper functions
def get_next_quarter_end():
    current_dateTime = datetime.now()
    forward_dateTime = datetime.now() + relativedelta(months=3)
    next_quarter_end = forward_dateTime + QuarterEnd()
    return next_quarter_end

def euclidean_division(x, y):
    quotient = x // y
    remainder = x % y
    return quotient, remainder

def euclidean_page_number(commitment_meta_data_total):
    quotient, remainder = euclidean_division(commitment_meta_data_total, 200)
    if remainder == 0:
        return quotient
    else:
        return quotient + 1

def convert_to_dataframe(data: list) -> pd.DataFrame:
    return pd.DataFrame(data)

# Fetch commitment data
print(f"\n🔍 Fetching commitment data for {preqin_firm_name} (ID: {preqin_firm_id_input})...")
base_url = "https://api.preqin.com"
api_headers = {"Authorization": f"Bearer {access_token}"}
investor_endpoint = f"/api/Investor/commitment/{strategy_lookup}/{preqin_firm_id_input}"
commitment_url = base_url + investor_endpoint

try:
    commitment_response = requests.get(commitment_url, headers=api_headers, timeout=30)

    if commitment_response.status_code == 200:
        commitment_meta_data = commitment_response.json()
        commitment_meta_data_total = commitment_meta_data["meta"]["total"]
        print(f"✓ Found {commitment_meta_data_total} total commitments")
    else:
        print(f"✗ Request failed with status {commitment_response.status_code}")
        print(f"Response: {commitment_response.text}")
        exit(1)
except Exception as e:
    print(f"✗ Error fetching commitment data: {str(e)}")
    exit(1)

# Fetch all pages
page_no = euclidean_page_number(commitment_meta_data_total)
print(f"📄 Fetching {page_no} pages of data...")

commitment_consolidated_data = []
for x in range(1, page_no + 1):
    page_endpoint_x = f"?Page={x}"
    commitment_url_x = base_url + investor_endpoint + page_endpoint_x

    try:
        commitment_response_x = requests.get(commitment_url_x, headers=api_headers, timeout=30)

        if commitment_response_x.status_code == 200:
            commitment_data_x = commitment_response_x.json()
            commitment_consolidated_data.extend(commitment_data_x["data"])
            print(f"  ✓ Page {x}/{page_no} fetched successfully")
        else:
            print(f"  ✗ Page {x} failed with status {commitment_response_x.status_code}")
    except Exception as e:
        print(f"  ✗ Error fetching page {x}: {str(e)}")

commitment_df = convert_to_dataframe(commitment_consolidated_data)
print(f"✓ Successfully retrieved {len(commitment_df)} commitments")

# Analyze performance
print("\n📈 Analyzing performance data...")

investor_fund_commitment_output = []
for i in range(len(commitment_df)):
    commitment = commitment_df.loc[i]
    commitment_info = {
        'fundId': commitment['fundId'],
        'fundName': commitment['fundName'],
        'fundManagerName': commitment['fundManagerName'],
        'committedMn': commitment['committedMn']
    }
    investor_fund_commitment_output.append(commitment_info)

investor_fund_commitment_output = pd.DataFrame(investor_fund_commitment_output)
investor_fund_commitment_output['fundId'] = investor_fund_commitment_output['fundId'].astype(int)

merged_df = pd.merge(investor_fund_commitment_output, fundperformance_df, left_on='fundId', right_on='FUND ID', how='left')

# Drop superfluous columns
merged_df = merged_df.drop(columns=['FUND ID', 'FIRM ID', 'NAME', 'PREQIN QUARTILE RANK', 'STRATEGY', 'ASSET CLASS', 'PRIMARY REGION FOCUS',
'FUND CURRENCY', 'STATUS', 'VINTAGE / INCEPTION YEAR', 'FUND SIZE (USD MN)', 'TARGET SIZE (USD MN)', 'FINAL CLOSE SIZE (USD MN)',
'HARD CAP (USD MN)', 'INITIAL TARGET (USD MN)'])

merged_df['committedMn'] = merged_df['committedMn'].replace('', float(DEFAULT_COMMITMENT_MN)).astype(float)

default_count = (merged_df['committedMn'] == float(DEFAULT_COMMITMENT_MN)).sum()
if default_count > 0:
    print(f"  Note: {default_count} positions used default commitment amount of ${DEFAULT_COMMITMENT_MN}M")

# Calculate CAS figures
merged_df['Paid-in capital'] = merged_df['committedMn'] * (merged_df['CALLED (%)']/100)
merged_df['Distributed Capital'] = merged_df['Paid-in capital'] * (merged_df['DPI (%)']/100)
merged_df['NAV'] = merged_df['Paid-in capital'] * (merged_df['RVPI (%)']/100)
merged_df['Unfunded'] = merged_df['committedMn'] - merged_df['Paid-in capital']
merged_df['Unfunded'] = merged_df['Unfunded'].clip(lower=0)
merged_df['Total exposure'] = merged_df['NAV'] + merged_df['Unfunded']

performance_output = merged_df

# Filter for complete data
performance_output['missing_data'] = performance_output.isnull().any(axis=1).apply(lambda x: 0 if x else 1)
performance_output = performance_output[performance_output['missing_data'] == 1]

print(f"✓ Filtered to {len(performance_output)} positions with complete data")

# Create max fund size dataframe
max_fund_size_df = fundperformance_df[['FUND ID', 'FUND SIZE (USD MN)',
                    'TARGET SIZE (USD MN)', 'FINAL CLOSE SIZE (USD MN)',
                    'HARD CAP (USD MN)', 'INITIAL TARGET (USD MN)']].copy()
max_fund_size_df.loc[:, 'Max Value'] = max_fund_size_df.iloc[:, 1:].max(axis=1)

# Build RockSling output
print("\n🏗️  Building RockSling output...")

def merge_and_rename(base_df, source_df, columns_to_merge, new_column_name, merge_on='FUND ID'):
    result = pd.merge(base_df, source_df[columns_to_merge], on=merge_on, how='left').fillna('')
    old_column = [col for col in columns_to_merge if col != merge_on][0]
    if old_column != new_column_name:
        result.rename(columns={old_column: new_column_name}, inplace=True)
    return result

rocksling_output_df = pd.DataFrame()
rocksling_output_df["FUND ID"] = performance_output["fundId"]
rocksling_output_df["Fund"] = performance_output["fundName"].values
rocksling_output_df["Fund Manager"] = performance_output["fundManagerName"].values

rocksling_output_df = merge_and_rename(rocksling_output_df, fundperformance_df, ['FUND ID', 'FUND CURRENCY'], 'Fund Currency')
rocksling_output_df['Fund Currency'] = rocksling_output_df['Fund Currency'].replace('', 'USD')

rocksling_output_df = merge_and_rename(rocksling_output_df, fundperformance_df, ['FUND ID', 'STRATEGY'], 'Fund Strategy')
rocksling_output_df = merge_and_rename(rocksling_output_df, fundperformance_df, ['FUND ID', 'VINTAGE / INCEPTION YEAR'], 'Fund Vintage')
rocksling_output_df = merge_and_rename(rocksling_output_df, fundperformance_df, ['FUND ID', 'PRIMARY REGION FOCUS'], 'Fund Region')
rocksling_output_df = merge_and_rename(rocksling_output_df, max_fund_size_df, ['FUND ID', 'Max Value'], 'Fund Size')

# Merge position data
for col_old, col_new in [('committedMn', 'Position Commitment'), ('NAV', 'Position NAV'),
                          ('Unfunded', 'Position Unfunded'), ('Paid-in capital', 'Position Contributions'),
                          ('Distributed Capital', 'Position Distributions')]:
    rocksling_output_df = pd.merge(rocksling_output_df, performance_output[['fundId', col_old]],
                                    left_on='FUND ID', right_on='fundId', how='left').fillna('')
    rocksling_output_df.drop(columns=['fundId'], inplace=True)
    rocksling_output_df.rename(columns={col_old: col_new}, inplace=True)

rocksling_output_df["Expected Final TVPI"] = [""] * len(rocksling_output_df)
rocksling_output_df["Position Post-accounts Contributions"] = [""] * len(rocksling_output_df)
rocksling_output_df["Position Post-accounts Distributions"] = [""] * len(rocksling_output_df)
rocksling_output_df["Input Currency"] = ["USD"] * len(rocksling_output_df)
rocksling_output_df["Subportfolio"] = rocksling_output_df["Fund"]

print(f"✓ Created output with {len(rocksling_output_df)} positions")

# Display summary statistics
print("\n" + "="*80)
print("PORTFOLIO SUMMARY STATISTICS")
print("="*80)

print(f"\n📊 PORTFOLIO OVERVIEW:")
print(f"  Total Positions: {len(rocksling_output_df)}")
print(f"  Unique Fund Managers: {rocksling_output_df['Fund Manager'].nunique()}")

print(f"\n💰 FINANCIAL SUMMARY (USD Millions):")
print(f"  Total Commitment: ${rocksling_output_df['Position Commitment'].sum():,.2f}M")
print(f"  Total NAV: ${rocksling_output_df['Position NAV'].sum():,.2f}M")
print(f"  Total Unfunded: ${rocksling_output_df['Position Unfunded'].sum():,.2f}M")
print(f"  Total Contributions: ${rocksling_output_df['Position Contributions'].sum():,.2f}M")
print(f"  Total Distributions: ${rocksling_output_df['Position Distributions'].sum():,.2f}M")

print(f"\n📈 TOP 10 STRATEGIES:")
strategy_counts = rocksling_output_df['Fund Strategy'].value_counts()
for strategy, count in strategy_counts.head(10).items():
    if strategy:
        print(f"  {strategy}: {count} positions")

print(f"\n🌍 TOP 5 REGIONS:")
region_counts = rocksling_output_df['Fund Region'].value_counts()
for region, count in region_counts.head(5).items():
    if region:
        print(f"  {region}: {count} positions")

print(f"\n⚠️  DATA QUALITY:")
print(f"  Positions with missing strategy: {(rocksling_output_df['Fund Strategy'] == '').sum()}")
print(f"  Positions with missing vintage: {(rocksling_output_df['Fund Vintage'] == '').sum()}")
print(f"  Positions with missing fund size: {(rocksling_output_df['Fund Size'] == '').sum()}")

print("\n" + "="*80)
print("✅ TEST COMPLETED SUCCESSFULLY!")
print("="*80)
