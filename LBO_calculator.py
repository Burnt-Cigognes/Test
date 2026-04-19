# -*- coding: utf-8 -*-
"""
Created on Tue Jan 16 19:07:43 2024

@author: chris
"""

import numpy as np
import pandas as pd


def generate_cash_flow_table(time: int,
                             curr_sales: float,
                             sales_rate: float,
                             ebit_sales_rate: float,
                             tax_rate: float,
                             depreciation_rate: float,
                             capex_rate: float,
                             wcr_rate: float) -> pd.DataFrame:
    """
    Projects income statement and cash flow line items over a holding period.

    Starting from current sales, each line item is derived as a fixed percentage
    of projected sales for that year. Working capital changes are computed as the
    year-on-year difference in the WCR balance, reflecting cash consumed or
    released as the business grows.

    Parameters:
    -----------
    time : int
        Number of years to project (i.e. the planned exit horizon)
    curr_sales : float
        Current (Year 0) revenue, used as the compounding base ($M)
    sales_rate : float
        Annual sales growth rate (e.g., 0.05 for 5%)
    ebit_sales_rate : float
        EBIT as a proportion of sales, i.e. the operating margin (e.g., 0.10 for 10%)
    tax_rate : float
        Effective corporate tax rate applied to EBIT (e.g., 0.25 for 25%)
    depreciation_rate : float
        Depreciation as a proportion of sales (e.g., 0.03 for 3%)
    capex_rate : float
        Capital expenditure as a proportion of sales (e.g., 0.04 for 4%)
    wcr_rate : float
        Working Capital Requirement as a proportion of sales (e.g., 0.05 for 5%)

    Returns:
    --------
    pd.DataFrame
        DataFrame with line items as rows and projection years (1 to time) as columns.
        Rows: Sales Projections, EBIT, Taxes, Net Income, Depreciation, EBITDA,
              Capex Projections, Chg in Working Capital, Free Cash Flows
    """
    # Step 0 - Create empty numpy arrays to store data
    time_table = np.arange(start=1,
                            stop = time + 1,
                            step=1
                            )
    
    # Step 1 - Sales Projection table
    sales_table = np.ones(time) * curr_sales * (1 + sales_rate) ** time_table
    ebit_table = sales_table * ebit_sales_rate
    tax_table = - ebit_table * tax_rate
    net_income_table = ebit_table + tax_table
    depreciation_table = sales_table * depreciation_rate
    capex_table = - sales_table * capex_rate
    
    current_wcr = curr_sales * wcr_rate
    wcr_table = sales_table * wcr_rate
    
    # We add the current WCR (year 0) to the projected WCR in the future using concatenation
    wcr_table = np.r_[current_wcr, wcr_table] 
    chg_wcr_table = -np.diff(wcr_table)
    
    free_cf_table = net_income_table + depreciation_table + capex_table + chg_wcr_table
    
    ebitda_table = ebit_table + depreciation_table
    
    # Step 2 - Create DataFrame with all projections
    df_projections = pd.DataFrame({'Sales Projections': sales_table,
                                   'EBIT': ebit_table,
                                   'Taxes': tax_table,
                                   'Net Income': net_income_table,
                                   'Depreciation': depreciation_table,
                                   'EBITDA': ebitda_table,
                                   'Capex Projections': capex_table,
                                   'Chg in Working Capital': chg_wcr_table,
                                   'Free Cash Flows': free_cf_table}).T
    #df_projections.columns = ['Year ' + str(i) for i in range(1, time+1)]
    df_projections.columns = list(time_table)
    
    return df_projections
    
    
def obj(debt_repayments: np.ndarray) -> float:
    """
    Objective function for optimization - maximizes total debt amount (sum of repayments).
    
    Parameters:
    -----------
    debt_repayments : np.ndarray
        1D array of debt repayment amounts for each year
    
    Returns:
    --------
    float
        Negative sum of debt repayments (for minimization in scipy.optimize)
    """
    
    return -get_total_debt_from_repayments(debt_repayments)

def get_total_debt_from_repayments(debt_repayments: np.ndarray) -> float:
    """
    Returns the total initial debt amount implied by a repayment schedule.

    Parameters:
    -----------
    debt_repayments : np.ndarray
        1D array of debt repayment amounts for each year

    Returns:
    --------
    float
        Total debt principal (sum of all repayments)
    """
    return np.sum(debt_repayments)

def debt_remaining_vector(debt_repayments: np.ndarray) -> np.ndarray:
    """
    Computes the outstanding debt balance at the start of each year.

    Works like a mortgage: total debt is borrowed upfront, then reduced by each
    year's repayment. The balance at the start of year t equals total debt minus
    all repayments made in prior years.

    Parameters:
    -----------
    debt_repayments : np.ndarray
        1D array of debt repayment amounts for each year

    Returns:
    --------
    np.ndarray
        1D array of outstanding debt balances at the start of each year
    """
    # Cumulative sum of repayments made up to and including each year
    debt_repayments_cs = np.cumsum(debt_repayments)

    # Debt SoY = Total debt − repayments made before this year
    # Adding back debt_repayments reverses the current year's repayment,
    # which was included in the cumulative sum but not yet paid at year start
    
    # Example: 50 repayment per year for 4 years= 200 total debt overall
    # repayment vector = [50, 50, 50, 50]
    # cumulative repayment vector = [50, 100, 150, 200]
    # remaining debt = [200 - 50 + 50 = 200, 200 - 100 + 50 = 150, 200 - 150 + 50 = 100, 200 - 200 + 50 = 50]
    
    return get_total_debt_from_repayments(debt_repayments) - debt_repayments_cs + debt_repayments

def dscr(debt_repayments: np.ndarray,
         fcf_vector: np.ndarray,
         interest_rate: float,
         tax_rate: float) -> np.ndarray:
    
    vector_debt = debt_remaining_vector(debt_repayments)
    interest_vector = vector_debt * interest_rate
    tax_shield_vector = interest_vector * tax_rate
    
    debt_service_vector = debt_repayments + interest_vector - tax_shield_vector
    
    return fcf_vector / debt_service_vector

def interest_coverage(debt_repayments: np.ndarray,
                      ebitda_vector: np.ndarray,
                      interest_rate: float) -> np.ndarray:
    """
    Calculates Interest Coverage Ratio (EBITDA/Cash Interest) for each year.
    
    Parameters:
    -----------
    debt_repayments : np.ndarray
        1D array of debt repayment amounts for each year
    ebitda_vector : np.ndarray
        1D array of EBITDA values for each year
    interest_rate : float
        Annual interest rate on debt (e.g., 0.08 for 8%)
    
    Returns:
    --------
    np.ndarray
        1D array of Interest Coverage Ratio values for each year
    """
    
    vector_debt = debt_remaining_vector(debt_repayments)
    interest_vector = vector_debt * interest_rate
    
    return ebitda_vector / interest_vector




def debt_table(fcf_table: np.ndarray,
               ebitda_table: np.ndarray,
               repayment_vector: np.ndarray,
               interest_rate: float,
               tax_rate: float) -> pd.DataFrame:
    """
    Generates a detailed debt amortization and coverage metrics table.
    
    Parameters:
    -----------
    fcf_table : np.ndarray
        1D array of free cash flows for each year
    ebitda_table : np.ndarray
        1D array of EBITDA values for each year
    repayment_vector : np.ndarray
        1D array of debt repayment amounts for each year
    interest_rate : float
        Annual interest rate on debt (e.g., 0.08 for 8%)
    tax_rate : float
        Corporate tax rate for tax shield calculation
    
    Returns:
    --------
    pd.DataFrame
        DataFrame with years as index and columns:
        FCF, EBITDA, Debt Repayment, Cumulative Repayment, Debt SoY, Debt EoY,
        Interest, Tax Shield, Debt Servicing, Cash, Interest Coverage, DSCR
    """
    
    n_year = repayment_vector.shape[0]
    df = pd.DataFrame()
    
    df['Year'] = np.arange(1, n_year + 1)
    df = df.set_index('Year')
    
    df['FCF'] = fcf_table
    df['EBITDA'] = ebitda_table
    df['Debt Repayment'] = repayment_vector
    df['Cumulative Repayment'] = np.cumsum(repayment_vector)
    df['Debt SoY'] = np.sum(repayment_vector) - df['Cumulative Repayment'] + df['Debt Repayment']
    df['Debt EoY'] = df['Debt SoY'].shift(-1)
    df.loc[n_year, 'Debt EoY'] = 0
    
    df['Interest'] = df['Debt SoY'] * interest_rate
    df['Tax Shield'] = df['Interest'] * tax_rate
    df['Debt Servicing'] = df['Debt Repayment'] + df['Interest'] - df['Tax Shield']
    df['Cash'] = (df['FCF'] - df['Debt Servicing']).cumsum()
    
    df['Interest Coverage'] = df['EBITDA'] / df['Interest']
    df['DSCR'] = df['FCF'] / df['Debt Servicing']
                                                                      
                                                                      
    df = df[['FCF', 'EBITDA', 'Debt SoY', 'Debt Repayment', 
             'Cumulative Repayment', 'Interest', 'Tax Shield', 
             'Debt Servicing', 'Cash', 'Debt EoY', 'Interest Coverage', 'DSCR']]
    
    return df

def calculate_irr(rate: float, cashflow_array: np.ndarray) -> float:
    """
    Calculates the Net Present Value (NPV) of a cash flow stream at a given discount rate.
    Used to find IRR via root-finding algorithm (IRR is the rate where NPV = 0).
    
    Parameters:
    -----------
    rate : float
        Discount rate to apply (e.g., 0.10 for 10%)
    cashflow_array : np.ndarray
        1D array of cash flows from year 0 to n (first element is typically negative for initial investment)
    
    Returns:
    --------
    float
        Net Present Value of the cash flow stream at the given rate
    """
    
    time_array = np.arange(cashflow_array.shape[0])
    discount_array = (1 + rate) ** -time_array
    
    cf_discount_array = np.sum(cashflow_array * discount_array)
    
    return cf_discount_array
