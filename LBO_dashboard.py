# -*- coding: utf-8 -*-
"""
Created on Tue Jan 16 19:39:00 2024

@author: chris
"""

import streamlit as st
import numpy as np
from scipy.optimize import minimize, NonlinearConstraint
from scipy.optimize import root
import numpy as np
import LBO_calculator as lbo
import LBO_visualization as lboviz

# ============================================================================
# CONFIGURATION & CONSTANTS
# ============================================================================

DEFAULT_EXIT_YEARS = 5
DEFAULT_CURRENT_SALES = 100
DEFAULT_SALES_GROWTH = 5.0
DEFAULT_OPERATING_MARGIN = 5.0
DEFAULT_DEPRECIATION_RATE = 5.0
DEFAULT_CAPEX_RATE = 5.0
DEFAULT_WORKING_CAPITAL = 5.0
DEFAULT_TAX_RATE = 15.0
DEFAULT_INTEREST_RATE = 5.0

DEFAULT_DSCR_COVENANT = 1.25
DEFAULT_ICR_COVENANT = 4.5
DEBT_REPAYMENT_RATIO = 0.30
DEBT_REPAYMENT_MINIMUM = 0.1


st.set_page_config(
    page_title="LBO Financial Modelling Tool - MSc Finance GEM 2026",
    page_icon="🧊",
    layout="wide")

st.title("LBO Financial Modelling")

# ============================================================================
# SIDEBAR: USER INPUTS
# ============================================================================

st.sidebar.title("Cash Flow Parameters")

# Transaction & timing parameters
time = st.sidebar.slider(
    'Exit time (in years from now)',
    value=DEFAULT_EXIT_YEARS,
    min_value=1,
    max_value=15,
    step=1
)

# Income statement parameters
current_sales = st.sidebar.number_input(
    'Current sales ($M)',
    value=DEFAULT_CURRENT_SALES
)

sales_growth_rate = st.sidebar.slider(
    'Sales Growth (%)',
    value=DEFAULT_SALES_GROWTH,
    min_value=0.0,
    max_value=30.0,
    step=0.5
) / 100

operating_margin = st.sidebar.slider(
    'Operating Margin (%)',
    value=DEFAULT_OPERATING_MARGIN,
    min_value=0.0,
    max_value=50.0,
    step=0.5
) / 100

# Balance sheet parameters
depreciation_rate = st.sidebar.slider(
    'Depreciation (% of sales)',
    value=DEFAULT_DEPRECIATION_RATE,
    min_value=0.0,
    max_value=10.0,
    step=0.5
) / 100

capex = st.sidebar.slider(
    'CapEx (% of sales)',
    value=DEFAULT_CAPEX_RATE,
    min_value=0.0,
    max_value=10.0,
    step=0.5
) / 100

working_capital = st.sidebar.slider(
    'Working Capital (%)',
    value=DEFAULT_WORKING_CAPITAL,
    min_value=0.0,
    max_value=50.0,
    step=0.5
) / 100

# Financing parameters
tax_rate = st.sidebar.number_input(
    'Effective tax rate (%)',
    min_value=0.0,
    max_value=100.0,
    value=DEFAULT_TAX_RATE,
    step=0.5
) / 100

interest_rate = st.sidebar.number_input(
    'Interest rate (%)',
    min_value=0.0,
    max_value=30.0,
    value=DEFAULT_INTEREST_RATE,
    step=0.5
) / 100

# ============================================================================
# CASH FLOW PROJECTIONS
# ============================================================================

st.header("Cash Flow Projections")

table_cf = lbo.generate_cash_flow_table(
    time=time,
    curr_sales=current_sales,
    sales_rate=sales_growth_rate,
    ebit_sales_rate=operating_margin,
    tax_rate=tax_rate,
    depreciation_rate=depreciation_rate,
    capex_rate=capex,
    wcr_rate=working_capital
)

# Format and display cash flow table
numeric_cols = table_cf.select_dtypes(include=np.number).columns
table_cf[numeric_cols] = table_cf[numeric_cols].round(decimals=2)
st.dataframe(table_cf)

# Extract key arrays for optimization
fcf_array = np.array(table_cf.loc['Free Cash Flows', :])
ebitda_array = np.array(table_cf.loc['EBITDA', :])

# ============================================================================
# DEBT OPTIMIZATION SETUP
# ============================================================================

st.header("Debt Modelling")

col1, col2 = st.columns(2)
with col1:
    deal_value = st.number_input(
        'Deal value ($M)',
        min_value=0.0,
        max_value=1500.0,
        value=None,
        step=10.0
    )

with col2:
    current_ebit = st.number_input(
        'Current EBIT ($M)',
        min_value=0.0,
        max_value=1500.0,
        value=None,
        step=10.0
    )

st.markdown("### Select Financial Covenants")

# Initialize optimization parameters
# Seed the optimizer with a sensible starting repayment schedule:
# each year repays 30% of that year's FCF, floored at $0.1M to avoid zero-repayment years
# that could cause numerical issues in coverage ratio calculations.
initial_guess = np.maximum(fcf_array * DEBT_REPAYMENT_RATIO, DEBT_REPAYMENT_MINIMUM)

# Constrain each year's repayment to be non-negative (debt cannot be drawn down again).
# None as the upper bound means no per-year cap — the optimizer can repay as much as it likes.
bounds = [(0, None) for _ in range(time)]
list_constraints = []

# DSCR Covenant
dscr_limit = 0.0
covenant_dscr_selected = st.checkbox('Debt Service Coverage Ratio (DSCR)')
if covenant_dscr_selected:
    dscr_limit = st.number_input(
        'Minimum DSCR',
        min_value=0.0,
        max_value=5.0,
        value=DEFAULT_DSCR_COVENANT,
        step=0.01
    )
    dscr_constraint = NonlinearConstraint(
        fun=lambda x: lbo.dscr(x, fcf_array, interest_rate, tax_rate),
        lb=dscr_limit,
        ub=np.inf
    )
    list_constraints.append(dscr_constraint)

# Interest Coverage Ratio Covenant
icr_limit = 0.0
icr_selected = st.checkbox('Interest Coverage Ratio (EBITDA / Cash Interest)')
if icr_selected:
    icr_limit = st.number_input(
        'Minimum Interest Coverage',
        min_value=0.0,
        max_value=10.0,
        value=DEFAULT_ICR_COVENANT,
        step=0.01
    )
    icr_constraint = NonlinearConstraint(
        fun=lambda x: lbo.interest_coverage(x, ebitda_array, interest_rate),
        lb=icr_limit,
        ub=np.inf
    )
    list_constraints.append(icr_constraint)

# Maximum Debt Covenant
debt_limit = deal_value
max_debt_selected = st.checkbox('Maximum Debt Covenant')
if max_debt_selected:
    debt_limit = st.number_input(
        'Maximum debt allowed ($M)',
        min_value=0.0,
        max_value=15000.0,
        value=deal_value,
        step=50.0
    )

# This constraint is always active: total debt must be non-negative and cannot exceed
# the deal value (or the user-specified cap if the Maximum Debt Covenant is selected).
# When no maximum debt covenant is chosen, debt_limit defaults to the full deal value,
# ensuring the model never suggests borrowing more than the company is worth.
max_debt_constraint = NonlinearConstraint(
    fun=lbo.get_total_debt_from_repayments,
    lb=0.0,
    ub=debt_limit
)
list_constraints.append(max_debt_constraint)


# ============================================================================
# OPTIMIZATION EXECUTION
# ============================================================================

if st.button('Run Optimization'):
    
    # Capture intermediate iterations for visualization
    intermediate_solutions = []
    
    # Run debt optimization
    result = minimize(
        fun=lbo.obj,
        x0=initial_guess,
        bounds=bounds,
        constraints=list_constraints,
        callback=lambda x: intermediate_solutions.append(x.copy()) # Store intermediate solutions so we can visualize them later
    )
    
    st.info(f'Optimization Result: {result.message}')
    
    # Build comprehensive debt amortization table
    debt_table_recap = lbo.debt_table(
        fcf_table=fcf_array,
        ebitda_table=ebitda_array,
        repayment_vector=np.array(result.x),
        interest_rate=interest_rate,
        tax_rate=tax_rate
    )
    
    # Format and display debt table
    numeric_cols = debt_table_recap.select_dtypes(include=np.number).columns
    debt_table_recap[numeric_cols] = debt_table_recap[numeric_cols].round(decimals=2)
    
    st.subheader("Debt Amortization Schedule")
    st.dataframe(debt_table_recap)
    
    # ============================================================================
    # IRR CALCULATION
    # ============================================================================
    st.header("Return Analysis")

    # IRR represents the equity return from the LBO transaction
    # It is the discount rate that makes NPV of equity cash flows = 0

    # The objective function returns the negative of total debt (for minimization),
    # so negate it to recover the actual optimal debt amount.
    optimal_debt = -result['fun']
    # Equity invested = whatever the PE sponsor contributes on top of the debt financing
    initial_equity_investment = deal_value - optimal_debt

    # Equity cash flows capture only the sponsor's perspective:
    # cash out at entry, cash back at exit. No intermediate dividends are modelled.
    equity_cash_flows = np.zeros(time)

    # Year 0: Initial equity investment (negative = cash outflow)
    equity_cash_flows[0] = -initial_equity_investment

    # Year N: Exit proceeds (positive = cash inflow)
    # The exit enterprise value is computed by applying the same EV/EBIT multiple
    # paid at entry (deal_value / current_ebit) to the final projected year's EBIT.
    # Equity value = Enterprise Value − remaining debt still outstanding at exit.
    exit_ebit = np.array(table_cf.loc['EBIT', :])[-1]
    exit_enterprise_value = exit_ebit * (deal_value / current_ebit)
    remaining_debt = np.array(debt_table_recap.loc[:, 'Debt SoY'])[-1]
    equity_cash_flows[-1] = exit_enterprise_value - remaining_debt

    # IRR is the discount rate at which the NPV of equity cash flows equals zero.
    # scipy.optimize.root finds this rate numerically, starting from the interest rate
    # as an initial guess (a reasonable lower bound for equity returns in an LBO).
    irr_result = root(
        fun=lbo.calculate_irr,
        x0=interest_rate,
        args=(equity_cash_flows,)
    )

    irr = irr_result.x[0]

    # Display results
    st.subheader("IRR Analysis")
    col1, col2 = st.columns(2)

    with col1:
        st.metric("Equity IRR", f"{irr:.2%}")

    with col2:
        st.metric("Optimal Debt", f"${optimal_debt:.2f}M")
        

    lboviz.visualize_optimization(result, debt_table_recap, fcf_array, ebitda_array, time,
                  interest_rate, tax_rate, dscr_limit, icr_limit, debt_limit, intermediate_solutions)
    
else:
    st.info('Click "Run Optimization" to execute the debt optimization model.')


