import numpy as np

import matplotlib.pyplot as plt
import streamlit as st

import LBO_calculator as lbo

def visualize_optimization(minimization_result, debt_table_recap, fcf_array, ebitda_array, time,
                           interest_rate, tax_rate, dscr_limit, icr_limit, debt_limit, intermediate_solutions):
    # ============================================================================
    # CONSTRAINT ANALYSIS & VISUALIZATION
    # ============================================================================
    st.header("Constraint Analysis & Visualization")
    
    # Extract data from debt_table_recap
    optimal_repayments = np.array(minimization_result.x)
    total_debt = lbo.get_total_debt_from_repayments(optimal_repayments)
    dscr_values = np.array(debt_table_recap.loc[:, 'DSCR'])
    interest_cov = np.array(debt_table_recap.loc[:, 'Interest Coverage'])
    
    # Create optimal solution visualization
    fig_optimal = visualize_optimal_solution(
        fcf_array=fcf_array,
        optimal_repayments=optimal_repayments,
        dscr_values=dscr_values,
        interest_cov=interest_cov,
        time=time,
        dscr_limit=dscr_limit,
        icr_limit=icr_limit,
        debt_limit=debt_limit
    )
    st.pyplot(fig_optimal)
    
    # Summary of constraint satisfaction
    st.subheader("Constraint Satisfaction Summary")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        dscr_satisfied = np.all(dscr_values >= dscr_limit)
        st.metric("DSCR", f"{dscr_values.min():.2f}x", f"Min: {dscr_limit:.2f}", 
                 delta_color="inverse" if not dscr_satisfied else "off")
    
    with col2:
        icr_satisfied = np.all(interest_cov >= icr_limit)
        st.metric("Interest Coverage", f"{interest_cov.min():.2f}x", f"Min: {icr_limit:.2f}",
                 delta_color="inverse" if not icr_satisfied else "off")
    
    with col3:
        debt_satisfied = round(total_debt, 2) <= debt_limit
        st.metric("Total Debt", f"${total_debt:.2f}M", f"Max: ${debt_limit:.2f}M",
                 delta_color="inverse" if not debt_satisfied else "off")
    
    # ============================================================================
    # OPTIMIZATION CONVERGENCE ANALYSIS
    # ============================================================================
    st.header("Optimization Convergence")
    st.write(f"Total iterations: {len(intermediate_solutions)}")
    
    # Create optimization progress visualization
    fig_progress, iteration_data = visualize_optimization_progress(
        intermediate_solutions=intermediate_solutions,
        fcf_array=fcf_array,
        ebitda_array=ebitda_array,
        time=time,
        interest_rate=interest_rate,
        tax_rate=tax_rate,
        dscr_limit=dscr_limit,
        icr_limit=icr_limit,
        debt_limit=debt_limit
    )
    st.pyplot(fig_progress)
    
    # Summary table of iterations
    st.subheader("Iteration Summary")
    st.table(iteration_data)


def visualize_optimal_solution(fcf_array, optimal_repayments, dscr_values, interest_cov, 
                               time, dscr_limit=0, icr_limit=0, debt_limit=None):
    """
    Creates a 2x2 visualization of the optimal debt solution and constraint satisfaction.
    
    Parameters:
    -----------
    fcf_array : np.ndarray
        1D array of free cash flows for each year
    optimal_repayments : np.ndarray
        1D array of optimal debt repayment amounts
    dscr_values : np.ndarray
        1D array of DSCR values for each year
    interest_cov : np.ndarray
        1D array of Interest Coverage Ratio values for each year
    time : int
        Number of years in projection
    dscr_limit : float
        Minimum required DSCR
    cov_ratio_value : float
        Minimum required Interest Coverage
    debt_limit : float
        Maximum allowed debt
    
    Returns:
    --------
    matplotlib.figure.Figure
        Figure containing the 2x2 constraint visualization
    """
    
    years = np.arange(1, time + 1)
    total_debt = lbo.get_total_debt_from_repayments(optimal_repayments)
    
    fig, axes = plt.subplots(2, 2, figsize=(14, 10))
    fig.suptitle('Optimal Debt Solution - Constraint Satisfaction Analysis', 
                 fontsize=14, fontweight='bold')
    
    # Plot 1: Debt Repayment Schedule
    ax1 = axes[0, 0]
    initial = fcf_array * 0.30
    ax1.bar(years - 0.2, initial, 0.4, label='Initial Guess', alpha=0.5, color='lightblue')
    ax1.bar(years + 0.2, optimal_repayments, 0.4, label='Optimal Solution', alpha=0.8, color='darkgreen')
    ax1.set_xlabel('Year')
    ax1.set_ylabel('Debt Repayment ($M)')
    ax1.set_title('Debt Repayment Schedule: Initial vs Optimal')
    ax1.legend()
    ax1.grid(axis='y', alpha=0.3)
    
    # Plot 2: DSCR Constraint
    ax2 = axes[0, 1]
    color_dscr = 'green' if np.all(np.round(dscr_values, 2) >= dscr_limit) else 'orange'
    ax2.plot(years, dscr_values, 'o-', linewidth=2.5, markersize=8, color=color_dscr, label='DSCR')
    ax2.axhline(y=dscr_limit, color='r', linestyle='--', linewidth=2, label=f'Min DSCR ({dscr_limit})')
    ax2.fill_between(years, dscr_limit, dscr_values.max() + 0.2, alpha=0.15, color='green')
    
    violations = np.round(dscr_values, 2) < dscr_limit
    if np.any(violations):
        ax2.scatter(years[violations], dscr_values[violations], s=200, color='red', 
                   marker='X', zorder=5, label='Violation')
    
    ax2.set_xlabel('Year')
    ax2.set_ylabel('DSCR Ratio')
    ax2.set_title(f'DSCR Constraint (Active: {dscr_limit != 0})')
    ax2.legend()
    ax2.grid(alpha=0.3)
    
    # Plot 3: Interest Coverage Constraint
    ax3 = axes[1, 0]
    color_icr = 'green' if np.all(interest_cov >= icr_limit) else 'orange'
    ax3.plot(years, interest_cov, 'o-', linewidth=2.5, markersize=8, color=color_icr, label='Interest Coverage')
    ax3.axhline(y=icr_limit, color='r', linestyle='--', linewidth=2, label=f'Min Coverage ({icr_limit}x)')
    ax3.fill_between(years, icr_limit, interest_cov.max() + 1, alpha=0.15, color='green')
    
    violations_icr = interest_cov < icr_limit
    if np.any(violations_icr):
        ax3.scatter(years[violations_icr], interest_cov[violations_icr], s=200, color='red',
                   marker='X', zorder=5, label='Violation')
    
    ax3.set_xlabel('Year')
    ax3.set_ylabel('EBITDA / Interest Ratio')
    ax3.set_title(f'Interest Coverage Constraint (Active: {icr_limit != 0})')
    ax3.legend()
    ax3.grid(alpha=0.3)
    
    # Plot 4: Total Debt Constraint
    ax4 = axes[1, 1]
    total_debt_cum = np.cumsum(optimal_repayments)
    
    color_debt = 'green' if total_debt <= debt_limit else 'orange'
    ax4.bar(years, total_debt_cum, alpha=0.8, color=color_debt, label='Cumulative Debt')
    ax4.axhline(y=debt_limit, color='r', linestyle='--', linewidth=2, label=f'Max Debt Limit (${debt_limit:.2f}M)')
    ax4.fill_between(years, 0, debt_limit, alpha=0.1, color='green')
    
    if round(total_debt, 2) > debt_limit:
        ax4.axhspan(debt_limit, total_debt_cum.max(), alpha=0.2, color='red', label='Violation')
    
    ax4.set_xlabel('Year')
    ax4.set_ylabel('Cumulative Debt Outstanding ($M)')
    ax4.set_title(f'Maximum Debt Constraint (Active: {debt_limit})')
    ax4.legend()
    ax4.grid(axis='y', alpha=0.3)
    
    plt.tight_layout()
    return fig


def visualize_optimization_progress(intermediate_solutions, fcf_array, ebitda_array, 
                                    time, interest_rate, tax_rate, dscr_limit=0,
                                    icr_limit=0, debt_limit=None):
    """
    Creates a 3xN visualization showing the optimization convergence across iterations.
    
    Parameters:
    -----------
    intermediate_solutions : list
        List of intermediate debt repayment solutions from optimization callback
    fcf_array : np.ndarray
        1D array of free cash flows for each year
    ebitda_array : np.ndarray
        1D array of EBITDA values for each year
    time : int
        Number of years in projection
    interest_rate : float
        Annual interest rate on debt
    tax_rate : float
        Corporate tax rate
    dscr_limit : float
        Minimum required DSCR
    icr_limit : float
        Minimum required Interest Coverage
    debt_limit : float
        Maximum allowed debt
    
    Returns:
    --------
    matplotlib.figure.Figure
        Figure containing the 3xN optimization progress visualization
    tuple
        Tuple of (figure, iteration_data_list) containing figure and summary data
    """
    import matplotlib.pyplot as plt
    
    years = np.arange(1, time + 1)
    
    # Select iterations to display: first, middle, and last
    iterations_to_show = []
    iterations_to_show.append(0)  # First iteration
    if len(intermediate_solutions) > 2:
        iterations_to_show.append(len(intermediate_solutions) // 2)  # Middle iteration
    iterations_to_show.append(len(intermediate_solutions) - 1)  # Last iteration
    
    # Create subplots for each selected iteration
    n_iterations = len(iterations_to_show)
    fig, axes = plt.subplots(3, n_iterations, figsize=(5*n_iterations, 12))
    if n_iterations == 1:
        axes = axes.reshape(3, 1)
    
    fig.suptitle('Optimization Convergence: Debt Repayment Schedule Evolution', 
                 fontsize=14, fontweight='bold', y=0.995)
    
    iteration_data = []
    
    for col_idx, iter_idx in enumerate(iterations_to_show):
        if iter_idx >= len(intermediate_solutions):
            continue
        
        repayment = intermediate_solutions[iter_idx]
        total = lbo.get_total_debt_from_repayments(repayment)
        
        # Calculate metrics for this iteration
        debt_soy = lbo.debt_remaining_vector(repayment)
        interest_expense = debt_soy * interest_rate
        tax_shield = interest_expense * tax_rate
        debt_service = repayment + interest_expense - tax_shield
        dscr_iter = fcf_array / debt_service
        interest_cov_iter = ebitda_array / interest_expense
        
        # Row 0: Debt Repayment Schedule
        ax = axes[0, col_idx]
        ax.bar(years, repayment, alpha=0.8, color='steelblue')
        ax.set_ylabel('Debt Repayment ($M)' if col_idx == 0 else '')
        ax.set_title(f'Iteration {iter_idx}\nTotal Debt: ${total:.2f}M')
        ax.grid(axis='y', alpha=0.3)
        
        # Row 1: DSCR Constraint
        ax = axes[1, col_idx]
        color_dscr = 'green' if np.all(np.round(dscr_iter, 2) >= dscr_limit) else 'orange'
        ax.plot(years, dscr_iter, 'o-', linewidth=2.5, markersize=6, color=color_dscr)
        ax.axhline(y=dscr_limit, color='r', linestyle='--', linewidth=1.5, label=f'Min: {dscr_limit}')
        ax.fill_between(years, dscr_limit, dscr_iter.max() + 0.2, alpha=0.1, color='green')
        ax.set_ylabel('DSCR Ratio' if col_idx == 0 else '')
        ax.set_ylim([0.5, max(dscr_iter.max() + 0.3, 2)])
        ax.grid(alpha=0.3)
        if col_idx == n_iterations - 1:
            ax.legend(fontsize=8)
        
        # Highlight violations
        violations = np.round(dscr_iter, 2) < dscr_limit
        if np.any(violations):
            ax.scatter(years[violations], dscr_iter[violations], s=100, color='red', 
                      marker='X', zorder=5)
        
        # Row 2: Interest Coverage Constraint
        ax = axes[2, col_idx]
        color_icr = 'green' if np.all(interest_cov_iter >= icr_limit) else 'orange'
        ax.plot(years, interest_cov_iter, 'o-', linewidth=2.5, markersize=6, color=color_icr)
        ax.axhline(y=icr_limit, color='r', linestyle='--', linewidth=1.5, label=f'Min: {icr_limit}')
        ax.fill_between(years, icr_limit, interest_cov_iter.max() + 1, alpha=0.1, color='green')
        ax.set_xlabel('Year')
        ax.set_ylabel('Interest Coverage' if col_idx == 0 else '')
        ax.grid(alpha=0.3)
        if col_idx == n_iterations - 1:
            ax.legend(fontsize=8)
        
        # Highlight violations
        violations_icr = interest_cov_iter < icr_limit
        if np.any(violations_icr):
            ax.scatter(years[violations_icr], interest_cov_iter[violations_icr], s=100, color='red',
                      marker='X', zorder=5)
        
        # Collect iteration data        
        dscr_ok = "✓" if round(dscr_iter.min(), 2) >= dscr_limit else "✗"
        icr_ok = "✓" if round(interest_cov_iter.min(), 2) >= icr_limit else "✗"
        debt_ok = "✓" if round(total, 2) <= debt_limit else "✗"
        
        iteration_data.append({
            'Iteration': iter_idx,
            'Total Debt ($M)': round(total, 2),
            'Min DSCR': round(dscr_iter.min(), 2),
            'DSCR ✓': dscr_ok,
            'Min Interest Cov': round(interest_cov_iter.min(), 2),
            'ICR ✓': icr_ok,
            'Debt ✓': debt_ok
        })
    
    plt.tight_layout()
    return fig, iteration_data
