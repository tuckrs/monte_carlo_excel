"""
Visualization functions for Monte Carlo simulation results.
"""
import numpy as np
import matplotlib.pyplot as plt
from typing import Dict, List, Optional, Tuple
import seaborn as sns

def create_histogram(
    data: np.ndarray,
    title: str = "Simulation Results",
    bins: int = 50,
    figsize: Tuple[int, int] = (10, 6)
) -> plt.Figure:
    """Create a histogram of simulation results.
    
    Args:
        data: Simulation results
        title: Plot title
        bins: Number of histogram bins
        figsize: Figure size
        
    Returns:
        Matplotlib figure
    """
    fig, ax = plt.subplots(figsize=figsize)
    
    # Plot histogram with kernel density estimate
    sns.histplot(data, bins=bins, kde=True, ax=ax)
    
    # Add percentile lines
    percentiles = [5, 50, 95]
    colors = ['r', 'g', 'r']
    labels = ['5th', 'Median', '95th']
    
    for p, c, l in zip(percentiles, colors, labels):
        val = np.percentile(data, p)
        ax.axvline(val, color=c, linestyle='--', label=f'{l}: {val:.2f}')
    
    ax.set_title(title)
    ax.set_xlabel("Value")
    ax.set_ylabel("Frequency")
    ax.legend()
    
    return fig

def create_tornado_chart(
    results: np.ndarray,
    inputs: Dict[str, callable],
    figsize: Tuple[int, int] = (10, 6)
) -> plt.Figure:
    """Create a tornado chart for sensitivity analysis.
    
    Args:
        results: Simulation results
        inputs: Dictionary of input distributions
        figsize: Figure size
        
    Returns:
        Matplotlib figure
    """
    # Calculate base case (mean)
    base = np.mean(results)
    
    # Calculate impact of each variable
    impacts = []
    
    for name in inputs.keys():
        # Get percentiles for this input
        low_val = np.percentile(results[results[name] < np.median(results[name])], 50)
        high_val = np.percentile(results[results[name] > np.median(results[name])], 50)
        
        impacts.append({
            'name': name,
            'low': low_val - base,
            'high': high_val - base
        })
    
    # Sort by total impact
    impacts.sort(key=lambda x: abs(x['high'] - x['low']), reverse=True)
    
    # Create tornado chart
    fig, ax = plt.subplots(figsize=figsize)
    
    y_pos = np.arange(len(impacts))
    
    # Plot bars
    for i, impact in enumerate(impacts):
        ax.barh(
            y_pos[i],
            impact['high'] - impact['low'],
            left=min(impact['low'], impact['high']),
            height=0.8,
            color='lightblue'
        )
    
    # Customize plot
    ax.set_yticks(y_pos)
    ax.set_yticklabels([i['name'] for i in impacts])
    ax.axvline(0, color='black', linestyle='-', linewidth=0.5)
    
    ax.set_title("Sensitivity Analysis")
    ax.set_xlabel("Impact on Output")
    
    return fig

def create_scatter_matrix(
    data: Dict[str, np.ndarray],
    figsize: Optional[Tuple[int, int]] = None
) -> plt.Figure:
    """Create a scatter matrix to visualize correlations.
    
    Args:
        data: Dictionary of variable names and their values
        figsize: Optional figure size
        
    Returns:
        Matplotlib figure
    """
    df = pd.DataFrame(data)
    
    if figsize is None:
        figsize = (2 * len(data), 2 * len(data))
    
    fig = sns.pairplot(df, diag_kind="kde")
    fig.fig.set_size_inches(figsize)
    
    return fig.fig
