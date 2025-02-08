"""
Tests for the core Monte Carlo simulation engine.
"""
import numpy as np
import pytest
from src.core.simulation import MonteCarloEngine, SimulationConfig

def test_basic_simulation():
    """Test basic Monte Carlo simulation with normal distribution."""
    config = SimulationConfig(iterations=10000, seed=42)
    engine = MonteCarloEngine(config)
    
    # Define a simple model that adds two normal distributions
    def model(x1, x2):
        return x1 + x2
    
    # Define input distributions
    mu1, sigma1 = 0, 1
    mu2, sigma2 = 2, 1.5
    
    input_distributions = {
        'x1': lambda u: np.sqrt(2) * sigma1 * np.erfinv(2*u - 1) + mu1,
        'x2': lambda u: np.sqrt(2) * sigma2 * np.erfinv(2*u - 1) + mu2
    }
    
    # Run simulation
    results = engine.run_simulation(model, input_distributions)
    
    # Check statistical properties
    assert abs(np.mean(results) - (mu1 + mu2)) < 0.1
    assert abs(np.std(results) - np.sqrt(sigma1**2 + sigma2**2)) < 0.1

def test_latin_hypercube_sampling():
    """Test simulation with Latin Hypercube sampling."""
    config = SimulationConfig(iterations=10000, seed=42)
    engine = MonteCarloEngine(config)
    
    def model(x):
        return x
    
    input_distributions = {
        'x': lambda u: u  # Identity function for uniform distribution
    }
    
    # Run simulation with LHS
    results_lhs = engine.run_simulation(model, input_distributions, use_lhs=True)
    
    # Run simulation with standard sampling
    results_std = engine.run_simulation(model, input_distributions, use_lhs=False)
    
    # LHS should provide more uniform coverage
    bins = np.linspace(0, 1, 21)
    hist_lhs, _ = np.histogram(results_lhs, bins)
    hist_std, _ = np.histogram(results_std, bins)
    
    # Check that LHS has more uniform distribution (lower variance in bin counts)
    assert np.var(hist_lhs) < np.var(hist_std)

def test_correlated_inputs():
    """Test simulation with correlated input variables."""
    config = SimulationConfig(iterations=10000, seed=42)
    engine = MonteCarloEngine(config)
    
    def model(x1, x2):
        return np.column_stack([x1, x2])
    
    # Define correlation matrix
    correlation = np.array([[1.0, 0.7],
                          [0.7, 1.0]])
    
    input_distributions = {
        'x1': lambda u: u,  # Uniform(0,1)
        'x2': lambda u: u   # Uniform(0,1)
    }
    
    # Run simulation with correlation
    results = engine.run_simulation(
        model,
        input_distributions,
        correlation_matrix=correlation
    )
    
    # Check correlation coefficient
    computed_corr = np.corrcoef(results[:, 0], results[:, 1])[0, 1]
    assert abs(computed_corr - 0.7) < 0.1

def test_stopping_criteria():
    """Test early stopping criteria."""
    def stopping_criteria(results):
        # Stop if mean converges within tolerance
        if len(results) < 1000:
            return False
        return abs(np.mean(results[-1000:]) - np.mean(results)) < 0.01
    
    config = SimulationConfig(
        iterations=10000,
        seed=42,
        stopping_criteria=stopping_criteria
    )
    engine = MonteCarloEngine(config)
    
    def model(x):
        return x
    
    input_distributions = {
        'x': lambda u: u  # Uniform(0,1)
    }
    
    results = engine.run_simulation(model, input_distributions)
    
    # Check that simulation stopped early
    assert len(results) < 10000
