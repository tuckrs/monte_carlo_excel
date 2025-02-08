"""
Core Monte Carlo simulation engine.
"""
import numpy as np
from typing import Callable, Dict, Optional, Union, List
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from ..utils.sampling import latin_hypercube_sampling

@dataclass
class SimulationConfig:
    """Configuration for Monte Carlo simulation."""
    iterations: int
    seed: Optional[int] = None
    num_threads: int = -1  # -1 means use all available cores
    stopping_criteria: Optional[Callable[[np.ndarray], bool]] = None

class MonteCarloEngine:
    """Core Monte Carlo simulation engine with support for various sampling methods."""
    
    def __init__(self, config: SimulationConfig):
        """Initialize the Monte Carlo simulation engine.
        
        Args:
            config: Simulation configuration parameters
        """
        self.config = config
        self.rng = np.random.default_rng(seed=config.seed)
        self.num_threads = (
            config.num_threads if config.num_threads > 0 
            else max(1, len(os.sched_getaffinity(0)) - 1)
        )
    
    def run_simulation(
        self,
        model: Callable[..., np.ndarray],
        input_distributions: Dict[str, Callable[[], np.ndarray]],
        correlation_matrix: Optional[np.ndarray] = None,
        use_lhs: bool = False
    ) -> np.ndarray:
        """Run Monte Carlo simulation with the specified model and input distributions.
        
        Args:
            model: Function that takes input samples and returns output
            input_distributions: Dictionary mapping variable names to their sampling functions
            correlation_matrix: Optional correlation matrix for input variables
            use_lhs: Whether to use Latin Hypercube Sampling
            
        Returns:
            Array of simulation results
        """
        num_vars = len(input_distributions)
        
        if use_lhs:
            # Generate Latin Hypercube samples
            samples = latin_hypercube_sampling(
                self.config.iterations,
                num_vars,
                self.rng
            )
        else:
            # Generate random samples
            samples = np.random.random((self.config.iterations, num_vars))
            
        if correlation_matrix is not None:
            # Apply correlation structure
            samples = self._apply_correlation(samples, correlation_matrix)
        
        # Transform samples according to input distributions
        input_data = {}
        for idx, (name, dist_func) in enumerate(input_distributions.items()):
            input_data[name] = dist_func(samples[:, idx])
        
        # Run simulation in parallel
        with ThreadPoolExecutor(max_workers=self.num_threads) as executor:
            chunk_size = self.config.iterations // self.num_threads
            futures = []
            
            for i in range(0, self.config.iterations, chunk_size):
                chunk_data = {
                    name: data[i:i + chunk_size]
                    for name, data in input_data.items()
                }
                futures.append(
                    executor.submit(self._run_chunk, model, chunk_data)
                )
            
            results = np.concatenate([f.result() for f in futures])
            
        if self.config.stopping_criteria and self.config.stopping_criteria(results):
            # Early stopping if criteria met
            return results
            
        return results
    
    def _apply_correlation(
        self,
        samples: np.ndarray,
        correlation_matrix: np.ndarray
    ) -> np.ndarray:
        """Apply correlation structure to samples using Cholesky decomposition."""
        chol = np.linalg.cholesky(correlation_matrix)
        return samples @ chol.T
    
    def _run_chunk(
        self,
        model: Callable[..., np.ndarray],
        chunk_data: Dict[str, np.ndarray]
    ) -> np.ndarray:
        """Run a chunk of simulations."""
        return model(**chunk_data)
