"""
Sampling utilities for Monte Carlo simulation.
"""
import numpy as np
from typing import Optional

def latin_hypercube_sampling(
    n_samples: int,
    n_dims: int,
    rng: Optional[np.random.Generator] = None
) -> np.ndarray:
    """Generate Latin Hypercube samples.
    
    Args:
        n_samples: Number of samples to generate
        n_dims: Number of dimensions
        rng: Random number generator (optional)
        
    Returns:
        Array of shape (n_samples, n_dims) containing Latin Hypercube samples
    """
    if rng is None:
        rng = np.random.default_rng()
        
    # Generate the intervals
    intervals = np.linspace(0, 1, n_samples + 1)
    
    # Create center points of the intervals
    centers = (intervals[1:] + intervals[:-1]) / 2
    
    # Generate points for each dimension
    points = np.zeros((n_samples, n_dims))
    
    for i in range(n_dims):
        # Randomly permute the sample indices
        perm = rng.permutation(n_samples)
        # Add some random noise within each interval
        points[:, i] = centers[perm] + rng.uniform(-0.5, 0.5, n_samples) / n_samples
        
    return np.clip(points, 0, 1)  # Ensure points are within [0, 1]
