"""
Probability distributions for Monte Carlo simulation.
"""
import numpy as np
from typing import List, Callable, Union
from scipy import stats

def get_distribution(
    dist_type: str,
    params: List[float]
) -> Callable[[np.ndarray], np.ndarray]:
    """Get a distribution function based on type and parameters.
    
    Args:
        dist_type: Type of distribution
        params: Distribution parameters
        
    Returns:
        Function that takes uniform random samples and returns distributed samples
    """
    dist_type = dist_type.lower()
    
    if dist_type == "normal":
        mu, sigma = params
        return lambda u: stats.norm.ppf(u, loc=mu, scale=sigma)
        
    elif dist_type == "lognormal":
        mu, sigma = params
        return lambda u: stats.lognorm.ppf(u, s=sigma, scale=np.exp(mu))
        
    elif dist_type == "uniform":
        low, high = params
        return lambda u: low + (high - low) * u
        
    elif dist_type == "triangular":
        low, mode, high = params
        return lambda u: stats.triang.ppf(
            u,
            c=(mode - low) / (high - low),
            loc=low,
            scale=high - low
        )
        
    elif dist_type == "beta":
        a, b, low, high = params
        return lambda u: stats.beta.ppf(u, a, b) * (high - low) + low
        
    elif dist_type == "gamma":
        shape, scale = params
        return lambda u: stats.gamma.ppf(u, shape, scale=scale)
        
    elif dist_type == "weibull":
        shape, scale = params
        return lambda u: stats.weibull_min.ppf(u, shape, scale=scale)
        
    elif dist_type == "custom":
        # Custom distribution using interpolation of empirical CDF
        values, probs = params
        return lambda u: np.interp(u, probs, values)
        
    else:
        raise ValueError(f"Unknown distribution type: {dist_type}")

def fit_distribution(
    data: np.ndarray,
    dist_type: str = "auto"
) -> tuple[str, List[float]]:
    """Fit a distribution to empirical data.
    
    Args:
        data: Empirical data
        dist_type: Distribution type to fit, or "auto" for automatic selection
        
    Returns:
        Tuple of (distribution type, parameters)
    """
    if dist_type == "auto":
        # Try different distributions and select best fit
        distributions = ["normal", "lognormal", "gamma", "weibull"]
        best_aic = np.inf
        best_dist = None
        best_params = None
        
        for dist in distributions:
            try:
                if dist == "normal":
                    params = stats.norm.fit(data)
                    aic = stats.norm.nnlf(params, data)
                elif dist == "lognormal":
                    params = stats.lognorm.fit(data)
                    aic = stats.lognorm.nnlf(params, data)
                elif dist == "gamma":
                    params = stats.gamma.fit(data)
                    aic = stats.gamma.nnlf(params, data)
                elif dist == "weibull":
                    params = stats.weibull_min.fit(data)
                    aic = stats.weibull_min.nnlf(params, data)
                
                if aic < best_aic:
                    best_aic = aic
                    best_dist = dist
                    best_params = params
                    
            except Exception:
                continue
                
        if best_dist is None:
            raise ValueError("Could not fit any distribution to data")
            
        return best_dist, list(best_params)
        
    else:
        # Fit specified distribution
        if dist_type == "normal":
            params = stats.norm.fit(data)
        elif dist_type == "lognormal":
            params = stats.lognorm.fit(data)
        elif dist_type == "gamma":
            params = stats.gamma.fit(data)
        elif dist_type == "weibull":
            params = stats.weibull_min.fit(data)
        else:
            raise ValueError(f"Cannot fit distribution type: {dist_type}")
            
        return dist_type, list(params)
