"""
Excel add-in integration using xlwings.
"""
import os
import numpy as np
import xlwings as xw
from typing import Dict, List, Optional, Tuple
from ..core.simulation import MonteCarloEngine, SimulationConfig
from ..distributions import get_distribution
from ..visualization import create_histogram, create_tornado_chart

class MonteCarloAddin:
    """Excel add-in for Monte Carlo simulation."""
    
    def __init__(self):
        """Initialize the Excel add-in."""
        self.engine = None
        self.current_results = None
        self.current_inputs = None
    
    def setup_ribbon(self) -> None:
        """Setup Excel ribbon with simulation controls."""
        # This will be called by xlwings when loading the add-in
        pass
    
    def run_simulation(
        self,
        input_range: str,
        output_range: str,
        num_iterations: int = 10000,
        use_lhs: bool = False,
        correlation_matrix: Optional[np.ndarray] = None
    ) -> None:
        """Run Monte Carlo simulation using inputs from Excel.
        
        Args:
            input_range: Excel range containing input parameters
            output_range: Excel range for output
            num_iterations: Number of iterations
            use_lhs: Whether to use Latin Hypercube Sampling
            correlation_matrix: Optional correlation matrix
        """
        wb = xw.books.active
        
        # Read input parameters
        inputs = wb.range(input_range).value
        self.current_inputs = self._process_inputs(inputs)
        
        # Create simulation config
        config = SimulationConfig(iterations=num_iterations)
        self.engine = MonteCarloEngine(config)
        
        # Run simulation
        self.current_results = self.engine.run_simulation(
            model=self._model_function,
            input_distributions=self.current_inputs,
            correlation_matrix=correlation_matrix,
            use_lhs=use_lhs
        )
        
        # Write results
        self._write_results(output_range)
    
    def create_charts(
        self,
        chart_type: str = "histogram",
        target_range: str = None
    ) -> None:
        """Create visualization charts.
        
        Args:
            chart_type: Type of chart to create
            target_range: Excel range for chart placement
        """
        if self.current_results is None:
            raise ValueError("No simulation results available")
            
        wb = xw.books.active
        sheet = wb.sheets.active
        
        if chart_type == "histogram":
            fig = create_histogram(self.current_results)
        elif chart_type == "tornado":
            fig = create_tornado_chart(
                self.current_results,
                self.current_inputs
            )
        else:
            raise ValueError(f"Unknown chart type: {chart_type}")
            
        if target_range:
            sheet.pictures.add(
                fig,
                name=f"MCSim_{chart_type}",
                update=True,
                left=wb.range(target_range).left,
                top=wb.range(target_range).top
            )
    
    def _process_inputs(
        self,
        inputs: List[List]
    ) -> Dict[str, callable]:
        """Process input parameters from Excel.
        
        Args:
            inputs: List of input parameters [name, distribution, params]
            
        Returns:
            Dictionary of distribution functions
        """
        distributions = {}
        for name, dist_type, *params in inputs:
            dist = get_distribution(dist_type, params)
            distributions[name] = dist
        return distributions
    
    def _model_function(self, **kwargs) -> np.ndarray:
        """Model function that processes inputs and returns outputs."""
        # This will be customized based on the Excel model
        return sum(kwargs.values())
    
    def _write_results(self, output_range: str) -> None:
        """Write simulation results to Excel.
        
        Args:
            output_range: Excel range for output
        """
        wb = xw.books.active
        results_range = wb.range(output_range)
        
        # Write summary statistics
        summary = [
            ["Statistic", "Value"],
            ["Mean", np.mean(self.current_results)],
            ["Std Dev", np.std(self.current_results)],
            ["Min", np.min(self.current_results)],
            ["Max", np.max(self.current_results)],
            ["5th Percentile", np.percentile(self.current_results, 5)],
            ["95th Percentile", np.percentile(self.current_results, 95)]
        ]
        
        results_range.value = summary
