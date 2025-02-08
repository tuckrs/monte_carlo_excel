# Monte Carlo Excel Add-in

A powerful Monte Carlo simulation engine implemented as an Excel add-in, providing advanced statistical analysis, optimization, and visualization capabilities.

## Features

- Core Monte Carlo Simulation Engine with multiple sampling techniques
- Support for 100+ probability distributions
- Advanced correlation and dependency modeling
- Time-series analysis and forecasting
- Portfolio optimization and Efficient Frontier Analysis
- Comprehensive visualization and reporting tools
- High-performance multi-threaded computation

## Project Structure

```
monte_carlo_excel/
├── src/                    # Source code
│   ├── core/              # Core simulation engine
│   ├── excel/             # Excel integration
│   ├── distributions/      # Probability distributions
│   ├── optimization/       # Optimization algorithms
│   ├── visualization/      # Plotting and reporting
│   └── utils/             # Utility functions
├── tests/                 # Test files
├── docs/                  # Documentation
└── examples/              # Example workbooks
```

## Development Setup

1. Requirements:
   - Python 3.8+
   - Microsoft Excel
   - Required Python packages (see requirements.txt)

2. Installation:
   ```bash
   pip install -r requirements.txt
   ```

3. Building the Add-in:
   ```bash
   python setup.py build
   ```

## Testing

Run the test suite:
```bash
pytest tests/
```

## License

This project is licensed under the MIT License - see the LICENSE file for details.
