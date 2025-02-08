from setuptools import setup, find_packages

setup(
    name="monte_carlo_excel",
    version="0.1.0",
    packages=find_packages(),
    install_requires=[
        "numpy>=1.24.0",
        "pandas>=2.0.0",
        "scipy>=1.10.0",
        "xlwings>=0.30.0",
        "statsmodels>=0.14.0",
        "scikit-learn>=1.2.0",
        "arch>=5.0.0",
        "pyomo>=6.5.0",
        "cvxopt>=1.3.0",
        "matplotlib>=3.7.0",
        "seaborn>=0.12.0",
        "plotly>=5.13.0",
    ],
    extras_require={
        "dev": [
            "pytest>=7.3.0",
            "pytest-cov>=4.0.0",
            "black>=23.0.0",
            "mypy>=1.0.0",
            "pylint>=2.17.0",
            "sphinx>=6.0.0",
            "sphinx-rtd-theme>=1.2.0",
        ]
    },
    author="Your Name",
    author_email="your.email@example.com",
    description="A powerful Monte Carlo simulation Excel add-in",
    long_description=open("README.md").read(),
    long_description_content_type="text/markdown",
    keywords="monte-carlo, simulation, excel, finance, risk-analysis",
    url="https://github.com/yourusername/monte_carlo_excel",
    classifiers=[
        "Development Status :: 3 - Alpha",
        "Intended Audience :: Financial and Insurance Industry",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
    ],
    python_requires=">=3.8",
)
