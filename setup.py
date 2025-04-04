from setuptools import setup, find_packages

setup(
    name="ai_financial_analyzer",
    version="0.1.0",
    packages=find_packages(),
    install_requires=[
        "pandas>=1.3.0",
        "numpy>=1.20.0",
        "openpyxl>=3.0.7",
        "requests>=2.26.0",
    ],
    entry_points={
        "console_scripts": [
            "financial-analyzer=financial_analyzer:main",
        ],
    },
    author="Your Name",
    author_email="your.email@example.com",
    description="AI-powered Excel financial model analyzer",
    long_description=open("README.md").read(),
    long_description_content_type="text/markdown",
    url="https://github.com/yourusername/ai-financial-analyzer",
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.7",
)