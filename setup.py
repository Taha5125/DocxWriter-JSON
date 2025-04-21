"""
Setup script for DocxWriter package.
"""

from setuptools import setup, find_packages

setup(
    name="docxwriter",
    version="0.1.0",
    description="A Python library for creating Word documents from JSON data",
    author="Karar Haider",
    author_email="",  # Add your email here if desired
    packages=find_packages(),
    classifiers=[
        "Development Status :: 3 - Alpha",
        "Intended Audience :: Developers",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.7",
    entry_points={
        "console_scripts": [
            "docxwriter=docxwriter.__main__:main",
        ],
    },
)
