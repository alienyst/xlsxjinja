"""Setup configuration for xlsxjinja package."""

import os
from io import open

from setuptools import setup

CUR_DIR = os.path.abspath(os.path.dirname(__file__))
README = os.path.join(CUR_DIR, "README.md")
with open(README, "r", encoding="utf-8") as fd:
    long_description = fd.read()

setup(
    name="xlsxjinja",
    version="1.0.0",
    author="Ali Ns",
    author_email="aliimrandtb@gmail.com",
    url="https://github.com/alienyst/xlsxjinja",
    packages=["xlsxjinja"],
    install_requires=[
        "openpyxl>=3.1.0",
        "jinja2>=2.10",
    ],
    extras_require={
        "image": ["Pillow>=9.0"],
        "all": ["Pillow>=9.0"],
    },
    python_requires=">=3.7",
    description="Generate Excel (.xlsx) files from templates using Jinja2",
    long_description=long_description,
    long_description_content_type="text/markdown",
    platforms=["Any platform"],
    license="MIT",
    keywords=[
        "Excel",
        "xlsx",
        "spreadsheet",
        "workbook",
        "template",
        "jinja2",
        "report",
        "generator",
    ],
    classifiers=[
        "Development Status :: 5 - Production/Stable",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Topic :: Office/Business :: Financial :: Spreadsheet",
        "Topic :: Software Development :: Libraries :: Python Modules",
    ],
)
