from setuptools import setup, find_packages

setup(
    name="ms-graph",
    version="0.0.7",
    description="A lightweight Python wrapper for Microsoft Graph API.",
    long_description=open("README.md", encoding="utf-8").read(),
    long_description_content_type="text/markdown",
    author="runway28R",
    url="https://github.com/runway28R/ms-graph",
    license="MIT",
    packages=find_packages(),
    python_requires=">=3.12",
    install_requires=[
        "requests>=2.31",
        "msal>=1.27",
    ],
    keywords=["microsoft graph", "graph api", "msal", "email", "office365"],
    classifiers=[
        "Programming Language :: Python :: 3.12",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
)
