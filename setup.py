from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="aspose-total-python-java",
    version="1.0.0",
    author="Your Name",
    author_email="your.email@example.com",
    description="Document processor using Aspose.Total for Python via Java",
    long_description=long_description,
    long_description_content_type="text/markdown",
    packages=find_packages(),
    classifiers=[
        "Development Status :: 3 - Alpha",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: MIT License",
        "Operating System :: MacOS",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
    ],
    python_requires=">=3.8",
    install_requires=[
        "JPype1>=1.4.1",
        "pathlib2>=2.3.7",
        "typing-extensions>=4.0.0",
    ],
    entry_points={
        "console_scripts": [
            "aspose-convert=main:main",
        ],
    },
)