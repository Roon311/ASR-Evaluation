# Automated Speech Recognition (ASR) Evaluation Tool
Evaluating ASR files through SCTK/SCLITE

# Introduction
The Automated Speech Recognition (ASR) Evaluation Tool is designed to streamline the process of evaluating ASR files using the SCTK/SCLITE toolkit. This tool aids in assessing the accuracy and performance of different ASR systems, providing valuable insights into their capabilities.


# Requirements
Before using the ASR Evaluation Tool, ensure you have the following prerequisites installed:
1. **Python**: Make sure you have Python installed on your system.
2. Packages: Install the required packages using the following command:
```bash
pip install pandas openpyxl
```
3.SCTK: To install compile and install sctk, from the main directory type the commands:

	% make config
	% make all
	% make check
	% make install
	% make doc

4.Docker: A Dockerfile is included to build and use SCTK without compiling the codebase for your particular platform.
```bash
 docker build -t sctk .
```
The command used to run sctk on linux and mac on this repository is:
```bash
docker run -it -v $PWD:/var/sctk sctk sclite -i wsj -r ref.txt -h hyp.txt
```
# SCLITE file formating
