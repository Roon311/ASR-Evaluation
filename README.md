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
You need to add an id for each speaker, and identify the sentence of the speaker:
```text
Hello how are you? (spkr1-001)

I am fine. (spkr2-001)

What's the weather like today? (spkr1-002)
```
                     SYSTEM SUMMARY PERCENTAGES by SPEAKER

       ,----------------------------------------------------------------.
       |                             1.txt                              |
       |----------------------------------------------------------------|
       | SPKR   | # Snt # Wrd | Corr    Sub    Del    Ins    Err  S.Err |
       |--------+-------------+-----------------------------------------|
       | spkr1  |    2      9 |100.0    0.0    0.0    0.0    0.0    0.0 |
       |--------+-------------+-----------------------------------------|
       | spkr2  |    1      3 |100.0    0.0    0.0    0.0    0.0    0.0 |
       |================================================================|
       | Sum/Avg|    3     12 |100.0    0.0    0.0    0.0    0.0    0.0 |
       |================================================================|
       |  Mean  |  1.5    8.5 |100.0    0.0    0.0    0.0    0.0    0.0 |
       |  S.D.  |  0.7    2.1 |  0.0    0.0    0.0    0.0    0.0    0.0 |
       | Median |  1.5    8.5 |100.0    0.0    0.0    0.0    0.0    0.0 |
       `----------------------------------------------------------------'

# Code Functions
1. **Formatting the Files:** This section includes functions for adding periods to lines and formatting text files.
2. **Forming the Lists:** Functions in this section calculate various metrics (e.g., WER, substitutions, deletions) for ASR systems and store the results in lists.
3. **Calculating WER:** This section contains a function to calculate Word Error Rate (WER) using a Docker container.
4. **Generating Reports:** Functions in this section generate reports for ASR system comparisons, including confusion matrices and other statistics.
5. **Error Analysis:** This part of the script extracts and analyzes errors from the ASR outputs.
6. **Creating Tables:** Functions for creating tables with the collected metrics and error analysis results.
7. **Main Logic:** The main logic section calls the functions in the correct order to perform the desired tasks.
8. **Main Function:** The script's main part specifies the input parameters and calls the functions accordingly.

The script is designed to automate the process of evaluating ASR system performance, generating reports, and analyzing errors. You can customize the input parameters and run the script to perform these tasks on your data.
