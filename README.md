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



## Usage
1. Set up your directory structure and ensure SCTK is installed.
2. Update your SCLITE-formatted files with speaker IDs and sentence identification.
3. Run the script using the provided commands in the "Main Logic" section.
4. View the generated reports and tables in the output files.
5. Customize parameters and functions as needed for your specific evaluation requirements.


### File Formatting Functions:
1. **add_period_to_lines(file_path):**
   - Appends a period to lines in the specified file if they don't already end with '.', '?', or '!'.

2. **oneLiner(filepath):**
   - Reads a file and converts its content into a single line.

3. **fix_hyp_length(hyp_path, ref_path):**
   - Adjusts the length of hypothesis lines by adding hyphens at the end to match the reference line length.

4. **file_format_checker(file_path):**
   - Checks if a file contains punctuation marks ('.', ',', ';', '?', '!').

5. **lineDivider(path):**
   - Splits text into sentences using regular expressions and adjusts line breaks.

6. **fileFormatter(path):**
   - Modifies the format of the file by adding line numbers to each sentence.

7. **remove_punc(filepath):**
   - Removes specified punctuation characters from a file.

8. **small_files(filepath):**
   - Converts the content of a file to lowercase.

### ASR Evaluation Functions:
9. **subfolder_average(accsub, subsub, delesub, inssub, wersub, swesub, count):**
   - Calculates average ASR evaluation metrics per subfolder.

10. **appender(W1, W2, W3, E1, E2, E3, S1, S2, S3, D1, D2, D3, I1, I2, I3, A1, A2, A3, accsub, subsub, delesub, inssub, wersub, swesub, num):**
    - Appends ASR evaluation metrics to respective lists based on the ASR system.

11. **calculate_wer(reference_file, hypothesis_file):**
    - Uses the NIST SCTK toolkit to calculate ASR metrics (Accuracy, Substitution, Deletion, Insertion, WER, SWER).

12. **calc_folder(base_directory, hypothesisfoldername='hypothesis', referencefoldername='reference'):**
    - Iterates through a directory structure, calculates ASR metrics, and appends them to lists.

13. **compreports(reference_file, hypothesis_file):**
    - Generates ASR comparison reports using the NIST SCTK toolkit.

14. **generate_all_reports(base_directory, hypothesisfoldername='hypothesis', referencefoldername='reference'):**
    - Generates ASR comparison reports for all files in a directory structure.

15. **fix_report_dir(base_directory, report_directory="SCLITE_reports", hypothesisfoldername='hypothesis'):**
    - Moves and organizes ASR comparison reports into a separate directory.

16. **obtain_errors(report_directory="SCLITE_reports"):**
    - Extracts ASR errors (confusion pairs) from ASR comparison reports.

17. **extract_most_reperror(dfs, numofmostrepeated=2):**
    - Extracts the most repeated ASR errors from DataFrames.

18. **Table Creation:**
   -  Functions like `make_compar_table` and `make_final_table` that create tables with various metrics for ASR systems, organized by subfolders.
   - The tables are saved as Excel files.

19. **Error Extraction and Analysis:**
   - Functions like `error_extract_table` and `error_summary_table` seem to be related to extracting and summarizing errors in ASR.
   - `ErrorsinDF` function appears to perform error classification, including checking for noun-article mismatches, compound word splits, hesitations, phonetic errors, deletion errors, and more.

20. **Most Common Error Analysis:**
   - `make_mostError_table` function generates tables for the most common error types and their frequencies for each ASR system.

21. **Text Processing:**
   - There are functions like `check_word_noun`, `check_compound_words`, `check_hesitations` for checking specific linguistic features in the text.

22. **Language Processing Libraries:**
   - Spacy is used for natural language processing tasks.
   - Enchant seems to be used for language checking.

23. **Cologne Phonetic Algorithm:**
   - The `cologne_phonetics` function applies the Cologne phonetic algorithm to words.

24. **Excel Styling and Formatting:**
   - Styling of DataFrames is done before saving them to Excel files.

25. **Counter and Data Manipulation:**
   - `Counter` from the `collections` module is used for counting occurrences of specific error types.
   - DataFrames are manipulated using Pandas.


## Example
```python
# Modify input parameters
base_directory = "SCLITE_Test10"
report_directory = "SCLITE_reports4"
hypothesis_foldername = 'hypothesis'
reference_foldername = 'reference'
num_of_most_repeated = 2

# Run the ASR evaluation logic
asrcomparelogic(base_directory, hypothesis_foldername, reference_foldername)

# Run the error reporting logic
errorlogic(base_directory, report_directory, hypothesis_foldername, num_of_most_repeated)

print('Process Complete')
```
The script is designed to automate the process of evaluating ASR system performance, generating reports, and analyzing errors. You can customize the input parameters and run the script to perform these tasks on your data.
