# Document Anonymizer

This script anonymizes sensitive terms in Word (.docx), Excel (.xlsx), and CSV (.csv) files. It replaces the sensitive terms with randomly generated strings of the same length and creates a separate mapping file to help track the original terms and their anonymized counterparts.


## Installation

1. Make sure you have Python 3.6 or higher installed on your system. You can download Python from the [official website](https://www.python.org/downloads/).

2. Clone or download this repository to your local machine.

3. Open a terminal or command prompt, navigate to the project directory, and install the required packages using the following command:

pip install -r requirements.txt


## Usage

1. Prepare a text file containing the list of sensitive terms you want to anonymize, one term per line.

2. Run the script using the following command:

python anonymizer.py

3. You will be prompted to select the text file containing the list of sensitive terms.

4. Next, you will be prompted to select the document (.docx, .xlsx, or .csv) that you want to anonymize.

5. The script will create a new anonymized document in the "Output Files" directory within the project directory. The anonymized file will have the same name as the original file, with "_anonymized" appended to the filename.

6. The script will also create a text file named `originalfilename_anonymized_terms_mapping.txt` in the "Output Files" directory. This file contains the mappings of the original sensitive terms to their anonymized counterparts.

7. Check the "Output Files" directory for the anonymized document and the mappings file.