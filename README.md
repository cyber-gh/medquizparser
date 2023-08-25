# README: Language Detection and Excel Writing

This program reads questions and answers from an `input.txt` file, detects the language of the questions, and writes them to an Excel file named `mastersheet.xlsx`.

## Prerequisites

1. **Python**: Ensure you have Python installed on your computer. If not, download and install it from [python.org](https://www.python.org/downloads/).

2. **Required Libraries**: The program uses a few Python libraries. You can install them using the following commands:

\```bash
pip install openpyxl langdetect
\```

3. **Input File**: Ensure you have an `input.txt` file in the same directory as the program. This file should contain the questions and answers you want to process.

## Steps to Execute the Program

1. **Save the Code**: Copy the provided source code and save it to a file named `new_parser.py` in a directory of your choice.

2. **Open Terminal or Command Prompt**: Navigate to the directory where you saved `new_parser.py`.

3. **Run the Program**: Execute the following command:

```bash
python new_parser.py
```

4. **Check the Output**: After running the program, you should see a file named `mastersheet.xlsx` in the same directory. This Excel file will contain the processed questions, answers, and explanations.

## Notes

- If you encounter any errors related to missing libraries, ensure you've installed all the required libraries as mentioned in the Prerequisites section.
- The program assumes a specific format for the `input.txt` file. Ensure your questions and answers follow this format for accurate processing.

---

Feel free to share this README with anyone who needs to execute the program. They can follow the steps even if they're not familiar with programming.
