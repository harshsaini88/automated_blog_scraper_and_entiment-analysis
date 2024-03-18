# How To Run

Follow these steps to run the project:

1. **Download:**
   - Download the project zip file.

2. **Extract:**
   - Extract the contents of the zip file to a folder on your computer.

3. **Install Dependencies:**
   - Open your command line interface (CLI).
   - Navigate to the folder where you extracted the project files.
   - Run this command:
     ```
     pip install pandas requests beautifulsoup4 openpyxl
     ```

4. **Run the Script:**
   - Still in the CLI, navigate into the 'nlp' folder of the extracted project files.
   - Run the script:
     ```
     python main.py
     ```

5. **Provide Input:**
   - Make sure you have an Excel file named `input.xlsx` in the 'nlp' folder.
   - In the Excel file, have a column named 'URL' containing the URLs of articles you want to analyze.

6. **View Output:**
   - After the script finishes running, it generates an Excel file named `output.xlsx` in the same 'nlp' folder.
   - Open this file to see the analysis results for each article.
