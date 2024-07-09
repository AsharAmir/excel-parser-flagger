# Excel Discrepancy Highlighter

This script highlights discrepancies. Made for https://www.reddit.com/u/trueblue1000/

### Steps to Run:

1. **Setup:**

   - Create a new folder for the project.
   - Place your Excel file (`client.xlsx` or rename accordingly) and the `main.py` script in this folder.

2. **Modify File Path:**

   - Open `main.py` (you can use Notepad if you don't have a code editor).
   - Locate the `filepath` variable in `main.py`.
   - Change `'client.xlsx'` to your Excel file's name if it's different or specify the complete path if it's located elsewhere.

3. **Install Dependencies:**

   - Open a terminal or command prompt.
   - Navigate to your project folder.
   - Run the following command to install required dependencies:

     ```
     pip install -r requirements.txt
     ```

4. **Run the Script:**

   - After installing dependencies, run the script by executing:

     ```
     python main.py
     ```

5. **Find Your Output:**
   - The modified Excel file with highlighted discrepancies will be saved in the same folder as `flagged_colorcoded.xlsx`.

### Notes:

- Ensure Python is installed on your system.
- Modify the file path in `main.py` according to your Excel file's location or name before running the script.
