# 🦾🤖 ETL Project to Save Baseline Dataframe on TOTVS MSSQL Database

### 🦾 Automating Baseline Data Processing in MSSQL for TOTVS ERP

This project is an **ETL (Extract, Transform, Load)** automation tool designed to manage baseline data in the **TOTVS ERP Protheus** system. It automates the data extraction from Excel files, validates the baseline codes, removes obsolete records, and inserts the latest baseline information into a **MSSQL** database.

## 🌟 Features

- **File Handling**: Automatically reads and deletes temporary Excel files used for ETL.
- **Baseline Management**: Validates baseline data, checks if records already exist, and updates them.
- **Progress Tracking**: User-friendly **Tkinter GUI** with a progress bar.
- **Exception Handling**: Graceful error handling with detailed error messages for database or file operations.
- **MSSQL Integration**: Connects securely to MSSQL using credentials stored in a configuration file.

## 🛠️ Tech Stack

- **Python**: Core language used for automation.
- **Tkinter**: GUI framework for interactive user interface.
- **Pandas**: Data processing and transformation of Excel files.
- **PyODBC**: Database connectivity to MSSQL.
- **OpenPyXL**: Reads Excel files for baseline data extraction.

## 🚀 Getting Started

Follow these steps to set up and run the project on your local machine.

### 1. Clone the Repository

```bash
git clone https://github.com/seu-usuario/nome-do-repositorio.git
cd nome-do-repositorio
```

### 2. Requirements
Make sure you have the following installed:

- **Python 3.8+**
- **MSSQL Server**
- **Pandas**
- **PyODBC**
- **Tkinter (for GUI applications)**
- **OpenPyXL (for reading Excel files)**

## 3. Installation
Install the required dependencies using `pip`:

```bash
pip install -r requirements.txt
```

Alternatively, install the dependencies manually:
```bash
pip install pandas pyodbc openpyxl
```

## 4. MSSQL Configuration
Ensure your MSSQL credentials are stored in a text file at the following location:
```bash
\\192.175.175.4\desenvolvimento\REPOSITORIOS\resources\application-properties\USER_PASSWORD_MSSQL_PROD.txt
```

The file should contain the following information (each separated by a semicolon):
```bash
username;password;database;server
```

## 5. Running the Application
To start the ETL process, run the following command:
```bash
python main.py
```

The **Tkinter GUI** will pop up, and the application will begin processing the baseline data automatically. The progress will be displayed in the GUI.

## 6. Environment Variables
Ensure that the `QP_BASELINE` environment variable is set before running the application. This variable is used to specify the QP code for the baseline.

Example:
```bash
export QP_BASELINE=QP-E1234
```

## 📋 How It Works
1. **Setup MSSQL:** The application reads MSSQL credentials from the configuration file.
2. **Data Extraction:** Reads and validates baseline data from an Excel file.
3. **Database Operations:**
- Verifies if the baseline already exists in the database.
- Deletes existing records, if found.
- Inserts the new baseline data.
4. **Progress Updates:** The application shows progress in the GUI.
5. **File Cleanup:** After completion, the source Excel file is deleted from the system.

## 🐛 Troubleshooting
- **FileNotFoundError:** If the MSSQL credentials file is missing, ensure the path is correct.
- **Database Connection Error:** Check MSSQL server connectivity and ensure the credentials are correct.
- **Excel File Error:** Verify that the Excel file follows the correct format and is located in the temp directory.

## ⚖️ License
This project is licensed under the MIT License - see the [LICENSE](https://www.mit.edu/~amini/LICENSE.md) file for details.

## ✨ Acknowledgements

- [Tkinter Documentation](https://docs.python.org/3/library/tkinter.html) for the GUI framework.
- [Pandas Documentation](https://pandas.pydata.org/docs/) for efficient data manipulation.
- [PyODBC Documentation](https://github.com/mkleehammer/pyodbc) for MSSQL connectivity.

# 👨‍💻 Author
Developed by [Eliezer Moraes Silva](https://www.linkedin.com/in/eliezer-moraes-silva-80b68010b/). Feel free to connect!
