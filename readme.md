# About Project

The project is designed to automate the updates of excel reports

# Instaliation

```git clone https://github.com/Polonez1/excel_files_refresher```

install requirements

```pip install -r requirements.txt```

# Doc

1. Open reports_directories.json
2. paste you reports path by days (when do you want to update. If every day, then paste everywhere)
3. Save
4. run ```python main.py``` module


# Notes

- all reports are updated visible
- after update (not sharepoint) saves 'read only'
- Sharepoint requires windows auth (your PC must have sharepoint access)
- automatically agrees with "Trust source" in the excel report

# Plans for the project

It is planned to make excel report control through the terminal