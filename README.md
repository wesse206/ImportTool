# Excel To SQL ImportTool
### Enable IIS
#### Turn Windows Features on (only ensure items are ticked, leave other options as default):
```
Internet Information Services >
	All Application Development Features
	Common HTTP Features >
		Default Document
		Directory Browsing
		HTTP Errors
		Static Content
	Performance Features >
		Static Content Compression
```

### Install Python
<img src="https://www.python.org/static/img/python-logo.png"/>
<br />
Get the Python installer from <a href="https://www.python.org/downloads/windows/">python.org</a>
<br />
Install Python in the local user, and ensure the ADD TO PATH checkbox is ticked.


### Install and enable Python virtual environment
#### Run in powershell in ImportTool folder (not elevated) 
Create the virtual environment and activate it:
```
python -m venv venv
.\venv\Scripts\Activate.ps1
```
You should see a `(venv)` in front of the directory in Powershell. 
<br />
Install all required Python packages in the venv:
```
python -m pip install -r requirements.txt
```

### Configure `web.config`
Edit `web.config` with the correct folder. Change all `C:\ImportTool` to `C:\path\to\tool\ImportTool`. 
<br />
Ensure that the config is edited without errors, else the IIS page will only return HTTP Code 500.

### Enable wfastcgi in IIS
#### Powershell as admin in ImportTool folder (`Win X` + `A`) 
Install `wfastcgi` to IIS to be able to host the app
```
.\venv\Scripts\wfastcgi-enable.exe
```

### Configure permissions for IIS
Give all permissions for the following folders to `IIS_IUSRS`: 
- ImportTool
- Python (where system installed it)

## Config File Structure
The config file is saved as `config.ini` in the ImportTool folder. <br />
The structure of the file is as follows:
```
[Database]
server=INSTANCE_NAME
database=DB
username=USER
password=PASS
table=WORKING_TABLE
```
Where INSTANCE_NAME points to the SQL Server Instance, DB points to the SQL Database, USER and PASS points to the SQL Database User and Password respectively and WORKING_TABLE points to the SQL Table that should be used

## SQL Table Structure
The SQL Table should contain the same number of fields that needs to be imported with all field types set up as `varchar`

## Excel File Structure
The first (therefore default) spreadsheet of the `.xlsx` workbook will be used for the import. The sheet should have headers that match the SQL Table fields.