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
