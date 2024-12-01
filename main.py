import os
import configparser

from flask import Flask, flash, json, render_template, request, redirect, url_for, jsonify
from werkzeug.utils import secure_filename

from connectDB import connectDB
from ImportTool import ImportTool

def runImport(filename):
    config = configparser.ConfigParser()

    OK = True

    try:
        config.read('config.ini')

        if not ('Database' in config.sections()):
            OK = False

        if OK:
            workbook = filename
            server = config['Database']['server']
            database = config['Database']['database']
            username = config['Database']['username']
            password = config['Database']['password']    
            table = config['Database']['table']    
    except:
        print('Config file is incorrect')
        OK = False

    if OK:
        conn = connectDB(server, database, username, password)
        tool = ImportTool(workbook, conn)
        return tool.importSheet(table)

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def upload_file():

    responseTable = None
    responseHeadings = None
    if request.method == 'POST':
        files = request.files.getlist("file")  
        table = request.form['table']
        # Iterate for each file in the files List, and Save them 
        for file in files:
            file.save(os.path.join('.',  f'{file.filename}'))
            responseHeadings, responseTable = runImport(f'{file.filename}')
            responseHeadings = responseHeadings.split(',')
            os.remove(f'{file.filename}')

    return render_template('./index.html', headings=responseHeadings, table=responseTable)