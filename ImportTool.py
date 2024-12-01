import openpyxl

class ImportTool():
    def __init__(self, path, conn):
        # Start Import Tool
        self.conn = conn
        self.cursor = self.conn.cursor()
        self.book = openpyxl.load_workbook(path)

    
    def loadSheet(self, sheetname = None):
        if sheetname == None:
            sheet = self.book.active
        else:
            sheet = self.book[sheetname]
        self.sheet = sheet
        return sheet

    def importToSQL(self, table, fields, values):
        self.cursor.execute(f"insert into {table} ({fields}) values ({values})")
        self.cursor.commit()

    def selectSQL(self, table):
        self.cursor.execute(f"select * from {table}")
        return self.cursor.fetchall()

    def importSheet(self, table, sheet = None):
        if sheet == None:
            self.loadSheet()
        else:
            self.loadSheet(sheet)

        header = True
        fields = ''
        i = 0
        for row in self.sheet:
            if header:
                header 
                for cell in row:
                    cell = cell.value
                    if cell.find('\'') > -1:
                        cell = cell[:cell.find('\'')] + '\'' + cell[cell.find('\''):]
                    fields += cell + ','
                header = False
                fields = fields[:-1]
                print('Importing fields (' + fields + ')')
            else:
                i += 1
                values = ''
                for cell in row:
                    
                    cell = cell.value
                    try:
                        if cell.find('\'') > -1:
                            cell = cell[:cell.find('\'')] + '\'' + cell[cell.find('\''):]
                    except:
                        pass
                    values += f"'{cell}',"
                values = values[:-1]

                self.importToSQL(table, fields, values)
        
        print('Imported', i, 'records')
        return fields, self.selectSQL(table)
            
    def cleanup(self):
        self.conn.close()
        self.book.close()

    def __str__(self):
        # self.cursor.execute('select * from test')
        # records = self.cursor.fetchall()

        # table = ''
        # for row in records:
        #     for field in row:
        #         table += f'{field}' + '\t'
        #     table += '\n'
        
        # return f'{table}'
        pass
