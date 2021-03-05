import database
import sys
from PyQt5.uic import loadUi
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QDialog, QApplication, QMainWindow, QMessageBox
from PyQt5.QtWidgets import QMenu, QAction, QFileDialog

#MVC: Model, View, Controller structure


#Model
class Model(object):
    """Model part in the MVC structure. interacts with the database"""

    def __init__(self, dbname=None):
        self.dbname = dbname
        self.conn = database.start_connection()

    def insert_data(self, table, values):
        database.insert_one(self.conn, table, values)

    def insert_many(self, table, loc):
        database.insert_using_excel(self.conn, table, loc)

    def view_row(self, table, designation):
        return database.view_one(self.conn, table, designation).fetchone()

    def view_table(self, table):
        return database.view_all(self.conn, table)

    def update_row(self, table, values, designation):
        database.update_one(self.conn, table, values, designation)

    def delete_row(self, table, designation):
        database.delete_one(self.conn, table, designation)

    def get_designations(self, table):
        return database.get_designations(self.conn, table)

    def get_columns(self, table):
        return database.get_columns(self.conn, table)

#View/Controller
class View(QMainWindow):
    """ View and Controller for the GUI application """
    def __init__(self, Model):
        super(View, self).__init__()
        self.model = Model
        loadUi("layout.ui", self)
        self.setWindowIcon(QtGui.QIcon('images/osdag.png'))
        self.category = ("I-section Beam", "Angles", "Channels")
        self.designation = None
        self.MenuBar()
        self.stackedWidget_func()
        self.combobox_func()
        

    def stackedWidget_func(self):
        self.stackedWidget.setCurrentIndex(2)
        self.add_item.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(0))
        self.view_item.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(1))
        self.view_tables.clicked.connect(lambda: self.stackedWidget.setCurrentIndex(3))

    def MenuBar(self):
        helpAction = QAction("&HOME", self)
        helpAction.triggered.connect(lambda: self.stackedWidget.setCurrentIndex(2))
        mainMenu = self.menuBar()
        self.setMenuBar(mainMenu)
        fileMenu = QMenu("&FILE", self)
        mainMenu.addMenu(fileMenu)
        fileMenu.addAction(helpAction)


    def combobox_func(self):
        self.add_item_combobox.addItems(self.category)
        self.view_tables_combobox.addItems(self.category)
        self.view_item_combobox1.addItems(self.category)
        self.add_item_combobox.activated.connect(self.add_steel_item)
        self.view_tables_combobox.activated.connect(self.func_view_table)
        self.view_item_combobox1.activated[str].connect(self.table_view_item)
        self.excel.clicked.connect(self.handle_excel)
        self.submit.clicked.connect(self.add_item_in_database)
        self.delete1.clicked.connect(self.delete_item)
        self.update.clicked.connect(self.update_item)
        


    # Enter data into the database
    
    def add_item_in_database(self):
        option = self.add_item_combobox.currentText()
        if option == "I-section Beam":
            option = "Beams"
        row_val = list()
        print(row_val)
        for col in range(self.header_length-1):
            col_val = self.add_item_table.item(0, col)
            if col_val is None or col_val.text() == '':
                row_val.append(None)
            else:
                col_val = col_val.text()
                try:
                    col_val = float(col_val)
                except ValueError:
                    pass
                row_val.append(col_val)
        if any(row_val) and row_val[0] is not None and row_val[0] != '':
            try:
                self.model.insert_data(option, row_val)
                self.confirm("added")
            except database.ItemAlreadyStored:
                self.item_already_present_error()
        else:
            self.on_submit()

    #Populate the table widget for the add_item stack
    def add_steel_item(self):
        option = self.add_item_combobox.currentText()
        if option == "I-section Beam":
            option = "Beams"
        header = self.model.get_columns(option)
        self.header_length = len(header)
        self.add_item_table.setColumnCount(self.header_length-1)
        self.add_item_table.setHorizontalHeaderLabels(header[1:])
        self.add_item_table.setRowCount(1)

    #Populate the table widget for the view_item stack
    def table_view_item(self):
        option = self.view_item_combobox1.currentText()
        if option == "I-section Beam":
            option = "Beams"
        result = self.model.get_designations(option)
        self.view_item_combobox2.clear()
        self.view_item_combobox2.addItems(result)
        self.view_item_combobox2.activated[str].connect(self.xyz)

    def xyz(self):
        option = self.view_item_combobox1.currentText()
        if option == "I-section Beam":
            option = "Beams"
        designation = self.view_item_combobox2.currentText()
        #This function is being called multiple times(don't know why)
        #so to avoid this, if cond is implemented.
        if self.designation != designation:
            header = self.model.get_columns(option)
            self.designation = designation
            result = self.model.view_row(option, designation)
            self.view_item_result = result
            self.view_item_table.setColumnCount(len(result)-1)
            self.view_item_table.setHorizontalHeaderLabels(header[1:])
            self.view_item_table.setRowCount(1)
            for col in range(1, len(result)):
                self.view_item_table.setItem(0, col-1, QtWidgets.QTableWidgetItem(str(result[col])))

    #Populate the table widget for the view_table stack
    def func_view_table(self):
        option = self.view_tables_combobox.currentText()
        if option == "I-section Beam":
            option = "Beams"
        header = self.model.get_columns(option)
        data = self.model.view_table(option)
        self.view_table_table.setColumnCount(len(header))
        self.view_table_table.setHorizontalHeaderLabels(header)
        self.view_table_table.setRowCount(1)
        index = 0
        for row in data:
            for col in range(len(header)):
                self.view_table_table.setItem(index, col, QtWidgets.QTableWidgetItem(str(row[col])))
            index += 1
            self.view_table_table.insertRow(self.view_table_table.rowCount())
            # print(row)
            # print("^^^^^^^")
            # print(self.view_table_table.rowCount())
        self.view_table_table.setEditTriggers(QtWidgets.QTableWidget.NoEditTriggers)

    def handle_excel(self):
        option = self.add_item_combobox.currentText()
        if option == "I-section Beam":
            option = "Beams"
        try:
            file_name, _filter = QFileDialog.getOpenFileName(self, 'Open File', '.', "Excel Files (*.xlsx *.xlsm *.xltx *.xltm)")
            self.model.insert_many(option, file_name)
            self.confirm("added")
        except Exception as e:
            self.handle_other_errors(e)

    # for updating the data in the GUI

    def update_item(self):
        option = self.view_item_combobox1.currentText()
        if option == "I-section Beam":
            option = "Beams"
        values = list()
        for col in range(self.view_item_table.columnCount()):
            col_val = self.view_item_table.item(0,col).text()
            if col_val == "None" or col_val == '':
                values.append(None)
            else:
                try:
                    col_val = float(col_val)
                except ValueError:
                    pass
                values.append(col_val)
        try:
            self.model.update_row(option, values, self.designation)
            self.confirm("updated")
        except Exception as e:
            self.handle_other_errors(e)

    # for deleting the data from the GUI

    def delete_item(self):
        option = self.view_item_combobox1.currentText()
        if option == "I-section Beam":
            option = "Beams"
        designation = self.view_item_table.item(0,0).text()
        try:
            self.model.delete_row(option, designation)
            self.confirm("deleted")
        except Exception as e:
            self.handle_other_errors(e)

    # On clicking submit button on the add items page 

    def on_submit(self):
        msg = QMessageBox(self)
        msg.setText("No Value supplied/designation is empty.\nEnter some values in the fields and then click submit")
        msg.setWindowTitle("Error Encountered")
        msg.setIcon(QMessageBox.Warning)
        msg.setStandardButtons(QMessageBox.Ok)
        msg.setDefaultButton(QMessageBox.Ok)
        msg.exec_()

    # handling any error 

    def item_already_present_error(self):
        msg = QMessageBox(self)
        msg.setText("This section is already present.\nTo update the value please use the `View/Update/Delete` tab")
        msg.setWindowTitle("Error Encountered")
        msg.setIcon(QMessageBox.Warning)
        msg.setStandardButtons(QMessageBox.Ok)
        msg.setDefaultButton(QMessageBox.Ok)
        msg.exec_()

    def handle_other_errors(self, e):
        msg = QMessageBox(self)
        msg.setText("Oops ! Some error has occured")
        exception = e.__class__.__name__
        msg.setDetailedText(exception+"\n"+str(e))
        msg.setWindowTitle("Error !!")
        msg.setIcon(QMessageBox.Critical)
        msg.setStandardButtons(QMessageBox)
        msg.setDefaultButton(QMessageBox.Ok)
        msg.exec_()

    def confirm(self,message):
        msg = QMessageBox(self)
        msg.setText("Changes done successfully !")
        msg.setInformativeText("Do vist view steel section tab to view your changes ")
        msg.setWindowTitle("Steel item {} Successfully".format(message))
        msg.setStandardButtons(QMessageBox.Ok)
        msg.setDefaultButton(QMessageBox.Ok)
        msg.exec_()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_window = View(Model('steel_sections.sqlite'))
    main_window.show()
    sys.exit(app.exec_())

