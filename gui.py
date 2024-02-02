import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QPushButton, QLabel, QTextEdit, QComboBox
from sqlite3 import connect
import pandas as pd

class ChelGUI(QMainWindow):
    def __init__(self):
        super().__init__()

        self.conn = None

        self.setWindowTitle("Endo Care")
        self.setGeometry(100, 100, 600, 400)

        self.initUI()

    def initUI(self):
        self.central_widget = QWidget(self)
        self.setCentralWidget(self.central_widget)

        layout = QVBoxLayout()
        self.central_widget.setLayout(layout)

        self.function_label = QLabel("Select Function:")
        layout.addWidget(self.function_label)

        self.function_combobox = QComboBox(self)
        layout.addWidget(self.function_combobox)

        # Add available functions to the combo box
        available_functions = [
            "Login ExFM",
            "Select Database",
            "Run SQL",
            "GDKT/Trouble Report",
            "Quotation",
            "Weekly Report",
            "Exit"
        ]
        self.function_combobox.addItems(available_functions)

        self.run_button = QPushButton("Run Function")
        self.run_button.clicked.connect(self.run_function)
        layout.addWidget(self.run_button)

        self.result_label = QLabel("Function Result:")
        layout.addWidget(self.result_label)

        self.result_text_edit = QTextEdit(self)
        self.result_text_edit.setReadOnly(True)
        layout.addWidget(self.result_text_edit)

    def run_function(self):
        function = self.function_combobox.currentText()
        result = ""

        if function.strip() == "":
            result = "Please select a function."
        elif "SELECT" in function.upper():
            try:
                result_df = pd.read_sql(function, self.conn)
                result = result_df.to_string()
            except Exception as e:
                result = str(e)
        else:
            result = f"'{function}' is not supported in the GUI. Please run it from the command line."

        self.result_text_edit.setPlainText(result)


    def connect_to_database(self, db_name):
        try:
            self.conn = connect(db_name)
        except Exception as e:
            self.result_text_edit.setPlainText(str(e))


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ChelGUI()
    ex.show()
    sys.exit(app.exec_())
