import sys
from PyQt5 import QtWidgets, QtGui
import filling_SV

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    ui = filling_SV.Ui_MainWindow()
    sys.exit(app.exec_())