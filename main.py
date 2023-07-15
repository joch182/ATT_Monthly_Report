from PyQt5 import QtWidgets
import UI
import data_importer

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Form = QtWidgets.QWidget()
    ui = UI.Ui_ATT_Monthly_Report()
    ui.setupUi(Form)
    umtsImporter = data_importer.CountersDataImporter()
    lteImporter = data_importer.CountersDataImporter()
    ui.umts_counters_1.clicked.connect(umtsImporter.umts_data_import1)
    ui.umts_counters_2.clicked.connect(umtsImporter.umts_data_import2)
    ui.runUMTS.clicked.connect(lambda: umtsImporter.umts_report_run())
    ui.lte_counters_1.clicked.connect(lteImporter.lte_data_import1)
    ui.lte_counters_2.clicked.connect(lteImporter.lte_data_import2)
    ui.lte_counters_mocn.clicked.connect(lteImporter.lte_data_import_mocn)
    ui.runLTE.clicked.connect(lambda: lteImporter.lte_report_run())
    
    Form.show()
    sys.exit(app.exec_())