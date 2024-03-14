import os
import openpyxl
from copy import copy
import logging
import plugins.base

class pluginClass(plugins.base.basePlugin):
    def __init__(self):
        pass
    def run(self, pluginInput, pluginOutput):
        py_logger = logging.getLogger(__name__)
        py_logger.setLevel(logging.INFO)
        py_handler = logging.FileHandler(f"{pluginOutput}/{__name__}.log", mode='w')
        py_formatter = logging.Formatter("%(name)s %(asctime)s %(levelname)s %(message)s")
        py_handler.setFormatter(py_formatter)
        py_logger.addHandler(py_handler)

        def makeReport(wbSources, wbSampleFileName, wbReportFileName):
            """
            Сформировать отчет.
            """
            pass          

        

        wbSources = list()
        wbSampleFileName = ""
        wbReportFileName = ""
        wbSources.clear()

        files = os.listdir(pluginInput)
        for iFile in files:
            file_name, file_extension = os.path.splitext(iFile)
            if (file_extension == ".xlsx") or (file_extension == ".xls"):
                if file_name.find("образец_сводный на проверку") != -1:
                    wbSampleFileName = pluginInput + "/" + iFile
                    wbReportFileName = pluginOutput + "/" + "сводный на проверку.xlsx"
                else:
                    wbSources.append(pluginInput + "/" + iFile)

        try:
            py_logger.info(f"Запускаем обработку файлов ")
            makeReport(wbSources, wbSampleFN, wbReportFileName)
        except:
            py_logger.error(f"Ошибка при обработке файлов", exc_info=True)
