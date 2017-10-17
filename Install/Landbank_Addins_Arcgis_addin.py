import pythonaddins

relPath = os.path.dirname(__file__)

class ButtonClass1(object):
    """Implementation for Landbank_Addins_Arcgis_addin.button (Button)"""
    def __init__(self):
        self.enabled = True
        self.checked = False
    def onClick(self):
        toolPath = relPath + "\\Landbank.pyt"
        pythonaddins.GPToolDialog(toolPath, "polygonToExcel")

class ButtonClass2(object):
    """Implementation for Landbank_Addins_Arcgis_addin.button_1 (Button)"""
    def __init__(self):
        self.enabled = True
        self.checked = False
    def onClick(self):
        toolPath = relPath + "\\Landbank.pyt"
        pythonaddins.GPToolDialog(toolPath, "excelToPolygon")
