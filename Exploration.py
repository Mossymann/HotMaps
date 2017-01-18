from qgis.core import *
import sip
API_NAMES = ["QDate", "QDateTime", "QString", "QTextStream", "QTime", "QUrl", "QVariant"]
API_VERSION = 2
for name in API_NAMES:
    sip.setapi(name, API_VERSION)
    
from PyQt4.QtGui import *
app = QgsApplication([], True)
app.setPrefixPath("C:/OSGeo4W64/apps/qgis", True)
app.initQgis()

import processing
from processing.core.Processing import Processing
Processing.initialize()
Processing.updateAlgsList()
# ------------------------------------------------------------------------------------------------------------
from PyQt4.QtCore import *
from xlrd import open_workbook
from PyQt4.QtCore import QVariant
import os
import shutil

'''
==============================================================================================================
Technology codes:
                    1 --> Wood pellets central
                    2 --> Heat pump air/water
                    3 --> Gas central
                    4 --> District Heating (DH)
==============================================================================================================                    
'''
def JoinLayers(input_joinLyr, joinField, input_vectorLyr, targetField, outputPath):
    vectorLyr = QgsVectorLayer(input_vectorLyr, "vectorLyr", "ogr")
    joinLyr = QgsVectorLayer(input_joinLyr, "joinLyr", "ogr")
    if not vectorLyr.isValid():
        print "Layer %s did not load" % vectorLyr.name()
    if not joinLyr.isValid():
        print "Layer %s did not load" % joinLyr.name()
    res = processing.runalg("qgis:joinattributestable",vectorLyr,joinLyr, targetField, joinField, outputPath)
    resultLyr = QgsVectorLayer(res['OUTPUT_LAYER'], "joined layer", "ogr")
    QgsMapLayerRegistry.instance().addMapLayers([vectorLyr, joinLyr, resultLyr])
    if resultLyr.isValid():
        QgsMapLayerRegistry.instance().removeMapLayers([vectorLyr, joinLyr, resultLyr])
        return True
    else:
        QgsMapLayerRegistry.instance().removeAll
        return False
    
#     -------------------------------------------------------------------------------------------------------
#     here another method:
#     -------------------------------------------------------------------------------------------------------
#     QgsMapLayerRegistry.instance().addMapLayers([vectorLyr,joinLyr], True)
#     vectorLyrCRS = vectorLyr.crs()
#     info = QgsVectorJoinInfo()
#     info.joinLayerId = joinLyr.id()
#     info.joinFieldName = joinField
#     info.targetFieldName = targetField
#     info.memoryCache = True
#     vectorLyr.addJoin(info)
#     QgsVectorFileWriter.writeAsVectorFormat(vectorLyr,outputPath, "CP120", vectorLyrCRS, "ESRI Shapefile")
#     outputLyr = QgsVectorLayer(outputPath, 'Joined', 'ogr')
#     QgsMapLayerRegistry.instance().removeMapLayers([vectorLyr, joinLyr])
#     return outputLyr
#     -------------------------------------------------------------------------------------------------------
 
def regCost(p,x0,y0,x1,y1):
    if x0 != x1:
        return float(y1 - ( x1 - p ) * ( y1 - y0 ) / ( x1 - x0 ))
    else:
        return float(y1)



def investment(joinLyrPath, econVal, powSteps, priceSteps):
    vectorLyr = QgsVectorLayer(joinLyrPath,'resLyr' , "ogr")
    vpr = vectorLyr.dataProvider()
    vpr.addAttributes([QgsField("CheapTech", QVariant.Int)])
    vpr.addAttributes([QgsField("AnnualCost", QVariant.Double)])
    vectorLyr.updateFields()
    
    for i in range(4): # technologies
        if powSteps[i] == []:
            continue
        lt = econVal[i][0]
        z = econVal[i][1]
        opc = econVal[i][2]
        fp = econVal[i][3]
        accf = float( ( z * (1+z)**lt ) / ( (1+z)**lt - 1 ) ) 
        numFields = vectorLyr.dataProvider().fields().count() # number of fields starting from 1
        pfid = numFields - 6 # join process adds 3 columns based on the input file. "Power" locates in his position
        
        for feature in vectorLyr.getFeatures():
            p = float(feature.attributes()[pfid])
            if i == 2: # for gas systems
                if feature.attributes()[numFields - 4] != "1": # 1: there is connection to gas
                    continue
            if i == 3: # for DH systems
                if feature.attributes()[numFields - 3] != "1": # 1: there is connection to DH
                    continue    
            try:
                x0 = max(x for x in powSteps[i] if x < p)
                y0 = priceSteps[i][powSteps[i].index(x0)]
            except:
                x0 = 0
                y0 = 0
            try:
                x1 = min(x for x in powSteps[i] if x > p)
                y1 = priceSteps[i][powSteps[i].index(x1)]
            except:
                x1 = p
                y1 = p * y0/x0

                
            capCost = regCost(p,x0,y0,x1,y1)
            temp = accf * capCost + p * opc + float(feature.attributes()[numFields - 7]) * fp 
            if feature.attributes()[numFields - 1] == "":
                vectorLyr.startEditing()
                vectorLyr.changeAttributeValue(feature.id(), numFields - 1, temp)
                vectorLyr.changeAttributeValue(feature.id(), numFields - 2, i)
                vectorLyr.commitChanges() 
            elif temp < feature.attributes()[numFields - 1]:
                vectorLyr.startEditing()
                vectorLyr.changeAttributeValue(feature.id(), numFields - 1, temp)
                vectorLyr.changeAttributeValue(feature.id(), numFields - 2, i)
                vectorLyr.commitChanges()
            else:
                pass
    QgsMapLayerRegistry.instance().addMapLayer(vectorLyr)
    QgsMapLayerRegistry.instance().removeMapLayer(vectorLyr)        
 
    

projectPath = str(QInputDialog.getText(None, "Project Path","Please enter the project path:")[0])
input_vectorLyr = str(QInputDialog.getText(None, "Input Shapefile", "Please enter the path to the input shapefile:")[0])
input_block_pd = str(QInputDialog.getText(None, "Power & Demand", "Please enter the path to the .csv file containing buildings installed power and demand:")[0])
input_gdha = str(QInputDialog.getText(None, "Access to Gas & DH Grids", "Please enter the path to the .csv file containing gas & DH grid access:")[0])
input_techcapcost = str(QInputDialog.getText(None, "Technology Properties", "Please enter the path to the .csv file containing properties of each technology:")[0])

# projectPath = "C:/HotMaps"
# input_vectorLyr = "C:/HotMaps/BBlocks/BAUBLOCKOGD.shp"
# input_block_pd = "C:/HotMaps/Buildings.csv"
# input_gdha = "C:/HotMaps/ZBez.csv"
# input_techcapcost = "C:/HotMaps/TechCapCosts.xlsx"

projectPath.replace("\\", "/")
input_vectorLyr.replace("\\", "/")
input_block_pd.replace("\\", "/")
input_gdha.replace("\\", "/")
input_techcapcost.replace("\\", "/")
tempDir = projectPath + os.sep + "Temp"
outputDir = projectPath + os.sep + "Output"
if not os.path.exists(projectPath):
    os.makedirs(projectPath)
if not os.path.exists(tempDir):
    os.makedirs(tempDir)
if not os.path.exists(outputDir):
    os.makedirs(outputDir)


joinLyrPathTemp = tempDir + os.sep + "tempJoinLyr.shp"
joinLyrPath = outputDir + os.sep + "joinLyr.shp"


a = JoinLayers(input_block_pd,"BLKNR", input_vectorLyr, "BLKNR", joinLyrPathTemp)
if not a:
    msg = QMessageBox.information(None, "Failure", "The joining process failed!")
    #return
a = JoinLayers(input_gdha,"ZBEZ", joinLyrPathTemp, "ZBEZ", joinLyrPath)
if not a:
    msg = QMessageBox.information(None, "Failure", "The joining process failed!")
    #return

econVal = [] 
powSteps = []
priceSteps = []
temp = []
temp1 = []
temp2 = []  
wb = open_workbook(input_techcapcost)
max_steps = wb.sheets()[0].nrows

# reading demand & power from the excel file
for i in range(4):
    for j in range (3,max_steps):
        if wb.sheets()[0].cell(j, 3*i + 2).value != '':
            temp1.append(float(wb.sheets()[0].cell(j, 3*i + 2).value))
            temp2.append(float(wb.sheets()[0].cell(j, 3*i + 3).value))
    powSteps.append(temp1)
    priceSteps.append(temp2)
    temp1 = []
    temp2 = []

# reading lifetime, interest rate, operation costs and energy price from the excel file
for i in range(4):    # technologies
    for j in range(4):   # elements
        temp.append(float(wb.sheets()[1].cell(j+3,i+2).value))
    econVal.append(temp)
    temp = []

del(temp,temp1,temp2)
# wb.close()

investment(joinLyrPath, econVal, powSteps, priceSteps)

joinLyr = QgsVectorLayer(joinLyrPath,"outputLyr","ogr")
QgsMapLayerRegistry.instance().addMapLayer(joinLyr)
joinLyr.loadNamedStyle(outputDir + os.sep + "myStyle.qml")
f = QFileInfo(outputDir + os.sep + "myProject.qgs")
p = QgsProject.instance().write(f)

QgsMapLayerRegistry.instance().removeMapLayer(joinLyr)
# os.remove(joinLyrPathTemp)
# if os.path.exists(tempDir):
#     shutil.rmtree(tempDir)
print("done!")
