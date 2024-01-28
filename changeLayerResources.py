import arcpy
import sys
import os

path2 = arcpy.GetParameterAsText(0)#r"D:\New Demo\Databases\NEW_PROJECT.gdb"#
arcpy.env.workspace=path2

aprx = arcpy.mp.ArcGISProject('current')
m = aprx.activeMap

listLayers = m.listLayers("*")

for lyr in listLayers:
    connprop = lyr.connectionProperties
    connprop['connection_info']['database']=path2
    lyr.updateConnectionProperties(lyr.connectionProperties,connprop)

