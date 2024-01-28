#!/usr/bin/python
# -*- coding: utf-8 -*-
import math
import arcpy
import xlrd
import os
import sys
try:
    sys.path.append(sys.argv[0])
except:
    pass
try:
    sys.path.append(r"D:\Red_Oak_Project\python\fanalv1")
except:
    pass
import platmatrixmutipagemain as pmm
import platmatrixsinglepagemain as pms

def returnTractList(wrkbook):
    tractlist=[]
    wksht = wrkbook.sheet_by_name("TRACT_LIST")
    for x in range(1,wksht.nrows):
        cell = wksht.cell_value(x,0)
        if not cell is None and cell !="":
            if not cell.strip() in tractlist:
                tractlist.append(cell)

    return tractlist

def returnPageLength(centerlineshape):
    dx = abs(centerlineshape.firstPoint.X-centerlineshape.lastPoint.X)
    dy = abs(centerlineshape.firstPoint.Y-centerlineshape.lastPoint.Y)
    if dx>=dy: return dx,"EW"
    if dy>dx: return dy,"NS"

def getCenterlineShape(centerline,cenname):
    desc = arcpy.Describe(centerline.name)
    sr = desc.spatialReference
    centerlineshape = None
    name = None
    where = "CENTERLINE_NAME LIKE '{}'".format(cenname)
    with arcpy.da.SearchCursor(centerline.name,["SHAPE@","CENTERLINE_NAME"],where,spatial_reference=sr) as sc:
        for row in sc:
            if row:
                centerlineshape = row[0]
                name = row[1]
                
    try:
        del sc
    except:
        print("No Delete Cursor")
    return centerlineshape,name

def getBoundaryShape(boundarypoly,tractnumber):
    desc = arcpy.Describe(boundarypoly.name)
    sr = desc.spatialReference
    where = "TRACT_NUMBER LIKE '%{}%'".format(tractnumber)
    centerlineshape = None
    name = None
    recagent = None
    docref = None
    surfown = None
    calledac = None
    ogsurv = None
    county = None
    state = None
    with arcpy.da.SearchCursor(boundarypoly.name,["SHAPE@","TRACT_NUMBER","RECORD_AGENT","DOC_REFERENCE","SURFACE_OWNER","CALLED_ACREAGE","ORIGINAL_SURVEY","COUNTY","STATE"],where,spatial_reference=sr) as sc:
        for row in sc:
            if row:
                centerlineshape = row[0]
                name = row[1]
                recagent = row[list(sc.fields).index("RECORD_AGENT")]
                docref = row[list(sc.fields).index("DOC_REFERENCE")]
                surfown = row[list(sc.fields).index("SURFACE_OWNER")]
                calledac = row[list(sc.fields).index("CALLED_ACREAGE")]
                ogsurv = row[list(sc.fields).index("ORIGINAL_SURVEY")]
                county = row[list(sc.fields).index("COUNTY")]
                state = row[list(sc.fields).index("STATE")]
                
    try:
        del sc
    except:
        print("No Delete Cursor")
    return centerlineshape,name,recagent,docref,surfown,calledac,ogsurv,county,state
 


def main():
    arcpy.env.addOutputsToMap=False
    aprx = arcpy.mp.ArcGISProject('current')
    m = aprx.activeMap
    centerline = m.listLayers("CENTERLINE")
    if len(centerline)<1: return "No Centerline Layer"
    centerline = centerline[0]
    boundarypoly = m.listLayers("BOUNDARY_POLY")
    if len(boundarypoly)<1: return "No Boundary Layer"
    boundarypoly = boundarypoly[0]

    desc = arcpy.Describe(centerline.name)
    sr = desc.spatialReference

    wrkspc = centerline.connectionProperties['connection_info']['database']
    wrkspcfolder = os.path.abspath(os.path.join(wrkspc, os.pardir))
    cen = "New Pipeline Project"
    path = r"D:\Red_Oak_Project\workbooks\RunList.xlsx"
    wrkbook = xlrd.open_workbook(path)
    tractlist=returnTractList(wrkbook)

    censhape,cenname =  getCenterlineShape(centerline,cen)
    if censhape is None: return "Pipeshape is Null, check CL Name"
    if tractlist:
        if len(tractlist)>0:
            for tractnumber in tractlist:
                boundrow = getBoundaryShape(boundarypoly,tractnumber) #boundaryshape,tractnumber,recagent,docref,surfown,calledac,ogsurv,county,state 
                if len(boundrow)<1: return "No Boundary Found For {0}".format(tractnumber)
                boundaryshape = boundrow[0]
                if boundaryshape is None: return "No Boundary Found For {0}".format(tractnumber)
                tractcenterline = censhape.intersect(boundaryshape,2)
                pglength,direction = returnPageLength(tractcenterline)
                if pglength>6304:
                    message=pmm.platmultipage(cenname,tractnumber,direction)
                    print(message)
                else:
                    message=pms.platsinglepage(cenname,tractnumber)
                    print(message)