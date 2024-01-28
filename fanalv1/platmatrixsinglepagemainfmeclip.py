#!/usr/bin/python
# -*- coding: utf-8 -*-
import arcpy
import xlrd
import arcpy
from math import sqrt
import arcpy
import os
import sys
import xlwt
import math
import shutil
import numpy as np
import platomaticfuntions as pf
import dimension as dm
global scaledict
scaledict = {.01:"1200",.005:"2400",.0025:"4800",.002:"6000",.001:"12000"}
global fmelocation
sys.path.append(r"C:\Program Files\FME\fmeobjects\python36")
sys.path.append(r"C:\Program Files\FME")
sys.path.append(r"C:\Program Files\FME\plugins")
sys.path.append(r"C:\Program Files\FME\python")
sys.path.append(r"C:\Program Files\FME\python\python36")
import fmeobjects
global wrkspacepath
wrkspacepath = r"D:\Red_Oak_Project\workspaces\tractToPlatMultiPageFmeClip.fmw"
global clipwrkspacepath
clipwrkspacepath = r"D:\Red_Oak_Project\workspaces\dwgtotesttractcliptodwg.fmw"
global templatedir 
templatedir = r"D:\Red_Oak_Project\templates\singlepagetemplates"
global basemaster
basemaster=r"D:\Red_Oak_Project\master\RedOak Alignment Master_Closed_Polygons.dwg"

def platsinglepage(cen,tractnumber,bigscale=None,scale=None):
    arcpy.env.addOutputsToMap=False
    aprx = arcpy.mp.ArcGISProject('current')
    m = aprx.activeMap
    monument = m.listLayers("MONUMENT")
    if len(monument)<1: return "No Monument Layer"
    monument = monument[0]
    dimppoint = m.listLayers("DIM_POINT")
    if len(dimppoint)<1: return "No Dim Layer"
    dimppoint = dimppoint[0]
    boundarypoly = m.listLayers("BOUNDARY_POLY")
    if len(boundarypoly)<1: return "No Boundary Layer"
    boundarypoly = boundarypoly[0]
    centerline = m.listLayers("CENTERLINE")
    if len(centerline)<1: return "No Centerline Layer"
    centerline = centerline[0]
    rowpoly = m.listLayers("ROW_POLY")
    if len(rowpoly)<1: return "No ROW Layer"
    rowpoly = rowpoly[0]
    atwspoly = m.listLayers("ATWS_POLY")
    if len(atwspoly)<1: return "No ATWS Layer"
    atwspoly = atwspoly[0]
    twspoly = m.listLayers("TWS_POLY")
    if len(twspoly)<1: return "No TWS Layer"
    twspoly = twspoly[0]
    circles = m.listLayers("CIRCLES")
    if len(circles)<1: return "No TWS Layer"
    circles = circles[0]
    centerlineseg = m.listLayers("CENTERLINE_SEGMENT")
    if len(centerlineseg)<1: return "No Centerline Segment Layer"
    centerlineseg = centerlineseg[0]
    existingpipeline = m.listLayers("EXISTING_PIPELINE")
    if len(existingpipeline)<1: return "No Existing Pipeline Layer"
    existingpipeline = existingpipeline[0]
    dimensionlinelayer = m.listLayers("DIMENSION")
    if len(dimensionlinelayer)<1: return "No Dimension Line Layer"
    dimensionlinelayer = dimensionlinelayer[0]
    viewportpointlayer = m.listLayers("VIEWPORT_POINT")
    if len(viewportpointlayer)<1: return "No Viewport Points"
    viewportpointlayer = viewportpointlayer[0]
    abstractsurveylayer = m.listLayers("ABSTRACT_SURVEYS")
    if len(abstractsurveylayer)<1: return "Abstract Surveys"
    abstractsurveylayer = abstractsurveylayer[0]
    crosslabel = m.listLayers("CROSSINGLABELS")
    if len(crosslabel)<1: return "No Cross Label in Base"
    crosslabel=crosslabel[0]
    endpoints = m.listLayers("ENDPOINTS")
    if len(endpoints)<1: return "No Cross Label in Base"
    endpoints=endpoints[0]
    
    desc = arcpy.Describe(centerline.name)
    sr = desc.spatialReference

    wrkspc = centerline.connectionProperties['connection_info']['database']
    wrkspcfolder = os.path.abspath(os.path.join(wrkspc, os.pardir))

    resourcegdb = r"D:\Red_Oak_Project\pristine_gdb\RED_OAK_PROJECT_TEMPLATE.gdb"


    censhape,cenname =  pf.getCenterlineShape(centerline,cen)

    ### station tracts and run down the sorted list of tuples with meas_station and tract_number as tuple in list of tuple


    #################################################################################
    ##
    ## Create workbook to write table too.
    ##
    ###################################################################################
    writebook = xlwt.Workbook()
    writebookpath = os.path.join(wrkspcfolder,tractnumber+".xls")
    writesheet = writebook.add_sheet("LINETABLE")
    #################################################################################
    ##
    ## Create plat geodatabase
    ##
    ##################################################################################
    newgdb = os.path.join(wrkspcfolder,tractnumber+".gdb")
    dwgfilename=newgdb[:-4]+".dwg"
    parameters = {}
    parameters['GDBPATH']=newgdb
    parameters['TRACTNUMBER']=tractnumber
    parameters['DestDataset_REALDWG']=dwgfilename
    parameters['SCALE']=str(bigscale)
    parameters['DWGTEMPLATE']=pf.returnTemplateName(templatedir,bigscale)
    parameters['BASEMASTER']=basemaster
    wrkspcrunner = fmeobjects.FMEWorkspaceRunner()
    try:
        print("Creating Geometry")
        wrkspcrunner.runWithParameters(clipwrkspacepath,parameters)
    except:
        return "No Geometry Created With FME"

    try:
        del wrkspcrunner
    except:
        print("No Delete Workspace Runner")
    ##################################################################################
    ##
    ## Get boundary , function returns shape and tract number
    ##
    ###################################################################################
    boundrow = pf.getBoundaryShape(boundarypoly,tractnumber) #boundaryshape,tractnumber,recagent,docref,surfown,calledac,ogsurv,county,state 
    if len(boundrow)<1: return "No Boundary Found For {0}".format(tractnumber)
    boundaryshape = boundrow[0]
    #boundrow = [boundaryshape,tractnumber,recagent,docref,surfown,calledac,ogsurv,county,state]
    name = pf.insertCenterlineSegment(newgdb,boundarypoly,boundrow,cenname,tractnumber,bound=True)
    boundarybuffer = boundaryshape.buffer(5000)
    ajnames = pf.getAjoiningTracts(boundarypoly,boundarybuffer,cenname,newgdb,tractnumber)
    mons = pf.getROWPolyShape(boundarybuffer,monument,cenname,type="monument")
    for mon in mons:
        name = mon[1]
        point = mon[0]
        lbl = mon[2]
        name = pf.insertCenterlineSegment(newgdb,monument,point,cenname,tractnumber,label=lbl)
    ########################################################################################
    ##
    ## Get EXPL and write tract number and pack
    ##
    ###########################################################################################
    print("Return Any Existing Pipelines")
    expls = pf.getROWPolyShape(boundaryshape,existingpipeline,cenname,type="geomfeat")
    for expl in expls:
        name = expl[1]
        line = expl[0]
        explintersection = line.intersect(boundaryshape,2)
        if not explintersection is None:
            name = pf.insertCenterlineSegment(newgdb,existingpipeline,explintersection,cenname,tractnumber)

   
    #################################################################################
    ##
    ## Intersect ROW,ATWS and TWS poly w Boundary, if multigeometry insert each part as separate into gdb.
    ##
    ################################################################################
    
    rowintersections=[]
    rowintersections = pf.getROWPolyShape(None,rowpoly,cenname,ngdb=newgdb,type="newgdbfeat")
    atwsintersections=[]
    atwsintersections = pf.getROWPolyShape(None,atwspoly,cenname,ngdb=newgdb,type="newgdbfeat")
    twsintersections=[]
    twsintersections = pf.getROWPolyShape(None,twspoly,cenname,ngdb=newgdb,type="newgdbfeat")

    #####################################################################
    ##
    ## Get centeriine iside tract
    ##
    #####################################################################

    tractcenterlines = pf.getROWPolyShape(None,centerline,cenname,ngdb=newgdb,type="newgdbfeat")
    if len(tractcenterlines)<1: return "No Centerline"
    tractcenterlines=tractcenterlines[0]

    #######################################################################################
    ##
    ## get scale
    ##
    ##########################################################################################


    dimoffset = (1/scale)*.25
    ####################################################################################
    ##
    ## create tile
    ##
    #################################################################################

    x = tractcenterlines.centroid.X
    y = tractcenterlines.centroid.Y
    tileindexrow,tileindexfields = pf.createIndexPoly(newgdb,tractnumber,scale,x,y,sr)
    ################################################################################################
    ##
    ## use centerline, boundary poly to create corner ties, 
    ##
    ##################################################################################################
    print("Create Corner Ties")
    for prt in range(tractcenterlines.partCount):

        newdim = dm.Dimension()
        newdim.boundaryshape=boundaryshape
        newdim.dimensionlayer=dimensionlinelayer
        newdim.dimscale=dimoffset
        newdim.newgdb = newgdb

        tractcenterline = arcpy.Polyline(tractcenterlines.getPart(prt))
        polylines = pf.explodePoly(boundaryshape)
        firstline = pf.returnPropline(polylines,tractcenterline,firstpoint=True,findclose=False)
        if firstline is None: firstline=pf.returnPropline(polylines,tractcenterline,firstpoint=True,findclose=True)
        lastline = pf.returnPropline(polylines,tractcenterline,firstpoint=False,findclose=False)
        if lastline is None: lastline=pf.returnPropline(polylines,tractcenterline,firstpoint=False,findclose=True)

        firstlinescale = pf.scaleGeom(firstline,150)
        lastlinescale = pf.scaleGeom(lastline,150)

        firstbuff = firstline.buffer(20)
        lastbuff = lastline.buffer(20)
        firstlinebuff = firstlinescale.buffer(20)
        lastlinebuff = lastlinescale.buffer(20)
        wheremon = """UPPER(LABEL) NOT LIKE '%60D%'"""
        ends = pf.getROWPolyShape(None,endpoints,cenname,where=None,ngdb=newgdb,type="endpoint")
        startmons = pf.getROWPolyShape(firstbuff,monument,cenname,type="monument",where=wheremon)
        if len(startmons)<1:
            startmons = pf.getROWPolyShape(firstlinebuff,monument,cenname,type="monument",where=wheremon)
        if len(startmons)<1:
            startmons = pf.getROWPolyShape(firstlinebuff,monument,cenname,type="monument",where=None)
        if len(startmons)>0:
            pobgeom = pf.ReturnClosestMonument(ends,tractcenterline.firstPoint,switch="ends")
            if len(pobgeom)<1: return "No Start Point"
            pobgeom=pobgeom[1]
            startmon = pf.ReturnClosestMonument(startmons,pobgeom.getPart(0),switch="mons")
            montostartaz,montostartdis = pf.returnInverse(startmon[1].getPart(0),pobgeom.getPart(0))
            if montostartdis<(((1/scale)*1.3)):
                newdim.CreateSpiderDimension(startmon[1],pobgeom,pob=True)
            else:
                newdim.createCornerTies(montostartaz,startmon[1],pobgeom,pob=True)
        if len(startmons)>0: wheremon = """OBJECTID <> {} AND UPPER(LABEL) NOT LIKE '%60D%'""".format(startmon[-1])
        endmons = pf.getROWPolyShape(lastbuff,monument,cenname,type="monument",where=wheremon)
        if len(endmons)<1:
            if len(startmons)>0: wheremon="""OBJECTID <> {} AND UPPER(LABEL) NOT LIKE '%60D%'""".format(startmon[-1])
            if len(startmons)<1: wheremon= """UPPER(LABEL) NOT LIKE '%60D%'"""
            endmons = pf.getROWPolyShape(lastbuff,monument,cenname,type="monument",where=wheremon)
        if len(endmons)<1:
            endmons = pf.getROWPolyShape(lastlinebuff,monument,cenname,type="monument",where=wheremon)
        if len(endmons)<1:
            endmons = pf.getROWPolyShape(lastlinebuff,monument,cenname,type="monument",where=None)
        if len(endmons)>0:
            potgeom = pf.ReturnClosestMonument(ends,tractcenterline.lastPoint,switch="ends")
            if len(potgeom)<1: return "No End Point"
            potgeom=potgeom[1]
            endmon = pf.ReturnClosestMonument(endmons,potgeom.getPart(0),switch="mons")
            endtomonaz,endtomondis = pf.returnInverse(potgeom.getPart(0),endmon[1].getPart(0))
            if endtomondis<(((1/scale)*1.3)):
                newdim.CreateSpiderDimension(potgeom,endmon[1],pob=False)
            else:
                newdim.createCornerTies(endtomonaz,potgeom,endmon[1],pob=False)

    ###############################################################################################
    ##
    ## Insert ROW Dim points 25' either side of line.
    ##
    ##############################################################################################
        if len(rowintersections)>0:
            rowtotalshape = rowintersections[0]
            for i in range(1,len(rowintersections)):
                rowtotalshape = rowtotalshape.union(rowintersections[i])
            if len(twsintersections)>0:
                insertlist,char = pf.createROWDim(tractcenterlines,newgdb,dimppoint,viewportpointlayer,twss=twsintersections,row=rowtotalshape,char='A',tileindex=[tileindexrow[0]])
                listthrough,dictin = pf.createDimDicts(atwsintersections)
                if len(dictin)>0:
                    pf.createAtwsViewports(newgdb,viewportpointlayer,dictin,listthrough,char)
            if len(twsintersections)<1:
                insertlist,char = pf.createROWDim(tractcenterlines,newgdb,dimppoint,viewportpointlayer,char='A',tileindex=[tileindexrow[0]])
                listthrough,dictin = pf.createDimDicts(atwsintersections)
                if len(dictin)>0:
                    pf.createAtwsViewports(newgdb,viewportpointlayer,dictin,listthrough,char)
            if len(insertlist)<1: return "No Dims Created"
   
    ##############################################################################################
    ##
    ## Insert Circles at vertices, 
    ##
    #############################################################################################
    for x in range(tractcenterlines.partCount):
        pnts = tractcenterlines.getPart(x)
        for x in range(1,len(pnts)-1):
            centerpoint = arcpy.PointGeometry(arcpy.Point(pnts.getObject(x).X,pnts.getObject(x).Y),sr,True,True)
            name = pf.insertCenterlineSegment(newgdb,circles,centerpoint,cenname,tractnumber)

    ##############################################################################
    ##
    ## Insert line segments and write table to excel
    ##
    ################################################################################
    centerlinesegs = pf.getROWPolyShape(None,centerlineseg,cenname,ngdb=newgdb,type="censeg")
    if len(centerlinesegs)<1: return "No Centerline Segments"
    censegoids = pf.insertSegments(centerlinesegs,writesheet)
    ####################################################################
    ##
    ## Get drawing scale from tractcenterline
    ##
    #####################################################################

    

    ##########################################
    ##
    ## Create Index Tile and Export Annotation to database, remove index tile layer
    ##
    ###########################################
    print("Create Layout Views")
    
    tileindexshape=tileindexrow[0]
    tileindextractnumber=tileindexrow[1]
    tileid=tileindexrow[2]
    tilescale=tileindexrow[3]
    tileoid=tileindexrow[4]
    indexpoly  = os.path.join(newgdb,"TILE_INDEX")
    m.addDataFromPath(indexpoly)
    tileindex = m.listLayers("TILE_INDEX")
    if len(tileindex)<1: return "No Tile Index Layer"
    tileindex = tileindex[0]
    if not tileindex is None:
        newlabel = arcpy.TiledLabelsToAnnotation_cartography(m, polygon_index_layer="TILE_INDEX", out_geodatabase=newgdb, out_layer="LAND_", \
                                                                anno_suffix="LABEL_", reference_scale_value=scaledict[scale], reference_scale_field="", \
                                                                tile_id_field="", coordinate_sys_field="", map_rotation_field="", feature_linked="STANDARD", \
                                                                generate_unplaced_annotation="GENERATE_UNPLACED_ANNOTATION",which_layers="SINGLE_LAYER",single_layer=boundarypoly)
    newgrouplayer = newlabel.getOutput(0)
    labelname = [lyr.name for lyr in newgrouplayer.listLayers("*") if "POLYLABEL" in lyr.name]
    if len(labelname)<1: return "No Label Name to Rename"
    labelname = labelname[0]
    labelsplit = labelname.split("_")
    labelsplit = labelsplit[0]+"_"+labelsplit[1]
    land = m.listLayers("LAND_")
    if len(land)<1: return "No Land Labels"
    land = land[0]
    m.removeLayer(land)
    boundarypoly.showLabels=True
    arcpy.Rename_management(os.path.join(newgdb,labelname),os.path.join(newgdb,labelsplit))
    
    if not tileindex is None:
        newlabel = arcpy.TiledLabelsToAnnotation_cartography(m, polygon_index_layer="TILE_INDEX", out_geodatabase=newgdb, out_layer="LAND_", \
                                                            anno_suffix="LABEL_", reference_scale_value=scaledict[scale], reference_scale_field="", \
                                                            tile_id_field="", coordinate_sys_field="", map_rotation_field="", feature_linked="STANDARD", \
                                                            generate_unplaced_annotation="GENERATE_UNPLACED_ANNOTATION",which_layers="SINGLE_LAYER",single_layer=abstractsurveylayer.name)
    newgrouplayer = newlabel.getOutput(0)
    labelname = [lyr.name for lyr in newgrouplayer.listLayers("*") if "SURVEYSLABEL" in lyr.name]
    if len(labelname)<1: return "No Label Name to Rename"
    labelname = labelname[0]
    labelsplit = labelname.split("_")
    labelsplit = labelsplit[0]+"_"+labelsplit[1]
    land = m.listLayers("LAND_")
    if len(land)<1: return "No Land Labels"
    land = land[0]
    m.removeLayer(land)
    abstractsurveylayer.showLabels=True
    arcpy.Rename_management(os.path.join(newgdb,labelname),os.path.join(newgdb,labelsplit))

    if not tileindex is None:
        newlabel = arcpy.TiledLabelsToAnnotation_cartography(m, polygon_index_layer="TILE_INDEX", out_geodatabase=newgdb, out_layer="LAND_", \
                                                            anno_suffix="LABEL_", reference_scale_value=scaledict[scale], reference_scale_field="", \
                                                            tile_id_field="", coordinate_sys_field="", map_rotation_field="", feature_linked="STANDARD", \
                                                            generate_unplaced_annotation="GENERATE_UNPLACED_ANNOTATION",which_layers="SINGLE_LAYER",single_layer=crosslabel.name)
    newgrouplayer = newlabel.getOutput(0)
    labelname = [lyr.name for lyr in newgrouplayer.listLayers("*") if "CROSSINGLABELS" in lyr.name]
    if len(labelname)<1: return "No Label Name to Rename"
    labelname = labelname[0]
    labelsplit = labelname.split("_")
    labelsplit = labelsplit[0]
    land = m.listLayers("LAND_")
    if len(land)<1: return "No Land Labels"
    land = land[0]
    m.removeLayer(land)
    crosslabel.showLabels=True
    arcpy.Rename_management(os.path.join(newgdb,labelname),os.path.join(newgdb,labelsplit))
    m.removeLayer(tileindex)
    aprx.save()
    writebook.save(writebookpath)
    parameters = {}
    parameters['SourceDataset_FILEGDB']=newgdb
    parameters['SCALE']=str(1/scale)
    parameters['DWGNAME']=dwgfilename
    parameters['DWGTEMPLATE']=dwgfilename
    parameters['WRKBOOK']=writebookpath
    
    wrkspcrunner = fmeobjects.FMEWorkspaceRunner()
    try:
        print("Exporting To CAD Through FME")
        wrkspcrunner.runWithParameters(wrkspacepath,parameters)
    except fmeobjects.FMEException as err:
        print("FMEExeption {}".format(err))
    try:
        del newgdb
    except:
        print("No Delete New GDB")
    try:
        del wrkspcrunner
    except:
        print("No Delete Workspace Runner")

    return "Succesfully Created Plat For {}".format(tractnumber)
