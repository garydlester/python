#!/usr/bin/python
# -*- coding: utf-8 -*-
import arcpy
from math import sqrt
import arcpy
import os
import sys
import xlwt
import math
import shutil
try:
    sys.path.append(r"D:\Projects\AutoWellGIS\Python")
except:
    pass

import dimensionTool
from datetime import datetime

def returnAzimuth(shape):
    point1 = shape.firstPoint
    point2 = shape.lastPoint
    dX = point2.X-point1.X
    dY = point2.Y-point1.Y
    az = math.atan2(dX,dY)*180/math.pi
    if az<0:
        az = az+360
        return az
    return az
    

def returnInverse(point1,point2):
    dX = point2.X-point1.X
    dY = point2.Y-point1.Y
    dis = sqrt(dX**2+dY**2)
    az = math.atan2(dX,dY)*180/math.pi
    if az<0:
        az = az+360
        return az,dis
    return az,dis

def ddToDms(dd):
    negative = dd<0
    dd=abs(dd)
    minutes,seconds = divmod(dd*3600,60)
    degrees,minutes = divmod(minutes,60)
    if negative:
        if degrees>0:
            degrees = -degrees
        elif minutes>0:
            minutes = -minutes
        else:
            seconds = -seconds
    return (int(degrees),int(minutes),round(seconds,3))


def returnBearingString(azimuth):
    bearing=None
    dmsbearing=None
    if azimuth>270 and azimuth<=360:
        bearing = 360 - azimuth
        dmsbearing=ddToDms(bearing)
        dmsbearing = u"""N {0}\xb0{1}'{2}" W""".format(int(dmsbearing[0]),int(dmsbearing[1]),int(round(dmsbearing[2],0)))
        return dmsbearing
    if azimuth>=0 and azimuth<=90:
        dmsbearing=ddToDms(azimuth)
        dmsbearing = u"""N {0}\xb0{1}'{2}" E""".format(int(dmsbearing[0]),int(dmsbearing[1]),int(round(dmsbearing[2],0)))
        return dmsbearing
    if azimuth>90 and azimuth<=180:
        bearing= 180 - azimuth
        dmsbearing=ddToDms(bearing)
        dmsbearing = u"""S {0}\xb0{1}'{2}" E""".format(int(dmsbearing[0]),int(dmsbearing[1]),int(round(dmsbearing[2],0)))
        return dmsbearing
    if azimuth>180 and azimuth<=270:
        bearing = azimuth-180
        dmsbearing=ddToDms(bearing)
        dmsbearing = u"""S {0}\xb0{1}'{2}" W""".format(int(dmsbearing[0]),int(dmsbearing[1]),int(round(dmsbearing[2],0)))
        return dmsbearing


def getLeaseRoadeShape(road):
    sr = arcpy.SpatialReference(6578,6360)
    where = "NAME LIKE '%PLAT%'"
    with arcpy.da.SearchCursor(road.name,["SHAPE@","NAME","WIDTH"],where,spatial_reference=sr) as sc:
        for row in sc:
            if row:
                roadShape = row
            else:
                roadShape = None
    return roadShape

def getLandOids(land,roadShape):
    oidList = []
    sr = arcpy.SpatialReference(6578)
    pipe = roadShape[0]
    with arcpy.da.SearchCursor(land.name,["SHAPE@","OBJECTID"],spatial_reference=sr) as sc:
        for row in sc:
            if row:
               shape = row[0]
               if shape.disjoint(pipe)==False:
                   if not row[1] in oidList:
                       oidList.append(row[1])
    return oidList


def insertRoads(land,road,roadShape):
    sr = arcpy.SpatialReference(6578,6360)
    landOids = getLandOids(land,roadShape)
    if landOids:
        if len(landOids)>0:
            for oid in landOids:
                print(oid)
                whereLand  = """OBJECTID = {}""".format(oid)
                with arcpy.da.InsertCursor(road.name,["SHAPE@","NAME","WIDTH","OBJECTID"]) as ic:
                    with arcpy.da.SearchCursor(land.name,"SHAPE@",whereLand,spatial_reference=sr) as sc:
                        for row in sc:
                            shape = row[0]
                            intersect = shape.intersect(roadShape[0],2)
                            print(intersect)
                            newRow = intersect,roadShape[1]+" LINE",roadShape[2],6000
                    ic.insertRow(newRow)
        
            return roadShape[1]
        else:
            return None


def getPipes(road,name):
    sr = arcpy.SpatialReference(6578,6360)
    where = """NAME = '{} LINE'""".format(name)
    roadDict = {}
    with arcpy.da.SearchCursor(road.name,["SHAPE@","OBJECTID","WIDTH"],where,spatial_reference=sr) as sc:
        for row in sc:
            if row:
                roadDict[row[1]]=row[0],row[2]
            else:
                roadDict = None
    return roadDict

def createElements(mxd):
    elementList={}
    elementDict={}
    listElements = arcpy.mapping.ListLayoutElements(mxd)
    for elem in listElements:
        if elem.name and elem.name!="":
            if not elem.name in elementList:
                elementList[elem.name]=elem.type

    if len(elementList)>0:
        for k,v in elementList.items():
            elementDict[k]=arcpy.mapping.ListLayoutElements(mxd,v,k)[0]
    return elementDict         

def getLandAtts(roadCenter,land):
    sr = arcpy.SpatialReference(6578,6360)
    pipe = roadCenter
    with arcpy.da.SearchCursor(land.name,["SHAPE@","*"],spatial_reference=sr) as sc:
        for row in sc:
            if row:
                shape = row[0]
                if shape.contains(pipe):
                    return row,sc.fields
                else:
                    pass
            else:
                return None,None
def proj2map(data_frame,proj_x,proj_y):
    """Convert projected coordinates to map coordinates"""
    # This code relies on the data_frame specified having
    # its anchor point at lower left

    #get the data frame dimensions in map units
    df_map_w = 7.899999999999636
    df_map_h = 8.5
    df_map_x = 0.28880000000026484
    df_map_y = 3.5

    #get the data frame projected coordinates
    df_min_x = data_frame.extent.XMin
    df_min_y = data_frame.extent.YMin
    df_max_x = data_frame.extent.XMax
    df_max_y = data_frame.extent.YMax
    df_proj_w = data_frame.extent.width
    df_proj_h = data_frame.extent.height

    #ensure the coordinates are in the dataframe
    if proj_x < df_min_x or proj_x > df_max_x:
        raise ValueError ('X coordinate is not within the data frame: %.1f - (%.1f, %.1f)' % (proj_x,df_min_x,df_max_x))

    if proj_y < df_min_y or proj_y > df_max_y:
        raise ValueError ('Y coordinate is not within the data frame: %.1f - (%.1f, %.1f)' % (proj_y,df_min_y,df_max_y))

    #scale the projected coordinates to map units from the lower left of the data frame
    map_x = (proj_x - df_min_x) / df_proj_w * df_map_w + df_map_x
    map_y = (proj_y - df_min_y) / df_proj_h * df_map_h + df_map_y

    return map_x,map_y


def addMeasureValues(leaseRoad):
    sr = leaseRoad.spatialReference
    with arcpy.da.UpdateCursor(leaseRoad.name,["SHAPE@","OBJECTID"],spatial_reference=sr) as sc:
        for row in sc:
            threeDsquaredDistCum=0
            pointArray=arcpy.Array()
            shp=row[0]
            line=shp.getPart(0)
            first=line.getObject(0)
            firstPoint=arcpy.Point()
            firstPoint.X=first.X
            firstPoint.Y=first.Y
            firstPoint.Z=first.Z
            firstPoint.M=0
            pointArray.add(firstPoint)     
            for i in range(1,len(line)):
                newPoint=arcpy.Point()
                point1= line.getObject(i)
                point2=line.getObject(i-1)
                X=point1.X
                Y=point1.Y
                Z=point1.Z
                X2=point2.X
                Y2=point2.Y
                Z2=point2.Z
                squaredZDist=(Z-Z2)**2
                squaredDist = (sqrt((X-X2)**2+(Y-Y2)**2))**2
                threeDsquaredDist=sqrt(squaredZDist+squaredDist)
                threeDsquaredDistCum=threeDsquaredDistCum+threeDsquaredDist
                newPoint.X=X
                newPoint.Y=Y
                newPoint.Z=Z
                newPoint.M=threeDsquaredDistCum
                pointArray.add(newPoint)
                print (threeDsquaredDistCum)
            newGeometry=arcpy.Polyline(pointArray,None,True)
            row=newGeometry,6666
            sc.updateRow(row)


def returnTractOrderDict(road,landPoly):
    tractList=[]
    orderDict = {}
    adjoinerDict={}
    firstTract = None
    lastTract = None
    with arcpy.da.SearchCursor(landPoly.name,["SHAPE@","OBJECTID"],spatial_reference=road.spatialReference) as sc:
        for row in sc:
            shape = row[0]
            centerPoint = shape.centroid
            if shape.disjoint(road.firstPoint)==False:
                firstTract=row[1]
            if shape.disjoint(road.lastPoint)==False:
                lastTract=row[1]
            if shape.disjoint(road.firstPoint)==True and shape.disjoint(road.lastPoint)==True:
                if shape.disjoint(road)==False:
                    orderDict[row[1]]=road.measureOnLine(centerPoint,False)


    while len(orderDict)>0:
        minDist=(min(orderDict.items(), key=lambda x:x[1]))
        tractList.append(minDist[0])
        orderDict.pop(minDist[0])
    tractList.insert(0,firstTract)
    if lastTract not in tractList:
        tractList.append(lastTract)
    for x  in range(len(tractList)):
        if x == 0 and len(tractList)>1:
            forward = tractList[1]
            backward = None
            adjoinerDict[tractList[x]]=[forward,backward]
        if x == len(tractList)-1 and len(tractList)>1:
            forward = None
            backward = tractList[len(tractList)-2]
            adjoinerDict[tractList[x]]=[forward,backward]
        if len(tractList)==1:
            forward=None
            backward=None
            adjoinerDict[tractList[x]]=[forward,backward]
        if x != 0 and x != len(tractList)-1:
            forward = tractList[x+1]
            backward = tractList[x-1]
            adjoinerDict[tractList[x]]=[forward,backward]
    return adjoinerDict,tractList
        

def insertSegments(shape,leaseRoadSegment,landDetails,workbook,k):
    sr = shape.spatialReference
    where = """PARENTOID={}""".format(k)
    style = xlwt.XFStyle()
    numStyle = xlwt.XFStyle()
    numStyle2=xlwt.XFStyle()
    style.alignment.wrap=1
    numStyle.alignment.wrap = 1
    numStyle.num_format_str='0'
    numStyle2.alignment.wrap = 1
    numStyle2.num_format_str='0.00'
    
    sec = landDetails[11]
    blk = landDetails[5]
    sheet = workbook.add_sheet("""SEC {0} BLK {1}""".format(sec,blk))
    sheet.write(0,0,"Line Number",style)
    sheet.write(0,1,"Bearing",style)
    sheet.write(0,2,"Distance",style)
    for x in range(3):
        col = sheet.col(x)
        col.width = 16 * 256
    
    ic = arcpy.da.InsertCursor(leaseRoadSegment.name,["OBJECTID","SHAPE@","LINE_NUMBER","BEARING_STRING","PARENTOID"])
    pnts = shape.getPart(0)
    cnt = 1
    row = sheet.row(cnt)
    row.height = 33 * 256
    dist = 0.00
    segoids=[]
    for x in range(len(pnts)-1):
        point1 = pnts.getObject(x)
        point2 = pnts.getObject(x+1)
        polyArray = arcpy.Array()
        polyArray.add(point1)
        polyArray.add(point2)
        polyShape = arcpy.Polyline(polyArray,sr,True,True)
        dist = dist+round(polyShape.length,2)
        az = returnAzimuth(polyShape)
        bear  = returnBearingString(az)
        bear2 = returnBearingString(az)
        print(bear)
        sheet.write(cnt,0,int(cnt),numStyle)
        sheet.write(cnt,1,bear,style)
        sheet.write(cnt,2,round(polyShape.length,2),numStyle2)
        newRow = 6000,polyShape,cnt,bear2,k
        insertrow=ic.insertRow(newRow)
        segoids.append(insertrow)
        cnt+=1
    sheet.write(cnt,1,"Total:",style)
    sheet.write(cnt,2,round(dist,2),numStyle2)
    leaseRoadSegment.definitionQuery=where
    return segoids

#################################################################################################################
##
##  Map Stuff
##
################################################################################################################
def main(companyName=None,srid=None,afe=None):
    exportlist=["MONUMENT","TX_BLOCK_POLY","LEASE_ROAD_SEGMENTS","ROAD","SECTION_LABELS","SEC_CORN_LABELS","LEASE_ROAD","DIMENSION","BOUNDARY_LINE","BOUNDARY_POLY"]
    mxd = arcpy.mapping.MapDocument('current')
    df = arcpy.mapping.ListDataFrames(mxd,"Layers")[0]
    arcpy.env.addOutputsToMap=False
    ###################################################################################################################
    ##
    ## data frames will exit with message on fail
    ##
    ###################################################################################################################
    detailFrame = arcpy.mapping.ListDataFrames(mxd,"Detail")
    detailFrame=detailFrame[0] if len(detailFrame)>0 else 0
    if detailFrame==0: return "No Detail Data Frame"
    detailEndFrame = arcpy.mapping.ListDataFrames(mxd,"DetailEnd")
    detailEndFrame = detailEndFrame[0] if len(detailEndFrame)>0 else 0
    if detailEndFrame==0: return "No Detail End Frame"


    ###################################################################################################################
    ##
    ##  Layers will exit with message on fail
    ##
    ####################################################################################################################

    boundaryPoly = arcpy.mapping.ListLayers(mxd,"BOUNDARY_POLY",df)
    boundaryPoly = boundaryPoly[0] if len(boundaryPoly)>0 else 0
    if boundaryPoly==0: return "No Bounday Poly"
    boundaryLine = arcpy.mapping.ListLayers(mxd,"BOUNDARY_LINE",df)
    boundaryLine = boundaryLine[0] if len(boundaryLine)>0 else 0
    if boundaryLine==0: return "No Boundary Line"
    monument = arcpy.mapping.ListLayers(mxd,"MONUMENTS",df)
    monument = monument[0] if len(monument)>0 else 0
    if monument==0: return "No Monument"
    dimension = arcpy.mapping.ListLayers(mxd,"DIMENSION",df)
    dimension = dimension[0] if len(dimension)>0 else 0
    if dimension==0: return "No Dimension"
    detaildimension = arcpy.mapping.ListLayers(mxd,"DIMENSION",detailFrame)
    detaildimension = detaildimension[0] if len(detaildimension)>0 else 0
    if detaildimension == 0: return "No Detail Dimension"
    leaseRoad = arcpy.mapping.ListLayers(mxd,"LEASE_ROAD",df)[0]
    detailenddimension = arcpy.mapping.ListLayers(mxd,"DIMENSION",detailEndFrame)
    detailenddimension = detailenddimension[0] if len(detailenddimension)>0 else 0
    if detailenddimension == 0: return "No Detail End Dimension"
    leaseRoad = arcpy.mapping.ListLayers(mxd,"LEASE_ROAD",df)[0]
    leaseroaddetail = arcpy.mapping.ListLayers(mxd,"LEASE_ROAD",detailFrame)
    leaseroaddetail=leaseroaddetail[0] if len(leaseroaddetail)>0 else 0
    if leaseroaddetail==0: return "No Lease Road Detail"
    leaseRoadSegment = arcpy.mapping.ListLayers(mxd,"LEASE_ROAD_SEGMENTS",df)
    leaseRoadSegment=leaseRoadSegment[0] if len(leaseRoadSegment)>0 else 0
    if leaseRoadSegment==0: return "No Lease Road Segment"
    detailLeaseRoad = arcpy.mapping.ListLayers(mxd,"LEASE_ROAD",detailFrame)
    detailLeaseRoad=detailLeaseRoad[0] if len(detailLeaseRoad)>0 else 0
    if detailLeaseRoad==0: return "No Detail Lease Road"

    fieldlist=["SHAPE@","BEGINX","BEGINY","ENDX","ENDY","DISTANCE"]
    wrkspc = leaseRoad.workspacePath

    parentDir = os.path.abspath(os.path.join(wrkspc, os.pardir))

    fileGdb = os.path.join(os.path.abspath(os.path.join(boundaryPoly.workspacePath, os.pardir)))



    roadShape = getLeaseRoadeShape(leaseRoad)
    print(roadShape)

    name = insertRoads(boundaryPoly,leaseRoad,roadShape)
    print(name)
    roadDict = getPipes(leaseRoad,name)

    adjoinerDict,tractList = returnTractOrderDict(roadShape[0],boundaryPoly)

    workbook = xlwt.Workbook(encoding='utf-8')
    for k,v in roadDict.items():
        shape = v[0]
        if shape:
            sr = shape.spatialReference
            roadDef = """OBJECTID = {}""".format(k)
            leaseRoad.definitionQuery=roadDef
            detailLeaseRoad.definitionQuery=roadDef
            ic = arcpy.da.InsertCursor(dimension.name,fieldlist+["PARENTOID"])
            #####################################################################################################################################
            ##
            ## Begin insert dimension
            ##
            #####################################################################################################################################
            firstpoint = dimensionTool.ReturnFirstPoint(df,leaseRoad,roadDef)
            firstmon = dimensionTool.ReturnClosestMonument(df.spatialReference,monument,firstpoint)
            firstmonoid = firstmon[0]
            ## insert next dimension but do not call to same monumnet so pass where clause
            where = """OBJECTID <> {0}""".format(firstmonoid)
            firstmon = firstmon[1][1] ## first monument point
            lastpoint = dimensionTool.ReturnLastPoint(df,leaseRoad,roadDef)
            lastmon = dimensionTool.ReturnClosestMonument(df.spatialReference,monument,lastpoint,where=where)[1][1] ## last monument point, use these two points and line end points to make  dimension
            dimoids=[]
            if returnInverse(firstmon,firstpoint)[1]<100: # if the dim distance is less than 100 feet use spider dimension use detail window
                newrow = dimensionTool.CreateDimension(firstmon,firstpoint,df.spatialReference,spider=True)
                newrow = newrow + (k,)
                newoid = ic.insertRow(newrow)
                if not newoid in dimoids: dimoids.append(newoid)

            if returnInverse(firstmon,firstpoint)[1]<1200 and returnInverse(firstmon,firstpoint)[1]>=100: # if the dim distance is between 100 and 1200 feet use detail window
                newrow = dimensionTool.CreateDimension(firstmon,firstpoint,df.spatialReference)
                newrow = newrow + (k,)
                newoid = ic.insertRow(newrow)
                if not newoid in dimoids: dimoids.append(newoid)

            else: 
                newrow = dimensionTool.CreateDimension(firstmon,firstpoint,df.spatialReference)
                newrow = newrow + (k,)
                newoid = ic.insertRow(newrow)
                if not newoid in dimoids: dimoids.append(newoid)
            
            if returnInverse(lastpoint,lastmon)[1]<100:
                newrow = dimensionTool.CreateDimension(lastpoint,lastmon,df.spatialReference,spider=True)
                newrow = newrow + (k,)
                newoid = ic.insertRow(newrow)
                if not newoid in dimoids: dimoids.append(newoid)

            if returnInverse(lastpoint,lastmon)[1]<1200 and returnInverse(lastpoint,lastmon)[1]>=100:
                newrow = dimensionTool.CreateDimension(lastpoint,lastmon,df.spatialReference)
                newrow = newrow + (k,)
                newoid = ic.insertRow(newrow)
                if not newoid in dimoids: dimoids.append(newoid)
            else:
                newrow = dimensionTool.CreateDimension(lastpoint,lastmon,df.spatialReference)
                newrow = newrow + (k,)
                newoid = ic.insertRow(newrow)
                if not newoid in dimoids: dimoids.append(newoid)
            if len(dimoids)>0: selectdim = dimension.setSelectionSet("NEW",dimoids)
            print(selectdim)
            ### need to get dimension property for length here, then apply approriate methods to  clone data frame if necessary.
            arcpy.AddMessage("yes")
            center = arcpy.PointGeometry(shape.centroid,sr,True,True)
            landDetails,cursor = getLandAtts(center,boundaryPoly)
            print(landDetails)
            if landDetails:
                segoids = insertSegments(shape,leaseRoadSegment,landDetails,workbook,k)
                select = leaseRoad.setSelectionSet("NEW",{k})
                selectsegs = leaseRoadSegment.setSelectionSet("NEW",segoids)
                
                ### Need to change everything to run of tract number

                gdbName = """SEC_{0}_BLK_{1}""".format(landDetails[11],landDetails[5])
                if os.path.exists(os.path.join(parentDir,gdbName+".gdb")):
                    shutil.rmtree(os.path.join(parentDir,gdbName+".gdb"))
                    newGdb = arcpy.CreateFileGDB_management(parentDir,gdbName,"CURRENT")
                    newGdb = newGdb.getOutput(0)
                    for lyr in arcpy.mapping.ListLayers(mxd,"*",df):
                        if lyr.name in exportlist:
                            newFc = os.path.join(newGdb,lyr.name)
                            arcpy.CopyFeatures_management(lyr.name,newFc)
                else:
                    newGdb = arcpy.CreateFileGDB_management(parentDir,gdbName,"CURRENT")
                    newGdb = newGdb.getOutput(0)
                    for lyr in arcpy.mapping.ListLayers(mxd,"*",df):
                        if lyr.name in exportlist:
                            newFc = os.path.join(newGdb,lyr.name)
                            arcpy.CopyFeatures_management(lyr.name,newFc)

            #######################################################
            ######
            ##### set extents and fill in map elements also create dimensions
            ######
            ######################################################

        else:
            print("Shape is None")
            arcpy.AddMessage("Shape is None")
            return "There is something wrong with the centerline"
    leaseRoad.definitionQuery=""
    leaseRoadSegment.definitionQuery=""
    newBook = os.path.join(parentDir,name+".xls")
    workbook.save(newBook)
    return "Successful Run"
            
#returnmessage=main(companyName = arcpy.GetParameterAsText(0),srid = arcpy.GetParameterAsText(1),afe = arcpy.GetParameterAsText(2))
returnmessage=main(companyName = "New Company",srid = "12345",afe = "A222555")
print(returnmessage )