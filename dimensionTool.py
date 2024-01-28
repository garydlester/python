#!/usr/bin/python
# -*- coding: utf-8 -*-
import arcpy
import sys
import math
from math import sqrt


def scaleGeom(geom,scale,reference=None):
    if geom is None: return None
    if reference is None:
        reference=geom.centroid
    refgeom = arcpy.PointGeometry(reference)
    newparts=[]
    for pind in range(geom.partCount):
        part = geom.getPart(pind)
        newPart = []
        for ptind in range(part.count):
            apnt = part.getObject(ptind)
            if apnt is None:
                newPart.append(apnt)
                continue
            bdist=refgeom.distanceTo(apnt)
            bpnt = arcpy.Point(reference.X+bdist,reference.Y)
            adist = refgeom.distanceTo(bpnt)
            cdist = arcpy.PointGeometry(apnt).distanceTo(bpnt)
            angle = math.acos((adist**2+bdist**2-cdist**2)/(2*adist*bdist))
            scaleDist = bdist * scale 
            if apnt.Y<reference.Y: angle = angle * -1

            scalex = scaleDist*math.cos(angle)+reference.X 
            scaley = scaleDist*math.sin(angle)+reference.Y
            #print(scalex,scaley)
            newPart.append(arcpy.Point(scalex,scaley))
        newparts.append(newPart)
    return arcpy.Geometry(geom.type,arcpy.Array(newparts),geom.spatialReference)

def copyParallelLeft(polyline,offsetWidth):
    sr = arcpy.SpatialReference(6578,6360)
    part = polyline.getPart(0)
    rArray = arcpy.Array()
    for pnt in part:
        dl = polyline.measureOnLine(pnt)
        ptX0 = polyline.positionAlongLine(dl-offsetWidth).firstPoint
        ptX1 = polyline.positionAlongLine(dl+offsetWidth).firstPoint
        dx = float(ptX1.X)-float(ptX0.X)
        dy = float(ptX1.Y)-float(ptX0.Y)
        lenv = math.hypot(dx,dy)
        sX = -dy * offsetWidth/lenv ; sY = dx*offsetWidth/lenv
        point = arcpy.Point(pnt.X+sX,pnt.Y+sY)
        rArray.add(point)
    offsetGeom = arcpy.Polyline(rArray,sr,False,False)
    return offsetGeom

def copyParallelRight(polyline,offsetWidth):
    sr = polyline.spatialReference
    part = polyline.getPart(0)
    rArray = arcpy.Array()
    for pnt in part:
        dl = polyline.measureOnLine(pnt)
        ptX0 = polyline.positionAlongLine(dl-offsetWidth).firstPoint
        ptX1 = polyline.positionAlongLine(dl+offsetWidth).firstPoint
        dx = float(ptX1.X)-float(ptX0.X)
        dy = float(ptX1.Y)-float(ptX0.Y)
        lenv = math.hypot(dx,dy)
        sX = -dy * offsetWidth/lenv ; sY = dx*offsetWidth/lenv
        point = arcpy.Point(pnt.X-sX,pnt.Y-sY)
        rArray.add(point)
    offsetGeom = arcpy.Polyline(rArray,sr,False,False)
    return offsetGeom

def ReturnFirstPoint(df,lyr,where):
    sr = df.spatialReference
    fields = ["SHAPE@"]
    with arcpy.da.SearchCursor(lyr.name,fields,where_clause=where,spatial_reference=sr) as sc:
        firstPoint = [row[0].firstPoint for row in sc][0]
        return firstPoint

def ReturnLastPoint(df,lyr,where):
    sr = df.spatialReference
    fields = ["SHAPE@"]
    with arcpy.da.SearchCursor(lyr.name,fields,where_clause=where,spatial_reference=sr) as sc:
        lastPoint = [row[0].lastPoint for row in sc][0]
        return lastPoint

def ReturnClosestMonument(sr,lyr,point,where=None):
    sr = sr
    fields = ["SHAPE@"]
    fields.append("OBJECTID")
    inX = point.X
    inY = point.Y
    closestDict = {}
    if where==None:
        with arcpy.da.SearchCursor(lyr.name,fields,spatial_reference=sr) as sc:
            for row in sc:
                point2 = row[0].getPart(0)
                outX = point2.X
                outY = point2.Y
                dist=sqrt((inX-outX)**2+(inY-outY)**2)
                closestDict[row[1]]=dist,point2
        minDist=(min(closestDict.items(), key=lambda x:x[1]))
        return minDist
    else:
        with arcpy.da.SearchCursor(lyr.name,fields,where_clause=where,spatial_reference=sr) as sc:
            for row in sc:
                point2 = row[0].getPart(0)
                outX = point2.X
                outY = point2.Y
                dist=sqrt((inX-outX)**2+(inY-outY)**2)
                closestDict[row[1]]=dist,point2
        minDist=(min(closestDict.items(), key=lambda x:x[1]))
        return minDist

def ReturnInverse(point1,point2):
        dX = point2.X-point1.X
        dY = point2.Y-point1.Y
        dis = sqrt(dX**2+dY**2)
        az = math.atan2(dX,dY)*180/math.pi
        if az<0:
            az = az+360
            return az,dis
        return az,dis
###  Need To Finish
def CreateDimension(point1,point2,sr,spider=False):
    pobaz,poblen = ReturnInverse(point1,point2)
    #print(pobaz,poblen)
    if spider==True:
        offsetaz = pobaz + 90
        if offsetaz>=360:
            offsetaz = offsetaz-360
        pobarray = arcpy.Array()
        pobarray.append(point1)
        pobarray.append(point2)
        pobpoly = arcpy.Polyline(pobarray,sr,False,False)
        pobmid = arcpy.PointGeometry(pobpoly.centroid,sr,False,False)
        pobpolymid = pobmid.pointFromAngleAndDistance(offsetaz,250,"PLANAR")
        pobarray = arcpy.Array()
        pobarray.append(point1)
        pobarray.append(pobpolymid.getPart(0))
        pobarray.append(point2)
        pobpoly = arcpy.Polyline(pobarray,sr,False,False)
        insertrow = pobpoly,pobpoly.firstPoint.X,pobpoly.firstPoint.Y,pobpoly.lastPoint.X,pobpoly.lastPoint.Y,poblen
        return insertrow
        
    else:
        pobarray = arcpy.Array()
        pobarray.append(point1)
        pobarray.append(point2)
        pobpoly = arcpy.Polyline(pobarray,sr,False,False)
        offsetpobpoly = copyParallelLeft(pobpoly,50)
        offsetpobpoly = scaleGeom(offsetpobpoly,.9)
        pobarray = arcpy.Array()
        pobarray.append(point1)
        pobarray.append(offsetpobpoly.firstPoint)
        pobarray.append(offsetpobpoly.lastPoint)
        pobarray.append(point2)
        pobpoly = arcpy.Polyline(pobarray,sr,False,False)
        insertrow = pobpoly,pobpoly.firstPoint.X,pobpoly.firstPoint.Y,pobpoly.lastPoint.X,pobpoly.lastPoint.Y,poblen
        return insertrow

def CreateDimensionWithSplineLeaders(point1,point2,sr,spline=False):
    pobaz,poblen = ReturnInverse(point1,point2)
        #print(pobaz,poblen)
    if spline==True:
        offsetaz = pobaz + 90
        if offsetaz>=360:
            offsetaz = offsetaz-360
        ### create baseline to throw point from middle
        pobarray = arcpy.Array()
        pobarray.append(point1)
        pobarray.append(point2)
        pobpoly = arcpy.Polyline(pobarray,sr,False,False)
        pobmid = arcpy.PointGeometry(pobpoly.centroid,sr,False,False)
        ## point1 to point2 point1 to point3
        print(poblen)
        pobpolymid = pobmid.pointFromAngleAndDistance(offsetaz,300,"PLANAR")


        #### pob spline geometry
        pobarray=arcpy.Array()
        pobarray.append(pobpolymid.getPart(0))
        pobarray.append(point1)
        print(pobarray)
        pobpoly = arcpy.Polyline(pobarray,sr,False,False)
        pobpolyscale = scaleGeom(pobpoly,.6)
        poboffset = copyParallelLeft(pobpoly,300)
        poboffsetscale = scaleGeom(poboffset,.6,pobpoly.lastPoint)
        poboffsetscalearray = poboffsetscale.getPart(0)
        print(poboffsetscalearray)
        #### poc spline geometry
        pocarray = arcpy.Array()
        pocarray.append(pobpolymid.getPart(0))
        pocarray.append(point2)
        pocpoly = arcpy.Polyline(pocarray,sr,False,False)
        pocpolyscale = scaleGeom(pocpoly,.6)
        
        pocoffset = copyParallelRight(pocpoly,300)
        pocoffsetscale = scaleGeom(pocoffset,.6,pocpoly.lastPoint)




        insertrow = pocpoly,pobpoly.firstPoint.X,pobpoly.firstPoint.Y,pobpoly.lastPoint.X,pobpoly.lastPoint.Y,poblen
        return insertrow
        
    else:
        pobarray = arcpy.Array()
        pobarray.append(point1)
        pobarray.append(point2)
        pobpoly = arcpy.Polyline(pobarray,sr,False,False)
        offsetpobpoly = copyParallelLeft(pobpoly,50)
        offsetpobpoly = scaleGeom(offsetpobpoly,.9)
        pobarray = arcpy.Array()
        pobarray.append(point1)
        pobarray.append(offsetpobpoly.firstPoint)
        pobarray.append(offsetpobpoly.lastPoint)
        pobarray.append(point2)
        pobpoly = arcpy.Polyline(pobarray,sr,False,False)
        insertrow = pobpoly,pobpoly.firstPoint.X,pobpoly.firstPoint.Y,pobpoly.lastPoint.X,pobpoly.lastPoint.Y,poblen
        return insertrow

    
def main():
    mxd = arcpy.mapping.MapDocument('current')
    platidentifier = "PLAT"
    leaseroadidentifier = 'NAME'
    leaseroadwhere = """{0} LIKE '%{1}%'""".format(leaseroadidentifier,platidentifier)
    fieldlist=["SHAPE@","BEGINX","BEGINY","ENDX","ENDY","DISTANCE"]
   

    
    
    ##leaseroadwhere = """NAME LIKE '%PLAT%'"""
    print(leaseroadwhere)
    df = arcpy.mapping.ListDataFrames(mxd,"Layers")
    if len(df)>0:
        df = df[0]
        print("Layers Frame, good")
    else:
        print("There is no data frame called Layers")
        return
    

    detailframe = arcpy.mapping.ListDataFrames(mxd,"Detail")
    if len(detailframe)>0:
        detailframe = detailframe[0]
        print("Detail Frame, good")
    else:
        print("There is no data frame called Detail")
        return

    dimensionlayer = arcpy.mapping.ListLayers(mxd,"DIMENSION",df)
    if len(dimensionlayer)>0:
        dimensionlayer = dimensionlayer[0]
        print("Dimension Layer, good")
    else:
        print("There is no layer called Dimension")
        return
    
    leaseroadlayer = arcpy.mapping.ListLayers(mxd,"LEASE_ROAD",df)
    if len(leaseroadlayer)>0:
        leaseroadlayer = leaseroadlayer[0]
        print("Lease Road Layer, good")
    else:
        print("There is no layer called Lease Road")
        return
    monumentlayer = arcpy.mapping.ListLayers(mxd,"MONUMENTS",df)
    if len(monumentlayer)>0:
        monumentlayer = monumentlayer[0]
        print("Monument Layer, good")
    else:
        print("There is no layer called Monumnets")
        return
  


    ic = arcpy.da.InsertCursor(dimensionlayer.name,fieldlist)

    firstpoint = ReturnFirstPoint(df,leaseroadlayer,leaseroadwhere)
    firstmon = ReturnClosestMonument(df.spatialReference,monumentlayer,firstpoint)[1][1]
    newrow = CreateDimension(firstmon,firstpoint,df.spatialReference)
    ic.insertRow(newrow)
    

    lastpoint = ReturnLastPoint(df,leaseroadlayer,leaseroadwhere)
    lastmon = ReturnClosestMonument(df.spatialReference,monumentlayer,lastpoint)[1][1]
    newrow = CreateDimension(lastpoint,lastmon,df.spatialReference)
    ic.insertRow(newrow)

    
   
    