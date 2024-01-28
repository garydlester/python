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



global scaledict
scaledict = {.01:"1200",.005:"2400",.0025:"4800",.002:"6000",.001:"12000"}



def ReturnClosestMonument(mons,point):
    inX = point.X
    inY = point.Y
    closest = []
    for row in mons:
        point2 = row[0].getPart(0)
        outX = point2.X
        outY = point2.Y
        dist=sqrt((inX-outX)**2+(inY-outY)**2)
        closest.append((dist,row[0],row[1],row[2]))
    minDist=(min(closest, key=lambda x:x[0]))
    return minDist
    
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


def explodePoly(boundaryshape):
    polylines = []
    shape = boundaryshape
    for i in range(shape.partCount):
        pnts = shape.getPart(i)
        for x in range(len(pnts)-1):
            array = arcpy.Array()
            point1 = arcpy.Point(pnts.getObject(x).X,pnts.getObject(x).Y,pnts.getObject(x).Z)
            point2 = arcpy.Point(pnts.getObject(x+1).X,pnts.getObject(x+1).Y,pnts.getObject(x+1).Z)
            array.add(point1)
            array.add(point2)
            polyline = arcpy.Polyline(array,boundaryshape.spatialReference,True,True)
            if not polyline in polylines: polylines.append(polyline)

    return polylines

def returnPropline(polylines,centerlineshape,firstpoint=False):
    for polyline in polylines:
        pntgeom = None
        if firstpoint==True: 
            pntgeom = arcpy.PointGeometry(arcpy.Point(centerlineshape.firstPoint.X,centerlineshape.firstPoint.Y,centerlineshape.firstPoint.Z),polyline.spatialReference,True,True)
            if polyline.disjoint(pntgeom)==False:
                return polyline
        if firstpoint==False: 
            pntgeom = arcpy.PointGeometry(arcpy.Point(centerlineshape.lastPoint.X,centerlineshape.lastPoint.Y,centerlineshape.lastPoint.Z),polyline.spatialReference,True,True)
            if polyline.disjoint(pntgeom)==False:
                return polyline
    return None

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
        bear = int(dmsbearing[0])
        minute = "{:02d}".format(int(dmsbearing[1]))
        second = "{:02d}".format(int(round(dmsbearing[2],0)))
        dmsbearing = u"""N{0}\xb0{1}'{2}"W""".format(bear,minute,second)
        return dmsbearing
    if azimuth>=0 and azimuth<=90:
        dmsbearing=ddToDms(azimuth)
        bear = int(dmsbearing[0])
        minute = "{:02d}".format(int(dmsbearing[1]))
        second = "{:02d}".format(int(round(dmsbearing[2],0)))
        dmsbearing = u"""N{0}\xb0{1}'{2}"E""".format(bear,minute,second)
        return dmsbearing
    if azimuth>90 and azimuth<=180:
        bearing= 180 - azimuth
        dmsbearing=ddToDms(bearing)
        bear = int(dmsbearing[0])
        minute = "{:02d}".format(int(dmsbearing[1]))
        second = "{:02d}".format(int(round(dmsbearing[2],0)))
        dmsbearing = u"""S{0}\xb0{1}'{2}"E""".format(bear,minute,second)
        return dmsbearing
    if azimuth>180 and azimuth<=270:
        bearing = azimuth-180
        dmsbearing=ddToDms(bearing)
        bear = int(dmsbearing[0])
        minute = "{:02d}".format(int(dmsbearing[1]))
        second = "{:02d}".format(int(round(dmsbearing[2],0)))
        dmsbearing = u"""S{0}\xb0{1}'{2}"W""".format(bear,minute,second)
        return dmsbearing

def createCornerTies(newgdb,az,dimoffsetscale,dimension,start,end,pob=False):
    az180 =  az-180
    if az180<0: az180=az180+360
    ic = arcpy.da.InsertCursor(os.path.join(newgdb,dimension.name),["SHAPE@","BEARING"])
    editor = arcpy.da.Editor(newgdb)
    editor.startEditing(True)
    editor.startOperation()
    if 270 <= az < 360 or 0 <= az < 90:
        if pob==True:
            az1 = az-45
            if az1<0: az1 = az1+360
            az2 = az-30
            if az2<0: az2 = az2+360
        else:
            az1 = az+45
            if az1>360: az1 = az1-360
            az2 = az+30
            if az2>360: az2 = az2-360
        point1 = start.getPart(0)
        print(point1)
        point2 = start.pointFromAngleAndDistance(az1,dimoffsetscale,"PLANAR")
        print(point2)
        point3 = start.pointFromAngleAndDistance(az2,(dimoffsetscale*2),"PLANAR")
        print(point3)
        array = arcpy.Array([point1,point2.getPart(0),point3.getPart(0)])
        dimline = arcpy.Polyline(array,start.spatialReference,True,True)
        row = (dimline,None)
        ic.insertRow(row)
        
        if pob==True:
            az3 = az180+45
            if az3>360: az3=az3-360
            az4 = az180+30
            if az4>360: az4=az4-360
        else:
            az3 = az180-45
            if az3<0: az3=az3+360
            az4 = az180-30
            if az4<0: az4=az4+360
        point4 = end.getPart(0)
        point5 = end.pointFromAngleAndDistance(az3,dimoffsetscale,"PLANAR")
        point6 = end.pointFromAngleAndDistance(az4,(dimoffsetscale*2),"PLANAR")
        array2 = arcpy.Array([point4,point5.getPart(0),point6.getPart(0)])
        dimline2 = arcpy.Polyline(array2,start.spatialReference,True,True)
        row2=(dimline2,None)
        ic.insertRow(row2)
        array3=arcpy.Array([point3.getPart(0),point6.getPart(0)])
        azb,dis = returnInverse(point1,point4)
        bear = returnBearingString(azb)
        bearstring = u"""{0} {1}'""".format(bear,round(dis,2))
        dimline3 = arcpy.Polyline(array3,start.spatialReference,True,True)
        row3=(dimline3,bearstring)
        ic.insertRow(row3)
    else:
        if pob==True:
            az1 = az+45
            if az1>360: az1=az1-360
            az2 = az+30
            if az2>360: az2 = az2-360
        else:
            az1 = az-45
            if az1<0: az1=az1+360
            az2 = az-30
            if az2<0: az2 = az2+360
        point1 = start.getPart(0)
        print(point1)
        point2 = start.pointFromAngleAndDistance(az1,dimoffsetscale,"PLANAR")
        print(point2)
        point3 = start.pointFromAngleAndDistance(az2,(dimoffsetscale*2),"PLANAR")
        print(point3)
        array = arcpy.Array([point1,point2.getPart(0),point3.getPart(0)])
        dimline = arcpy.Polyline(array,start.spatialReference,True,True)
        row = (dimline,None)
        ic.insertRow(row)
        
        if pob==True:
            az3 = az180-45
            if az3<0: az3=az3+360
            az4 = az180-30
            if az4<0: az4=az4+360
        else:
            az3 = az180+45
            if az3>360: az3=az3-360
            az4 = az180+30
            if az4>360: az4=az4-360
        point4 = end.getPart(0)
        point5 = end.pointFromAngleAndDistance(az3,dimoffsetscale,"PLANAR")
        point6 = end.pointFromAngleAndDistance(az4,(dimoffsetscale*2),"PLANAR")
        array2 = arcpy.Array([point4,point5.getPart(0),point6.getPart(0)])
        dimline2 = arcpy.Polyline(array2,start.spatialReference,True,True)
        row2=(dimline2,None)
        ic.insertRow(row2)
        array3=arcpy.Array([point3.getPart(0),point6.getPart(0)])
        dimline3 = arcpy.Polyline(array3,start.spatialReference,True,True)
        azb,dis = returnInverse(point1,point4)
        bear = returnBearingString(azb)
        bearstring = u"""{0} {1}'""".format(bear,round(dis,2))
        row3=(dimline3,bearstring)
        ic.insertRow(row3)
    editor.stopOperation()
    editor.stopEditing(True)
    try:
        del ic
    except:
        print("No Delete Cursor")
    try:
        del editor
    except:
        print("No Delete Editor")
    
def createROWDim(censhape,newgdb,dimensionlayer):
    editor =  arcpy.da.Editor(newgdb)
    editor.startEditing(True)
    editor.startOperation()
    with arcpy.da.InsertCursor(os.path.join(newgdb,dimensionlayer.name),["SHAPE@","DIM_GROUP","DIM_NUMBER","LABEL"]) as ic:
        for i in range(censhape.partCount):
            linelist = []
            array = censhape.getPart(i)
            for x in range(len(array)-1):
                cenpoint1 = array.getObject(x)
                cenpoint2 = array.getObject(x+1)
                newarray = arcpy.Array([cenpoint1,cenpoint2])
                centerseg = arcpy.Polyline(newarray,censhape.spatialReference,True,True)
                linelist.append((centerseg.length,centerseg))
            sortedlinelist = max(linelist, key=lambda x:x[0])
            insertlist = []
            if len(sortedlinelist)>0:
                censeg = sortedlinelist[1]
                az = returnAzimuth(censeg)
                osrang  = az + 90
                if osrang>=360: osrang = osrang-360
                oslang = az - 90
                if oslang<0: oslang = oslang+360
                pointR = censeg.positionAlongLine(.4,use_percentage=True)
                dimRgroup = i+1
                pos = 2
                label = "25'"
                insertlist.append((pointR,dimRgroup,pos,label))
                ic.insertRow([pointR,dimRgroup,pos,label])
                pointL = censeg.positionAlongLine(.5,use_percentage=True)
                dimLgroup = i+5
                pos = 1
                label = "25'"
                insertlist.append((pointL,dimLgroup,pos,label))
                ic.insertRow([pointL,dimLgroup,pos,label])
                ospointR = pointR.pointFromAngleAndDistance(osrang,25,method='PLANAR')
                dimRgroup = i+1
                pos = 1
                label = "25'"
                insertlist.append((ospointR,dimRgroup,pos,label))
                ic.insertRow([ospointR,dimRgroup,pos,label])
                ospointL = pointL.pointFromAngleAndDistance(oslang,25,method='PLANAR')
                dimLgroup = i+5
                pos = 2
                label = "25'"
                insertlist.append((ospointL,dimLgroup,pos,label))
                ic.insertRow([ospointL,dimLgroup,pos,label])
    editor.stopOperation()
    editor.stopEditing(True)
    try:
        del ic
    except:
        print("No Delete Cursor")

    try:
        del editor
    except:
        print("No Delete Editor")
    return insertlist


def insertSegments(newgdb,tractcenterline,centerlinesegment,tractnumber,sheet,ismultipart):
    sr = tractcenterline.spatialReference
    borders = xlwt.Borders()
    borders.left = 1
    borders.right = 1
    borders.bottom = 1
    borders.top = 1
    style = xlwt.XFStyle()
    style2 = style
    style2.alignment.horz = xlwt.Alignment.HORZ_RIGHT
    style2.borders = borders
    style.borders = borders
    style.alignment.wrap=1
    style.alignment.horz = xlwt.Alignment.HORZ_CENTER
    for x in range(3):
        if x==0:
            col = sheet.col(x)
            col.width = 6 * 256
        if x==1:
            col = sheet.col(x)
            col.width = 14 * 256
        if x==2:
            col = sheet.col(x)
            col.width = 11 * 256
    sheet.write_merge(0,0,0,2,"LINE TABLE",style)
    sheet.write(1,0,"LINE",style)
    sheet.write(1,1,"BEARING",style)
    sheet.write(1,2,"DIST",style)
    if ismultipart==False:
        editor = arcpy.da.Editor(newgdb)
        editor.startEditing(True)
        editor.startOperation()
        ic = arcpy.da.InsertCursor(os.path.join(newgdb,centerlinesegment.name),["SHAPE@","LINE_NUMBER","BEARING","TRACT_NUMBER"])
        pnts = tractcenterline.getPart(0)
        cnt = 2
        row = sheet.row(cnt) 
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
            sheet.write(cnt,0,"L{0}".format(cnt-1),style)
            sheet.write(cnt,1,bear,style)
            sheet.write(cnt,2,"{:.2f}'".format(round(polyShape.length,2)),style)
            newRow = polyShape,cnt,bear2,tractnumber
            insertrow=ic.insertRow(newRow)
            segoids.append(insertrow)
            cnt+=1
        sheet.write(cnt,1,"Total:",style2)
        sheet.write(cnt,2,"{:.2f}'".format(round(dist,2)),style)
        editor.stopOperation()
        editor.stopEditing(True)
        try:
            del ic
        except:
            print("No Delete Cursor")
        try:
            del editor
        except:
            print("No Delete Editor")
    else:
        editor = arcpy.da.Editor(newgdb)
        editor.startEditing(True)
        editor.startOperation()
        ic = arcpy.da.InsertCursor(os.path.join(newgdb,centerlinesegment.name),["SHAPE@","LINE_NUMBER","BEARING","TRACT_NUMBER"])
        cnt = 2
        row = sheet.row(cnt)
        dist = 0.000
        segoids=[]
        for i in range(tractcenterline.partCount):
            pnts = tractcenterline.getPart(i)
            for x in range(len(pnts)-1):
                point1 = pnts.getObject(x)
                point2 = pnts.getObject(x+1)
                polyArray = arcpy.Array()
                polyArray.add(point1)
                polyArray.add(point2)
                polyShape = arcpy.Polyline(polyArray,sr,True,True)
                dist = dist + polyShape.length
                az = returnAzimuth(polyShape)
                bear  = returnBearingString(az)
                bear2 = returnBearingString(az)
                print(bear)
                sheet.write(cnt,0,"L{0}".format(cnt-1),style)
                sheet.write(cnt,1,bear,style)
                sheet.write(cnt,2,"{:.2f}'".format(round(polyShape.length,2)),style)
                newRow = polyShape,cnt,bear2,tractnumber
                insertrow=ic.insertRow(newRow)
                segoids.append(insertrow)
                cnt+=1
        sheet.write(cnt,1,"Total:",style)
        sheet.write(cnt,2,"{:.2f}'".format(round(dist,2)),style2)
        editor.stopOperation()
        editor.stopEditing(True)          
        try:
            del ic
        except:
            print("No Delete Cursor")
        try:
            del editor
        except:
            print("No Delete Editor")
    return segoids

def lenToScale(length):
    scale = None
    if length>=0 and length<=725:
        scale = .01
        return scale
    if length>725 and length<=1450:
        scale = .005
        return scale
    if length>1450 and length<=2960:
        scale = .0025
        return scale
    if length>2960 and length<=3625:
        scale = .002
        return scale
    if length>3625:
        scale = .001
        return scale

def returnPlatScale(tractcenterline):
    scale = None
    dX = abs(tractcenterline.firstPoint.X-tractcenterline.lastPoint.X)
    dY = abs(tractcenterline.firstPoint.Y-tractcenterline.lastPoint.Y)
    if dY>=dX:
        scale = lenToScale(dY)
        return scale
    else:
        scale = lenToScale(dX)
        return scale
    
def returnHWFromSCale(scale):
    if scale == .01:
        h=725
        w=725
        return h,w
    if scale == .005:
        h=1450
        w=1450
        return h,w
    if scale == .0025:
        h=2960
        w=2960
        return h,w
    if scale == .002:
        h=3625
        w=3625
        return h,w
    if scale == .001:
        h=7250
        w=7250
        return h,w
    else:
        return None,None

def createIndexPoly(newgdb,tractname,scale,X,Y,sr):
    h,w=returnHWFromSCale(scale)
    if not h is None:
        editor = arcpy.da.Editor(newgdb)
        editor.startEditing(True)
        editor.startOperation()
        with arcpy.da.InsertCursor(os.path.join(newgdb,"TILE_INDEX"),["SHAPE@","TRACT_NUMBER","TILE_ID","TILE_SCALE"]) as ic:
            arrPnts = arcpy.Array()  
            # point 1  
            pnt = arcpy.Point(X-w/2,Y-h/2)  
            arrPnts.add(pnt)  
            # point 2  
            pnt = arcpy.Point(X-w/2,Y+h/2)  
            arrPnts.add(pnt)  
            # point 3  
            pnt = arcpy.Point(X+w/2,Y+h/2)  
            arrPnts.add(pnt)  
            # point 4  
            pnt = arcpy.Point(X+w/2,Y-h/2)  
            arrPnts.add(pnt)  
            # point 5 (close diamond)  
            pnt = arcpy.Point(X-w/2,Y-h/2)  
            arrPnts.add(pnt)  
            pol = arcpy.Polygon(arrPnts,sr,True,True)
            row = [pol,tractname,None,scale]
            oid = ic.insertRow(row)
        editor.stopOperation()
        editor.stopEditing(True)
        fields = ic.fields
        retrow = row+[oid]
        try:
            del ic
        except:
            print("No Delete Cursor")
        try:
            del editor
        except:
            print("No Delete Editor")
        return retrow,fields


def getBoundaryShape(boundarypoly,tractnumber):
    desc = arcpy.Describe(boundarypoly.name)
    sr = desc.spatialReference
    where = "TRACT_NUMBER = '{}'".format(tractnumber)
    centerlineshape = None 
    name = None 
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
            else:
                centerlineshape = None
                name = None
                recagent = None
                docref = None
                surfown = None
                calledac = None
                ogsurv = None
                county = None
                state = None
    try:
        del sc
    except:
        print("No Delete Cursor")
    return centerlineshape,name,recagent,docref,surfown,calledac,ogsurv,county,state
 
def getROWPolyShape(boundarypolyshape,rowpoly,cenname,monument=False):
    shapes = []
    name = None 
    desc = arcpy.Describe(rowpoly.name)
    sr = desc.spatialReference
    if monument==False:
        where = "CENTERLINE_NAME LIKE '{}'".format(cenname)
        with arcpy.da.SearchCursor(rowpoly.name,["SHAPE@","CENTERLINE_NAME"],where,spatial_reference=sr) as sc:
            for row in sc:
                if row[0].disjoint(boundarypolyshape)==False:
                    add = (row[0],row[1])
                    if not add in shapes:
                        shapes.append(add)
    else:
        where = "CENTERLINE_NAME LIKE '{}'".format(cenname)
        with arcpy.da.SearchCursor(rowpoly.name,["SHAPE@","CENTERLINE_NAME","LABEL"],where,spatial_reference=sr) as sc:
            for row in sc:
                if row[0].disjoint(boundarypolyshape)==False:
                    add = (row[0],row[1],row[2])
                    if not add in shapes:
                        shapes.append(add)
    try:
        del sc
    except:
        print("No Delete Cursor")
    return shapes

def getCenterlineShape(centerline,cenname):
    desc = arcpy.Describe(centerline.name)
    sr = desc.spatialReference
    where = "CENTERLINE_NAME LIKE '{}'".format(cenname)
    with arcpy.da.SearchCursor(centerline.name,["SHAPE@","CENTERLINE_NAME"],where,spatial_reference=sr) as sc:
        for row in sc:
            if row:
                centerlineshape = row[0]
                name = row[1]
            else:
                centerlineshape = None
                name = None
    try:
        del sc
    except:
        print("No Delete Cursor")
    return centerlineshape,name


def returnTractList(wrkbook):
    tractlist=[]
    wksht = wrkbook.sheet_by_name("TRACT_LIST")
    for x in range(1,wksht.nrows):
        cell = wksht.cell_value(x,0)
        if not cell is None and cell !="":
            if not cell.strip() in tractlist:
                tractlist.append(cell)

    return tractlist

     


def insertCenterlineSegment(newgdb,centerlinelayer,censhape,cenname,tractnumber,label=None,bound=False):
    editor = arcpy.da.Editor(newgdb)
    editor.startEditing(True)
    editor.startOperation()
    if label is None and bound==False:
        with arcpy.da.InsertCursor(os.path.join(newgdb,centerlinelayer.name),["SHAPE@","CENTERLINE_NAME"]) as ic:
            newrow = censhape,cenname+" "+tractnumber
            ic.insertRow(newrow)
    if not label is None and bound==False:
        with arcpy.da.InsertCursor(os.path.join(newgdb,centerlinelayer.name),["SHAPE@","CENTERLINE_NAME","LABEL"]) as ic:
            newrow = censhape,cenname+" "+tractnumber,label
            ic.insertRow(newrow)

    if bound==True:
        with arcpy.da.InsertCursor(os.path.join(newgdb,centerlinelayer.name),["SHAPE@","TRACT_NUMBER","RECORD_AGENT","DOC_REFERENCE","SURFACE_OWNER","CALLED_ACREAGE","ORIGINAL_SURVEY","COUNTY","STATE"]) as ic:
            newrow = censhape
            ic.insertRow(newrow)
    editor.stopOperation()
    editor.stopEditing(True)
    try:
        del ic
    except:
        print("No Delete Cursor")
    try:
        del editor
    except:
        print("No Delete Editor")
    return cenname+" "+tractnumber

def getAjoiningTracts(boundarylayer,boundarybuffer,cenname,newgdb,tractnumber):
    names = []
    where = """TRACT_NUMBER <> '{0}'""".format(tractnumber)
    desc  = arcpy.Describe(boundarylayer.name)
    sr = desc.spatialReference
    with  arcpy.da.SearchCursor(boundarylayer.name,["SHAPE@","TRACT_NUMBER","RECORD_AGENT","DOC_REFERENCE","SURFACE_OWNER","CALLED_ACREAGE","ORIGINAL_SURVEY","COUNTY","STATE"],where_clause=where,spatial_reference=sr) as sc:
        for row in sc:
            shape = row[0]
            if shape.disjoint(boundarybuffer)==False:
                name = insertCenterlineSegment(newgdb,boundarylayer,row,cenname,tractnumber,bound=True)
                if not name in names: names.append(name)
    try:
        del sc
    except:
        print("No Delete Cursor")

    return names

def main():
    mxd = arcpy.mapping.MapDocument('current')
    df = mxd.activeDataFrame
    monument = arcpy.mapping.ListLayers(mxd,"MONUMENT",df)
    if len(monument)<1: return "No Monument Layer"
    monument = monument[0]
    dimppoint = arcpy.mapping.ListLayers(mxd,"DIM_POINT",df)
    if len(dimppoint)<1: return "No Dim Layer"
    dimppoint = dimppoint[0]
    boundarypoly = arcpy.mapping.ListLayers(mxd,"BOUNDARY_POLY",df)
    if len(boundarypoly)<1: return "No Boundary Layer"
    boundarypoly = boundarypoly[0]
    centerline = arcpy.mapping.ListLayers(mxd,"CENTERLINE",df)
    if len(centerline)<1: return "No Centerline Layer"
    centerline = centerline[0]
    rowpoly = arcpy.mapping.ListLayers(mxd,"ROW_POLY",df)
    if len(rowpoly)<1: return "No ROW Layer"
    rowpoly = rowpoly[0]
    atwspoly = arcpy.mapping.ListLayers(mxd,"ATWS_POLY",df)
    if len(atwspoly)<1: return "No ATWS Layer"
    atwspoly = atwspoly[0]
    twspoly = arcpy.mapping.ListLayers(mxd,"TWS_POLY",df)
    if len(twspoly)<1: return "No TWS Layer"
    twspoly = twspoly[0]
    circles = arcpy.mapping.ListLayers(mxd,"CIRCLES",df)
    if len(circles)<1: return "No TWS Layer"
    circles = circles[0]
    centerlineseg = arcpy.mapping.ListLayers(mxd,"CENTERLINE_SEGMENT",df)
    if len(centerlineseg)<1: return "No Centerline Segment Layer"
    centerlineseg = centerlineseg[0]
    existingpipeline = arcpy.mapping.ListLayers(mxd,"EXISTING_PIPELINE",df)
    if len(existingpipeline)<1: return "No Existing Pipeline Layer"
    existingpipeline = existingpipeline[0]
    dimensionlinelayer = arcpy.mapping.ListLayers(mxd,"DIMENSION",df)
    if len(dimensionlinelayer)<1: return "No Dimension Line Layer"
    dimensionlinelayer = dimensionlinelayer[0]

    desc = arcpy.Describe(centerline.name)
    sr = desc.spatialReference

    wrkspc = centerline.workspacePath
    wrkspcfolder = os.path.abspath(os.path.join(wrkspc, os.pardir))
    print(wrkspcfolder)
    cenname = "Red Oak Pipeline"
    path = r"D:\Red_Oak_Project\workbooks\RunList.xlsx"

    resourcegdb = r"D:\Red_Oak_Project\pristine_gdb\RED_OAK_PROJECT_TEMPLATE.gdb"

    wrkbook = xlrd.open_workbook(path)
    tractlist=returnTractList(wrkbook)

    censhape,cenname =  getCenterlineShape(centerline,cenname)

    ### station tracts and run down the sorted list of tuples with meas_station and tract_number as tuple in list of tuple

    if tractlist:
        if len(tractlist)>0:
            for tractnumber in tractlist:
                #################################################################################
                ##
                ## Create workbook to write table too.
                ##
                ###################################################################################
                writebook = xlwt.Workbook()
                writebookpath = os.path.join(wrkspcfolder,tractnumber+".xls")
                writesheet = writebook.add_sheet(tractnumber)
                #################################################################################
                ##
                ## Create plat geodatabase
                ##
                ##################################################################################
                newgdb = os.path.join(wrkspcfolder,tractnumber+".gdb")
                if os.path.exists(newgdb):
                    shutil.rmtree(newgdb)
                    shutil.copytree(resourcegdb,newgdb)
                else:
                    shutil.copytree(resourcegdb,newgdb)

                ##################################################################################
                ##
                ## Get boundary , function returns shape and tract number
                ##
                ###################################################################################
                boundrow = getBoundaryShape(boundarypoly,tractnumber) #boundaryshape,tractnumber,recagent,docref,surfown,calledac,ogsurv,county,state 
                if len(boundrow)<1: return "No Boundary Found For {0}".format(tractnumber)
                boundaryshape = boundrow[0]
                #boundrow = [boundaryshape,tractnumber,recagent,docref,surfown,calledac,ogsurv,county,state]
                name = insertCenterlineSegment(newgdb,boundarypoly,boundrow,cenname,tractnumber,bound=True)
                boundarybuffer = boundaryshape.buffer(5000)
                ajnames = getAjoiningTracts(boundarypoly,boundarybuffer,cenname,newgdb,tractnumber)
                print(ajnames)
                mons = getROWPolyShape(boundarybuffer,monument,cenname,monument=True)
                for mon in mons:
                    name = mon[1]
                    point = mon[0]
                    lbl = mon[2]
                    name = insertCenterlineSegment(newgdb,monument,point,cenname,tractnumber,label=lbl)
                ########################################################################################
                ##
                ## Get EXPL and write tract number and pack
                ##
                ###########################################################################################
                expls = getROWPolyShape(boundaryshape,existingpipeline,cenname,monument=False)
                for expl in expls:
                    name = expl[1]
                    line = expl[0]
                    explintersection = line.intersect(boundaryshape,2)
                    if not explintersection is None:
                        name = insertCenterlineSegment(newgdb,existingpipeline,explintersection,cenname,tractnumber)

                ##################################################################################
                ##
                ## These are lists of polyshapes that intersect the boundary poly they need to be clipped
                ##
                ##################################################################################
                rows = getROWPolyShape(boundaryshape,rowpoly,cenname)
                atws = getROWPolyShape(boundaryshape,atwspoly,cenname)
                twss = getROWPolyShape(boundaryshape,twspoly,cenname)

                #################################################################################
                ##
                ## Intersect ROW,ATWS and TWS poly w Boundary, if multigeometry insert each part as separate into gdb.
                ##
                ################################################################################
                rowintersections=[]
                for poly in rows:
                    name = None
                    polygon = poly[0]
                    name  = poly[1]
                    print(polygon)
                    rowintersection = polygon.intersect(boundaryshape,4)
                    if rowintersection: name = insertCenterlineSegment(newgdb,rowpoly,rowintersection,cenname,tractnumber)
                    if not name is None and name not in rowintersections: rowintersections.append(name)

                #######################################################################################
                ########################################################################################
                for poly in atws:
                    name = None
                    polygon = poly[0]
                    name  = poly[1]
                    print(polygon)
                    rowintersection = polygon.intersect(boundaryshape,4)
                    if rowintersection: name = insertCenterlineSegment(newgdb,atwspoly,rowintersection,cenname,tractnumber)
                    if not name is None and name not in rowintersections: rowintersections.append(name)

                ##############################################################################################
                ############################################################################################
                for poly in twss:
                    name = None
                    polygon = poly[0]
                    name  = poly[1]
                    print(polygon)
                    rowintersection = polygon.intersect(boundaryshape,4)
                    if rowintersection: name = insertCenterlineSegment(newgdb,twspoly,rowintersection,cenname,tractnumber)
                    if not name is None and name not in rowintersections: rowintersections.append(name)

                #####################################################################
                ##
                ## Get centeriine iside tract
                ##
                #####################################################################
            
                tractcenterline = censhape.intersect(boundaryshape,2)
                scale = returnPlatScale(tractcenterline)
                dimoffset = (1/scale)*.1

                ################################################################################################
                ##
                ## use centerline, boundary poly to create corner ties, 
                ##
                ##################################################################################################

                polylines = explodePoly(boundaryshape)
                firstline = returnPropline(polylines,tractcenterline,firstpoint=True)
                lastline = returnPropline(polylines,tractcenterline,firstpoint=False)

                firstlinescale = scaleGeom(firstline,150)
                lastlinescale = scaleGeom(lastline,150)

                firstlinebuff = firstlinescale.buffer(20)
                lastlinebuff = lastlinescale.buffer(20)

                startmons = getROWPolyShape(firstlinebuff,monument,cenname,monument=True)
                if len(startmons)>0:
                    startmon = ReturnClosestMonument(startmons,tractcenterline.firstPoint)
                    montostartaz,montostartdis = returnInverse(startmon[1].getPart(0),tractcenterline.firstPoint)
                    pobgeom = arcpy.PointGeometry(tractcenterline.firstPoint,tractcenterline.spatialReference,True,True)
                    createCornerTies(newgdb,montostartaz,dimoffset,dimensionlinelayer,startmon[1],pobgeom,pob=True)

                endmons = getROWPolyShape(lastlinebuff,monument,cenname,monument=True)
                if len(endmons)>0:
                    endmon = ReturnClosestMonument(endmons,tractcenterline.lastPoint)
                    endtomonaz,endtomondis = returnInverse(tractcenterline.lastPoint,endmon[1].getPart(0))
                    potgeom = arcpy.PointGeometry(tractcenterline.lastPoint,tractcenterline.spatialReference,True,True)
                    createCornerTies(newgdb,endtomonaz,dimoffset,dimensionlinelayer,potgeom,endmon[1],pob=False)

                ###############################################################################################
                ##
                ## Insert ROW Dim points 25' either side of line.
                ##
                ##############################################################################################
                insertlist = createROWDim(tractcenterline,newgdb,dimppoint)
                if len(insertlist)<1: return "No Tract Centerline Created"
                #############################################################################################
                ##
                ## Insert segments and create line tables
                ##
                ###########################################################################################
                newtractname = insertCenterlineSegment(newgdb,centerline,tractcenterline,cenname,tractnumber)
                if len(newtractname)<1: return "Segments were not inserted"
                ##############################################################################################
                ##
                ## Insert Circles at vertices, 
                ##
                #############################################################################################
                for x in range(tractcenterline.partCount):
                    pnts = tractcenterline.getPart(x)
                    for x in range(1,len(pnts)-1):
                        centerpoint = arcpy.PointGeometry(arcpy.Point(pnts.getObject(x).X,pnts.getObject(x).Y),sr,True,True)
                        name = insertCenterlineSegment(newgdb,circles,centerpoint,cenname,tractnumber)
                        print(name)

                ##############################################################################
                ##
                ## Insert line segments and write table to excel
                ##
                ################################################################################
                
                censegoids = insertSegments(newgdb,tractcenterline,centerlineseg,tractnumber,writesheet,tractcenterline.isMultipart)
                print(censegoids)
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
                x = tractcenterline.centroid.X 
                y = tractcenterline.centroid.Y 
                tileindexrow,tileindexfields = createIndexPoly(newgdb,tractnumber,scale,x,y,sr)
                tileindexshape=tileindexrow[0]
                tileindextractnumber=tileindexrow[1]
                tileid=tileindexrow[2]
                tilescale=tileindexrow[3]
                tileoid=tileindexrow[4]
                indexpoly  = os.path.join(newgdb,"TILE_INDEX")
                layerfile = arcpy.mapping.Layer(indexpoly)
                arcpy.mapping.AddLayer(df,layerfile,add_position='AUTO_ARRANGE')
                tileindex = arcpy.mapping.ListLayers(mxd,"TILE_INDEX",df)
                if len(tileindex)<1: return "No Tile Index Layer"
                tileindex = tileindex[0]
                if not tileindex is None:
                    arcpy.RefreshActiveView
                    arcpy.RefreshTOC
                    mxd.save()
                    arcpy.TiledLabelsToAnnotation_cartography(map_document=mxd.filePath, data_frame="Layers", polygon_index_layer="TILE_INDEX", out_geodatabase=newgdb, out_layer="LAND_", anno_suffix="LABEL_", reference_scale_value=scaledict[scale], reference_scale_field="", tile_id_field="", coordinate_sys_field="", map_rotation_field="", feature_linked="STANDARD", generate_unplaced_annotation="GENERATE_UNPLACED_ANNOTATION")
                
                arcpy.mapping.RemoveLayer(df,tileindex)
                arcpy.RefreshTOC()
                try:
                    del newgdb
                except:
                    print("No Delete New GDB")
                try:
                    del layerfile
                except:
                    print("No Delete Layer File")
                writebook.save(writebookpath)
