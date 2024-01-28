#!/usr/bin/python
# -*- coding: utf-8 -*-
import math
import arcpy
import xlrd
from math import sqrt
import os
import sys
import xlwt
import shutil
import numpy as np

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
wrkspacepath = r"D:\Red_Oak_Project\workspaces\tractToPlatMultiPage.fmw"



def returnPageLength(centerlineshape):
    dx = abs(centerlineshape.firstPoint.X-centerlineshape.lastPoint.X)
    dy = abs(centerlineshape.firstPoint.Y-centerlineshape.lastPoint.Y)
    if dx>=dy: return dx
    if dy>dx: return dy

def closest(dist):
    if 0<dist<=2175: return 300
    if 2175<dist<=2900: return 400
    if 2900<dist<=3625: return 500
    if 6525<dist<=7250:
        return 1000
    else:
        return None
 
def findPageRange(length):
    n=100
    length = length*1.15
    for x in range(1,n):
        newscale1=length/x
        scalevalue = closest(newscale1)
        if scalevalue:
            pages = math.ceil(length/(7.25*scalevalue))
            pagelength = length/pages
            return(pagelength,pages,scalevalue)
  
    scalevalue=1000
    pages = math.ceil(length/(7.25*scalevalue))
    pagelength = (length/1.15)/pages
    return(pagelength,pages,scalevalue)

def createAtwsViewports(newgdb,viewportpointlayer,dictin,listthrough,char):
    editor = arcpy.da.Editor(newgdb)
    editor.startEditing(True)
    editor.startOperation()
    with arcpy.da.InsertCursor(os.path.join(newgdb,viewportpointlayer.name),["SHAPE@","DETAIL_LETTER","DETAIL_SIZE"]) as ic:
        for item in listthrough:
            k=item[0]
            v=item[1]
            size = 1
            orderinc = 1
            cumcntr = 0
            cumx = 0
            cumy = 0
            if k in dictin.keys():
                dictin.pop(k)
                shape1 = v
                cumx = cumx+shape1.centroid.X
                cumy = cumy+shape1.centroid.Y
                cumcntr+=1
                for item2 in listthrough:
                    k2  = item2[0]
                    v2 = item2[1]
                    if k2 in dictin.keys():
                        shape2 = v2
                        if shape1.distanceTo(shape2)<300:
                            size = 4
                            cumx = cumx+shape2.centroid.X 
                            cumy = cumy+shape2.centroid.Y
                            cumcntr+=1
                            dictin.pop(k2)
                unionx = cumx/cumcntr
                uniony = cumy/cumcntr
                viewpoint = arcpy.PointGeometry(arcpy.Point(unionx,uniony,0),shape1.spatialReference,True,True)
                char = chr(ord(char)+orderinc)
                orderinc+=1
                viewrow = [viewpoint,char,size]
                ic.insertRow(viewrow)
    editor.stopOperation()
    editor.stopEditing(True)
    try:
        del editor,ic
    except:
        print("No Delete Editor Or Cursor")


def createDimDicts(atwsintersections):
    listthrough=[]
    dictin={}
    cntr=1
    for atwsintersect in atwsintersections:
        for x in range(atwsintersect.partCount):
            polygon = arcpy.Polygon(atwsintersect.getPart(x),atwsintersect.spatialReference,True,True)
            dictin[cntr]=polygon
            listthrough.append((cntr,polygon))
            cntr+=1
    
    return listthrough,dictin

def convertToNumpy(geom):
    pointlist=[]
    if geom.type=='point':
        nparray = np.array([geom.getPart(0).X,geom.getPart(0).Y],dtype='float64')
        return nparray
    if  geom.type=='polyline':
        for pnt in geom.getPart(0):
            pointlist.append([pnt.X,pnt.Y])
        nparray = np.array(pointlist,dtype='float64')
        return nparray

def ReturnClosestMonument(mons,point):
    inX = point.X
    inY = point.Y
    closest = []
    for row in mons:
        point2 = row[0].getPart(0)
        outX = point2.X
        outY = point2.Y
        dist=sqrt((inX-outX)**2+(inY-outY)**2)
        closest.append((dist,row[0],row[1],row[2],row[3]))
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

def returnPropline(polylines,centerlineshape,firstpoint=False,findclose=False):
    distances = []
    if findclose==False:
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
    else:
        for polyline in polylines:
            pntgeom = None
            if firstpoint==True: 
                pntgeom = arcpy.PointGeometry(arcpy.Point(centerlineshape.firstPoint.X,centerlineshape.firstPoint.Y,centerlineshape.firstPoint.Z),polyline.spatialReference,True,True)
                dist = polyline.distanceTo(pntgeom)
                distances.append((dist,polyline))

            if firstpoint==False: 
                pntgeom = arcpy.PointGeometry(arcpy.Point(centerlineshape.lastPoint.X,centerlineshape.lastPoint.Y,centerlineshape.lastPoint.Z),polyline.spatialReference,True,True)
                dist = polyline.distanceTo(pntgeom)
                distances.append((dist,polyline))
        mindist=min(distances,key=lambda x:x[0])
        return mindist[1]


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
    degrees = int(dd)
    minutes = int((dd-degrees)*60)
    seconds = (dd-degrees-minutes/60)*3600
    return (degrees,minutes,seconds)

def returnBearingString(azimuth):
    bearing=None
    dmsbearing=None
    if azimuth>270 and azimuth<=360:
        bearing = 360 - azimuth
        dmsbearing=ddToDms(bearing)
        bear = int(dmsbearing[0])
        minute = "{:02d}".format(int(dmsbearing[1]))
        second = "{:02d}".format(int(dmsbearing[2]))
        dmsbearing = u"""N{0}\xb0{1}'{2}"W""".format(bear,minute,second)
        return dmsbearing
    if azimuth>=0 and azimuth<=90:
        dmsbearing=ddToDms(azimuth)
        bear = int(dmsbearing[0])
        minute = "{:02d}".format(int(dmsbearing[1]))
        second = "{:02d}".format(int(dmsbearing[2]))
        dmsbearing = u"""N{0}\xb0{1}'{2}"E""".format(bear,minute,second)
        return dmsbearing
    if azimuth>90 and azimuth<=180:
        bearing= 180 - azimuth
        dmsbearing=ddToDms(bearing)
        bear = int(dmsbearing[0])
        minute = "{:02d}".format(int(dmsbearing[1]))
        second = "{:02d}".format(int(dmsbearing[2]))
        dmsbearing = u"""S{0}\xb0{1}'{2}"E""".format(bear,minute,second)
        return dmsbearing
    if azimuth>180 and azimuth<=270:
        bearing = azimuth-180
        dmsbearing=ddToDms(bearing)
        bear = int(dmsbearing[0])
        minute = "{:02d}".format(int(dmsbearing[1]))
        second = "{:02d}".format(int(dmsbearing[2]))
        dmsbearing = u"""S{0}\xb0{1}'{2}"W""".format(bear,minute,second)
        return dmsbearing

def createCornerTies(newgdb,az,dimoffsetscale,dimension,start,end,pob=False):
    az180 =  az-180
    if az180<0: az180=az180+360
    ic = arcpy.da.InsertCursor(os.path.join(newgdb,dimension.name),["SHAPE@","BEARING"])
    editor = arcpy.da.Editor(newgdb)
    editor.startEditing(True)
    editor.startOperation()
    if 315 <= az < 360 or 0 <= az < 135:
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
        point2 = start.pointFromAngleAndDistance(az1,dimoffsetscale,"PLANAR")
        point3 = start.pointFromAngleAndDistance(az2,(dimoffsetscale*2),"PLANAR")
        array = arcpy.Array([point3.getPart(0),point2.getPart(0),point1])
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
        array2 = arcpy.Array([point6.getPart(0),point5.getPart(0),point4])
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
    
    if 135 <= az < 315:
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
        point2 = start.pointFromAngleAndDistance(az1,dimoffsetscale,"PLANAR")
        point3 = start.pointFromAngleAndDistance(az2,(dimoffsetscale*2),"PLANAR")
        array = arcpy.Array([point3.getPart(0),point2.getPart(0),point1])
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
        array2 = arcpy.Array([point6.getPart(0),point5.getPart(0),point4])
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

def CreateSpiderDimension(newgdb,dimension,startpoint,endpoint,scale,pob=False):
    desc = arcpy.Describe(dimension.name)
    sr = desc.spatialReference
    point1 = startpoint.getPart(0)
    point2 = endpoint.getPart(0)
    pobaz,poblen = returnInverse(point1,point2)
    bigscale = 1/scale
    leadergroup = 1
    if pob == True: leadergroup = 2
    #print(pobaz,poblen)
    bear = returnBearingString(pobaz)
    editor = arcpy.da.Editor(newgdb)
    editor.startEditing(True)
    editor.startOperation()
    ic = arcpy.da.InsertCursor(os.path.join(newgdb,dimension.name),["SHAPE@","BEARING","PARENTOID"])
    leader = startpoint.pointFromAngleAndDistance(45,(bigscale/2),"PLANAR")
    lander = leader.pointFromAngleAndDistance(90,(bigscale/20),"PLANAR")
    array1 = arcpy.Array([point1,leader.getPart(0),lander.getPart(0)])
    array2 = arcpy.Array([point2,leader.getPart(0),lander.getPart(0)])
    startpoly = arcpy.Polyline(array1,sr,False,False)
    endppoly = arcpy.Polyline(array2,sr,False,False)
    bearstring = """{} {}'""".format(bear,round(poblen,2))
    row1 = [startpoly,bearstring,leadergroup]
    row2 = [endppoly,None,leadergroup]
    ic.insertRow(row1)
    ic.insertRow(row2)
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

def createROWDim(tractcenshape,newgdb,dimensionlayer,viewportpointlayer,twss=None,row=None,char=None,tileindex=[]):
    insertlist = []
    ordinc = 0 
    di=0
    twsdimgroup=200
    for tile in tileindex:
        editor =  arcpy.da.Editor(newgdb)
        editor.startEditing(True)
        editor.startOperation()
        with arcpy.da.InsertCursor(os.path.join(newgdb,dimensionlayer.name),["SHAPE@","DIM_GROUP","DIM_NUMBER","LABEL"]) as ic:
            if char is None: char = 'A'
            for i in range(tractcenshape.partCount):
                censhape = arcpy.Polyline(tractcenshape.getPart(i),tractcenshape.spatialReference,True,True)
                censhape = censhape.intersect(tile,2)
                linelist = []  
                cumx = 0
                cumy = 0
                denom = 0
                array = censhape.getPart(0)
                for x in range(len(array)-1):
                    cenpoint1 = array.getObject(x)
                    cenpoint2 = array.getObject(x+1)
                    newarray = arcpy.Array([cenpoint1,cenpoint2])
                    centerseg = arcpy.Polyline(newarray,censhape.spatialReference,True,True)
                    linelist.append((centerseg.length,centerseg))
                sortedlinelist = max(linelist, key=lambda x:x[0])
                if len(sortedlinelist)>0:
                    censeg = sortedlinelist[1]
                    az = returnAzimuth(censeg)
                    osrang  = az + 90
                    if osrang>=360: osrang = osrang-360
                    oslang = az - 90
                    if oslang<0: oslang = oslang+360
                    pointR = censeg.positionAlongLine(.495,use_percentage=True)
                    cumx = cumx+pointR.getPart(0).X
                    cumy = cumy+pointR.getPart(0).Y
                    denom+=1
                    dimRgroup = di+1
                    pos = 2
                    label = "25'"
                    insertlist.append((pointR,dimRgroup,pos,label))
                    ic.insertRow([pointR,dimRgroup,pos,label])
                    pointL = censeg.positionAlongLine(.5,use_percentage=True)
                    cumx = cumx+pointL.getPart(0).X
                    cumy = cumy+pointL.getPart(0).Y
                    denom+=1
                    dimLgroup = di+20
                    pos = 1
                    label = "25'"
                    insertlist.append((pointL,dimLgroup,pos,label))
                    ic.insertRow([pointL,dimLgroup,pos,label])
                    ospointR = pointR.pointFromAngleAndDistance(osrang,25,method='PLANAR')
                    cumx = cumx+ospointR.getPart(0).X
                    cumy = cumy+ospointR.getPart(0).Y
                    denom+=1
                    dimRgroup = di+1
                    pos = 1
                    label = "25'"
                    insertlist.append((ospointR,dimRgroup,pos,label))
                    ic.insertRow([ospointR,dimRgroup,pos,label])
                    ospointL = pointL.pointFromAngleAndDistance(oslang,25,method='PLANAR')
                    cumx = cumx+ospointL.getPart(0).X
                    cumy = cumy+ospointL.getPart(0).Y
                    denom+=1
                    dimLgroup = di+20
                    pos = 2
                    label = "25'"
                    insertlist.append((ospointL,dimLgroup,pos,label))
                    ic.insertRow([ospointL,dimLgroup,pos,label])
                    di+=1
                if not twss is None:
                    
                    for tws  in twss:
                        for j in range(tws.partCount):
                            twspart = tws.getPart(j)
                            twsshape = arcpy.Polygon(twspart,tws.spatialReference,True,True)
                            twsshape = twsshape.intersect(tile,4)
                            twslines = explodePoly(twsshape)
                            linelengthnondis = [(line.length,line) for line in twslines if line.disjoint(row)==True]
                            linelengthdis = [(line.length,line) for line in twslines if line.disjoint(row)==False]
                            maxnondis = max(linelengthnondis, key=lambda x:x[0])
                            maxdis = max(linelengthdis, key=lambda x:x[0])
                            outsidepoint = maxnondis[1].positionAlongLine(.50,True)
                            x3 = outsidepoint.getPart(0).X 
                            y3 = outsidepoint.getPart(0).Y
                            x2 = maxdis[1].lastPoint.X 
                            y2 = maxdis[1].lastPoint.Y 
                            x1 = maxdis[1].firstPoint.X
                            y1 = maxdis[1].firstPoint.Y
                            dx = x2-x1
                            dy = y2-y1
                            d2 = dx*dx+dy*dy
                            nx = ((x3-x1)*dx+(y3-y1)*dy)/d2
                            x  = dx*nx+x1
                            y = dy*nx+y1
                            #print((x,y))
                            insidepoint = arcpy.Point(x,y)
                            insidepointgeometry = arcpy.PointGeometry(insidepoint,tws.spatialReference,True,True)
                            az,dist = returnInverse(outsidepoint.getPart(0),insidepoint)
                            label = "{0}'".format(int(round(dist,0)))
                            cumx = cumx+outsidepoint.getPart(0).X
                            cumy = cumy+outsidepoint.getPart(0).Y
                            denom+=1
                            outsiderow = [outsidepoint,twsdimgroup,1,label]
                            ic.insertRow(outsiderow)
                            cumx = cumx+insidepointgeometry.getPart(0).X
                            cumy = cumy+insidepointgeometry.getPart(0).Y
                            denom+=1
                            insiderow = [insidepointgeometry,twsdimgroup,2,label]
                            ic.insertRow(insiderow)
                            twsdimgroup+=1
                if cumx>0 and cumy>0:
                    avx = cumx/denom
                    avy = cumy/denom
                    avpoint = arcpy.PointGeometry(arcpy.Point(avx,avy),tws.spatialReference,True,True)
                    char = chr(ord(char)+ordinc)
                    viewrow = [avpoint,char,1]
                    if ordinc==0: ordinc = 1
                    with arcpy.da.InsertCursor(os.path.join(newgdb,viewportpointlayer.name),["SHAPE@","DETAIL_LETTER","DETAIL_SIZE"]) as vc:
                        vc.insertRow(viewrow)     
        
                                
        editor.stopOperation()
        editor.stopEditing(True)
        try:
            del vc
        except:
            print("No Delete View Cursor")

        try:
            del ic
        except:
            print("No Delete Cursor")

        try:
            del editor
        except:
            print("No Delete Editor")
    return insertlist,char

def insertSegments(newgdb,tractcenterline,centerlinesegment,tractnumber,sheet,ismultipart):
    sr = tractcenterline.spatialReference
    borders = xlwt.Borders()
    borders.left = 1
    borders.right = 1
    borders.bottom = 1
    borders.top = 1
    style = xlwt.XFStyle()
    style.font.height = 6*20
    style2 = style
    style2.font.height = 6*20
    style2.alignment.horz = xlwt.Alignment.HORZ_RIGHT
    style2.borders = borders
    style.borders = borders
    style.alignment.wrap=1
    style.alignment.horz = xlwt.Alignment.HORZ_CENTER
    for x in range(3):
        if x==0:
            col = sheet.col(x)
            col.width = 4 * 256
        if x==1:
            col = sheet.col(x)
            col.width = 8 * 256
        if x==2:
            col = sheet.col(x)
            col.width = 7 * 256
    sheet.write_merge(0,0,0,2,"LINE TABLE",style)
    sheet.write(1,0,"LINE",style)
    sheet.write(1,1,"BEARING",style)
    sheet.write(1,2,"LENGTH",style)
    if ismultipart==False:
        editor = arcpy.da.Editor(newgdb)
        editor.startEditing(True)
        editor.startOperation()
        ic = arcpy.da.InsertCursor(os.path.join(newgdb,centerlinesegment.name),["SHAPE@","LINE_NUMBER","BEARING","TRACT_NUMBER"])
        pnts = tractcenterline.getPart(0)
        cnt = 2
        dist = 0.000
        segoids=[]
        for x in range(len(pnts)-1):
            point1 = pnts.getObject(x)
            point2 = pnts.getObject(x+1)
            polyArray = arcpy.Array()
            polyArray.add(point1)
            polyArray.add(point2)
            polyShape = arcpy.Polyline(polyArray,sr,True,True)
            dist = dist+polyShape.length
            az = returnAzimuth(polyShape)
            bear  = returnBearingString(az)
            bear2 = returnBearingString(az)
            lnum = cnt -1
            sheet.write(cnt,0,"L{0}".format(lnum),style)
            sheet.write(cnt,1,bear,style)
            sheet.write(cnt,2,"{:.2f}'".format(round(polyShape.length,4)),style)
            newRow = polyShape,lnum,bear2,tractnumber
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
                lnum = cnt -1
                sheet.write(cnt,0,"L{0}".format(lnum),style)
                sheet.write(cnt,1,bear,style)
                sheet.write(cnt,2,"{:.2f}'".format(round(polyShape.length,4)),style)
                newRow = polyShape,lnum,bear2,tractnumber
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
    if 7250>length>3625:
        scale = .001
        return scale
    else:
        return 30

def returnPlatScale(tractcenterline):
    scale = None
    dX = abs(tractcenterline.firstPoint.X-tractcenterline.lastPoint.X)*1.25
    dY = abs(tractcenterline.firstPoint.Y-tractcenterline.lastPoint.Y)*1.25
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

def createIndexPoly(newgdb,tractname,scale,X,Y,sr,tileid=None):
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
            row = [pol,tractname,tileid,scale]
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

def getTileIndexShapes(newgdb,tileindexlayer):
    tiles = []
    with arcpy.da.SearchCursor(os.path.join(newgdb,tileindexlayer),"SHAPE@") as sc:
        for row in sc:
            tiles.append(row[0])
    return tiles

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
 
def getROWPolyShape(boundarypolyshape,rowpoly,cenname,monument=False,where=None):
    shapes = []
    desc = arcpy.Describe(rowpoly.name)
    sr = desc.spatialReference
    if monument==False:
        with arcpy.da.SearchCursor(rowpoly.name,["SHAPE@","CENTERLINE_NAME","OBJECTID"],where,spatial_reference=sr) as sc:
            for row in sc:
                if row[0].disjoint(boundarypolyshape)==False:
                    add = (row[0],row[1])
                    if not add in shapes:
                        shapes.append(add)
    else:
        with arcpy.da.SearchCursor(rowpoly.name,["SHAPE@","CENTERLINE_NAME","LABEL","OBJECTID"],where,spatial_reference=sr) as sc:
            for row in sc:
                if row[0].disjoint(boundarypolyshape)==False:
                    add = (row[0],row[1],row[2],row[3])
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

def platmultipage(cen,tractnumber,direction):
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

    desc = arcpy.Describe(centerline.name)
    sr = desc.spatialReference

    wrkspc = centerline.connectionProperties['connection_info']['database']
    wrkspcfolder = os.path.abspath(os.path.join(wrkspc, os.pardir))
    
    resourcegdb = r"D:\Red_Oak_Project\pristine_gdb\RED_OAK_PROJECT_TEMPLATE.gdb"


    censhape,cenname =  getCenterlineShape(centerline,cen)
    if censhape is None: return "Pipeshape is Null, check CL Name"

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
    if boundaryshape is None: return "No Boundary Found For {0}".format(tractnumber)
    #boundrow = [boundaryshape,tractnumber,recagent,docref,surfown,calledac,ogsurv,county,state]
    name = insertCenterlineSegment(newgdb,boundarypoly,boundrow,cenname,tractnumber,bound=True)
    boundarybuffer = boundaryshape.buffer(5000)
    ajnames = getAjoiningTracts(boundarypoly,boundarybuffer,cenname,newgdb,tractnumber)
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
        rowintersection = None
        name = None
        polygon = poly[0]
        name  = poly[1]
        rowintersection = polygon.intersect(boundaryshape,4)
        if rowintersection: rowintersections.append(rowintersection)
        if rowintersection: name = insertCenterlineSegment(newgdb,rowpoly,rowintersection,cenname,tractnumber)
        

    #######################################################################################
    ########################################################################################
    atwsintersections=[]
    for poly in atws:
        name = None
        rowintersection = None
        polygon = poly[0]
        name  = poly[1]
        rowintersection = polygon.intersect(boundaryshape,4)
        if rowintersection: name = insertCenterlineSegment(newgdb,atwspoly,rowintersection,cenname,tractnumber)
        if rowintersection: atwsintersections.append(rowintersection)
    

    ##############################################################################################
    ############################################################################################
    twsintersections=[]
    for poly in twss:
        name = None
        rowintersection = None
        polygon = poly[0]
        name  = poly[1]
        rowintersection = polygon.intersect(boundaryshape,4)
        if rowintersection: twsintersections.append(rowintersection)
        if rowintersection: name = insertCenterlineSegment(newgdb,twspoly,rowintersection,cenname,tractnumber)
        

    #####################################################################
    ##
    ## Get centeriine iside tract
    ##
    #####################################################################

    tractcenterline = censhape.intersect(boundaryshape,2)

    tractcenterlinelength = returnPageLength(tractcenterline)

    ###################################################################################################3
    ##
    ##  Get page length and number of pages
    ##
    #######################################################################################################
    pagelength,pages,scalevalue=findPageRange(tractcenterlinelength)

    #set scale
    scale=1/scalevalue

    #set dim offset coeificient for cormer ties
    dimoffset = (1/scale)*.1

    ################################################################################################
    ##
    ## use centerline, boundary poly to create corner ties, 
    ##
    ##################################################################################################

    polylines = explodePoly(boundaryshape)
    
    firstline = returnPropline(polylines,tractcenterline,firstpoint=True,findclose=False)
    if firstline is None: firstline=returnPropline(polylines,tractcenterline,firstpoint=True,findclose=True)
    lastline = returnPropline(polylines,tractcenterline,firstpoint=False,findclose=False)
    if lastline is None: lastline=returnPropline(polylines,tractcenterline,firstpoint=False,findclose=True)
    firstlinescale = scaleGeom(firstline,150)
    lastlinescale = scaleGeom(lastline,150)

    firstbuff = firstline.buffer(20)
    lastbuff = lastline.buffer(20)
    firstlinebuff = firstlinescale.buffer(20)
    lastlinebuff = lastlinescale.buffer(20)
    wheremon = """UPPER(LABEL) NOT LIKE '%60D%'"""
    startmons = getROWPolyShape(firstbuff,monument,cenname,monument=True,where=wheremon)
    if len(startmons)<1:
        startmons = getROWPolyShape(firstlinebuff,monument,cenname,monument=True,where=wheremon)
    if len(startmons)<1:
        startmons = getROWPolyShape(firstlinebuff,monument,cenname,monument=True,where=None)
    if len(startmons)>0:
        startmon = ReturnClosestMonument(startmons,tractcenterline.firstPoint)
        montostartaz,montostartdis = returnInverse(startmon[1].getPart(0),tractcenterline.firstPoint)
        if montostartdis<(1/scale)*.85:
            pobgeom = arcpy.PointGeometry(tractcenterline.firstPoint,tractcenterline.spatialReference,True,True)
            CreateSpiderDimension(newgdb,dimensionlinelayer,startmon[1],pobgeom,scale,pob=True)
        else:
            pobgeom = arcpy.PointGeometry(tractcenterline.firstPoint,tractcenterline.spatialReference,True,True)
            createCornerTies(newgdb,montostartaz,dimoffset,dimensionlinelayer,startmon[1],pobgeom,pob=True)
    if len(startmons)>0: wheremon = """OBJECTID <> {} AND UPPER(LABEL) NOT LIKE '%60D%'""".format(startmon[-1])
    endmons = getROWPolyShape(lastbuff,monument,cenname,monument=True,where=wheremon)
    if len(endmons)<1:
        if len(startmons)>0: wheremon="""OBJECTID <> {}""".format(startmon[-1])
        if len(startmons)<1: wheremon= """UPPER(LABEL) NOT LIKE '%60D%'"""
        endmons = getROWPolyShape(lastbuff,monument,cenname,monument=True,where=wheremon)

    if len(endmons)<1:
        endmons = getROWPolyShape(lastlinebuff,monument,cenname,monument=True,where=wheremon)
    if len(endmons)<1:
        endmons = getROWPolyShape(lastlinebuff,monument,cenname,monument=True,where=None)
    if len(endmons)>0:
        endmon = ReturnClosestMonument(endmons,tractcenterline.lastPoint)
        endtomonaz,endtomondis = returnInverse(tractcenterline.lastPoint,endmon[1].getPart(0))
        if endtomondis<(1/scale)*.85:
            potgeom = arcpy.PointGeometry(tractcenterline.lastPoint,tractcenterline.spatialReference,True,True)
            CreateSpiderDimension(newgdb,dimensionlinelayer,potgeom,endmon[1],scale,pob=False)
        else:
            potgeom = arcpy.PointGeometry(tractcenterline.lastPoint,tractcenterline.spatialReference,True,True)
            createCornerTies(newgdb,endtomonaz,dimoffset,dimensionlinelayer,potgeom,endmon[1],pob=False)

    ###################################################################################################
    ##
    ## Create indexpoly
    ##
    #####################################################################################################
    multipages=[]
    for x in range(pages):
        dist1 = x*pagelength
        dist2 = (x+1)*pagelength
        distalongline = tractcenterline.segmentAlongLine(dist1,dist2,False)
        tileindexrow,tileindexfields = createIndexPoly(newgdb,tractnumber,scale,distalongline.centroid.X,distalongline.centroid.Y,sr,tileid=x+1)
    ##########################################################################################################################
    ##
    ## get index polys loop through and dim for each tile
    ##
    ############################################################################################################################
    tiles = getTileIndexShapes(newgdb,"TILE_INDEX")
    if len(tiles)<1: return "No Tiles In Database"
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
            insertlist,char = createROWDim(tractcenterline,newgdb,dimppoint,viewportpointlayer,twss=twsintersections,row=rowtotalshape,tileindex=tiles)
            listthrough,dictin = createDimDicts(atwsintersections)
            if len(dictin)>0:
                createAtwsViewports(newgdb,viewportpointlayer,dictin,listthrough,char)
        if len(twsintersections)<1:
            insertlist,char = createROWDim(tractcenterline,newgdb,dimppoint,viewportpointlayer)
            listthrough,dictin = createDimDicts(atwsintersections)
            if len(dictin)>0:
                createAtwsViewports(newgdb,viewportpointlayer,dictin,listthrough,char)
        if len(insertlist)<1: return "No Dims Created"
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


    ##############################################################################
    ##
    ## Insert line segments and write table to excel
    ##
    ################################################################################
    
    censegoids = insertSegments(newgdb,tractcenterline,centerlineseg,tractnumber,writesheet,tractcenterline.isMultipart)

    #####################################################################
    ##
    ##  Add TIles to map and export labels.
    ##
    #######################################################################
    
    
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
    parameters['DWGNAME']=newgdb[:-4]+".dwg"
    parameters['DIR']=direction
    parameters['NUMPAGES']=str((pages+1))
    parameters['WRKBOOK']=writebookpath
    wrkspcrunner = fmeobjects.FMEWorkspaceRunner()
    try:
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