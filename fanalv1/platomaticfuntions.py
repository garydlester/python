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


def returnTemplateName(templatedir,scale):
    if not os.path.exists(templatedir): return None 
    for filename in os.listdir(templatedir):
        if str(scale) in filename:
            return os.path.join(templatedir,filename)



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



def ReturnClosestMonument(mons,point,switch=None):
    inX = point.X
    inY = point.Y
    closest = []
    if switch=="mons":
        for row in mons:
            point2 = row[0].getPart(0)
            outX = point2.X
            outY = point2.Y
            dist=sqrt((inX-outX)**2+(inY-outY)**2)
            closest.append((dist,row[0],row[1],row[2],row[3]))
        minDist=(min(closest, key=lambda x:x[0]))
        return minDist
    if switch=="ends":
        for row in mons:
            point2 = row[0].getPart(0)
            outX = point2.X
            outY = point2.Y
            dist=sqrt((inX-outX)**2+(inY-outY)**2)
            closest.append((dist,row[0]))
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
    seconds = round(seconds,2)
    if seconds>=60: 
        seconds = 0
        minutes = minutes+1
    if minutes>=60:
        minutes=0
        degrees=degrees+1
    return (degrees,minutes,seconds)

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
    if 315 <= az < 360 or 0 <= az < 135:
        if pob==True:
            az1 = az-60
            if az1<0: az1 = az1+360
            az2 = az-20
            if az2<0: az2 = az2+360
        else:
            az1 = az+60
            if az1>360: az1 = az1-360
            az2 = az+20
            if az2>360: az2 = az2-360
        point1 = start.getPart(0)
        point2 = start.pointFromAngleAndDistance(az1,dimoffsetscale*.3,"PLANAR")
        point3 = start.pointFromAngleAndDistance(az2,dimoffsetscale,"PLANAR")
        array = arcpy.Array([point3.getPart(0),point2.getPart(0),point1])
        dimline = arcpy.Polyline(array,start.spatialReference,True,True)
        row = (dimline,None)
        ic.insertRow(row)
        if pob==True:
            az3 = az180+60
            if az3>360: az3=az3-360
            az4 = az180+20
            if az4>360: az4=az4-360
        else:
            az3 = az180-60
            if az3<0: az3=az3+360
            az4 = az180-20
            if az4<0: az4=az4+360
        point4 = end.getPart(0)
        point5 = end.pointFromAngleAndDistance(az3,dimoffsetscale*.3,"PLANAR")
        point6 = end.pointFromAngleAndDistance(az4,dimoffsetscale,"PLANAR")
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
            az1 = az+60
            if az1>360: az1=az1-360
            az2 = az+20
            if az2>360: az2 = az2-360
        else:
            az1 = az-60
            if az1<0: az1=az1+360
            az2 = az-20
            if az2<0: az2 = az2+360
        point1 = start.getPart(0)
        point2 = start.pointFromAngleAndDistance(az1,dimoffsetscale*.3,"PLANAR")
        point3 = start.pointFromAngleAndDistance(az2,dimoffsetscale,"PLANAR")
        array = arcpy.Array([point3.getPart(0),point2.getPart(0),point1])
        dimline = arcpy.Polyline(array,start.spatialReference,True,True)
        row = (dimline,None)
        ic.insertRow(row)
        
        if pob==True:
            az3 = az180-60
            if az3<0: az3=az3+360
            az4 = az180-20
            if az4<0: az4=az4+360
        else:
            az3 = az180+60
            if az3>360: az3=az3-360
            az4 = az180+20
            if az4>360: az4=az4-360
        point4 = end.getPart(0)
        point5 = end.pointFromAngleAndDistance(az3,dimoffsetscale*.3,"PLANAR")
        point6 = end.pointFromAngleAndDistance(az4,dimoffsetscale,"PLANAR")
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
    bear = returnBearingString(pobaz)
    editor = arcpy.da.Editor(newgdb)
    editor.startEditing(True)
    editor.startOperation()
    ic = arcpy.da.InsertCursor(os.path.join(newgdb,dimension.name),["SHAPE@","BEARING","PARENTOID"])
    if 315 <= pobaz < 360 or 0 <= pobaz < 135: leader = startpoint.pointFromAngleAndDistance(315,(bigscale/2),"PLANAR")
    if 315 <= pobaz < 360 or 0 <= pobaz < 135: lander = leader.pointFromAngleAndDistance(270,(bigscale/20),"PLANAR")
    if 135 <= pobaz < 315: leader =startpoint.pointFromAngleAndDistance(45,(bigscale/2),"PLANAR")
    if 135 <= pobaz < 315: lander = leader.pointFromAngleAndDistance(90,(bigscale/20),"PLANAR")
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

def insertSegments(centerlinesegments,sheet):
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
    cnt = 2
    dist = 0.00
    segoids=[]
    for censeg in centerlinesegments:
        polyshape = censeg[0]
        lnum = censeg[1]
        az = float(censeg[2])
        bear  = returnBearingString(az)
        bear2 = returnBearingString(az)
        sheet.write(cnt,0,"L{0}".format(lnum),style)
        sheet.write(cnt,1,bear,style)
        sheet.write(cnt,2,"{:.2f}'".format(round(polyshape.length,4)),style)
        dist = dist+round(polyshape.length,2)
        segoids.append(lnum)
        cnt+=1
    sheet.write(cnt,1,"Total:",style2)
    sheet.write(cnt,2,"{:.2f}'".format(dist),style)

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
    dX = abs(tractcenterline.firstPoint.X-tractcenterline.lastPoint.X)*1.10
    dY = abs(tractcenterline.firstPoint.Y-tractcenterline.lastPoint.Y)*1.10
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

def getROWPolyShape(boundarypolyshape,rowpoly,cenname,where=None,ngdb=None,type=None):
    shapes = []
    desc = arcpy.Describe(rowpoly.name)
    sr = desc.spatialReference
    if type=="endpoint":
        with arcpy.da.SearchCursor(os.path.join(ngdb,rowpoly.name),["SHAPE@","NORTH","EAST"],where,spatial_reference=sr) as sc:
            for row in sc:
                add = (row[0],row[1],row[2])
                shapes.append(add)
    if type=="geomfeat":
        with arcpy.da.SearchCursor(rowpoly.name,["SHAPE@","CENTERLINE_NAME","OBJECTID"],where,spatial_reference=sr) as sc:
            for row in sc:
                if row[0].disjoint(boundarypolyshape)==False:
                    add = (row[0],row[1])
                    if not add in shapes:
                        shapes.append(add)
    if type=="monument":
        with arcpy.da.SearchCursor(rowpoly.name,["SHAPE@","CENTERLINE_NAME","LABEL","OBJECTID"],where,spatial_reference=sr) as sc:
            for row in sc:
                if row[0].disjoint(boundarypolyshape)==False:
                    add = (row[0],row[1],row[2],row[3])
                    if not add in shapes:
                        shapes.append(add)
    if type=="newgdbfeat":
         with arcpy.da.SearchCursor(os.path.join(ngdb,rowpoly.name),["SHAPE@","CENTERLINE_NAME","OBJECTID"],where,spatial_reference=sr) as sc:
            for row in sc:
                add = (row[0])
                shapes.append(add)
    if type=="censeg":
        with arcpy.da.SearchCursor(os.path.join(ngdb,rowpoly.name),["SHAPE@","LINE_NUMBER","BEARING"],where,spatial_reference=sr) as sc:
            for row in sc:
                add = (row[0],row[1],row[2])
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
