#!/usr/bin/python
# -*- coding: utf-8 -*-
##############################################################################################################################################################
##
## This library is the property of Gary Lester Please use it but do not distribute it as your own.
##
##############################################################################################################################################################
import arcpy
import math
from math import sqrt
import sys
try:
    sys.path.append(r"D:\Projects\AutoWellGIS\Python")
except:
    pass
import corpse
import xlwt

__all__=[item for item in dir(arcpy) if not item.startswith("_")]
locals().update(arcpy.__dict__)

##### classes for plat.py
class AzimuthBearingDistance(object):
    def __init__(self):
        self.widthangle = 90
        self.lengthangle = 90
        self.rotationfieldname = "ROTATION"

    def getSectionRotation(self,landpolyshape,blockpoly):
        sr = landpolyshape.spatialReference
        landpoint = landpolyshape.centroid
        print(landpoint)
        #arcpy.AddMessage(str(landPoint))
        desc = arcpy.Describe(blockpoly.name)
        with arcpy.da.SearchCursor(blockpoly.name,[desc.shapeFieldName+"@",self.rotationfieldname],spatial_reference=sr) as sc2:
            for row2 in sc2:
                    if row2[0].contains(landpoint):
                        print(row2)
                        rot = row2[1]
                        return rot



    def getNEWSByAzimuth(self,landpolyshape,landlineshape,rot):
        polyCenter = landpolyshape.centroid
        lineCenter = landlineshape.centroid
        az=self.returnInverse(polyCenter,lineCenter)[0]
        az = az + rot
        if az>360:
            az=az-360           
        if az < 360 and az >= 360 - self.lengthangle/2: #this sets the ratio for the rectangle lengthangle/2
            side="NORTH"
            return side
        if az <= 0 + self.lengthangle/2  and az >= 0:
            side="NORTH"
            return side
        if az < 270 + self.widthangle/2 and az >= 270 - self.widthangle/2:
            side="WEST"
            return side
        if az < 180 + self.lengthangle/2 and az >= 180 - self.lengthangle/2:
            side="SOUTH"
            return side
        if az > 90 - self.widthangle/2 and az < 90 + self.widthangle/2:
            side="EAST"
            return side


    #### my rectangles have many vertices along one line so I am having to add all the linsegments up to find the lonest side then area/longestside=shortestside
    def getTractSideRatio(self,landpolyshape):
        distances = []
        area = landpolyshape.area
        if area:
            pnts= pnts = landpolyshape.getPart(0)
            distance = 0
            for x in range(pnts.count-2):
                if distance==0: distance = self.returnInverse(pnts.getObject(x),pnts.getObject(x+1))[1]
                az1=self.returnInverse(pnts.getObject(x),pnts.getObject(x+1))[0]
                az2=self.returnInverse(pnts.getObject(x+1),pnts.getObject(x+2))[0]
                length = self.returnInverse(pnts.getObject(x+1),pnts.getObject(x+2))[1]
                topaz = self.returnInverse(pnts.getObject(x),pnts.getObject(x+1))[0]+2
                lowaz = self.returnInverse(pnts.getObject(x),pnts.getObject(x+1))[0]-2          
                if az2 <= topaz and az2 >= lowaz:
                    distance = distance + length
                    if x==pnts.count-2 : distances.append(distance)
                else:
                    distances.append(distance)
                    distance = 0
            length = max(distances)
            width = area/length
            w2 = width/2
            l2 = length/2
            widthangle = 2*int(round(math.degrees(math.atan2(w2, l2))))
            lengthangle = 2*int(round(math.degrees(math.atan2(l2, w2))))
        return widthangle,lengthangle

    #### this method will make a parallel offset from a polyline or polygon, takes a poyline or polygon object and a distance returns a geometry object
    def copyParallelLeft(self,polyline,offsetWidth):
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
            point = arcpy.Point(pnt.X+sX,pnt.Y+sY)
            rArray.add(point)
        offsetGeom = arcpy.Polyline(rArray,sr,False,False)
        return offsetGeom

    #### this method will make a parallel offset from a polyline or polygon, takes a poyline or polygon object and a distance returns a geometry object
    def copyParallelRight(self,polyline,offsetWidth):
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

    def ddToDms(self,dd):
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

    def returnBearingString(self,azimuth):
        bearing=None
        dmsbearing=None
        if azimuth>270 and azimuth<=360:
            bearing = 360 - azimuth
            dmsbearing=self.ddToDms(bearing)
            dmsbearing = """N {0}°{1}'{2}" W""".format(int(dmsbearing[0]),int(dmsbearing[1]),int(round(dmsbearing[2],0)))
            return dmsbearing
        if azimuth>=0 and azimuth<=90:
            dmsbearing=self.ddToDms(azimuth)
            dmsbearing = """N {0}°{1}'{2}" E""".format(int(dmsbearing[0]),int(dmsbearing[1]),int(round(dmsbearing[2],0)))
            return dmsbearing
        if azimuth>90 and azimuth<=180:
            bearing= 180 - azimuth
            dmsbearing=self.ddToDms(bearing)
            dmsbearing = """S {0}°{1}'{2}" E""".format(int(dmsbearing[0]),int(dmsbearing[1]),int(round(dmsbearing[2],0)))
            return dmsbearing
        if azimuth>180 and azimuth<=270:
            bearing = azimuth-180
            dmsbearing=self.ddToDms(bearing)
            dmsbearing = """S {0}°{1}'{2}" W""".format(int(dmsbearing[0]),int(dmsbearing[1]),int(round(dmsbearing[2],0)))
            return dmsbearing

    def returnAzimuth(self,shape):
        point1 = shape.firstPoint
        point2 = shape.lastPoint
        dX = point2.X-point1.X
        dY = point2.Y-point1.Y
        az = math.atan2(dX,dY)*180/math.pi
        if az<0:
            az = az+360
            return az
        return az
    
    def returnInverse(self,point1,point2):
        dX = point2.X-point1.X
        dY = point2.Y-point1.Y
        dis = sqrt(dX**2+dY**2)
        az = math.atan2(dX,dY)*180/math.pi
        if az<0:
            az = az+360
            return az,dis
        return az,dis

    def scaleGeom(self,shape,scale,reference=None):
        if shape is None: return None
        if reference is None: reference=shape.centroid
        refshape = arcpy.PointGeometry(reference)
        newparts=[]
        for i in range(shape.partCount):
            part = shape.getPart(i)
            newPart = []
            for t in range(part.count):
                apnt = part.getObject(t)
                if apnt is None:
                    newPart.append(apnt)
                    continue
                bdist=refshape.distanceTo(apnt)
                bpnt = arcpy.Point(reference.X+bdist,reference.Y)
                adist = refshape.distanceTo(bpnt)
                cdist = arcpy.PointGeometry(apnt).distanceTo(bpnt)
                angle = math.acos((adist**2+bdist**2-cdist**2)/(2*adist*bdist))
                scaleDist = bdist * scale 
                if apnt.Y<reference.Y: angle = angle * -1

                scalex = scaleDist*math.cos(angle)+reference.X 
                scaley = scaleDist*math.sin(angle)+reference.Y
                #print(scalex,scaley)
                newPart.append(arcpy.Point(scalex,scaley))
            newparts.append(newPart)
        return arcpy.Geometry(shape.type,arcpy.Array(newparts),shape.spatialReference)

class Label(AzimuthBearingDistance):
    def __init__(self):
        super(Label,self).__init__()
        #### self declarations that can be changed after instatiating the object. these are for the ReturnSetDict methods. 
        self.shlabrev = "SHL"
        self.ppabrev = "PP"
        self.ttpabrev = "TTP"
        self.btpabrev = "BTP"
        self.bhlabrev = "BHL"
        self.wellnamefield = "WELL_NAME"
        self.dirfield = "DIRECTION"
        self.direastfield = "DIRECTION_EAST"
        self.distfield = "DIST"
        self.disteastfield = "DIST_EAST"
        self.secfield = "SEC"
        self.blkfield = "BLK"
        self.takefield = "TAKE_POINT"
        ##### self. declarations that can be changed after instantiating the object these are for the ReturnBoreSum method
        self.takepointshapefield = "SHAPE@"
        self.takepointdescfield = "DESCRIPTION"
        self.takepointsfc = "TAKE_POINTS"
        self.sr = None
        ##### self. declarations that can be changed after instantiating the object these are for the ReturnWellBoreSummaryTablePoint() method
        self.wellboresumtableX = 10696969.1630
        self.wellboresumtableY = 1272691.5310
        self.wellname = ""
        self.sectiontieX = 10696795.2209
        self.sectiontieY = 1272691.531
        self.seclabelfields = ['SHAPE@','OBJECTID','SURFACE_OWNER','MINERAL_OWNER', 'BLK', 'TWP','COUNTY', 'STATE','ORIGINAL_SURVEY','ABSTRACT_NUMBER','SEC']
        self.seccornblockname = "SEC_CORN_LABELS"
        self.cadaction = "INSERT"
        self.weboresumblockname = "WELL_BORE_SUMMARY"
        self.corpseyearin = 1983
        self.corpseyearout = 1927
        self.corpseyearin2 = 1927
        self.corpseyearout2 = 1983
        self.corpsetype = 1
        self.corpsetype2 = 2
        self.corpsecoordzone = 4203

    def ReturnSetDicts(self,wellcalltable,where=None,multiwell=False,shl=False):
        secList=[]
        blockDict  = {}
        shlSetDict = {}
        ppSetDict = {}
        ttpSetDict = {}
        btpSetDict = {}
        bhlSetDict = {}
        shldict={}
        wellname = None
        if multiwell==False and shl==False:
            with arcpy.da.SearchCursor(wellcalltable.name,"*") as sc:
                for row in sc:
                    wellname = row[list(sc.fields).index(self.wellnamefield)]
                    if not row[list(sc.fields).index(self.secfield)] in secList: 
                        secList.append(row[list(sc.fields).index(self.secfield)])
                        blockDict[row[list(sc.fields).index(self.secfield)]]=row[list(sc.fields).index(self.blkfield)]
                    if row[list(sc.fields).index(self.takefield)]==self.shlabrev:
                        shlSetDict[row[list(sc.fields).index(self.dirfield)]]=[row[list(sc.fields).index(self.distfield)],row[list(sc.fields).index(self.secfield)]]
                        shlSetDict[row[list(sc.fields).index(self.direastfield)]]=[row[list(sc.fields).index(self.disteastfield)],row[list(sc.fields).index(self.secfield)]]
                    if row[list(sc.fields).index(self.takefield)]==self.ppabrev:
                        ppSetDict[row[list(sc.fields).index(self.dirfield)]]=[row[list(sc.fields).index(self.distfield)],row[list(sc.fields).index(self.secfield)]]
                        ppSetDict[row[list(sc.fields).index(self.direastfield)]]=[row[list(sc.fields).index(self.disteastfield)],row[list(sc.fields).index(self.secfield)]]
                    if row[list(sc.fields).index(self.takefield)]==self.ttpabrev:
                        ttpSetDict[row[list(sc.fields).index(self.dirfield)]]=[row[list(sc.fields).index(self.distfield)],row[list(sc.fields).index(self.secfield)]]
                        ttpSetDict[row[list(sc.fields).index(self.direastfield)]]=[row[list(sc.fields).index(self.disteastfield)],row[list(sc.fields).index(self.secfield)]]
                    if row[list(sc.fields).index(self.takefield)]==self.btpabrev:
                        btpSetDict[row[list(sc.fields).index(self.dirfield)]]=[row[list(sc.fields).index(self.distfield)],row[list(sc.fields).index(self.secfield)]]
                        btpSetDict[row[list(sc.fields).index(self.direastfield)]]=[row[list(sc.fields).index(self.disteastfield)],row[list(sc.fields).index(self.secfield)]]
                    if row[list(sc.fields).index(self.takefield)]==self.bhlabrev:
                        bhlSetDict[row[list(sc.fields).index(self.dirfield)]]=[row[list(sc.fields).index(self.distfield)],row[list(sc.fields).index(self.secfield)]]
                        bhlSetDict[row[list(sc.fields).index(self.direastfield)]]=[row[list(sc.fields).index(self.disteastfield)],row[list(sc.fields).index(self.secfield)]]
            return secList,blockDict,shlSetDict,ppSetDict,ttpSetDict,btpSetDict,bhlSetDict,wellname
        if multiwell==True:
            with arcpy.da.SearchCursor(wellcalltable.name,"*",where_clause=where) as sc:
                for row in sc:
                    wellname = row[list(sc.fields).index(self.wellnamefield)]
                    if not row[list(sc.fields).index(self.secfield)] in secList: 
                        secList.append(row[list(sc.fields).index(self.secfield)])
                        blockDict[row[list(sc.fields).index(self.secfield)]]=row[list(sc.fields).index(self.blkfield)]
                    if row[list(sc.fields).index(self.takefield)]==self.shlabrev:
                        shlSetDict[row[list(sc.fields).index(self.dirfield)]]=[row[list(sc.fields).index(self.distfield)],row[list(sc.fields).index(self.secfield)]]
                        shlSetDict[row[list(sc.fields).index(self.direastfield)]]=[row[list(sc.fields).index(self.disteastfield)],row[list(sc.fields).index(self.secfield)]]
                    if row[list(sc.fields).index(self.takefield)]==self.ppabrev:
                        ppSetDict[row[list(sc.fields).index(self.dirfield)]]=[row[list(sc.fields).index(self.distfield)],row[list(sc.fields).index(self.secfield)]]
                        ppSetDict[row[list(sc.fields).index(self.direastfield)]]=[row[list(sc.fields).index(self.disteastfield)],row[list(sc.fields).index(self.secfield)]]
                    if row[list(sc.fields).index(self.takefield)]==self.ttpabrev:
                        ttpSetDict[row[list(sc.fields).index(self.dirfield)]]=[row[list(sc.fields).index(self.distfield)],row[list(sc.fields).index(self.secfield)]]
                        ttpSetDict[row[list(sc.fields).index(self.direastfield)]]=[row[list(sc.fields).index(self.disteastfield)],row[list(sc.fields).index(self.secfield)]]
                    if row[list(sc.fields).index(self.takefield)]==self.btpabrev:
                        btpSetDict[row[list(sc.fields).index(self.dirfield)]]=[row[list(sc.fields).index(self.distfield)],row[list(sc.fields).index(self.secfield)]]
                        btpSetDict[row[list(sc.fields).index(self.direastfield)]]=[row[list(sc.fields).index(self.disteastfield)],row[list(sc.fields).index(self.secfield)]]
                    if row[list(sc.fields).index(self.takefield)]==self.bhlabrev:
                        bhlSetDict[row[list(sc.fields).index(self.dirfield)]]=[row[list(sc.fields).index(self.distfield)],row[list(sc.fields).index(self.secfield)]]
                        bhlSetDict[row[list(sc.fields).index(self.direastfield)]]=[row[list(sc.fields).index(self.disteastfield)],row[list(sc.fields).index(self.secfield)]]
            return secList,blockDict,shlSetDict,ppSetDict,ttpSetDict,btpSetDict,bhlSetDict,wellname
        if shl==True:
            with arcpy.da.SearchCursor(wellcalltable.name,"*",where_clause=where) as sc:
                for row in sc:
                    shldict[row[list(sc.fields).index(self.wellnamefield)]]={self.secfield:row[list(sc.fields).index(self.secfield)],\
                                                                            self.blkfield:row[list(sc.fields).index(self.blkfield)],\
                                                                            self.dirfield:row[list(sc.fields).index(self.dirfield)],\
                                                                            self.distfield:row[list(sc.fields).index(self.distfield)],\
                                                                            self.disteastfield:row[list(sc.fields).index(self.disteastfield)],\
                                                                            self.direastfield:row[list(sc.fields).index(self.direastfield)]}

    def ReturnBoreSumDict(self,wherelist):
        boreSumDict = {}
        for where in wherelist:
            whereHole="""{0}='{1}'""".format(self.takepointdescfield,where)
            with arcpy.da.SearchCursor(self.takepointsfc,[self.takepointshapefield,self.takepointdescfield],where_clause=whereHole,spatial_reference=self.sr) as sc:
                for row in sc:
                    boreSumDict[row[1]]=row[0]
        return boreSumDict

    def ReturnWellBoreSummaryTablePoint(self,boresumdict,boresum,wherelist):
        row = []
        row.append(arcpy.PointGeometry(arcpy.Point(self.wellboresumtableX,self.wellboresumtableY,0),self.sr,True,True))
        for x in range(len(wherelist)):
            if x<len(wherelist)-1:
                pointArray = arcpy.Array()
                pointArray.add(boresumdict[wherelist[x]].getPart(0))
                pointArray.add(boresumdict[wherelist[x+1]].getPart(0))
                bearPolyline = arcpy.Polyline(pointArray,self.sr,True,True)
                azimuth = self.returnAzimuth(bearPolyline)
                bearingString=self.returnBearingString(azimuth)
                row.append(bearingString)
                row.append(round(self.returnInverse(boresumdict[wherelist[x]].getPart(0),boresumdict[wherelist[x+1]].getPart(0))[1],2))
                
            if x>len(wherelist)-1:
                pointArray = arcpy.Array()
                pointArray.add(boresumdict[wherelist[x-1]].getPart(0))
                pointArray.add(boresumdict[wherelist[x]].getPart(0))
                bearPolyline = arcpy.Polyline(pointArray,self.sr,True,True)
                azimuth = self.returnAzimuth(bearPolyline)
                bearingString=self.returnBearingString(azimuth)
                row.append(bearingString)
                row.append(round(self.returnInverse(boresumdict[wherelist[x-1]].getPart(0),boresumdict[wherelist[x]].getPart(0))[1],2))
        row.append(self.weboresumblockname)
        row.append(self.cadaction)
        
    def ReturnTakeTables(self,wherelist):
        dict83={}
        dict27={}
        for where in wherelist:
            whereHole = """{0}='{1}'""".format(self.takepointdescfield, where)
            with arcpy.da.SearchCursor(self.takepointsfc,[self.takepointshapefield],where_clause=whereHole,spatial_reference=self.sr) as sc:
                for row in sc:
                    projectName = None
                    east_27 = None
                    north_27 = None
                    lat_27_dms = None
                    long_27_dms = None
                    lat_27 = None
                    long_27 = None
                    east_83 = None
                    north_83 = None
                    lat_83_dms = None
                    long_83_dms = None
                    lat_83 = None
                    long_83 = None
                    shape=row[0]
                    point = shape.getPart(0)
                    ex,y,z = corpse.conCoords(self.corpsetype2,self.corpseyearout,self.corpseyearin,self.corpsecoordzone,point.X,point.Y,point.Z)
                    east_27 = round(ex,3)
                    north_27 = round(y,3)
                    ex,y,z = corpse.conCoords(self.corpsetype,self.corpseyearout,self.corpseyearin,self.corpsecoordzone,point.X,point.Y,point.Z)
                    long_27_dms = """{0}°{1}'{2}" W""".format(self.ddToDms(ex)[0],self.ddToDms(ex)[1],self.ddToDms(ex)[2])
                    lat_27_dms = """{0}°{1}'{2}" N""".format(self.ddToDms(y)[0],self.ddToDms(y)[1],self.ddToDms(y)[2])
                    long_27 = ex
                    lat_27 = y
                    z27 = z
                    east_83 = round(point.X,3)
                    north_83 = round(point.Y,3)
                    ex,y,z = corpse.conCoords(self.corpsetype,self.corpseyearout2,self.corpseyearin2,self.corpsecoordzone,point.X,point.Y,point.Z)
                    long_83_dms = """{0}°{1}'{2}" W""".format(self.ddToDms(ex)[0],self.ddToDms(ex)[1],self.ddToDms(ex)[2])
                    lat_83_dms = """{0}°{1}'{2}" N""".format(self.ddToDms(y)[0],self.ddToDms(y)[1],self.ddToDms(y)[2])
                    long_83 = ex
                    lat_83 = y
                    z83 = z
                                       
                    if where==self.ppabrev:
                        projectName = self.wellname + " {0}".format(self.ppabrev)
                        dict27[where]=[projectName,east_27,north_27,lat_27_dms,long_27_dms,lat_27,long_27]
                        dict83[where]=[projectName,east_83,north_83,lat_83_dms,long_83_dms,lat_83,long_83]
                    if where==self.ttpabrev:
                        projectName = self.wellname + " {0}".format(self.ttpabrev)
                        dict27[where]=[projectName,east_27,north_27,lat_27_dms,long_27_dms,lat_27,long_27]
                        dict83[where]=[projectName,east_83,north_83,lat_83_dms,long_83_dms,lat_83,long_83]
                    if where==self.btpabrev:
                        projectName = self.wellname + " {0}".format(self.btpabrev)
                        dict27[where]=[projectName,east_27,north_27,lat_27_dms,long_27_dms,lat_27,long_27]
                        dict83[where]=[projectName,east_83,north_83,lat_83_dms,long_83_dms,lat_83,long_83]
                    if where==self.bhlabrev:
                        projectName = self.wellname + " {0}".format(self.bhlabrev)
                        dict27[where]=[projectName,east_27,north_27,lat_27_dms,long_27_dms,lat_27,long_27]
                        dict83[where]=[projectName,east_83,north_83,lat_83_dms,long_83_dms,lat_83,long_83]
                    if where==self.shlabrev:
                        projectName83 = self.wellname  + " ({0})' El.".format(z83)
                        projectName27 = self.wellname  + " ({0})' El.".format(z27)
                        dict27[where]=[projectName27,east_27,north_27,lat_27_dms,long_27_dms,lat_27,long_27]
                        dict83[where]=[projectName83,east_83,north_83,lat_83_dms,long_83_dms,lat_83,long_83]
                    

        return dict27,dict83

    def ReturnSectionTieDict(self,wherelist,setdict):
        sectionTieSumPoint = [arcpy.PointGeometry(arcpy.Point(self.sectiontieX,self.sectiontieY,0),self.sr,True,True),self.wellname]
        for where in wherelist:
            ns=None
            ew=None
            for k,v in setdict[where].iteritems():
                if k=="FNL" or k =="FSL":
                    ns = """{0}-{1}' {2},""".format(where,v[0],k)
                if k=="FEL" or k =="FWL":
                    ew = """ {0}' {1}-SECTION {2}""".format(v[0],k,v[1])
            sectionTieSumPoint.append(ns+ew)
        return sectionTieSumPoint


    def ReturnSectionLabels(self,landpoly):
        rows=[]
        with arcpy.da.SearchCursor(landpoly.name,self.seclabelfields,spatial_reference=self.sr) as sc:
            for row in sc:
                shape=row[0]
                point = arcpy.PointGeometry(shape.centroid,self.sr,True,True)
                newRow  =  [point] + list(row[1:])
                rows.append(newRow)
        return rows
    
    def ReturnSecCornerLabels(self,landpoly):
        rows=[]
        with arcpy.da.SearchCursor(landpoly.name,[self.takepointshapefield,self.secfield],spatial_reference=self.sr) as sc:
            for row in sc:
                shape = row[0]
                sec = row[1]
                insideShape = self.scaleGeom(shape,.90)
                array = insideShape.getPart(0)
                for x in range(len(array)-1):
                    pnt =  array.getObject(x)
                    X = pnt.X
                    Y = pnt.Y 
                    Z = pnt.Z 
                    pointGeom = arcpy.PointGeometry(arcpy.Point(X,Y,Z),self.sr,True,True)
                    newRow = 6000,pointGeom,sec,self.seccornblockname,self.cadaction
                    rows.append(newRow)
        return rows

class Dimension(AzimuthBearingDistance): 
    def __init__(self):
        super(Dimension,self).__init__()
        self.sr = None
        self.insertfields = []
        self.where = None
        self.shapefield="SHAPE@"
        self.objectidfield = "OBJECTID"
        self.linetagfieldname = "LINE_NUMBER"
    def ReturnFirstPoint(self,lyr,where):
        sref = self.sr
        fields = [self.shapefield]
        with arcpy.da.SearchCursor(lyr.name,fields,where_clause=where,spatial_reference=sref) as sc:
            firstPoint = [row[0].firstPoint for row in sc][0]
            return firstPoint

    # this method returns the lastt point of a subject line section line segment returns point geometry
    def ReturnLastPoint(self,lyr,where):
        sref = self.sr
        fields = [self.shapefield]
        with arcpy.da.SearchCursor(lyr.name,fields,where_clause=where,spatial_reference=sref) as sc:
            lastPoint = [row[0].lastPoint for row in sc][0]
            return lastPoint

    # this method returns the closest monument to a point, a where clause can be passed to ensure you are not closing on the same monument that you comenced from
    def ReturnClosestMonument(self,lyr,point,where=None):
        sref = self.sr
        fields = [self.shapefield]
        fields.append(self.objectidfield)
        inX = point.X
        inY = point.Y
        closestDict = {}
        if where==None:
            with arcpy.da.SearchCursor(lyr.name,fields,spatial_reference=sref) as sc:
                for row in sc:
                    point2 = row[0].getPart(0)
                    outX = point2.X
                    outY = point2.Y
                    dist=sqrt((inX-outX)**2+(inY-outY)**2)
                    closestDict[row[1]]=dist,point2
            minDist=(min(closestDict.items(), key=lambda x:x[1]))
            return minDist
        else:
            with arcpy.da.SearchCursor(lyr.name,fields,where_clause=where,spatial_reference=sref) as sc:
                for row in sc:
                    point2 = row[0].getPart(0)
                    outX = point2.X
                    outY = point2.Y
                    dist=sqrt((inX-outX)**2+(inY-outY)**2)
                    closestDict[row[1]]=dist,point2
            minDist=(min(closestDict.items(), key=lambda x:x[1]))
            return minDist

    # this method returns a distnce and azimmuth inverse. returns floats

    ### this method creates  the dimension shape, the dimension text is rendered through a label expression and based on x y returns a row object for insertion with a cursor.
    def CreateDimension(self,point1,point2,spider=False):
        pobaz,poblen = self.returnInverse(point1,point2)
        sref=self.sr
        #print(pobaz,poblen)
        if spider==True:
            offsetaz = pobaz + 90
            if offsetaz>=360:
                offsetaz = offsetaz-360
            pobarray = arcpy.Array()
            pobarray.append(point1)
            pobarray.append(point2)
            pobpoly = arcpy.Polyline(pobarray,sref,False,False)
            pobmid = arcpy.PointGeometry(pobpoly.centroid,sref,False,False)
            pobpolymid = pobmid.pointFromAngleAndDistance(offsetaz,250,"PLANAR")
            pobarray = arcpy.Array()
            pobarray.append(point1)
            pobarray.append(pobpolymid.getPart(0))
            pobarray.append(point2)
            pobpoly = arcpy.Polyline(pobarray,sref,False,False)
            insertrow = pobpoly,pobpoly.firstPoint.X,pobpoly.firstPoint.Y,pobpoly.lastPoint.X,pobpoly.lastPoint.Y,poblen
            return insertrow
            
        else:
            pobarray = arcpy.Array()
            pobarray.append(point1)
            pobarray.append(point2)
            pobpoly = arcpy.Polyline(pobarray,sref,False,False)
            offsetpobpoly = self.copyParallelLeft(pobpoly,50)
            offsetpobpoly = self.scaleGeom(offsetpobpoly,.9)
            pobarray = arcpy.Array()
            pobarray.append(point1)
            pobarray.append(offsetpobpoly.firstPoint)
            pobarray.append(offsetpobpoly.lastPoint)
            pobarray.append(point2)
            pobpoly = arcpy.Polyline(pobarray,sref,False,False)
            insertrow = pobpoly,pobpoly.firstPoint.X,pobpoly.firstPoint.Y,pobpoly.lastPoint.X,pobpoly.lastPoint.Y,poblen
            return insertrow
    
class PlatConstructor(AzimuthBearingDistance):
    def __init__(self):
        super(PlatConstructor,self).__init__()
        self.centerlineplatidentifier = "PLAT"
        self.centerlinenamefield = "NAME"
        self.srworld = 6318
        self.parentcenterlineidfield = "PARENTOID"
        self.bearingstringfieldname = "BEARING_STRING"
        self.linenumberfieldname =  "LINE_NUMBER"
        self.widthfieldname = "WIDTH"
        
    def GetCenterlineShape(self,layer):
        desc = arcpy.Desscribe(layer.name)
        sr = desc.spatialReference
        where = "{0} LIKE '%{1}%'".format(self.centerlinenamefield,self.centerlineplatidentifier)
        centershape = None
        name = None
        width = None
        if desc.fieldInfo.findFieldByName(self.widthfieldname)>-1:
            with arcpy.da.SearchCursor(layer.name,[desc.shapeFieldName,self.centerlinenamefield,self.widthfieldname],where,spatial_reference=sr) as sc:
                for row in sc:
                    if row:
                        centershape = row[0]
                        width = row[2]
                        name = row[1]
                        return centershape,width,name

            return centershape,width,name

    def GetLandOids(self,landpoly,centerlineshape):
        oidList = []
        desc = arcpy.Describe(landpoly.name)
        sr = desc.spatialReference
        with arcpy.da.SearchCursor(landpoly.name,[desc.shapeFieldName+"@",desc.OIDFieldName],spatial_reference=sr) as sc:
            for row in sc:
                if not row[0]: return
                shape = row[0]
                if shape.disjoint(centerlineshape)==False:
                    if not row[1] in oidList:
                        oidList.append(row[1])
        return oidList


    def InsertCenterlineSegments(self,landpoly,centerline,centerlineshape):
        sr = centerlineshape.spatialReference
        descland = arcpy.Describe(landpoly.name)
        desccenter = arcpy.Describe(centerline.name)
        landOids = self.getLandOids(landpoly,centerlineshape)
        centerlineshapes = []
        if landOids:
            if len(landOids)>0:
                for oid in landOids:
                    print(oid)
                    whereLand  = """{0} = {1}""".format(descland.OIDFieldName,oid)
                    with arcpy.da.InsertCursor(centerline.name,[desccenter.shapeFieldName+"@",self.centerlinenamefield,desccenter.OIDFieldName]) as ic:
                        with arcpy.da.SearchCursor(landpoly.name,descland.shapeFieldName+"@",whereLand,spatial_reference=sr) as sc:
                            for row in sc:
                                shape = row[0]
                                intersect = shape.intersect(centerlineshape,2)
                                print(intersect)
                                newRow = intersect,"LINE",6000
                        ic.insertRow(newRow)
                        centerlineshape.append(intersect)
                return centerlineshapes
            else:
                return None


    def GetCenterlineSegments(self,centerline):
        desc = arcpy.Describe(centerline.name)
        sr = desc.spatialReference
        where = """{0} = 'LINE'""".format(self.centerlinenamefield)
        centerdict = {}
        with arcpy.da.SearchCursor(centerline.name,[desc.shapefieldName+"@",desc.OIDFieldName],where,spatial_reference=sr) as sc:
            for row in sc:
                if row[0] is None: continue
                centerdict[row[1]]=row[0]
        return centerdict

    def CreateLabelElements(self,mxd):
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

    def GetTractAttributes(self,centerlineshape,landpoly):
        sr = centerlineshape.spatialReference
        desc = arcpy.Describe(landpoly.name)
        with arcpy.da.SearchCursor(landpoly.name,[desc.shapeFieldName+"@","*"],spatial_reference=sr) as sc:
            for row in sc:
                if row:
                    shape = row[0]
                    if shape.contains(centerlineshape):
                        return row,sc.fields
        return None

    def ProjectLayout2MapCoordinates(self,data_frame,proj_x,proj_y):
        """Convert projected coordinates to map coordinates"""
        # This code relies on the data_frame specified having
        # its anchor point at lower left

        #get the data frame dimensions in map units
        df_map_w = data_frame.elementWidth
        df_map_h = data_frame.elementHeight
        df_map_x = data_frame.elementPositionX
        df_map_y = data_frame.elementPositionY

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


    def AddMeasureValuesToCenterline(self,centerline):
        desc = arcpy.Describe(centerline.name)
        sr = desc.spatialReference
        with arcpy.da.UpdateCursor(centerline.name,[desc.shapeFieldName+"@",desc.OIDFieldName],spatial_reference=sr) as sc:
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


    def GetCenterlineNumbers(self,centerlineshape):
        sr_state = centerlineshape.spatialReference
        sr_world = arcpy.SpatialReference(self.srworld)
        shp=centerlineshape
        length = shp.length
        frstPoint=shp.firstPoint
        arcFirstPoint=frstPoint
        lstPoint=shp.lastPoint
        arcLastPoint=lstPoint
        center=shp.centroid
        frstX = frstPoint.X
        frstY = frstPoint.Y
        lstX = lstPoint.X
        lstY = lstPoint.Y
        cenX = center.X
        cenY = center.Y
        frstGeom=arcpy.PointGeometry(arcFirstPoint,sr_state)
        lstGeom=arcpy.PointGeometry(arcLastPoint,sr_state)
        frstProjPointGeom=frstGeom.projectAs(sr_world)
        lstProjPointGeom=lstGeom.projectAs(sr_world)
        return length,frstPoint,lstPoint,cenX,cenY,frstGeom,lstGeom,frstProjPointGeom,lstProjPointGeom,frstX,frstY,lstX,lstY


    def ReturnTractOrderDict(self,centerlineshape,landPoly):
        tractList=[]
        orderDict = {}
        adjoinerDict={}
        firstTract = None
        lastTract = None
        desc = arcpy.Describe(landPoly.name)
        with arcpy.da.SearchCursor(landPoly.name,[desc.shapeFieldName+"@",desc.OIDFieldName],spatial_reference=centerlineshape.spatialReference) as sc:
            for row in sc:
                shape = row[0]
                centerPoint = shape.centroid
                if shape.disjoint(centerlineshape.firstPoint)==False:
                    firstTract=row[1]
                if shape.disjoint(centerlineshape.lastPoint)==False:
                    lastTract=row[1]
                if shape.disjoint(centerlineshape.firstPoint)==True and shape.disjoint(centerlineshape.lastPoint)==True:
                    if shape.disjoint(centerlineshape)==False:
                        orderDict[row[1]]=centerlineshape.measureOnLine(centerPoint,False)
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
            
    def InsertCenterlineLabelSegments(self,centerlinesegmentshape,labelsegmentlayer,sheet,k):
        desc = arcpy.Describe(labelsegmentlayer.name)
        sr = centerlinesegmentshape.spatialReference
        where = """{0}={1}""".format(self.parentcenterlineidfield,k)
        style = xlwt.XFStyle()
        numStyle = xlwt.XFStyle()
        numStyle2=xlwt.XFStyle()
        style.alignment.wrap=1
        numStyle.alignment.wrap = 1
        numStyle.num_format_str='0'
        numStyle2.alignment.wrap = 1
        numStyle2.num_format_str='0.00'
        
        sheet.write(0,0,"Line Number",style)
        sheet.write(0,1,"Bearing",style)
        sheet.write(0,2,"Distance",style)
        for x in range(3):
            col = sheet.col(x)
            col.width = 16 * 256
        
        ic = arcpy.da.InsertCursor(labelsegmentlayer.name,[desc.OIDFieldName,desc.shapeFieldName+"@",self.linenumberfieldname,self.bearingstringfieldname,self.parentcenterlineidfield])
        pnts = centerlinesegmentshape.getPart(0)
        cnt = 1
        row = sheet.row(cnt)
        row.height = 33 * 256
        dist = 0.0
        for x in range(len(pnts)-1):
            point1 = pnts.getObject(x)
            point2 = pnts.getObject(x+1)
            polyArray = arcpy.Array()
            polyArray.add(point1)
            polyArray.add(point2)
            polyShape = arcpy.Polyline(polyArray,sr,True,True)
            dist = dist+polyShape.length
            az = self.returnAzimuth(polyShape)
            bear  = self.returnBearingString(az)
            bear2 = self.returnBearingString(az)
            print(bear)
            sheet.write(cnt,0,int(cnt),numStyle)
            sheet.write(cnt,1,bear,style)
            sheet.write(cnt,2,round(polyShape.length,2),numStyle2)
            newRow = 6000,polyShape,cnt,bear2,k
            ic.insertRow(newRow)
            cnt+=1
        sheet.write(cnt,1,"Total:",style)
        sheet.write(cnt,2,round(dist,2),numStyle2)
        labelsegmentlayer.definitionQuery=where
        return dist


label = Label()
dime = Dimension()
plat = PlatConstructor()