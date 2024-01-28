
import math
from math import sqrt
import arcpy
import os
__all__=[item for item in dir(arcpy) if not item.startswith('_')]
locals().update(arcpy.__dict__)

class Dimension():
    def __init__(self):
        self.boundaryshape=None 
        self.LorR=None
        self.newgdb=None 
        self.dimscale=None 
        self.dimensionlayer=None 
        self.leadergroup = 1
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

    def ddToDms(self,dd):
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

    def returnBearingString(self,azimuth):
        bearing=None
        dmsbearing=None
        if azimuth>270 and azimuth<=360:
            bearing = 360 - azimuth
            dmsbearing=self.ddToDms(bearing)
            bear = int(dmsbearing[0])
            minute = "{:02d}".format(int(dmsbearing[1]))
            second = "{:02d}".format(int(round(dmsbearing[2],0)))
            dmsbearing = u"""N{0}\xb0{1}'{2}"W""".format(bear,minute,second)
            return dmsbearing
        if azimuth>=0 and azimuth<=90:
            dmsbearing=self.ddToDms(azimuth)
            bear = int(dmsbearing[0])
            minute = "{:02d}".format(int(dmsbearing[1]))
            second = "{:02d}".format(int(round(dmsbearing[2],0)))
            dmsbearing = u"""N{0}\xb0{1}'{2}"E""".format(bear,minute,second)
            return dmsbearing
        if azimuth>90 and azimuth<=180:
            bearing= 180 - azimuth
            dmsbearing=self.ddToDms(bearing)
            bear = int(dmsbearing[0])
            minute = "{:02d}".format(int(dmsbearing[1]))
            second = "{:02d}".format(int(round(dmsbearing[2],0)))
            dmsbearing = u"""S{0}\xb0{1}'{2}"E""".format(bear,minute,second)
            return dmsbearing
        if azimuth>180 and azimuth<=270:
            bearing = azimuth-180
            dmsbearing=self.ddToDms(bearing)
            bear = int(dmsbearing[0])
            minute = "{:02d}".format(int(dmsbearing[1]))
            second = "{:02d}".format(int(round(dmsbearing[2],0)))
            dmsbearing = u"""S{0}\xb0{1}'{2}"W""".format(bear,minute,second)
            return dmsbearing

    def createCornerTies(self,az,start,end,pob=False):
        az180 =  az-180
        if az180<0: az180=az180+360
        ic = arcpy.da.InsertCursor(os.path.join(self.newgdb,self.dimensionlayer.name),["SHAPE@","BEARING"])
        editor = arcpy.da.Editor(self.newgdb)
        editor.startEditing(True)
        editor.startOperation()
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
        point2 = start.pointFromAngleAndDistance(az1,self.dimscale*.3,"PLANAR")
        point3 = start.pointFromAngleAndDistance(az2,self.dimscale,"PLANAR")
        array = arcpy.Array([point3.getPart(0),point2.getPart(0),point1])
        dimline = arcpy.Polyline(array,start.spatialReference,True,True)
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
        point5 = end.pointFromAngleAndDistance(az3,self.dimscale*.3,"PLANAR")
        point6 = end.pointFromAngleAndDistance(az4,self.dimscale,"PLANAR")
        array2 = arcpy.Array([point6.getPart(0),point5.getPart(0),point4])
        dimline2 = arcpy.Polyline(array2,start.spatialReference,True,True)
        array3=arcpy.Array([point3.getPart(0),point6.getPart(0)])
        azb,dis = self.returnInverse(point1,point4)
        bear = self.returnBearingString(azb)
        bearstring = u"""{0} {1}'""".format(bear,round(dis,2))
        dimline3 = arcpy.Polyline(array3,start.spatialReference,True,True)
        if point2.disjoint(self.boundaryshape)==True:
            row = (dimline,None)
            ic.insertRow(row)
            row2=(dimline2,None)
            ic.insertRow(row2)
            row3=(dimline3,bearstring)
            ic.insertRow(row3)
        else:
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
            point2 = start.pointFromAngleAndDistance(az1,self.dimscale*.3,"PLANAR")
            point3 = start.pointFromAngleAndDistance(az2,self.dimscale,"PLANAR")
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
            point5 = end.pointFromAngleAndDistance(az3,self.dimscale*.3,"PLANAR")
            point6 = end.pointFromAngleAndDistance(az4,self.dimscale,"PLANAR")
            array2 = arcpy.Array([point6.getPart(0),point5.getPart(0),point4])
            dimline2 = arcpy.Polyline(array2,start.spatialReference,True,True)
            row2=(dimline2,None)
            ic.insertRow(row2)
            array3=arcpy.Array([point3.getPart(0),point6.getPart(0)])
            dimline3 = arcpy.Polyline(array3,start.spatialReference,True,True)
            azb,dis = self.returnInverse(point1,point4)
            bear = self.returnBearingString(azb)
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

    def CreateSpiderDimension(self,startpoint,endpoint,pob=False):
        desc = arcpy.Describe(self.dimensionlayer.name)
        sr = desc.spatialReference
        point1 = startpoint.getPart(0)
        point2 = endpoint.getPart(0)
        pobaz,poblen = self.returnInverse(point1,point2)
        azminus = pobaz-45
        if azminus<0: azminus=azminus+360
        azplus = pobaz+45
        if azplus>360: azplus = azplus-360
        if pob==False:
            self.leadergroup = self.leadergroup + 1
        bear = self.returnBearingString(pobaz)
        editor = arcpy.da.Editor(self.newgdb)
        editor.startEditing(True)
        editor.startOperation()
        if pob==True: leaderaz = 270
        if pob==False: leaderaz = 90
        ic = arcpy.da.InsertCursor(os.path.join(self.newgdb,self.dimensionlayer.name),["SHAPE@","BEARING","PARENTOID"])
        leader = startpoint.pointFromAngleAndDistance(azplus,(self.dimscale/1.5),"PLANAR")
        lander = leader.pointFromAngleAndDistance(leaderaz,(self.dimscale/10),"PLANAR")
        if lander.disjoint(self.boundaryshape)==True:
            array1 = arcpy.Array([point1,leader.getPart(0),lander.getPart(0)])
            array2 = arcpy.Array([point2,leader.getPart(0),lander.getPart(0)])
            startpoly = arcpy.Polyline(array1,sr,False,False)
            endppoly = arcpy.Polyline(array2,sr,False,False)
            bearstring = """{}-{}'""".format(bear,round(poblen,2))
            row1 = [startpoly,bearstring,self.leadergroup]
            row2 = [endppoly,None,self.leadergroup]
            ic.insertRow(row1)
            ic.insertRow(row2)
        if lander.disjoint(self.boundaryshape)==False:
            if pob==True: leaderaz = 90
            if pob==False: leaderaz = 270
            leader = startpoint.pointFromAngleAndDistance(azminus,(self.dimscale/1.5),"PLANAR")
            lander = leader.pointFromAngleAndDistance(leaderaz,(self.dimscale/10),"PLANAR")
            array1 = arcpy.Array([point1,leader.getPart(0),lander.getPart(0)])
            array2 = arcpy.Array([point2,leader.getPart(0),lander.getPart(0)])
            startpoly = arcpy.Polyline(array1,sr,False,False)
            endppoly = arcpy.Polyline(array2,sr,False,False)
            bearstring = """{}-{}'""".format(bear,round(poblen,2))
            row1 = [startpoly,bearstring,self.leadergroup]
            row2 = [endppoly,None,self.leadergroup]
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
