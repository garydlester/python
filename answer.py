from django.contrib.gis import gdal

def rotateXY(x,y,xc=0,yc=0,angle=0,units = "DEGREES"):
    import math
    x = x - xc
    y = y - yc
    if units=="DEGREES":
        angle = math.radians(angle)
    xr = (x*math.cos(angle))-(y*math.sin(angle)) + xc 
    yr = (y*math.sin(angle))+(y*math.cos(angle)) + yc
    return xr,yr


def createIndexPoly(X,Y,h=5,w=5,sr=None):
    angle = 20
    if not h is None and not w is None:
        ### set up editor
        
        arrPnts = arcpy.Array()  
        # point 1  
        pnt = arcpy.Point(X-w/2,Y-h/2)  
        pnt = arcpy.Point(rotateXY(pnt.X,pnt.Y,X,Y,angle)[0],rotateXY(pnt.X,pnt.Y,X,Y,angle)[1])
        arrPnts.add(pnt)  
        # point 2  
        pnt = arcpy.Point(X-w/2,Y+h/2)  
        pnt = arcpy.Point(rotateXY(pnt.X,pnt.Y,X,Y,angle)[0],rotateXY(pnt.X,pnt.Y,X,Y,angle)[1])
        arrPnts.add(pnt)  
        # point 3  
        pnt = arcpy.Point(X+w/2,Y+h/2)  
        pnt = arcpy.Point(rotateXY(pnt.X,pnt.Y,X,Y,angle)[0],rotateXY(pnt.X,pnt.Y,X,Y,angle)[1])
        arrPnts.add(pnt)  
        # point 4  
        pnt = arcpy.Point(X+w/2,Y-h/2)  
        pnt = arcpy.Point(rotateXY(pnt.X,pnt.Y,X,Y,angle)[0],rotateXY(pnt.X,pnt.Y,X,Y,angle)[1])
        arrPnts.add(pnt)  
        # point 5 (close diamond)  
        pnt = arcpy.Point(X-w/2,Y-h/2)  
        pnt = arcpy.Point(rotateXY(pnt.X,pnt.Y,X,Y,angle)[0],rotateXY(pnt.X,pnt.Y,X,Y,angle)[1])
        arrPnts.add(pnt)  
        pol = arcpy.Polygon(arrPnts,sr,True,True)
        
        return pol



print(createIndexPoly(3000000,1000000))