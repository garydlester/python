import numpy as np

def convertToNumpy(geom):
    pointlist=[]
    if geom.type=='point':
        nparray = np.array([geom.getPart(0).X,geom.getPart(0).Y],dtype='float64')
        return nparray
    if  geom.type=='polyline' or geom.type=='polygon':
        for pnt in geom.getPart(0):
            pointlist.append([pnt.X,pnt.Y])
        nparray = np.array(pointlist,dtype='float64')
        return nparray