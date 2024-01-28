import numpy as np
import math
def closest(dist):
    if 7250<dist<=10875: return 500
    if 10875<dist<=14500: 
        return 1000
    # if 14500<dist<=18125: return 500
    # if 18125<dist<=21750: return 1000
    # if 21750<dist<=25375: return 500
    # if 25375<dist<=29000: 
    #     return 1000
    else:
        return None
        
def findPageRange(length):
    n=5
    length = length*1.15
    for x in range(1,n):
        newscale1=length/x
        scalevalue = closest(newscale1)
        if scalevalue:
            pages = math.ceil(length/(7.25*scalevalue))
            pagelength = (length/1.15)/pages
            return(pagelength,pages,scalevalue)
    scalevalue=1000
    pages = math.ceil(length/(7.25*scalevalue))
    pagelength = (length/1.15)/pages
    return(pagelength,pages,scalevalue)


length = 6000
pagelength,pages,scalevalue=findPageRange(length)
print(pagelength,pages,scalevalue)

#### split linesegments at length along line

#### get center of bounding box

#### create viewport tile at scale, pass scale and page number to 