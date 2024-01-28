from ctypes import *
import os

def conCoords(sysOutNum,outdatyear,indatyear,outzonecode,inX,inY,inZ):
     corpslib = windll.LoadLibrary(r"C:\Program Files\CORPSCON6\corpscon_v6.dll")
     test00 = corpslib.corpscon_default_config()
     SetNadconPath = corpslib.SetNadconPath
     SetVertconPath = corpslib.SetVertconPath
     SetGeoidPath = corpslib.SetGeoidPath
     SetInSystem = corpslib.SetInSystem
     SetOutSystem = corpslib.SetOutSystem
     SetInDatum = corpslib.SetInDatum
     SetOutDatum = corpslib.SetOutDatum
     SetInZone = corpslib.SetInZone
     SetOutZone = corpslib.SetOutZone
     SetInUnits = corpslib.SetInUnits
     SetOutUnits = corpslib.SetOutUnits
     SetInVDatum = corpslib.SetInVDatum
     SetOutVDatum = corpslib.SetOutVDatum
     SetInVUnits = corpslib.SetInVUnits
     SetOutVUnits = corpslib.SetOutVUnits
     SetGeoidCodeBase = corpslib.SetGeoidCodeBase
     SetUseGeoidCustomAreas = corpslib.SetUseGeoidCustomAreas
     SetGeoidCustomAreaListFile = corpslib.SetGeoidCustomAreaListFile
     SetXIn = corpslib.SetXIn
     SetYIn = corpslib.SetYIn
     SetZIn = corpslib.SetZIn
     GetXOut = corpslib.GetXOut
     GetYOut = corpslib.GetYOut
     GetZOut = corpslib.GetZOut

     import ctypes

     SetNadconPath.argtypes = [ctypes.c_char_p]
     SetNadconPath.retval = [ctypes.c_int]
     SetVertconPath.argtypes = [ctypes.c_char_p]
     SetVertconPath.retval = [ctypes.c_int]
     SetGeoidPath.argtypes = [ctypes.c_char_p]
     SetGeoidPath.retval = [ctypes.c_int]
     SetInSystem.argtypes = [ctypes.c_int]  
     SetInSystem.retval = [ctypes.c_int]  
     SetInDatum.argtypes = [ctypes.c_int]  
     SetInDatum.retval = [ctypes.c_int]  
     SetOutDatum.argtypes = [ctypes.c_int]
     SetOutDatum.retval = [ctypes.c_int]
     SetOutSystem.argtypes = [ctypes.c_int]  
     SetOutSystem.retval = [ctypes.c_int]  
     SetInZone.argtypes = [ctypes.c_int]      
     SetInZone .retval = [ctypes.c_int]  
     SetOutZone.argtypes = [ctypes.c_int]    
     SetOutZone .retval = [ctypes.c_int] 
     SetInUnits.argtypes = [ctypes.c_int]  
     SetInUnits .retval = [ctypes.c_int]   
     SetInVDatum.argtypes = [ctypes.c_int]     
     SetInVDatum .retval = [ctypes.c_int]     
     SetOutVDatum.argtypes = [ctypes.c_int]    
     SetOutVDatum .retval =[ctypes.c_int]    
     SetInVUnits.argtypes = [ctypes.c_int]   
     SetInVUnits .retval = [ctypes.c_int]   
     SetOutVUnits.argtypes = [ctypes.c_int]    
     SetOutVUnits .retval = [ctypes.c_int]  
     SetGeoidCodeBase.argtypes = [ctypes.c_int]    
     SetGeoidCodeBase.retval = [ctypes.c_int]  
     SetUseGeoidCustomAreas.argtypes = [ctypes.c_int]
     SetUseGeoidCustomAreas.retval = [ctypes.c_int]
     SetGeoidCustomAreaListFile.argtypes = [ctypes.c_char_p]
     SetGeoidCustomAreaListFile.retval = [ctypes.c_int]
     SetXIn.argtypes = [ctypes.c_double]  
     SetXIn.retval = [ctypes.c_int]  
     SetYIn.argtypes = [ctypes.c_double]  
     SetYIn.retval = [ctypes.c_int]  
     SetZIn.argtypes = [ctypes.c_double]  
     SetZIn.retval = [ctypes.c_int]  
     GetXOut.retval =[ctypes.c_double]  
     GetYOut.retval =[ctypes.c_double]  
     GetZOut.retval = [ctypes.c_double]  

     test1 = SetNadconPath(r"C:\Program Files\CORPSCON6\Nadcon")
     test2 = SetVertconPath(r"C:\Program Files\CORPSCON6\Vertcon")
     test3 = SetGeoidPath(r"C:\Program Files\CORPSCON6\Geoid")

    ###################################
    ##
    ## Set Geographic or Stateplane
    ##
    ##################################
     sysInNum = 2 
     insys = SetInSystem(sysInNum)
     outsys = SetOutSystem(sysOutNum)

     ###################################
     ##
     ## Set Datum 83 or 27
     ##
     ##################################
     datumInYear = indatyear
     indat = SetInDatum(datumInYear)
     datumOutYear = outdatyear
     outdat = SetOutDatum(datumOutYear)

     ###################################
     ##
     ## Set Zone ie 4202
     ##
     ##################################

     incode = outzonecode
     inzone = SetInZone(outzonecode)

     outcode = outzonecode
     outzone = SetOutZone(outcode)

     ###################################
     ##
     ## Set Units
     ##
     ##################################

     units = 1

     outunits = SetOutUnits(units)
     inunits = SetInUnits(units)


     ###################################
     ##
     ## Set V Datum
     ##
     ##################################

     invdatum = SetInVDatum(1988)
     outvdatum = SetOutVDatum(1988)


     ###################################
     ##
     ## Set V Units
     ##
     ##################################

     invunits = SetInVUnits(1)
     outvunits = SetOutVUnits(1)

     geoidbase = SetUseGeoidCustomAreas(1)#SetGeoidCodeBase(2003)
     geoidbasefile = SetGeoidCustomAreaListFile(r"C:\Program Files\CORPSCON6\Geoid\geoid12b.txt")

     intcorpse = corpslib.corpscon_initialize_convert()

     #inX = 2790955
     #inY = 503380 
     #inZ = 2800.00

     xin = c_double(inX)
     yin = c_double(inY)
     zin = c_double(inZ)

     xout = c_double()
     yout = c_double()
     zout = c_double()

     SetXIn(xin)
     SetYIn(yin)
     SetZIn(zin)

     corpslib.corpscon_convert()

     corpslib.GetXOut.restype = c_double
     corpslib.GetYOut.restype = c_double
     corpslib.GetZOut.restype = c_double

     pntX = corpslib.GetXOut()
     pntY = corpslib.GetYOut()
     pntZ =  corpslib.GetZOut()

     handle = corpslib._handle
     del corpslib
     ctypes.windll.kernel32.FreeLibrary(handle)
     return pntX,pntY,pntZ


#ex,y,z = conCoords(1,1983,1927,4203,979348.906947841,715374.233160408,3002.28)
#print(ex,y,z)
