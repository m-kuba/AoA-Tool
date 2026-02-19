import win32com.client
import math
import os
import time

#Solidworks API
swDocPART = 1   #1 - part / 2 - asembly / 3 - drawing
swSaveAsCurrenVersion = 0
swSaveAsOptions_Silent = 1

def run_aero_sweep():
    #configuration
    angleName = "AoA_Rear_Wing@AirfoilSketch"
    baseFileName = "GT3Wing.sldprt"
    newAngles = [2.0, 6.0, 10.0, 15.0]

    workingDirectory = os.getcwd()
    baseDirectory = os.path.join(workingDirectory, baseFileName)

    if not os.path.exists(baseDirectory):
        print(f"Error: file not found {baseDirectory}")
        return
    
    print ("Launching SolidWorks in the background...")
    swApp = win32com.client.Dispatch("SldWorks.Application")
    swApp.Visible = True    #Change to true to show SolidWorks while working