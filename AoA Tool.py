import win32com.client
import math
import os
import time

#Solidworks API
swDocPART = 1   #1 - part / 2 - asembly / 3 - drawing
swSaveAsCurrenVersion = 0
swSaveAsOptions_Silent = 1
swSaveAsOptions_Copy = 2

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

    try:
        print("Opening base model: {baseFileName}...")
        Model = swApp.openDoc6(baseDirectory, swDocPART, 1, "", 0, 0)   #(directory, file type, silent mode, configuration, error code, error code)

        if Model is None:
            print ("Error: fialed to open file")
            return
        
        swDimension = Model.Parameter(angleName)
        if swDimension is None:
            print("Critical error: angle {angleName} not found, check the name in SolidWorks")
            return
        
        print("Generationg variants...")

        for angleDeg in newAngles:
            angleRad = angleDeg * (math.pi / 180.0)     #Conversion to radians
            swDimension.SystemValue = angleRad
            Model.EditRebuild3()

            newFileName = f"GT3Wing_{angleDeg}deg.sldprt"
            saveDirectory = os.path.join(workingDirectory, newFileName)

            success = Model.Extension.SaveAs3(saveDirectory, swSaveAsCurrenVersion, swSaveAsOptions_Silent, None, 0, 0)

            if success:
                print(f"Generation successful: {newFileName}")
            
            else:
                print(f"File save error")

            time.sleep(0.5) #Pause for file handling

        print("Variants generation successful, new files should be present in folder")
    
    except Exception as e:
        print(f"\nUnexpecter error: {e}")

    finally:
        if 'Model' in locals() and Model is not None:
            swApp.CloseDoc(Model.GetTitle())
            print("Base model file closed")

if __name__ == "__main__":
    run_aero_sweep()