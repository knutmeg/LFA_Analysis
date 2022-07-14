import matplotlib.pyplot as plt
import pandas as pd
from pandas import ExcelWriter
import numpy as np
import matplotlib.pyplot as plt
import os

def LFA_driftcorr(xls_path):
    """Read all sheets of an Excel workbook and return a single DataFrame containing the drift-corrected results.
       Works with both single and multisheet workbooks- no need to differentiate between them."""
    
#initial settings --------------------------------------------------------------------------------------------------------------
    
    filename = os.path.splitext(os.path.basename(xls_path))[0] #gets rid of the extension part of the resulting string
    savename = filename + "_DC.xlsx" #sets a name that the drift corrected results will be saved to
    
    print("\nThe results of this drift correction will be saved as " + savename + " in your current folder. \n")
    
    print(f'Loading {xls_path} into pandas\n') #prints when loading file
    xl = pd.ExcelFile(xls_path) #allows for pandas manipulation of current sheet
    columns = None

    control_displaced = [] #stores names of sheets where the control is displaced. Implies the need to re-read the strip.
    collated_data = pd.DataFrame() # for storing the drift-corrected data in when done
    
    
#sheet-wise iteration and drift correction calcs -------------------------------------------------------------------------------

    for idx, name in enumerate(xl.sheet_names): # for each sheet
        print(f'Reading sheet #{idx}: {name}') #print each sheet name when loading
        sheet = xl.parse(name) #contains all data for the excel sheet! Very important!!!!
        min_control_range = min(sheet[228:259]["Intensity [mV]"]) #control peak should around here [230-261) in sheet)
        min_sheet = min(sheet[0:300]["Intensity [mV]"]) # finds minimum of sheet total for comparison
        corr_check = sheet.loc[sheet["Intensity [mV]"][125:150].idxmax()] #checks if too far above/below baseline        
        baseline = sheet.loc[sheet["Intensity [mV]"][290:300].idxmin()]
    
        if ((corr_check[1] >= (baseline[1] + 50)) or (corr_check[1] <= (baseline[1] - 50))): #if control peak is where it's supposed to be and DCF is necessary, do drift correction.
    
            if (min_control_range == min_sheet):
                print("Calculating DCF...")
                
                DCF_test = sheet.loc[sheet["Intensity [mV]"][28:99].idxmax()] #finds pos/int values for test peak
                DCF_control = sheet.loc[sheet["Intensity [mV]"][222:253].idxmax()] #finds pos/int values for control peak
                
                pos_test = DCF_test[0]; int_test = DCF_test[1] #test position and intensity, respectively
                pos_control = DCF_control[0]; int_control = DCF_control[1] #control position and intensity, respectively
                
                DCF = (int_control - int_test)/(pos_control - pos_test) #calculates drift correction factor
                print("Applying DCF...")
                                          
                pos_corr_start = sheet.loc[sheet["Intensity [mV]"][150:175].idxmax()] #finds max of bridge area
                pos_corr = np.arange(0, (pos_corr_start[0] - 43.98), 0.04)[::-1] #calculates position correction factor from bridge; 0.04 is step size and 43.98 is starting from 44 mm down the test strip
                needs_corr = sheet["Intensity [mV]"][0:pos_corr_start.name + 2] #data that needs correcting
                corr_not_needed = sheet["Intensity [mV]"][pos_corr_start.name + 1:] #data that does not need correction
                
                corrected_values = [] #array to store corrected values in once calculated
                
                for i in range(len(pos_corr)): #calculates drift correction using "original+DCF*position" formula
                    corrected_values.append(needs_corr[i] + DCF *pos_corr[i]) # appends result to new list for use later
                    
                corrected_values.extend(corr_not_needed.tolist()) #adds uncorrected data onto end of list to get whole dataset    
                print("DCF: " + str(DCF) + "\n")
                
                #plt.figure()
                #plt.plot(sheet["Pos [mm]"], sheet["Intensity [mV]"], label = "pre-DC") #graphs the strip's data so you can see what's wrong
                #plt.plot(sheet["Pos [mm]"], corrected_values, label = "post-DC") #graphs the strip's data so you can see what's wrong
                #plt.legend(); plt.title(name)
            
#if the control peak is out of bounds, appends the name of the sheet to the displaced list, moves on.
            else: 
                control_displaced.append(name)
                print("Whoa there! This line's out of control! Better check it out.")
                
#Generating excel files- the output contains the drift corrections of all measurements that the input file had!-----------------
    
            collated_data[name]= corrected_values #adds the corrected values under column [name] to the collation dataframe
            with pd.ExcelWriter(savename) as writer:  
                collated_data.to_excel(writer, sheet_name= "Corrected Intensity (mV)")
    
#checks for any incorrectly run test strips and notifies the user if so. -------------------------------------------------------

        else:
            print("DCF not needed\n")
            #plt.figure()
            #plt.plot(sheet["Pos [mm]"], sheet["Intensity [mV]"], label = "no DC") #graphs the strip's data so you can see what's wrong
            #plt.legend(); plt.title(name)      
            
#Generating excel files- the output contains the drift corrections of all measurements that the input file had!-----------------
            collated_data[name]= sheet["Intensity [mV]"] #adds the corrected values under column [name] to the collation dataframe
            with pd.ExcelWriter(savename) as writer:  
                collated_data.to_excel(writer, sheet_name= "Corrected Intensity (mV)")

#once all relevant DCF is done, check for anything that needs re-reading. --------------------------------------------------
            
    if control_displaced: 
        
        print("\nDisplaced Control peaks detected! Please ensure the test strip was run correctly. All other data has been output to the current folder.")
        
        for name in control_displaced:
            print(name) #prints the name of the offending LFA strip's dataset
            sheet = xl.parse(name) #parses the name into the sheet again
            plt.plot(sheet["Pos [mm]"], sheet["Intensity [mV]"]) #graphs the strip's data so you can see what's wrong

#if no re-reading needs to be done, print this!
    else: 
        
        print("\nAnalysis complete! No displaced control peaks were detected.")
        print("Please see your current working directory for a complete, drift-corrected dataset.") 

    return savename, collated_data

#--------------------------------------------------------------------------------------------------------------------
#--------------------------------------------------------------------------------------------------------------------
#--------------------------------------------------------------------------------------------------------------------

def peak_analysis(savename, collated_data):
    
    """Uses the drift-corrected data to perform peak analysis on the LFA strips; only intakes .xlsx sheets! 
       If using a sheet without an index column, input None for the index argument."""
    
    savename = savename[:-5] + "_PA.xlsx" #sets a name that the drift corrected results will be saved to
    
    print("\nThe results of this peak analysis will be saved as " + savename + " in your current folder. \n")
   
    #defines empty lists and a DataFrame to store everything
    peak47_48 = []; baselines = [];
    test_height = []; max_height = [];
    standard_curve = pd.DataFrame(); conditions = [];
    
    #grabs peak and baseline data
    for i in collated_data.columns: 
        print("Now Analyzing: " + str(i))
        peak47_48.append(min(collated_data[i][75:100])) #finds the 47-48 nm peak values (test)
        baselines.append(min(collated_data[i][290:300])) #finds min of the last 10 points for use as a baseline

    #gets absolute test peak height
    for i in range(len(baselines)):
        test_height.append(baselines[i] - peak47_48[i])
    
    for i in range(len(test_height)):
        if i % 2 == 0: #prints every other result to avoid duplications
            max_height.append(max(test_height[i], test_height[i+1])) #calculates minimum of the two current points
            conditions.append(collated_data.columns[i]) #gets name of run for appending to the dataframe         

    #appends the results to a dataframe specifically for export
    standard_curve["Sample ID"] = conditions
    standard_curve["Peak Height"] = max_height
    
    #uses ExcelWriter to export the previous dataframe to the current working directory using the "savename" as a designator
    with pd.ExcelWriter(savename) as writer:  
        standard_curve.to_excel(writer, sheet_name= "Raw Peak Calculations")
        
    print("\nPeak analysis complete! Have a great day!")
    
    return standard_curve