# LFA_Analysis

A program used to perform drift correction and peak-analysis on data taken from a lateral flow assay (LFA). 

This program was originally created for use with Elabscience's Salbutamol Lateral Flow Assay Kit- the LFAs contained within these kits are approximately 70 by 20 mm in size, and the visible test strip area (not the sample port) was measured at 16 by 3 mm. Other LFAs may have different dimensions, and so this program may need some additional customization for adaptation to those assays. With any questions on how to do so, please contact mkornexl@ues.com.

Please note that the headings of the data used to create this program were automatically exported from the LFA reader used in the experiments (ESEQuant LF3, Qiagen), and that they may differ from reader to reader. In order to ensure seamless functionality across all LFA readers, either the input data's headers or the settings of the code itself should be modified for your specific circumstances. 

## Requirements

In order to run this program, you will need the base installation of Python 3 as well as the following packages, which can be installed via the Anaconda Prompt using the ```pip install``` command if they aren't already present:

- matplotlib
- numpy
- pandas
- os

Additionally, this program intakes and outputs single or multisheet Excel workbooks in the XLSX format (example of a multisheet workbook below), so it is recommended that you run this program on a system with Microsoft Excel 2007 or later installed.

![multipass](https://user-images.githubusercontent.com/30781270/179033532-7a2abaa9-443c-478a-ae61-17bb2a96f7fb.png)


## Running the Program

This program can be run in either the Anaconda Prompt or the Jupyter Notebooks command line. First, navigate to the directory that this program is located at in the terminal of your choice. Then, run the following command:

```
python LFA_Analysis_Code.py
```
The terminal will then ask you to input the full filepath of the workbook you want to analyze- once you've done so, it will automatically begin the analysis (example below). 

![Command Line](https://user-images.githubusercontent.com/30781270/179044041-011bdcf5-5a87-4f0c-8a31-9d3eebecb1f7.PNG)

![Command Line 2](https://user-images.githubusercontent.com/30781270/179044654-8ab4fcc6-bc3f-4f11-8cf9-ed837b245b4b.PNG)

Once the program has finished running, you will find two new Excel workbooks in your current working directory. One will be named *<input_filename>_DC*, and the other will be named *<input_filename>_PA*. These files contain the drift-corrected and peak-analysed datasets, respectively, and both can then be used in statistical analysis or other such procedures. 

**Note: if this script is run again in the same directory, any previous DC or PA files with the same name as the current input file will be overwritten unless they are open in Microsoft Excel at the time.**

### Usage notes

If the input workbook contains 100+ sheets, you may encounter a warning like the one below.

![Error](https://user-images.githubusercontent.com/30781270/179035647-1aa64b7f-6543-4b14-a9c8-079f90576c72.PNG)

Note that this is simply a warning regarding the efficiency of the code, and does not affect its overall function much aside from increasing the time it takes to run by ~2 minutes (as per the test workbook of 305 sheets). 
