import matplotlib.pyplot as plt
import pandas as pd
from pandas import ExcelWriter
import numpy as np
import matplotlib.pyplot as plt
import os

from Necessary_Functions import *

print("\nHello! Please input the file path to the Excel workbook of interest.\n")
filepath = str(input())

savename, DC_data = LFA_driftcorr(filepath)

peak_heights = peak_analysis(savename, DC_data)