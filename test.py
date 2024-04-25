from math import log10
from openpyxl import load_workbook, Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    
# -------------------
# Read Co-locate.xlsx
# -------------------

wb = load_workbook('Co-locate.xlsx', read_only=True, data_only=True)
ws_Bands = wb['Bands']
ws_OR_Block = wb['Override-blocking']
ws_OR_SpEm = wb['Override-SpEm']

   end(noreg)
    regions.append(line)
