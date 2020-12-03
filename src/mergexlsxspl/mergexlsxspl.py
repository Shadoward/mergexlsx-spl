# -*- coding: utf-8 -*-
###############################################################
# Author:       patrice.ponchant@furgo.com  (Fugro Brasil)    #
# Created:      03/12/2020                                    #
# Python :      3.x                                           #
###############################################################

# The future package will provide support for running your code on Python 2.6, 2.7, and 3.3+ mostly unchanged.
# http://python-future.org/quickstart.html
from __future__ import (absolute_import, division,
                        print_function, unicode_literals)
from builtins import *

##### Basic packages #####
import pandas as pd
import glob
import os, sys
import datetime

##### GUI packages #####
from gooey import Gooey, GooeyParser
from colored import stylize, attr, fg

# 417574686f723a205061747269636520506f6e6368616e74
##########################################################
#                       Main code                        #
##########################################################
# https://pythonpedia.com/en/knowledge-base/30635145/create-multiple-dataframes-in-loop

# this needs to be *before* the @Gooey decorator!
# (this code allows to only use Gooey when no arguments are passed to the script)
if len(sys.argv) >= 2:
    if not '--ignore-gooey' in sys.argv:
        sys.argv.append('--ignore-gooey')

# GUI Configuration
@Gooey(
    program_name='Merge XLSX from the splsensors tool',
    richtext_controls=True,
    #richtext_controls=True,
    terminal_font_family = 'Courier New', # for tabulate table nice formatation
    #dump_build_config=True,
    #load_build_config="gooey_config.json",
    default_size=(600, 500),
    timing_options={        
        'show_time_remaining':True,
        'hide_time_remaining_on_complete':True
        },
    header_bg_color = '#95ACC8',
    #body_bg_color = '#95ACC8',
    menu=[{
        'name': 'File',
        'items': [{
                'type': 'AboutDialog',
                'menuTitle': 'About',
                'name': 'mergexlsx-spl',
                'description': 'Merge XLSX from the splsensors tool',
                'version': '0.1.0',
                'copyright': '2020',
                'website': 'https://github.com/Shadoward/mergexlsx-spl',
                'developer': 'patrice.ponchant@fugro.com',
                'license': 'MIT'
                }]
        }]
    )

def main():
    desc = "Merge XLSX from the splsensors tool"    
    parser = GooeyParser(description=desc)
    
    main = parser.add_argument_group('Main', gooey_options={'columns': 1})
    main.add_argument(
        '-i', '--input',
        dest='inputFolder',
        metavar='Input Logs Folder',  
        help='Input folder to merge all the logs files. (*_FINAL_Log.xlsx)',      
        widget='DirChooser',
        gooey_options={'wildcard': "Logs SPL files (*.xlsx)|*.xlsx"})
    
    # Use to create help readme.md. TO BE COMMENT WHEN DONE
    # if len(sys.argv)==1:
    #    parser.print_help()
    #    sys.exit(1)   
        
    args = parser.parse_args()
    process(args)

def process(args):
    """
    Uses this if called as __main__.
    """
    
    inputFolder = args.inputFolder
    excel_names = glob.glob(inputFolder + '\\*_Log.xlsx')
    
    print('')
    print(f'Merging the following files.\n {excel_names}\nPlease wait.......')

    xl = pd.ExcelFile(excel_names[0])
    lsSH = xl.sheet_names
    d = {name: combine_excel_to_dfs(excel_names, name) for name in xl.sheet_names}
    d['Summary_Process_Log'] = pd.DataFrame()
    
    # Create a Pandas Excel writer using XlsxWriter as the engine.
    if os.path.exists(inputFolder + '\\sheets_combined.xlsx'):
        os.remove(inputFolder + '\\sheets_combined.xlsx')

    writer = pd.ExcelWriter(inputFolder + '\\sheets_combined.xlsx', engine='xlsxwriter')

    # Write each dataframe to a different worksheet.

    workbook  = writer.book

    for name, df in d.items():
        df.to_excel(writer, sheet_name=name)

    w = {name: writer.sheets[name] for name in xl.sheet_names}
    w['Summary_Process_Log'].hide_gridlines(2) 

    #### Set format       
    bold = workbook.add_format({'bold': True,
                                'font_name': 'Segoe UI',
                                'font_size': 10,
                                'valign': 'vcenter',})
    normal = workbook.add_format({'bold': False,
                                'font_name': 'Segoe UI',
                                'font_size': 10,
                                'valign': 'vcenter',})
    hlink = workbook.add_format({'bold': False,
                                'font_color': '#0250AE',
                                'underline': True,
                                'font_name': 'Segoe UI',
                                'font_size': 10,
                                'valign': 'vcenter',})

    cell_format = workbook.add_format({'text_wrap': True,
                                    'font_name': 'Segoe UI',
                                    'font_size': 10,
                                    'valign': 'vcenter',
                                    'align': 'left',
                                    'border_color': '#000000',
                                    'border': 1})

    header_format = workbook.add_format({'bold': True,
                                        'font_name': 'Segoe UI',
                                        'font_size': 12,
                                        'text_wrap': False,
                                        'valign': 'vcenter',
                                        'align': 'left',
                                        'fg_color': '#011E41',
                                        'font_color': '#FFFFFF',
                                        'border_color': '#FFFFFF',
                                        'border': 1})

    textFull = [bold, 'Full_List', normal, ': Full log list of all sensors without duplicated and skip files']
    textMissingSPL = [bold, 'Missing_SPL', normal, ': List of all sennsors that have missing SPL SPL file that']
    textMBES = [bold, 'MBES_NotMatching', normal, ': MBES log list of all files that do not match the SPL name; without duplicated and skip files']
    textSSS = [bold, 'SSS_NotMatching', normal, ': SSS log list of all files that do not match the SPL name; without duplicated and skip files']
    textSBP = [bold, 'SBP_NotMatching', normal, ': SBP log list of all files that do not match the SPL name; without duplicated and skip files']
    textMAG = [bold, 'MAG_NotMatching', normal, ': MAG log list of all files that do not match the SPL name; without duplicated and skip files']
    textSUHRS = [bold, 'SUHRS_NotMatching', normal, ': SUHRS log list of all files that do not match the SPL name; without duplicated and skip files']
    textDuplSPL = [bold, 'Duplicated_SPL_Name', normal, ': List of all duplicated SPL name']
    textDuplSensor = [bold, 'Duplicated_Sensor_Data', normal, ': List of all duplicated sensors files; Based on the start time']
    textSPLProblem = [bold, 'SPL_Problem', normal, ': List of all SPL session without a line name in the columns LineName, are empty or too small']
    textSkip = [bold, 'Skip_SSS_Files', normal, ': List of all SSS data that have a file size less than 1 MB']
    textsgy = [bold, 'Wrong_SBP_Time', normal, ': List of all SBP data that have a wrong timestamp']

    ListT = [textFull, textMissingSPL, textMBES, textSSS, textSBP, textMAG, textSUHRS, textDuplSPL, textDuplSensor, 
            textSPLProblem, textSkip, textsgy]
    ListHL = ['internal:Full_List!A1', 'internal:Missing_SPL!A1', 'internal:MBES_NotMatching!A1', 
            'internal:SSS_NotMatching!A1', 'internal:SBP_NotMatching!A1', 'internal:MAG_NotMatching!A1',
            'internal:SUHRS_NotMatching!A1','internal:Duplicated_SPL_Name!A1', 'internal:Duplicated_Sensor_Data!A1', 
            'internal:SPL_Problem!A1', 'internal:Skip_SSS_Files!A1', 'internal:Wrong_SBP_Time!A1']
                
    icount = 1
    for e, l in zip(ListT,ListHL):
        w['Summary_Process_Log'].write_rich_string(icount, 1, *e)
        w['Summary_Process_Log'].write_url(icount, 0, l, hlink, string='Link')
        icount += 1

    # Others Sheet
    for name, ws in w.items():
        if name != 'Summary_Process_Log':
            ws.write_url(0, 0, 'internal:Summary_Process_Log!A1', hlink, string='Summary')

    for (namedf, df), (namews, ws) in zip(d.items(), w.items()):
        if namedf != 'Summary_Process_Log':
            ws.autofilter(0, 0, df.shape[0], df.shape[1])
            ws.set_column(0, 0, 15, cell_format)
            ws.set_column(1, 4, 24, cell_format)
            ws.set_column(5, df.shape[1], 50, cell_format)
            for col_num, value in enumerate(df.columns.values):
                ws.set_row(0, 25)
                ws.write(0, col_num + 1, value, header_format)
    #    for i, width in enumerate(get_col_widths(df)): # Autosize will not work because of the "\n" in the text
    #        ws.set_column(i, i, width, cell_format)

    # Add a format.
    fWRONG = workbook.add_format({'bg_color': '#FFC7CE',
                                'font_color': '#9C0006'})
    fOK = workbook.add_format({'bg_color': '#C6EFCE',
                                'font_color': '#006100'})
    fBLANK = workbook.add_format({'bg_color': '#FFFFFF',
                                'font_color': '#000000'})
    fDUPL = workbook.add_format({'bg_color': '#2385FC',
                                'font_color': '#FFFFFF'})
    fWSPL = workbook.add_format({'bg_color': '#C90119',
                                'font_color': '#FFFFFF'})

    # Highlight the values (first is overwrite the others below.....)
    ListFC = [w['Full_List'], w['MBES_NotMatching'], w['SSS_NotMatching'], w['SBP_NotMatching'], w['MAG_NotMatching'], w['SUHRS_NotMatching']]

    # use the bigger df 
    color_range1 = "E2:E{}".format(d['Full_List'].shape[0]+1)
    color_range2 = "F2:J{}".format(d['Full_List'].shape[0]+1)

    for i in ListFC:
        i.conditional_format(color_range1, {'type': 'text',
                                            'criteria': 'containing',
                                            'value':    'SPLtoSmall',
                                            'format': fWSPL})
        i.conditional_format(color_range1, {'type': 'text',
                                            'criteria': 'containing',
                                            'value':    'NoLineNameFound',
                                            'format': fWSPL})
        i.conditional_format(color_range1, {'type': 'text',
                                            'criteria': 'containing',
                                            'value':    'EmptySPL',
                                            'format': fWSPL})
        i.conditional_format(color_range1, {'type': 'duplicate',
                                            'format': fDUPL})
        i.conditional_format(color_range2, {'type': 'blanks',
                                            'format': fBLANK})
        i.conditional_format(color_range2, {'type': 'text',
                                            'criteria': 'containing',
                                            'value':    '[WRONG]',
                                            #'criteria': '=NOT(ISNUMBER(SEARCH($E2,F2)))',
                                            'format': fWRONG})
        i.conditional_format(color_range2, {'type': 'text',
                                            'criteria': 'containing',
                                            'value':    '[OK]',
                                            'format': fOK}) 

    # Close the Pandas Excel writer and output the Excel file.
    writer.save() 


# https://stackoverflow.com/questions/48780464/how-to-combine-multiple-excel-files-having-multiple-equal-number-of-sheets-in-ea
def combine_excel_to_dfs(excel_names, sheet_name):
    sheet_frames = [pd.read_excel(x, sheet_name=sheet_name) for x in excel_names]
    combined_df = pd.concat(sheet_frames)
    combined_df = combined_df.drop(combined_df.columns[0], axis=1)
    return combined_df


##########################################################
#                        __main__                        #
########################################################## 
if __name__ == "__main__":
    now = datetime.datetime.now() # time the process
    # Preparing your script for packaging https://chriskiehl.com/article/packaging-gooey-with-pyinstaller
    # Prevent stdout buffering     
    #nonbuffered_stdout = os.fdopen(sys.stdout.fileno(), 'w') #https://stackoverflow.com/questions/45263064/how-can-i-fix-this-valueerror-cant-have-unbuffered-text-i-o-in-python-3/45263101
    #sys.stdout = nonbuffered_stdout
    main()
    print('')
    print("Process Duration: ", (datetime.datetime.now() - now)) # print the processing time. It is handy to keep an eye on processing performance.