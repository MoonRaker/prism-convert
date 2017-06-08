
import logging
import numpy as np
import os
import pandas as pd

import excel


"""
.. module:: prism_convert
   :platform: Windows
   :synopsis: A useful module indeed.

.. moduleauthor:: Derek Groenendyk
12/20/2016

"""

logging.basicConfig(filename='prism_convert.log', 
                    level=logging.DEBUG,
                    format='%(asctime)s %(name)-15s %(levelname)-8s %(message)s',
                    datefmt='%m-%d %H:%M',
                    filemode='w')
logger = logging.getLogger('convert_main')

# vis = True
# ex = excel.Excel(vis)


def import_data(directory, afile):
    """
    Reads the site input file and initilizes a SITE object.

    Returns
    -------
    sites: list
        SITE objects for each input location.
    """

    ex_fname = os.path.join(directory, afile)
    # ex_fname = os.path.join(directory, afile)

    vis = False
    ex = excel.Excel(ex_fname, vis, False)
    # wb = ex.wb.Sheets(1)

    # wb = ex.open_workbook(, vis)

    if ex == None:
        logger.critical('Please close workbook: ' + afile)
        raise SystemExit
    ws = ex.wb.Sheets(1)

    numrows = ws.UsedRange.Rows.Count
    site_names = [ws.Cells(row+2,3).Value for row in range(numrows-1)]

    raw_data = np.zeros((len(site_names), 5, 12))

    for i in range(len(site_names)):
        raw_data[i,0,:] = [float(ws.Cells(i+2,col).Value) for col in range(6,18)]
        raw_data[i,1,:] = [float(ws.Cells(i+2,col).Value) for col in range(19,31)]
        raw_data[i,2,:] = [float(ws.Cells(i+2,col).Value) for col in range(32,44)]
        raw_data[i,3,:] = [float(ws.Cells(i+2,col).Value) for col in range(45,57)]
        raw_data[i,4,:] = [float(ws.Cells(i+2,col).Value) for col in range(58,70)]

    dp = pd.Panel(raw_data, items=site_names, major_axis=['td', 'p', 'th', 'tl', 'ta'])

    logger.info('finished importing site data')
    ex.close_workbook(0)
    # wb.Close(0)

    return dp


def convert_data(dp):

    site_names = []
    precips = []
    temps = [] 

    for item in dp.items:
        if item != None:
            site_names.append(item)
            precips.append(dp[item].loc['p'].values)
            temps.append(dp[item].loc['t'].values)

    cdata_dict = {}
    for i in range(len(site_names)):
        cdata_dict[site_names[i]] = pd.DataFrame([precips[i], temps[i]], index=['td', 'p', 'th', 'tl', 'ta'])

    dp = pd.Panel(cdata_dict)

    return dp


def save_data(directory, year, cdata):

    vis = False

    dataout_path = os.path.join(directory, 'wx_prism.xlsx')

    # dataout_path = os.path.join(directory, 'CU_Plots.xlsx')
    # ex_fname = os.path.join(directory, afile)

    # ex = excel.Excel(dataout_path, True, True)
    # wb = ex.wb.Sheets(1)

    # wb = ex.open_workbook(, vis)

    # if ex == None:
    #     logger.critical('Please close workbook: ' + afile)
    #     raise SystemExit
    # ws = ex.wb.Sheets(1)

    if os.path.exists(dataout_path):
        # wb = ex.open_workbook(dataout_path, vis)
        ex = excel.Excel(dataout_path, vis, False)
    else:
        # wb = ex.createBook(dataout_path, vis)
        ex = excel.Excel(dataout_path, vis, True)

    if ex.wb.ReadOnly:
        logger.critical('Close workbook: wx_prism.xlsx')
        raise SystemExit        
    else:
        wb = ex.wb
    # ws = ex.wb.Sheets(1)

    # wb = ex.open_workbook(siteout_path, vis)
    # if wb_yearly == None:
    #     logger.critical('Close workbook: '+'sites_out.xlsx')
    #     raise SystemExit
    # else:
    #     wb_yearly.Close(0)
    
    # wb = ex.createBook(dataout_path, vis)

    k = -1
    for item in cdata.items:
        k += 1
        try:
            # ws = wb.Worksheets(item.replace(' ', ''))
            ws = wb.Worksheets(item)
        except:
            ws = wb.Worksheets.Add()
            # ws.Name = item.replace(' ', '')
            ws.Name = item
            columns = ['Year', 'Month', 'Precipitation', 'Temp_Avg', 'Temp_Min', 'Temp_Max',  'Dew_Point']
            for i in range(len(columns)):
                ws.Cells(1,i+1).Value = columns[i]

        numrows = ws.UsedRange.Rows.Count

        flag = False
        if numrows > 1:
            if int(ws.Cells(numrows,1).Value) < int(year):
                flag = True
        else:
            flag = True

        if flag:
            row = numrows
            for ind in range(12):
                row += 1
                ws.Cells(row,1).Value = year
                ws.Cells(row,2).Value = ind + 1
                ws.Cells(row,3).Value = round(cdata[item].loc['p'].values[ind],2)  
                ws.Cells(row,4).Value = int(cdata[item].loc['tl'].values[ind])
                ws.Cells(row,5).Value = int(cdata[item].loc['th'].values[ind])
                ws.Cells(row,6).Value = int(cdata[item].loc['ta'].values[ind])
                ws.Cells(row,7).Value = int(cdata[item].loc['td'].values[ind])         

    # wb.Save(1)
    # wb.Close(1)
    ex.close_workbook(1)


def main():

    directory = os.getcwd()
    dir_list = os.listdir(directory)
    ext_list = [os.path.splitext(afile)[1] for afile in dir_list]

    file_list = []
    for i in range(len(ext_list)):
        if ext_list[i] == '.dbf':
            file_list.append(dir_list[i])

    for afile in file_list:
        year = os.path.splitext(afile)[0][-4:]
        raw_data = import_data(directory, afile)
        # raw_data = convert_data(raw_data)
        save_data(directory, year, raw_data)





if __name__=="__main__":
    main()