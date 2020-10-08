import pandas as pd
from pathlib import Path
import sys


def file_diff(path_OLD, path_NEW, index_col):

    # check to see what file type and read into pandas
    if str(path_OLD).lower().endswith('.csv'):
        OLD = pd.read_csv(path_OLD, index_col=index_col).fillna(0)
    elif str(path_OLD).lower().endswith('.xlsx'):
        OLD = pd.read_excel(path_OLD, index_col=index_col).fillna(0)   

    if str(path_NEW).lower().endswith('.csv'):
        NEW = pd.read_csv(path_NEW, index_col=index_col).fillna(0)
    elif str(path_NEW).lower().endswith('xlsx'):
        NEW = pd.read_excel(path_NEW, index_col=index_col).fillna(0)

    # find differences
    dfDiff = NEW.copy()
    droppedRows = []
    newRows = []

    cols_OLD = OLD.columns
    cols_NEW = NEW.columns

    intersectCols = list(set(cols_OLD).intersection(cols_NEW))

    for row in dfDiff.index:
        if (row in OLD.index) and (row in NEW.index):
            for col in intersectCols:
                value_OLD = OLD.loc[row,col]
                value_NEW = NEW.loc[row,col]
                if value_OLD==value_NEW:
                    dfDiff.loc[row,col] = NEW.loc[row,col]
                else:
                    dfDiff.loc[row,col] = ('{}→{}').format(value_OLD,value_NEW)
        else:
            newRows.append(row)

    for row in OLD.index:
        if row not in NEW.index:
            droppedRows.append(row)
            dfDiff = dfDiff.append(OLD.loc[row,:])
    
    # remove columns that are different in dfDiff only 
    diffCols = cols_NEW.difference(cols_OLD)
    diffCols2 = cols_OLD.difference(cols_NEW)
    for diff in diffCols:
        del dfDiff[diff]
    for diff in diffCols2:
        del dfDiff[diff]
        
    dfDiff = dfDiff.sort_index().fillna('')

    # save output as excel and format
    fname = '{} vs {}.xlsx'.format(path_OLD.stem,path_NEW.stem)
    writer = pd.ExcelWriter(fname, engine='xlsxwriter')

    dfDiff.to_excel(writer, sheet_name='DIFF', index=True)
    NEW.to_excel(writer, sheet_name=path_NEW.stem, index=True)
    OLD.to_excel(writer, sheet_name=path_OLD.stem, index=True)

    # get xlsxwriter objects
    workbook  = writer.book
    worksheet = writer.sheets['DIFF']
    worksheet.hide_gridlines(2)
    worksheet.set_default_row(15)

    # define formats
    rmvd_fmt = workbook.add_format({'font_color': '#E0E0E0'})
    change_fmt = workbook.add_format({'font_color': '#FF0000', 'bg_color':'#B1B3B3'})
    add_fmt = workbook.add_format({'font_color': '#32CD32','bold':True})

    # set format over range
    ## highlight changed cells
    worksheet.conditional_format('A1:ALL100000', {'type': 'text',
                                            'criteria': 'containing',
                                            'value':'→',
                                            'format': change_fmt})

    # highlight new/changed rows
    for row in range(dfDiff.shape[0]):
        if row+1 in newRows:
            worksheet.set_row(row+1, 15, add_fmt)
        if row+1 in droppedRows:
            worksheet.set_row(row+1, 15, rmvd_fmt)

    # save
    writer.save()
    print('\nDone.\n')


def main(): 

    if len(sys.argv) < 3:
        print('\nPlease Provide Two File Paths To Compare')
        sys.exit()
    
    # get file name from args
    path_OLD = Path(sys.argv[1])
    path_NEW = Path(sys.argv[2])

    # get index col from data either csv or xlxs
    if sys.argv[2].lower().endswith('.csv'):
        df = pd.read_csv(path_NEW)
    elif sys.argv[2].lower().endswith('.xlsx'):
        df = pd.read_excel(path_NEW)
    else:
        print('\nThe File Paths Need The Extention .csv or .xlsx')
        sys.exit()

    if sys.argv[1].lower().endswith('.csv'):
        argv1 = '\n{} Loaded'.format(str(sys.argv[1]))
        argv2 = '\n{} Loaded'.format(str(sys.argv[2]))
        print(argv1, argv2)
    elif sys.argv[1].lower().endswith('.xlsx'):
        argv1 = '\n{} Loaded'.format(str(sys.argv[1]))
        argv2 = '\n{} Loaded'.format(str(sys.argv[2]))
        print(argv1, argv2)
    else:
        print('\nThe File Paths Need The Extention .csv or .xlsx')
        sys.exit()

    # set the index column as the first column.
    ## future: add argument to indicate index column if not in first row 
    index_col = df.columns[0]

    # print the index column name
    print('\nIndex column: {}'.format(index_col))

    # now we compare the files
    file_diff(path_OLD, path_NEW, index_col)

if __name__ == '__main__':
    main()
