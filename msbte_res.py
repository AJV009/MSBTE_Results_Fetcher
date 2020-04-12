import pandas as pd
import fire
from colorama import Fore, init
from os.path import abspath
from timeit import default_timer as timer
import win32com.client as win32
# pip install fire pandas lxml openpyxl
def res_loop(sample="407517",
             from_seat="",
             to_seat="",
             export_file_name="res_data_out"):
    '''
    Description:
    Takes in a range of seat number then exports marks to a excel file!
    
    usage:
    fire-res --<flag_with_val>
    
    Flags:
    --sample="<seat_number>"          =  give a working seat number with the perfect format of results!
    --from_seat="<seat_number>"       =  From seat number (for looping through)
    --to_seat="<seat_number>"         =  To seat number. e.x. till 407518
    --export_file_name="<file_name>"  =  Only the file name, .xlsx will be added automatically!
    --help                            =  Display this help!
    
    e.x. python msbte_res.py --from_seat="407515" --to_seat="407554" --export_file_name="dataout"
    '''
    st_time = timer()
    init()
    f_no = str(sample)[0:2]
    df = pd.read_html(f"https://msbte.org.in/SHTFNL20BTERESLIVE/SHTFNL20BTERESLIVE/SeatNumber/{f_no}/{sample}Marksheet.html")
    df_sub = df[1][0]
    sub_name = [ df_sub[i] for i in range(2, len(df_sub)) if type( df_sub[i] ) == str ]
    COLS =  ['ENROLLMENT NO', 'SEAT NO'] + sub_name + ['PERCENTAGE', 'TOTAL MARKS', 'CREDITS']
    DATA = pd.DataFrame( columns=COLS )
    for i in range(from_seat, to_seat+1):
        url = f"https://msbte.org.in/SHTFNL20BTERESLIVE/SHTFNL20BTERESLIVE/SeatNumber/{f_no}/{i}Marksheet.html"
        try:
            df = pd.read_html(url)
            subs = [ i for i in range(2, len( df[1][0] )) if type( df[1][0][i] ) == str]
            r1 = [ df[0][1][1], df[0][5][1] ] + (( df[1].loc[subs][7] ).to_list()) + [ df[2][3][1], df[2][4][1], df[2][5][1] ]
            print(Fore.CYAN,df[0][1][0],r1)
            DATA.loc[len(DATA)] = r1
        except:
            print(Fore.RED,f"seat_no: {i} not found or couldnt fetch!")
            continue
    file_name = f"{export_file_name}.xlsx"
    DATA.to_excel(file_name)
    excel = win32.dynamic.Dispatch('Excel.Application')
    wb = excel.Workbooks.Open(abspath(file_name))
    ws = wb.Worksheets("Sheet1")
    ws.Columns.AutoFit()
    wb.Save()
    excel.Application.Quit()
    print(Fore.GREEN,f"Data collected and exported to {file_name} file!")
    print(Fore.BLUE,f"TOTAL EXECUTION TIME: {timer() - st_time}")
    
if __name__ == '__main__':
    fire.Fire(res_loop)