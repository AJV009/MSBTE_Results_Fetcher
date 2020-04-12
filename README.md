# MSBTE_Result_Fetcher
A python script to fetch MSBTE polytechnic results and store in a excel file. If the script doesnt work 
please open a issue. As of now only seat numberare supported. The script will be improved after next SEM end exam.

## Dependencies Installation
```bash
$ pip install --upgrade 'setuptools<45.0.0'
$ pip install colorama pandas pyinstaller pytest-timeit pywin32 lxml openpyxl
```

## Usage
```bash
$ git clone https://github.com/AJV009/MSBTE_Results_Fetcher.git
$ cd MSBTE_Results_Fetcher
$ python msbte_res.py --from_seat="407517" --to_seat="407520" --export_file_name="dataout1"
```
### Try --help flag to find more info
```bash
$ python msbte_res.py --help
INFO: Showing help with the command 'msbte_res.py -- --help'.
NAME
    msbte_res.py - Description: Takes in a range of seat number then exports marks to a excel file!
SYNOPSIS
    msbte_res.py <flags>
DESCRIPTION
    usage:
    fire-res --<flag_with_val>
    
    Flags:
    --sample="<seat_number>"          =  give a working seat number with the perfect format of results!
    --from_seat="<seat_number>"       =  From seat number (for looping through)
    --to_seat="<seat_number>"         =  To seat number. e.x. till 407518
    --export_file_name="<file_name>"  =  Only the file name, .xlsx will be added automatically!
    --help                            =  Display this help!

    e.x. python msbte_res.py --from_seat="407520" --to_seat="407525" --export_file_name="dataout"
FLAGS
    --sample=SAMPLE
    --from_seat=FROM_SEAT
    --to_seat=TO_SEAT
    --export_file_name=EXPORT_FILE_NAME
```

## Using pyinstaller to compile to a win32 exe file

- Packed (32mb, SLOW)
```bash
$ pyinstaller --hidden-import='pkg_resources.py2_warn' --onefile msbte_res.py
```

- Unpacked (82mb, FAST)
```bash
$ pyinstaller --hidden-import='pkg_resources.py2_warn' msbte_res.py
```
Linux users can better create a env and run the script in it!

## NOTE! 
As of now its only for MSBTE! AND the url is static! (I meant I have placed it in the code, so if you wish to fetch the same type of results but with different URL and stuff, please youll have to make changes in the code. the URL var!)

### You can contribute by opening issues!
