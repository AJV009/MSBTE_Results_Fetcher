# MSBTE_Result_Formater
A python script to fetch MSBTE polytechnic results and store in a excel file. If the script doesnt work 
please open a issue. As of now only seat numberare supported. The script will be improved after next SEM end exam.

## Installation
```bash
$ pip install --upgrade 'setuptools<45.0.0'
$ pip install colorama pandas pyinstaller pytest-timeit pywin32 lxml openpyxl
```

## Usage
```bash
$ python msbte_res.py --from_seat="407517" --to_seat="407520" --export_file_name="dataout1"
```

## Using pyinstaller to compile

- Packed (32mb, SLOW)
```bash
$ pyinstaller --hidden-import='pkg_resources.py2_warn' --onefile msbte_res.py
```

- Unpacked (82mb, FAST)
```bash
$ pyinstaller --hidden-import='pkg_resources.py2_warn' msbte_res.py
```

### You can contribute by opening issues!
