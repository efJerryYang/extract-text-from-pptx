# Extract text from pptx

A simple tool to extract plain text from pptx file

You may rename it to `pptx2txt` or any other name you prefer.

Usage: 
```sh
python extract_text_from_pptx.py <pptx_file_path>
```
or
```sh
# remember to change the "#!/usr/bin/python3" in the source code to the path of your python interpreter
./extract_text_from_pptx.py <pptx_file_path>
```

Requirements:

```txt
lxml==4.9.2
olefile==0.46
Pillow==9.5.0
python-pptx==0.6.21
XlsxWriter==3.1.0
```
You can run the following to install the required dependencies:

```sh
pip install -r requirements.txt
```
