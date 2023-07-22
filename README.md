# name
- PythonExeVbac

## Overview
- Python program to run vbac.

## Requirement
- Python v3.11.0
- Download vbac.wsf and place it in the same folder as PythonExeVbac.exe.
  - [vbac.wf](https://github.com/vbaidiot/ariawase)
- Check the following items in the xlsm file.
  - File -> Options ->
  - Transcenter -> Transcenter Settings ->
  - Macro Settings -> VBA Project Trust access to object model

## Method of distributable format(EXE)
- Execute the following command.
```
$ pyinstaller .\PythonExeVbac.py --onefile --clean
```

## Usage
- Execute the following command.
```
$ PythonExeVbac.exe decombine [full path of xlsm file]

or 

$ PythonExeVbac.exe combine [full path of xlsm file]
```

## Features
- decombine
  - Extract macro code from xlsm file and output bas file in src folder.
  - bin folder is generated each time and the specified xlsm file is copied.
- combine
  - Generate xlsm file from bas file.
  - 

## Reference
- [vbac.wf](https://github.com/vbaidiot/ariawase)

## Author
- tkyasu999

## Licence
[MIT](https://github.com/tkyasu999/PythonExeVbac/blob/main/LICENSE)

~~~
Copyright (c) 2011 igeta

Permission is hereby granted, free of charge, to any person obtaining a
copy of this software and associated documentation files (the "Software"),
to deal in the Software without restriction, including without limitation
the rights to use, copy, modify, merge, publish, distribute, sublicense,
and/or sell copies of the Software, and to permit persons to whom the
Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included
in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS
OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL
THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,
ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE
OR OTHER DEALINGS IN THE SOFTWARE.
~~
