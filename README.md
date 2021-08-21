# S3 file download with Python and access with VBA

This simple project is using the following stacks:

- Python
- AWS S3
- VBA/Excel
- A Bitcoin API

With this stacks, we will use the Python script to download a JSON file from
S3 with boto3, and save it locally.

The Excel with VBA will call this Python script to perform this download,
and after this the VBA will read the local JSON file.

And finally the VBA will call an external API to get Bitcoin current bid,
the API is from https://bitpreco.com/api

To parse JSON on VBA we will be using a library called VBA-JSON from
https://github.com/VBA-tools/VBA-JSON

## Architecture

![./images/s3-file-python-vba.png](./images/s3-file-python-vba.png)

## VBA script to call python and read the saved JSON file locally

```vba
Public Sub callPythonScript()
    Dim objShell As Object
    Dim PythonExe$, PythonScript$, fileLocalPath$, fileContent$
    
    Set objShell = VBA.CreateObject("Wscript.Shell")

    'which python
    PythonExe = """C:\Users\USERNAME\AppData\Local\Programs\Python\Python37-32\python.exe"""
    PythonScript = """C:\Users\USERNAME\Documents\s3-file-python-vba\main.py"""

    'call the python script
    objShell.Run PythonExe & PythonScript, 2, True
    
    fileLocalPath = "C:\Users\USERNAME\Documents\s3-file-python-vba\file.json"
    fileContent = READ_FILE_CONTENT(fileLocalPath)

    Dim Json As Object
    Set Json = JsonConverter.ParseJson(fileContent)

    'delete the local file
    Kill fileLocalPath

    MsgBox Json("bitcoin")
End Sub

Public Function READ_FILE_CONTENT(ByVal fileName$) As String
    Dim textline$, text$
    
    Open fileName For Input As #1
    Do Until EOF(1)
        Line Input #1, textline
        text = text & textline
    Loop
    Close #1
    
    READ_FILE_CONTENT = text
End Function
```
