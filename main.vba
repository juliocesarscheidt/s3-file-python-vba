Public Sub execute()
    Call callPythonScript
    Call callReadFile
    Call callAPI
End Sub

Public Sub callPythonScript()
    Dim objShell As Object
    Dim PythonExe$, PythonScript$
    
    On Error GoTo ErrorHandler
    
    Set objShell = VBA.CreateObject("Wscript.Shell")

    'which python
    PythonExe = """C:\Users\USERNAME\AppData\Local\Programs\Python\Python37-32\python.exe"""
    PythonScript = """C:\Users\USERNAME\Documents\s3-file-python-vba\main.py"""

    'executa o script python
    objShell.Run PythonExe & PythonScript, 2, True

ErrorHandler:
    If Err.Number <> 0 Then
        Msg = "[ERRO] Número: " & Str(Err.Number) & " foi gerado por " & Err.Source & Chr(13) & Chr(13) & Err.Description
        MsgBox Msg, vbMsgBoxHelpButton, "Error", Err.HelpFile, Err.HelpContext
    End If
End Sub

Public Sub callReadFile()
    Dim fileLocalPath$, fileContent$

    fileLocalPath = "C:\Users\USERNAME\Documents\s3-file-python-vba\file.json"
    fileContent = READ_FILE_CONTENT(fileLocalPath)
    'MsgBox fileContent

    Dim Json As Object
    Set Json = JsonConverter.ParseJson(fileContent)

    'deleta o arquivo local
    Kill fileLocalPath

    MsgBox Json("last")
    Range("S9") = CCur(Replace(Json("last"), ".", ","))
End Sub

Public Sub callAPI()
    Dim apiUrl$, result$
    apiUrl = "https://api.bitpreco.com/btc-brl/ticker"

    'chama a API
    result = REQUEST(apiUrl, "GET")
    'MsgBox result

    Dim Json As Object
    Set Json = JsonConverter.ParseJson(result)

    MsgBox Json("last")
    Range("S15") = CCur(Replace(Json("last"), ".", ","))
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

Public Function REQUEST(ByVal apiUrl$, ByVal method$, Optional ByVal jsonDataString$, Optional ByVal bearerToken$) As String
    Dim objHTTP As Object
    Dim responseCode$, responseText$

    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")

    objHTTP.Open method, apiUrl, False
    objHTTP.setRequestHeader "Content-type", "application/json"

    'setting oauth token when provided
    If bearerToken <> "" Then
        objHTTP.setRequestHeader "Authorization", "Bearer " & bearerToken
    End If

    'setting payload when provided
    If jsonDataString <> "" Then
        objHTTP.Send (jsonDataString)
    Else
        objHTTP.Send
    End If

    responseCode = objHTTP.Status

    If responseCode >= 200 And responseCode <= 299 Then
        responseText = objHTTP.responseText

        responseText = Replace(responseText, Chr(10), "")
        responseText = Replace(responseText, Chr(13), "")
    End If

    Set objHTTP = Nothing

    REQUEST = responseText
End Function
