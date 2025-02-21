<%@LANGUAGE="VBSCRIPT" codepage="65001"%>
<!--#include file="upload.lib.asp"-->
<%
Response.ContentType = "application/json"
Response.CharSet = "utf-8"

' 檢查 CSRF Token
If Request.Form("csrf_token") <> Session("CSRF_Token") Then
    Response.Write "{""success"": false, ""error"": ""無效的請求""}"
    Response.End
End If

' 檢查檔案類型函數
Function IsValidImageType(contentType)
    IsValidImageType = (contentType = "image/jpeg" Or contentType = "image/png" Or contentType = "image/gif")
End Function

' 生成安全的檔案名
Function GenerateSecureFileName(extension)
    Dim secureFileName
    secureFileName = Year(Now) & Month(Now) & Day(Now) & Hour(Now) & Minute(Now) & Second(Now) & Int(Rnd * 1000)
    secureFileName = Replace(secureFileName, " ", "")
    secureFileName = Replace(secureFileName, ".", "")
    secureFileName = Replace(secureFileName, "/", "")
    secureFileName = Replace(secureFileName, "\", "")
    secureFileName = Replace(secureFileName, ":", "")
    secureFileName = Replace(secureFileName, "*", "")
    secureFileName = Replace(secureFileName, "?", "")
    secureFileName = Replace(secureFileName, """", "")
    secureFileName = Replace(secureFileName, "<", "")
    secureFileName = Replace(secureFileName, ">", "")
    secureFileName = Replace(secureFileName, "|", "")
    GenerateSecureFileName = secureFileName & extension
End Function

' 檢查檔案大小
Const MaxFileSize = 10485760 ' 10MB

On Error Resume Next

Dim uploadForm
Set uploadForm = New ASPForm
uploadForm.SizeLimit = MaxFileSize

If uploadForm.State = xfsCompleted Then
    Dim file
    Set file = uploadForm.Files("file")
    
    If Not file Is Nothing Then
        ' 檢查檔案類型
        If Not IsValidImageType(file.ContentType) Then
            Response.Write "{""success"": false, ""error"": ""只允許上傳圖片檔案(.jpg, .jpeg, .png, .gif)""}"
            Response.End
        End If
        
        ' 檢查檔案大小
        If file.Length > MaxFileSize Then
            Response.Write "{""success"": false, ""error"": ""檔案大小不能超過 10MB""}"
            Response.End
        End If
        
        ' 生成安全的檔案名
        Dim fileExtension
        fileExtension = LCase(Mid(file.FileName, InStrRev(file.FileName, ".")))
        
        If fileExtension <> ".jpg" And fileExtension <> ".jpeg" And fileExtension <> ".png" And fileExtension <> ".gif" Then
            Response.Write "{""success"": false, ""error"": ""不支援的檔案類型""}"
            Response.End
        End If
        
        Dim newFileName
        newFileName = GenerateSecureFileName(fileExtension)
        
        ' 確保目標目錄存在且有寫入權限
        Dim fs, uploadPath
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        uploadPath = Server.MapPath("images/")
        
        If Not fs.FolderExists(uploadPath) Then
            fs.CreateFolder(uploadPath)
        End If
        
        ' 保存檔案
        On Error Resume Next
        file.SaveAs Server.MapPath("images/" & newFileName)
        
        If Err.Number <> 0 Then
            Response.Write "{""success"": false, ""error"": ""儲存檔案時發生錯誤""}"
            Response.End
        End If
        
        ' 返回成功訊息
        Response.Write "{""success"": true, ""fileName"": """ & newFileName & """}"
    Else
        Response.Write "{""success"": false, ""error"": ""未收到檔案""}"
    End If
Else
    Response.Write "{""success"": false, ""error"": ""上傳失敗: " & uploadForm.State & """}"
End If

Set uploadForm = Nothing
%> 