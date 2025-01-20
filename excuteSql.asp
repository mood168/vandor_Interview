<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="2D34D3E4/db.asp"-->
<%
Dim rs    
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open "SELECT UserID, Username, Password FROM dbo.Users", conn, 1, 3
    
    Do While Not rs.EOF
        rs("Username") = SimpleEncrypt(rs("Username"))
        rs("Password") = SimpleEncrypt(rs("Password"))
        rs.Update
        rs.MoveNext
    Loop
    
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing

' 加密函數
Function SimpleEncrypt(inputText)
    Dim i, charCode
    Dim encryptedText
    encryptedText = ""
    
    For i = 1 To Len(inputText)
        charCode = AscW(Mid(inputText, i, 1))
        charCode = charCode + 3 ' 將字符碼增加1（可以根據需要調整）
        encryptedText = encryptedText & ChrW(charCode)
    Next
    
    SimpleEncrypt = encryptedText
End Function

' 簡單的字符替換解密函數
Function SimpleDecrypt(encryptedText)
    Dim i, charCode
    Dim decryptedText
    decryptedText = ""
    
    For i = 1 To Len(encryptedText)
        charCode = AscW(Mid(encryptedText, i, 1))
        charCode = charCode - 3 ' 將字符碼減少1（必須與加密時相反）
        decryptedText = decryptedText & ChrW(charCode)
    Next
    
    SimpleDecrypt = decryptedText
End Function

%>
