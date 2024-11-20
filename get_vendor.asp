<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="2D34D3E4/db.asp"-->
<%
Response.Clear
Response.CharSet = "utf-8"
Response.ContentType = "application/json"

' 檢查登入狀態
If Session("UserID") = "" Then
    Response.Write "{""success"":false,""message"":""未登入""}"
    Response.End
End If

' 取得廠商ID
Dim vendorId
vendorId = Request.QueryString("id")

If vendorId = "" Then
    Response.Write "{""success"":false,""message"":""未提供廠商ID""}"
    Response.End
End If

On Error Resume Next

' 取得廠商資料
Dim sql
sql = "SELECT * FROM Vendors WHERE VendorID = " & CLng(vendorId) & " AND IsActive = 1"

Dim rs
Set rs = conn.Execute(sql)

If Err.Number <> 0 Then
    Response.Write "{""success"":false,""message"":""資料庫錯誤: " & Server.HTMLEncode(Err.Description) & """}"
    Response.End
End If

If Not rs.EOF Then
    ' 組織 JSON 回應
    Response.Write "{""success"":true," & _
                   """VendorID"":" & rs("VendorID") & "," & _
                   """ParentCode"":""" & Trim(rs("ParentCode")) & """," & _
                   """ChildCode"":""" & Trim(rs("ChildCode")) & """," & _
                   """UniformNumber"":""" & Trim(rs("UniformNumber")) & """," & _
                   """VendorName"":""" & Replace(rs("VendorName"), """", "\""") & """," & _
                   """ContactPerson"":""" & Replace(rs("ContactPerson"), """", "\""") & """," & _
                   """Phone"":""" & Replace(rs("Phone") & "", """", "\""") & """," & _
                   """Address"":""" & Replace(rs("Address") & "", """", "\""") & """," & _
                   """Email"":""" & Replace(rs("Email") & "", """", "\""") & """," & _
                   """Website"":""" & Replace(rs("Website") & "", """", "\""") & """}"
Else
    Response.Write "{""success"":false,""message"":""找不到廠商資料""}"
End If

' 清理資源
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
%> 