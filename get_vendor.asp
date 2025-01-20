<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="2D34D3E4/db.asp"-->
<%
Response.CharSet = "utf-8"
Response.ContentType = "application/json"

Function JsonEscape(str)
    If IsNull(str) Or IsEmpty(str) Then
        JsonEscape = ""
    Else
        JsonEscape = Replace(Replace(Replace(str, "\", "\\"), """", "\"""), vbCrLf, "\n")
    End If
End Function

Function SafeField(field)
    If IsNull(field) Then
        SafeField = ""
    Else
        SafeField = field & ""
    End If
End Function

' 檢查登入狀態
If Session("UserID") = "" Then
    Response.Write "{""success"": false, ""message"": ""請先登入系統""}"
    Response.End
End If

' 取得電商ID
vendorId = Request.QueryString("id")
If vendorId = "" Then
    Response.Write "{""success"": false, ""message"": ""未提供電商ID""}"
    Response.End
End If

On Error Resume Next

' SQL 查詢電商資料
sql = "SELECT VendorID, ParentCode, ChildCode, UniformNumber, VendorName, " & _
      "ContactPerson, ISNULL(LogisticsContact, '') AS LogisticsContact, " & _
      "ISNULL(MarketingContact, '') AS MarketingContact, " & _
      "ISNULL(Phone, '') AS Phone, ISNULL(Address, '') AS Address, " & _
      "ISNULL(Email, '') AS Email, ISNULL(Website, '') AS Website " & _
      "FROM Vendors WHERE VendorID = " & vendorId & " AND IsActive = 1"

Set rs = conn.Execute(sql)

If Err.Number <> 0 Then
    Response.Write "{""success"": false, ""message"": ""資料庫查詢錯誤""}"
    Response.End
End If

If rs.EOF Then
    Response.Write "{""success"": false, ""message"": ""找不到電商資料""}"
Else
    ' 組織 JSON 回應
    Response.Write "{"
    Response.Write """success"": true,"
    Response.Write """VendorID"": " & rs("VendorID") & ","
    Response.Write """ParentCode"": """ & JsonEscape(rs("ParentCode")) & ""","
    Response.Write """ChildCode"": """ & JsonEscape(rs("ChildCode")) & ""","
    Response.Write """UniformNumber"": """ & JsonEscape(rs("UniformNumber")) & ""","
    Response.Write """VendorName"": """ & JsonEscape(rs("VendorName")) & ""","
    Response.Write """ContactPerson"": """ & JsonEscape(rs("ContactPerson")) & ""","
    Response.Write """LogisticsContact"": """ & JsonEscape(SafeField(rs("LogisticsContact"))) & ""","
    Response.Write """MarketingContact"": """ & JsonEscape(SafeField(rs("MarketingContact"))) & ""","
    Response.Write """Phone"": """ & JsonEscape(SafeField(rs("Phone"))) & ""","
    Response.Write """Address"": """ & JsonEscape(SafeField(rs("Address"))) & ""","
    Response.Write """Email"": """ & JsonEscape(SafeField(rs("Email"))) & ""","
    Response.Write """Website"": """ & JsonEscape(SafeField(rs("Website"))) & """"
    Response.Write "}"
End If

On Error Goto 0

rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
%> 