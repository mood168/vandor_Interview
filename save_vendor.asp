<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="2D34D3E4/db.asp"-->
<%
Response.Clear
Response.CharSet = "utf-8"
Response.ContentType = "application/json"

' 錯誤處理函數
Function HandleError(message)
    Response.Clear
    Response.Write "{""success"": false, ""message"": """ & Replace(message, """", "\""") & """}"
    Response.End
End Function

On Error Resume Next

' 取得表單資料
Dim vendorId, parentCode, childCode, uniformNumber, vendorName, contactPerson, phone, address, email, website

' 檢查是新增還是編輯表單
vendorId = Request.Form("vendorId")
parentCode = Request.Form("parentCode")
childCode = Request.Form("childCode")
uniformNumber = Request.Form("uniformNumber")
vendorName = Request.Form("vendorName")
contactPerson = Request.Form("contactPerson")
phone = Request.Form("phone")
address = Request.Form("address")
email = Request.Form("email")
website = Request.Form("website")

' 基本驗證
If parentCode = "" Then HandleError("母代號不能為空")
If childCode = "" Then HandleError("子代號不能為空")
If uniformNumber = "" Then HandleError("統一編號不能為空")
If vendorName = "" Then HandleError("廠商名稱不能為空")
If contactPerson = "" Then HandleError("聯絡人不能為空")

' SQL 注入防護
Function SafeSQL(str)
    If IsNull(str) Or str = "" Then
        SafeSQL = "NULL"
    Else
        SafeSQL = "'" & Replace(str, "'", "''") & "'"
    End If
End Function

Dim sql
If vendorId = "" Then
    ' 新增廠商
    sql = "INSERT INTO Vendors (ParentCode, ChildCode, UniformNumber, VendorName, ContactPerson, " & _
          "Phone, Address, Email, Website, IsActive, CreatedDate) VALUES (" & _
          SafeSQL(parentCode) & ", " & _
          SafeSQL(childCode) & ", " & _
          SafeSQL(uniformNumber) & ", " & _
          SafeSQL(vendorName) & ", " & _
          SafeSQL(contactPerson) & ", " & _
          SafeSQL(phone) & ", " & _
          SafeSQL(address) & ", " & _
          SafeSQL(email) & ", " & _
          SafeSQL(website) & ", " & _
          "1, GETDATE())"
Else
    ' 更新廠商
    sql = "UPDATE Vendors SET " & _
          "ParentCode = " & SafeSQL(parentCode) & ", " & _
          "ChildCode = " & SafeSQL(childCode) & ", " & _
          "UniformNumber = " & SafeSQL(uniformNumber) & ", " & _
          "VendorName = " & SafeSQL(vendorName) & ", " & _
          "ContactPerson = " & SafeSQL(contactPerson) & ", " & _
          "Phone = " & SafeSQL(phone) & ", " & _
          "Address = " & SafeSQL(address) & ", " & _
          "Email = " & SafeSQL(email) & ", " & _
          "Website = " & SafeSQL(website) & ", " & _
          "ModifiedDate = GETDATE() " & _
          "WHERE VendorID = " & CLng(vendorId)
End If

' 執行 SQL
conn.Execute sql

If Err.Number <> 0 Then
    HandleError("資料庫錯誤: " & Err.Description)
End If

' 成功回應
Response.Write "{""success"": true, ""message"": ""資料儲存成功""}"

Response.Redirect "vendors_management.asp"
conn.Close
Set conn = Nothing
%> 