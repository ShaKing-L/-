<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/WM.asp" -->
<%
Dim MM_editAction
MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
Dim MM_abortEdit
MM_abortEdit = false
%>
<%
' IIf implementation
Function MM_IIf(condition, ifTrue, ifFalse)
  If condition = "" Then
    MM_IIf = ifFalse
  Else
    MM_IIf = ifTrue
  End If
End Function
%>
<%
If (CStr(Request("MM_update")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the update
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_WM_STRING
    MM_editCmd.CommandText = "UPDATE dbo.menu SET m_id = ?, m_name = ?, m_prince = ?, m_information = ?, u_id = ? WHERE m_id = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 201, 1, 5, Request.Form("textfield")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 201, 1, 20, Request.Form("textfield2")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 5, 1, -1, MM_IIF(Request.Form("textfield3"), Request.Form("textfield3"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 201, 1, 20, Request.Form("textfield4")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 202, 1, 10, Request.Form("textfield5")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 200, 1, 5, Request.Form("MM_recordId")) ' adVarChar
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "商家menu - .asp"
    If (Request.QueryString <> "") Then
      If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0) Then
        MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
      Else
        MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
      End If
    End If
    Response.Redirect(MM_editRedirectUrl)
  End If
End If
%>
<%
Dim Recordset1__MMColParam
Recordset1__MMColParam = "1"
If (Request.QueryString("xgid") <> "") Then 
  Recordset1__MMColParam = Request.QueryString("xgid")
End If
%>
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_WM_STRING
Recordset1_cmd.CommandText = "SELECT * FROM dbo.menu WHERE m_id = ?" 
Recordset1_cmd.Prepared = true
Recordset1_cmd.Parameters.Append Recordset1_cmd.CreateParameter("param1", 200, 1, 5, Recordset1__MMColParam) ' adVarChar

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>无标题文档</title>
</head>

<body>
<form id="form1" name="form1" method="POST" action="<%=MM_editAction%>">
  <p align="center">菜品号：
    <label for="textfield"></label>
    <input name="textfield" type="text" id="textfield" value="<%=(Recordset1.Fields.Item("m_id").Value)%>" />
  </p>
  <p align="center">菜品名：
    <label for="textfield2"></label>
    <input name="textfield2" type="text" id="textfield2" value="<%=(Recordset1.Fields.Item("m_name").Value)%>" />
  </p>
  <p align="center">价格： 
    <label for="textfield3"></label>
    <input name="textfield3" type="text" id="textfield3" value="<%=(Recordset1.Fields.Item("m_prince").Value)%>" />
  </p>
  <p align="center">备注：
    <label for="textfield4"></label>
    <input name="textfield4" type="text" id="textfield4" value="<%=(Recordset1.Fields.Item("m_information").Value)%>" />
  </p>
  <p align="center">商家号：
    <label for="textfield5"></label>
    <input name="textfield5" type="text" id="textfield5" value="<%=(Recordset1.Fields.Item("u_id").Value)%>" />
  </p>
  <p align="center">
    <input type="submit" name="button" id="button" value="修改" />
  <a href="商家menu - .asp">退出修改</a></p>
  <div align="center">
    <input type="hidden" name="MM_update" value="form1" />
    <input type="hidden" name="MM_recordId" value="<%= Recordset1.Fields.Item("m_id").Value %>" />
  </div>
</form>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
