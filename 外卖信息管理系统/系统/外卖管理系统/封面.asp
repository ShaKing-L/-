<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/WM.asp" -->
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_WM_STRING
Recordset1_cmd.CommandText = "SELECT * FROM dbo.customer" 
Recordset1_cmd.Prepared = true

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
<table width="500" border="1" align="center">
  <tr align="center">
    <td>外卖管理系统</td>
  </tr>
</table>
<table width="300" border="1" align="center">
  <tr align="center">
    <td>选择你的身份</td>
  </tr>
  <tr align="center">
    <td><a href="用户login.asp">客户</a></td>
  </tr>
  <tr align="center">
    <td><a href="c_login.asp">商家</a></td>
  </tr>
  <tr>
    <td align="center"><a href="送餐员登录.asp">送餐员</a></td>
  </tr>
</table>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
