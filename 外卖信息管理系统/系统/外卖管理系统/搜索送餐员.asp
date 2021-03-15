<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/WM.asp" -->
<%
Dim Recordset1__MMColParam
Recordset1__MMColParam = "1"
If (Request.Form("word") <> "") Then 
  Recordset1__MMColParam = Request.Form("word")
End If
%>
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_WM_STRING
Recordset1_cmd.CommandText = "SELECT * FROM dbo.staff WHERE s_name = ?" 
Recordset1_cmd.Prepared = true
Recordset1_cmd.Parameters.Append Recordset1_cmd.CreateParameter("param1", 200, 1, 20, Recordset1__MMColParam) ' adVarChar

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>无标题文档</title>
<script type="text/javascript">
function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
function MM_jumpMenuGo(objId,targ,restore){ //v9.0
  var selObj = null;  with (document) { 
  if (getElementById) selObj = getElementById(objId);
  if (selObj) eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0; }
}
</script>
</head>

<body>
<table width="300" border="1" align="center">
  <tr align="center">
    <td>送餐员号</td>
    <td>送餐员名</td>
    <td>送餐员电话</td>
  </tr>
  <tr align="center">
    <td><%=(Recordset1.Fields.Item("s_id").Value)%></td>
    <td><%=(Recordset1.Fields.Item("s_name").Value)%></td>
    <td><%=(Recordset1.Fields.Item("s_tele").Value)%></td>
  </tr>
</table>
<form id="word" name="word" method="post" action="搜索送餐员.asp">
  <div align="center"></div>
  <label for="word"></label>
  <div align="center">继续搜索：
    <input type="text" name="word" id="word" />
    <input type="submit" name="button" id="button" value="搜索" />
  </div>
</form>
<form id="form2" name="form2" method="post" action="">
  <div align="center">
    <select name="jumpMenu" id="jumpMenu" onchange="MM_jumpMenu('parent',this,0)">
      <option value="商家界面.asp">返回商家界面</option>
      <option value="c_login.asp">返回登录</option>
      <option value="封面.asp">返回系统首页</option>
      <option value="送餐员.asp">退出搜索</option>
    </select>
    <input type="button" name="go_button" id= "go_button" value="前往" onclick="MM_jumpMenuGo('jumpMenu','parent',0)" />
  </div>
</form>
<p>&nbsp;</p>
<p>&nbsp;</p>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
