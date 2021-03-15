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
Recordset1_cmd.CommandText = "SELECT * FROM dbo.o_c_s_d_s_m WHERE s_id = ?" 
Recordset1_cmd.Prepared = true
Recordset1_cmd.Parameters.Append Recordset1_cmd.CreateParameter("param1", 200, 1, 20, Recordset1__MMColParam) ' adVarChar

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
Recordset1_numRows = Recordset1_numRows + Repeat1__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>无标题文档</title>
<script type="text/javascript">
function MM_jumpMenuGo(objId,targ,restore){ //v9.0
  var selObj = null;  with (document) { 
  if (getElementById) selObj = getElementById(objId);
  if (selObj) eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0; }
}
</script>
</head>

<body>
<table width="600" border="1" align="center">
  <tr align="center">
    <td>员工号</td>
    <td>员工名</td>
    <td>订单号</td>
    <td>菜名</td>
    <td>数量</td>
    <td>客户名</td>
    <td>客户地址</td>
    <td>客户电话</td>
  </tr>
  <% 
While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) 
%>
  <tr align="center">
    <td><%=(Recordset1.Fields.Item("s_id").Value)%></td>
    <td><%=(Recordset1.Fields.Item("s_name").Value)%></td>
    <td><%=(Recordset1.Fields.Item("o_id").Value)%></td>
    <td><%=(Recordset1.Fields.Item("m_name").Value)%></td>
    <td><%=(Recordset1.Fields.Item("number").Value)%></td>
    <td><%=(Recordset1.Fields.Item("c_name").Value)%></td>
    <td><%=(Recordset1.Fields.Item("c_address").Value)%></td>
    <td><%=(Recordset1.Fields.Item("c_tele").Value)%></td>
  </tr>
  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset1.MoveNext()
Wend
%>
</table>
<form id="form1" name="form1" method="post" action="搜索订单.asp">
  <label for="word"></label>
  <div align="center">
    <p>继续搜索：
      <input type="text" name="word" id="word" />
      <input type="submit" name="button" id="button" value="搜索" />
    </p>
    <p>
      <select name="jumpMenu" id="jumpMenu">
        <option value="送餐可看.asp">退出搜索</option>
        <option value="送餐员登录.asp">返回登录界面</option>
        <option value="封面.asp">返回系统主页</option>
      </select>
      <input type="button" name="go_button" id= "go_button" value="前往" onclick="MM_jumpMenuGo('jumpMenu','parent',0)" />
    </p>
  </div>
</form>
<p>&nbsp;</p>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
