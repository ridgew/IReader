<!--#include file="Const.asp"-->
<!--#include file="iBCore.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>iReader Menu</title>
<link rel="stylesheet" href="Reader.css" type="text/css" />
<style type="text/css">
<!--
.style1 {font-family: "Courier New", Courier, mono;font-style: italic;}
.style2 {font-family: "Courier New", Courier, mono;}
.sperator {border-left:1px solid #000000;}
-->
</style>
</head>
<body bgcolor="#DEE3F7">
<%

Dim ISBN
ISBN = Request.QueryString("ISBN")

If IsPostBack() Then
	Response.Cookies("charset") = Request.Form("charset")
End If

%>
<table width="100%" height="25"  border="0" cellpadding="2" cellspacing="0">
  <tr>
  	<td width="30" valign="middle" bgcolor="#FFFFFF" style="border-right:1px solid black;"><a href="http://www.vbyte.com" target="_top"><img src="images/ib.gif" alt="&#105;&#223; Networks" width="24" height="16" border="0" /></a></td>
    <td width="100" valign="middle" nowrap><img src="images/showtoc.gif" width="16" height="16" border="0" align="absmiddle" /> <a href="TOC.asp?ISBN=<%=ISBN%>" target="tocWin">装载本书目录</a>
    </td>
    <td width="100" align="center" valign="middle" nowrap class="sperator"><img src="images/download.gif" width="16" height="16" border="0" align="absmiddle" /> <a href="download.asp?ISBN=<%=ISBN%>" target="mainWin">下载本书</a></td>
    <td width="100" align="center" valign="middle" nowrap class="sperator"><img src="images/icon/37.gif" border="0" align="absmiddle" /> <a href="Vote.asp" target="mainWin">投票列表</a></td>
    <td width="100" align="center" valign="middle" nowrap class="sperator"><img src="images/icon/9.gif" border="0" align="absmiddle" /> <a href="TOC.asp" target="tocWin">使用指南</a></td>
    <td width="130" align="center" valign="middle" nowrap class="sperator"><img src="images/icon/13.gif" border="0" align="absmiddle" /> <a href="default.asp" target="_top">返回iReader首页</a></td>
    <td width="120" align="center" valign="middle" nowrap class="sperator"><img src="images/icon/27.gif" border="0" align="absmiddle" /> <a href="/my" target="_blank">iB会员中心</a></td>
    <td width="120" align="center" valign="middle" nowrap class="sperator"> <form method="post" style="margin:0px;padding:0px" name="frmCharset"><select name="charset" onchange="javascript:this.form.submit();"><option value="iso-8859-1"<%=getValue(Request.Cookies("charset") <> "gb2312"," selected","")%>>编码ISO-8859-1</option>
	<option value="gb2312"<%=getValue(Request.Cookies("charset") = "gb2312"," selected","")%>>编码GB2312</option>
	</select></form></td>
    <td align="right" valign="middle" nowrap>&copy; <a href="/" target="_blank">虚数传播网络</a> Since 2003.</td>
  </tr>
  <tr><td colspan="8" bgcolor="#6699CC"></td></tr>
</table>
</body>
</html>