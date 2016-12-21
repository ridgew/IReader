<!--#include file="Const.asp"-->
<!--#include file="iBCore.asp"-->
<%
Dim Rs,Conn,sql,strMsg,Id
Dim title,idx,Content
Id = Request.QueryString("id")
If Not IsNumber(Id) Then
	ShowMsgPage("default.asp")
End If

Call DBOpen(Conn,ConnectionString)
sql = "select top 1 * from [iReaderHelp] where hId="&Id
Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.Open sql,Conn,1,1
If Not Rs.Eof Then
	title = Rs("Title")
	idx = Rs("iConIndex")
	Content = Rs("Content")
	Else
		title = "N/A"
		idx = 9
		Content = "没有找到相关帮助。"
End If
Rs.Close()
Set Rs = Nothing

Call DbClose(Conn)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=title%> -- iReader Help</title>
<link rel="stylesheet" href="Reader.css" type="text/css" />
<style type="text/css">
<!--
.style1 {font-family: "Courier New", Courier, mono;font-style: italic;}
.style2 {font-family: "Courier New", Courier, mono;}
.sperator {border-left:1px solid #000000;}
p {color:#666666;font-size:14px;}
-->
</style>
</head>
<body>
<p>&nbsp;</p>
<p align="center">
<img src="images/icon/<%=idx%>.GIF" name="icon" border="0" id="icon" /> <%=title%>
<hr noshade size="1" align="center" width="80%" />
</p>

<div style="padding-left:30px;padding-right:30px;padding-top:10px;"><p align="left"><%=Content%></p></div>

</body>
</html>