<!--#include file="Const.asp"-->
<!--#include file="iBCore.asp"-->
<% Response.Buffer = False %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-frameset.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>iReader eBook Vote</title>
<link rel="stylesheet" href="Reader.css" type="text/css" />
<script language="javascript" type="text/javascript" src="Reader.js"></script>
<base target="mainWin">
</head>
<body>
<%
Dim iBDb,HelpHHC
'Set iBDb = new iBDataBase
		'iBDb.ConnString = ConnectionString
		'HelpHHC = "<UL>" & BuildHelpHHC(iBDb,0) & "</UL>"
'Set iBDb = Nothing

	'Response.Write HelpHHC
%>
</body>
</html>