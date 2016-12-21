<!--#include file="Const.asp"-->
<!--#include file="iBCore.asp"-->
<%
Dim Rs,Conn,sql,ISBN,IO
Dim FilePath
ISBN = Request.QueryString("ISBN")
If Len(ISBN)<>10 Then
	ShowMsgPage("default.asp")
End If

'If IsPostBack() And (Request.Form("Key") <> "" And Session("iBookDownKey") = Request.Form("Key")) Then
'2007年12月10日 取消POST下载限制
'------------------------------
If Request("Key") <> "" And (Request("key")="greengate" Or Session("iBookDownKey") = Request("Key")) Then
	Call DBOpen(Conn,ConnectionString)
	sql = "select top 1 BookPath,DownHits from [iReaderBooks] where ISBN='"&CheckStr(ISBN)&"' and STOK=True"
	Set Rs = Server.CreateObject("ADODB.Recordset")
	Rs.Open sql,Conn,3,3
	If Not Rs.Eof Then
			FilePath = GetValue(InStr(1,Rs("BookPath"),":",1)>0,Rs("BookPath"),bookRootPath & Rs("BookPath"))
			Rs("DownHits") = Rs("DownHits") + 1
			Rs.Update()
		Else
			Response.Write("该文件已经不存在！")
	End If
	Rs.Close()
	Set Rs = Nothing
	Call DbClose(Conn)

	Set IO = New iBFileIO
		If IO.FileExists(FilePath) Then
			Session("iBookDownKey") = ""
			IO.WriteFile FilePath,ISBN & Right(FilePath,4)
		Else
			Response.Write("该文件已经不存在！")
		End If
	Set IO= Nothing
Else
	Call ShowDownKey(ISBN)
End If


'***********************************
Sub ShowDownKey(ISBN)
	Session("iBookDownKey") = CreateWindowsGUID()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>下载ISBN: <%=ISBN%>的书籍</title>
<link rel="stylesheet" href="Reader.css" type="text/css" />
<style type="text/css">
<!--
.style1 {font-family: "Courier New", Courier, mono;font-style: italic;color:#666666;font-size:14px;line-height:1.8;}
.style2 {font-family: "Courier New", Courier, mono;color:#666666;font-size:14px;line-height:1.8;}
.sperator {border-left:1px solid #000000;}
-->
</style>
</head>
<body>
<p>&nbsp;</p>
<form method="post" enctype="application/x-www-form-urlencoded">
<table width="600" border="0" align="center" cellpadding="8" cellspacing="0">
  <tr>
    <td align="right" class="style2">随机Key (<span class="style1">Random Key</span>)：</td>
    <td class="style2"><span style="background-color:#f3f3f3;padding:5px;"><%=Session("iBookDownKey")%></span></td>
  </tr>
  <tr>
    <td align="right" class="style2">确认Key (<span class="style1">Confirm Key</span>)：</td>
    <td class="style2"><input name="key" type="text" id="key" value="<%=Request.Form("key")%>" size="40" maxlength="40" /></td>
  </tr>
  <tr>
    <td colspan="2" align="center"><input name="btnDown" type="submit" id="btnDown" value="下 载(Download)" />
	<%
	 If IsPostBack() Then
		 If Request.Form("key") <> Session("iBookDownKey") Then
			Response.Write("<font color=red>* 请输入正确的Key</font>")
		 End If
	 End If
	%><br><br>
	<a href="Download.asp?ISBN=<%=ISBN%>&key=greengate" target="_blank"><font color="green"><strong>绿色通道：点击断点续传快速下载</strong></font></a>
	</td>
  </tr>
</table>
</form>
</body>
</html>
<%
End Sub
%>