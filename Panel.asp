<!--#include file="Const.asp"-->
<!--#include file="iBCore.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<META NAME="Robots" CONTENT="noindex" />
<title>书籍控制面板</title>
<link rel="stylesheet" href="Reader.css" type="text/css" />
<style type="text/css">
<!--
.style1 {font-family: "Courier New", Courier, mono;font-style: italic;}
.opbtn {border:outset;cursor:pointer;}
a {text-decoration:none;color:#000000;}
-->
</style>
</head>
<body bgcolor="#DEE3F7">
<%
Dim ISBN,dbReady,pISBN,nISBN
Dim BookTitle,BookOnlineDate,BookID
Dim iBDB,Rs,iDay

Set iBDb = New iBDataBase
	iBDb.ConnString = ConnectionString
	iBDb.RsUpdate = True

ISBN = Request.QueryString("ISBN")
iBDb.Sql = "select top 1 BooKID,BookName,OnlineDate,Hits,YesterdayHits,TodayHits,ThisWeekHits,LastWeekHits,TimeFlag from [iReaderBooks] where ISBN='"&CheckStr(ISBN)&"' And STOK=True "

Set Rs = iBDb.GetRs()
If Not Rs.Eof Then
	BookTitle = Rs("BookName")
	BookOnlineDate = Rs("OnlineDate")
	BookID = Rs("BookID")
	dbReady = True

	'****************** Update Count ***********************
	If (CountBookRead = True) Then
		iDay = DateDiff("d", Rs("TimeFlag"), Now())
		If (iDay = 1) Then
			Rs("YesterdayHits") = Rs("TodayHits")
			Rs("TodayHits") = 1
		ElseIf (iDay = 7) Then
			Rs("LastWeekHits") = Rs("ThisWeekHits")
			Rs("ThisWeekHits") = 1
		End If
		Rs("Hits") = Rs("Hits") + 1
		Rs("TodayHits") = Rs("TodayHits") + 1
		Rs("ThisWeekHits") = Rs("ThisWeekHits") + 1
		Rs("TimeFlag") = Now()
		Rs.Update()
	End If
	'-----------------------------------------------

	pISBN = iBDb.GetScalar("select top 1 ISBN from [iReaderBooks] where BookID < "&BookID&" And STOK=True order by BookID desc")
	nISBN = iBDb.GetScalar("select top 1 ISBN from [iReaderBooks] where BookID > "&BookID&" And STOK=True order by BookID Asc")

	pISBN = GetValue(Len(pISBN)=10,"<a href=""/iReader/?ISBN="&pISBN&""" target=""_top"">上一本</a>","<font color=""gray"">上一本</font>")
	nISBN = GetValue(Len(nISBN)=10,"<a href=""/iReader/?ISBN="&nISBN&""" target=""_top"">下一本</a>","<font color=""gray"">下一本</font>")

Else
	dbReady = False
End If

	Rs.Close()
Set Rs = Nothing
Set iBDb = Nothing

'===============================================

If Len(ISBN)=10 And (dbReady = True) Then
%>
<script language="JavaScript">
<!--
function encodeURI(str) 
{
	var returnString;
		returnString=escape(str);
		returnString=returnString.replace(/\+/g,"%2B");
	return returnString;
}

function doQueryWord(oId)
{
	//var queryUrl = "http://www.baidu.com/s?lm=0&si=&rn=10&ie=gb2312&ct=1048576&wd=";
	var queryUrl = "http://www.iciba.com/search?s=";
	var obj = document.getElementById(oId);
	var objEncoding = document.getElementById("NeedEncode");
	if (obj.value && obj.value.length>0)
	{
		if (objEncoding.checked)
		{
			window.open(queryUrl + encodeURI(obj.value));
		}
		else
		{
			window.open(queryUrl + obj.value);
		}
	}
}
//-->
</script>
<table width="100%" height="25"  border="2" cellpadding="2" cellspacing="0">
  <tr>
  	<td valign="middle"><a href="/iReader" target="_top"><img src="images/folder.gif" alt="iReader Books" width="17" height="16" border="0" align="absmiddle" /></a>&nbsp;<a href="Reader.asp?ISBN=<%=ISBN%>" target="mainWin" title="ISBN:<%=ISBN%>"><span class="style1"><%=BookTitle%></span></a> &nbsp;&nbsp;&nbsp;到期时间：<span class="style1"><%=formatDT(DateAdd("d",EXPIREDDAYNUM,BookOnlineDate),"7")%></span></td>
    <td valign="middle" align="right" width="200">
	<span name="bkmk" id="bkmk" class="opbtn" title="bookmark"><a href="###">贴书签</a></span>
	<span name="voteb" id="voteb" class="opbtn" title="vote this book"><a href="###">投一票</a></span>
	<span name="prevb" id="prevb" class="opbtn" title="the preview book"><%=pISBN%></span>
	<span name="nextb" id="nextb" class="opbtn" title="the next book"><%=nISBN%></a></span></td>
	<td align="right" width="150">查词<input type="text" size="6" name="key" id="key" maxlength="32" /><input type="checkbox" id="NeedEncode" title="编码后发送" />
	<img src="images/searchd.gif" width="19" height="16" border="0" onclick="javascript:doQueryWord('key')" style="cursor:pointer" />
	</td>
  </tr>
</table>
<%
Else
	Response.Write("当前书目不存在")
End If
%>
</body>
</html>