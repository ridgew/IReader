<!--#include file="Const.asp"-->
<!--#include file="iBCore.asp"-->
<%
Server.ScriptTimeOut = 99999999
Response.Buffer = False 
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Table of Content</title>
<script language="javascript" type="text/javascript" src="Reader.js"></script>
<base target="mainWin">
<link rel="stylesheet" href="Reader.css" type="text/css" />
<script language="javascript">
   try {
   	if (top.location.host!=self.location.host) { top.location.href = self.location; }
    }
   catch (e) { }
  
  function selectstart()
    {
    window.event.cancelBubble = true;
    window.event.returnValue = false;
    return false;
    }

  function loadMsg()
  {
	 document.getElementById("MsgPanel").style.display = "none";
  }
</script>
</head>
<body onselectstart="selectstart()">
<div id="MsgPanel" style="Position:absolute;top:100px;left:30px;width:220px;height:35px;border:1px solid green;background-color:#f3f3f3;display:block;padding:15px;z-index:100;">TOC页面内容加载中……<br/><br/>如时间稍长,可能是首次加载,请稍等.</div>
<%
Dim ISBN,TOC,strHHContent,BooK
Dim Rs,sql,Conn
Dim dbBookPath,dbHHCPath,dbReady,dbHHCHTML
Dim blnReload,blnShowHHCObject

Call DBOpen(Conn,ConnectionString)

dbHHCHTML = ""
ISBN = Request.QueryString("ISBN")
blnReload = GetValue(Request.QueryString("do")="reload",True,False)
blnShowHHCObject = GetValue(Request.QueryString("show")="hhc",True,False)

Set Rs = Server.CreateObject("ADODB.Recordset")
If Len(ISBN) = 10 Then
	'**********  Display Book HHCHTML
	Sql = "select top 1 BookPath,TOCSys,HHCHTML from [iReaderBooks] where ISBN='"&CheckStr(ISBN)&"' And STOK=True "
	Rs.Open sql,Conn,3,3
		If Not Rs.Eof Then

			dbBookPath = Rs("BookPath")
			dbHHCPath = Rs("TOCSys")
			dbHHCHTML = Rs("HHCHTML")
			'2005年10月8日 13:02:10
			If (Not IsEmpty(Session("HHCHTML")) And Session("HHCISBN") = ISBN) Then
				If (IsNull(dbHHCHTML) Or blnReload = True) Then
					Rs("HHCHTML") = Session("HHCHTML")
					dbHHCHTML = Session("HHCHTML")
					Session("HHCHTML") = Empty
					Session("HHCISBN") = Empty
					Rs.Update()
				End If
			End If

			If (IsNull(dbHHCHTML) Or blnReload = True) Then
				Set Book = New iBook
				strHHContent = Book.GetHtmlContent("mk:@MSITStore:" & GetValue(InStr(1,dbBookPath,":",1)>0,dbBookPath,bookRootPath & dbBookPath) & "::" & dbHHCPath)
				Set TOC = New HHCParse
					strHHContent = TOC.GetHHCBodyContent(strHHContent)
					If (blnShowHHCObject = False) Then
						strHHContent = TOC.HHCObjecToLink("Reader.asp?ISBN="&ISBN&"&URI="&dbHHCPath,strHHContent)
					End If
					dbHHCHTML = strHHContent
				Set TOC = Nothing
				Set Book = Nothing
				Session("HHCHTML") = dbHHCHTML
				Session("HHCISBN") = ISBN
			End If

			Response.Write dbHHCHTML
		 Else
			Response.Write("<p>&nbsp;</p><p>&nbsp;</p><p align=""center"">当前书目不存在或没有装载书目</p>")
		End If
	Rs.Close()

   Else
   '**********  Display Book HHCHTML
	Sql = "select top 1 HelpHHCHtml from [iReaderConst]"
	Rs.Open sql,Conn,1,1
		Response.Write Rs(0)
	Rs.Close()
End If

Set Rs = Nothing
Call DbClose(Conn)
%>
<script language="JavaScript">
loadMsg();
if (parent) {
 if(parent.document.frames) {
   var obj = parent.document.frames["tocPanel"];
   try { obj.cols = "300,*"; } catch (e) { }
	}
}
</script>
</body>
</html>