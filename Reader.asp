<!--#include file="Const.asp"-->
<!--#include file="iBCore.asp"-->
<%
Server.ScriptTimeOut = 99999999
Response.Buffer = False 

Dim ISBN,Book,URI,Mark
Dim iBDB,Rs
Dim dbMainURI,dbTOCURI,dbBookPath

ISBN = Request.QueryString("ISBN")
URI = Request.QueryString("URI")
Mark = Request.QueryString("Mark")

If Len(ISBN)=10 Then
	If Session("ISBN") <> ISBN Then
		Set iBDb = New iBDataBase
			iBDb.ConnString = ConnectionString
			iBDb.SQL = "select top 1 BookPath,MainURI,TOContent from [iReaderBooks] where ISBN='"&CheckStr(ISBN)&"' And STOK=True "
		Set Rs = iBDb.GetRs()
		If Not Rs.Eof Then
			dbBookPath = Rs("BookPath")
			dbMainURI = Rs("MainURI")
			dbTOCURI = Rs("TOContent")
			'-----------------
			Session("ISBN") = ISBN
			Session("BookPath") = dbBookPath
			Session("MainURI") = dbMainURI
			Session("TOCURI") = dbTOCURI
		End If
			Rs.Close()
		Set Rs = Nothing
		Set iBDb = Nothing
	Else
		dbBookPath = Session("BookPath")
		dbMainURI = Session("MainURI")
		dbTOCURI = Session("TOCURI")
	End If

		'Response.Write "ms-its:" & GetValue(InStr(1,dbBookPath,":",1)>0,dbBookPath,bookRootPath & dbBookPath) & "::"
		'Response.Write "mk:@MSITStore:" & GetValue(InStr(1,dbBookPath,":",1)>0,dbBookPath,bookRootPath & dbBookPath) & "::"
		'Response.End

	Set Book = New iBook
		Book.ISBN = ISBN
		Book.BaseURI = "mk:@MSITStore:" & GetValue(InStr(1,dbBookPath,":",1)>0,dbBookPath,bookRootPath & dbBookPath) & "::"
		Book.MainURI = dbMainURI
		Book.TOCURI = dbTOCURI
		If IsBinaryContent(URI) Then
			Book.Binary = True
		ElseIf LCase(Right(URI,4)) = ".css" Then
			Response.AddHeader "Content-Type","text/css"
		ElseIf LCase(Right(URI,4)) = ".htm" Or LCase(Right(URI,5)) = ".html" Then
			
			Response.AddHeader "Content-Type","text/html"
			If Request.Cookies("charset") = "iso-8859-1" Then
				Response.Charset = "iso-8859-1"
			Else
				Response.Charset = "gb2312"
			End If
			
		End If
		Book.Write(URI)

'		If LCase(Right(URI,4)) = ".htm" Or LCase(Right(URI,5)) = ".html" Then
'			Call ICIBAWrite()
'		End If
	Set Book = Nothing
  Else
		'Response.Write("当前书目不存在")
		'Response.Write("<hr>")
		ShowMsgPage("default.asp")
End If

'即划即译
'http://web.iciba.com/partner/#jhjy
Sub ICIBAWrite()
	Response.Write("<div align=""center"" style=""background-color:#E6EFFF; padding:5px 5px 5px 5px; width:180px; display:inline"">"&vbCrlf)
	Response.Write("<span id=""Kingsoft_openDict2""></span>&nbsp;"&vbCrlf)
	Response.Write("<a href=""javascript:openDictF();"" id=""Kingsoft_openDict""></a>&nbsp;"&vbCrlf)
	Response.Write("<a href="""" id=""Kingsoft_helpDict"" title=""查看帮助"" target=""_blank"">帮助</a></div>")

	Response.Write("<link href=""http://web.iciba.com/sl/dict/main.css"" rel=""stylesheet"" type=""text/css"">"&vbCrlf&_
		"<SCRIPT>var uid = 23636,sid = 0;</SCRIPT>"&vbCrlf&_
		"<script language=""javascript"" type=""text/javascript"" id=""Kingsoft_insertmean""></script>"&vbCrlf&_
		"<script language=""javascript1.2"" type=""text/javascript"" src=""http://web.iciba.com/sl/dict/cb-gb2312.js""></script>")

End Sub
%>