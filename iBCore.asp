<%
Sub DBOpen(ByRef objConn,ByVal ConnectionString)
	Set objConn = Server.CreateObject("ADODB.Connection")
	 On Error Resume Next
	 objConn.Open ConnectionString
	 If Err.number<>0 then
	    'Response.Clear()
		'Response.Charset="gb2312"
	    Response.Write("�����ݿ����ʧ��!")
		Response.End()
	 End If
	 On Error Goto 0
End Sub

Sub DbClose(ByRef objConn)
	On Error Resume Next
	objConn.Close()
	Set objConn = Nothing
	On Error Goto 0
End Sub

'************************************************
'Main Function List
'2005��9��28�� ������ 
'--------------------------------------------------------------
Function FileExist(ByVal FilePath)
	Dim objFSO,blnReturn
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		blnReturn = objFSO.FileExists(FilePath)
	Set objFSO = Nothing
	FileExist = blnReturn
End Function


Function BuildHelpHHC(ByRef iBDb,ByVal pId)
	Dim Rs,iTotal,retStr
	iBDB.SQL = "select hId,iConIndex,Title from [iReaderHelp] where ParentID="&pId
	Set Rs = iBDb.GetRs()
		While Not Rs.Eof
			iTotal = iBDb.GetScalar("select Count(hId) as total from [iReaderHelp] where ParentId="&Rs("hId"))
			If CLng(iTotal) > 0 Then
				retStr = retStr & "<LI><span class=""box"" onclick=""iBTree(this,this.parentNode);""><img src=""images/plus.gif"" border=""0""> <img src=""images/folder.gif"" border=""0""></span> <a href=""Help.asp?id="&Rs("hId")&""">"&Rs("Title")&"</a>" & vbCRLf
				retStr = retStr & "<ul class=""folder"">" & vbCRLf
				retStr = retStr & BuildHelpHHC(iBDb,Rs("hId"))
				retStr = retStr & "</ul>" & vbCRLf
			Else
				retStr = retStr & "<LI><img src=""images/icon/"&Rs("iConIndex")&".gif"" border=""0""> <a href=""Help.asp?id="&Rs("hId")&""">"&Rs("Title")&"</a>" & vbCRLf
			End If
			Rs.MoveNext
		Wend
	Set Rs = Nothing
	BuildHelpHHC = retStr
End Function

Function IsBinaryContent(ByVal URI)
	If Trim(URI) = "" Then
		IsBinaryContent = False
		Exit Function
	End If

	Dim strBinFiles,idx,fileExt
	strBinFiles = ",html,htm,css,js,vbs,asp,php,jsp,aspx,asmx,"
	idx = InstrRev(URI,".",-1,1)
	IsBinaryContent = Not (InStr(1,strBinFiles,","&Mid(URI,idx+1)& ",",1)>0)
End Function

Function getValue(ByVal blnJudge,ByVal yesShow,ByVal noShow)
    if (blnJudge = true) then
	    getValue = yesShow
    else
	    getValue = noShow
	end if
End Function

Function RenameFile(ByVal strFilePath,ByVal strNewFilePath)
	Dim FSO
	Set FSO = Server.CreateObject("Scripting.FileSystemObject")
	If FSO.FileExists(strNewFilePath) Then
		Set FSO = Nothing
		RenameFile = False
	Else
		FSO.CopyFile strFilePath,strNewFilePath,true
		FSO.DeleteFile  strFilePath,True
		Set FSO = Nothing
		RenameFile = True
	End If
End Function

'��ȡ�ͻ��˵�IP
Public Function Web_GetClientIP()
	Web_GetClientIP = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
	if Web_GetClientIP = "" then Web_GetClientIP = Request.ServerVariables("REMOTE_ADDR")
End Function

'�ͻ��˽ű�:alert(Msg) I;
Public Sub Client_Alert(ByVal  Msg)
   Response.Write("<script language=""javascript"">alert(""" & Msg & """);</script>")
End Sub

'�ͻ��˽ű�:alert(Msg) II;
Public Sub  Client_Alert2(ByVal Msg,ByVal returnURL)
  Response.Write("<script language=""javascript"">alert(""" & Msg & """);location.href=""" &returnURL& """;</script>")
End Sub


'�ͻ��˽ű�:confrim(Msg) I;
Public Sub Client_Confirm(ByVal Msg,ByVal url)
   Response.Write("<script language=""javascript"">" &_
		"if (confirm(""" & Msg & """)) " &_
		" { location.href=""" &  url & """; }" &_
		"</script>")
End Sub

'�ͻ��˽ű�:confrim(Msg) II;
Public Sub Client_Confirm2(ByVal Msg,ByVal cfmurl,ByVal retrunURL)
   Response.Write("<script language=""javascript"">" &_
		"if (confirm(""" & Msg & """)) " &_
		" { location.href=""" &cfmurl& """; }" &_
		"else { location.href=""" & retrunURL & """; }</script>")
End Sub

'�ͻ��˽ű�:�ض�����ַ
Public Sub Client_Redirect(ByVal URL,ByVal CopyHistory)
   if CopyHistory then
	   Response.Write("<script language=""javascript"">top.location.href=""" &URL& """;</script>")
	else
	   Response.Write("<script language=""javascript"">top.location.replace(""" &URL& """);</script>")
	end if
End Sub

Sub ShowMsgPage(URLPath)
    Response.Clear()
	Response.Redirect(URLPath)
	Response.End
End Sub

'*****************************************
'��ʾ�ض���Ϣ
'*****************************************
Sub ShowMsg(strMsg,iSecond)
 if len(strMsg)>0 then
%>
<div id="cMsg" style="Position:absolute;top:240px;left:240px;width:350px;height:50px;border:1px solid green;background-color:#f3f3f3;display:block;padding:20px;z-index:100;" ondblclick="this.style.display='none';"><%=strMsg%></div>
<script language="javascript">
  window.setTimeout("document.getElementById('cMsg').style.display='none';",<%=iSecond%>*1000,"javascript");
</script>
<%
 end if
End Sub

'*****************************************
'��ʾ�ض���Ϣ3���ת���ض���ַ
'*****************************************
Sub DisplayMsgandGo(strMsg,urlGo)
 if len(strMsg)>0 then
    strMsg = strMsg &"<br>3 ����֮��ϵͳ�Զ�ת�򡭡�"
%>
<div id="cMsg" style="Position:absolute;top:240px;left:240px;width:350px;height:50px;border:1px solid green;background-color:#f3f3f3;display:block;padding:20px;z-index:100;" ondblclick="this.style.display='none';"><%=strMsg%></div>
<script language="javascript">
  window.setTimeout("document.getElementById('cMsg').style.display='none';location.href='<%=urlGo%>';",3000,"javascript");
</script>
<%
 end if
End Sub

'*****************************************
'��ʾ�ض���Ϣ3���,���нű�
'*****************************************
Sub DisplayMsgandDo(strMsg,strScripts)
 if len(strMsg)>0 then
%>
<div id="cMsg" style="Position:absolute;top:240px;left:240px;width:350px;height:50px;border:1px solid green;background-color:#f3f3f3;display:block;padding:20px;z-index:100;" ondblclick="this.style.display='none';"><%=strMsg%></div>
<script language="javascript">
  window.setTimeout("<%=strScripts%>",3000,"javascript");
</script>
<%
 end if
End Sub

''+++++++++++++++++++++++++++++++++++++++++++++++++++
''������:IsOuterPost
''����  :�ж��Ƿ�Ϊ�ⲿ�ύ����
''����  :��
''+++++++++++++++++++++++++++++++++++++++++++++++++++
Function IsOuterPost()
	dim server_v1,server_v2
	server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
	server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
	if mid(server_v1,8,len(server_v2))<>server_v2 then
		IsOuterPost=true
	else
		IsOuterPost=false
	end if
End Function

'*****************************************
'�Ƿ��Ǳ��ύ
'*****************************************
Function IsPostBack()
   IsPostBack = (UCase(Request.ServerVariables("REQUEST_METHOD")) = "POST")
End Function

'*********************************
'�жϱ����Ƿ���ѡ��״̬(��ѡ����ѡ)
'***********************************
Function isChecked(ByVal formName,ByVal itemValue,ByVal bMultiply)
    Dim bProcess,ProcessValue
	    bProcess = (Request.TotalBytes>0)

	if (bProcess=true) then
	   ProcessValue = Request.Form(formName)
	else
	   ProcessValue = Trim(formName)
	end if

    if not bMultiply then
	    if ProcessValue = itemValue then
		   isChecked = " checked"
		end if
    else
	    Dim strRequest
		   strRequest = Replace(ProcessValue,", ",",")
	    if InStr(1,(","&strRequest&","),","&itemValue&",",1)>0 then
		   isChecked = " checked"
		end if
	end if
End Function

'*********************************
'�жϱ����Ƿ���ѡ��״̬(�����б�)
'***********************************
Function isSelected(ByVal formName,ByVal itemValue)
	if Request(formName) = itemValue then
	   isSelected = " selected"
	else
	   isSelected = ""
	end if
End Function

'*********************************
'�жϱ����Ƿ���ѡ��
'***********************************
Function isThisValue(ByVal bValue,ByVal retStr)
  if (bValue=true) then
    isThisValue = retStr
  else
    isThisValue = ""
  end if
End Function

'*******************************************
'���Բ��� 2005��5��18�� ������ [R.W.]
'################################################################################################
const sDebugTemplate = "<div align=""center""><div style=""Background-color:#f3f3f3;color:#000000;width:75%;height:120px;font-family:'Times New Roman';font-size:14px;border:1px #cccccc dotted;padding:5px;""><fieldset style=""height:100%""><legend>===========������Ϣ==============</legend><div align=""left""  style=""text-indent:24px;font-family:fixedsys;font-size:12px;color:red;word-break:break-all;line-height:150%;padding-left:32px;padding-right:32px;"">[��Ϣ����]</div></fieldset></div></div>"

Sub Debug_String(ByVal message)
        Response.Write(Replace(sDebugTemplate,"[��Ϣ����]",Message))
end Sub

Sub Debug_Topic(ByVal message,ByVal topic)
        Response.Write(Replace(Replace(sDebugTemplate,"[��Ϣ����]",Message),"������Ϣ",topic))
End Sub

Sub Debug(ByVal Message)
	Debug_String(message)
End Sub

Function IsEmptyStr(strChk)
    if IsNull(strChk) then
	   IsEmptyStr = true
	   Exit Function
    else
       if Len(CStr(strChk))>=1 then
	      IsEmptyStr = false
	   else
	      IsEmptyStr = true
	   end if
	end if
End Function

Function IsNumber(str)
   if not IsEmptyStr(str) then
       isNumber = isNumeric(str)
   else
       isNumber = false
   end if
End Function

'*********************************
'��ԭ�ύ��ֵ(�ı�ѡ��ı�����)
'***********************************
Function GetFrmItemValue(ByVal formName)
	GetFrmItemValue = Server.HTMLEnCode(Request(formName))
End Function

'��ֹSQLע��
Function Checkstr(Str)
		If Isnull(Str) Then
			CheckStr = ""
			Exit Function 
		End If
		Str = Replace(Str,Chr(0),"")
		CheckStr = Replace(Str,"'","''")
End Function

Function jsTxt(str)
  jsTxt = Replace(Replace(str,"'","\'"),chr(34),"\"&Chr(34))
End Function

Function CreateWindowsGUID()
  CreateWindowsGUID = CreateGUID(8) & "-" & _
    CreateGUID(4) & "-" & _
    CreateGUID(4) & "-" & _
    CreateGUID(4) & "-" & _
    CreateGUID(12)
End Function

Function CreateGUID(tmpLength)
  Randomize Timer
  Dim tmpCounter,tmpGUID
  Const strValid = "0123456789ABCDEF"
  For tmpCounter = 1 To tmpLength
    tmpGUID = tmpGUID & Mid(strValid, Int(Rnd(1) * Len(strValid)) + 1, 1)
  Next
  CreateGUID = tmpGUID
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'��������formatDT
'��  �ã���ʽ��������ʾ
'��  ����Dtype ��ʾ���ͣ�DateTime Ҫ��ʽ����ʾ��ʱ��
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function formatDT(DateTime,Dtype)
	select case Dtype
	'2004-07-25 09:40:50
	case "0" formatDT = year(DateTime) & "-" & doublenum(Month(DateTime)) & "-" & doublenum(Day(DateTime)) & " " & doublenum(Hour(DateTime)) & ":" & doublenum(Minute(DateTime)) & ":" & doublenum(Second(DateTime))
	'2004-07-25 09:40
	case "1" formatDT = year(DateTime) & "-" & doublenum(Month(DateTime)) & "-" & doublenum(Day(DateTime)) & " " & doublenum(Hour(DateTime)) & ":" & doublenum(Minute(DateTime))
	'2004-07-25
	case "-1" formatDT = year(DateTime) & "-" & doublenum(Month(DateTime)) & "-" & doublenum(Day(DateTime))
	'07/25/03
	case "2" formatDT =  doublenum(Month(DateTime)) & "/" & doublenum(Day(DateTime))& "/" & Right(year(DateTime),2)
	'2004-07
	case "3" formatDT = year(DateTime) & "-" & doublenum(Month(DateTime))
	'07-25
	case "4" formatDT = doublenum(Month(DateTime)) & "-" & doublenum(Day(DateTime))
	'09:40:50
	case "5" formatDT = doublenum(Hour(DateTime)) & ":" & doublenum(Minute(DateTime)) & ":" & doublenum(Second(DateTime))
	'09:40
	case "6" formatDT = doublenum(Hour(DateTime)) & ":" & doublenum(Minute(DateTime))
	'2004��07��25��
	case "7" formatDT = year(DateTime) & "��" & doublenum(Month(DateTime)) & "��" & doublenum(Day(DateTime)) & "��"
	'2004��07��
	CASE "8" formatDT = year(DateTime) & "��" & doublenum(Month(DateTime)) & "��"
	'07��25��
	case "9" formatDT = doublenum(Month(DateTime)) & "��" & doublenum(Day(DateTime)) & "��"
	'07��25�� 09:40
	case "10" formatDT = doublenum(Month(DateTime)) & "��" & doublenum(Day(DateTime)) & "�� " & doublenum(Hour(DateTime)) & ":" & doublenum(Minute(DateTime))
	'Monday,Jul 25,2004
	case "11"
			 MonthArray = Array("January","February","March","April","May","June","July","August","September","October","November","December")
			 WeekArray =  Array("Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday")
			 formatDT = WeekArray(Weekday(DateTime )-1) & "," & MonthArray(Month(DateTime)-1) & " " &Day(DateTime) & "," &  Year(DateTime)
	end select
End Function

'ȡ��λ���ϵ�����
'���ȡָ����nλ������ʹ��Right����(���ַ����ұ߷���ָ����Ŀ���ַ���)
Function DoubleNum(fNum)
    if fNum > 9 then 
        doublenum = fNum 
    else 
        doublenum = "0" & fNum
    end if
End Function

'�滻�ַ�
Function ReplaceText(StrContent,PatternStr,reText)	
	Dim objRe,strReturn
	Set objRe = New RegExp
	objRe.Pattern = PatternStr
	objRe.Global = True
	objRe.IgnoreCase = True
	strReturn = objRe.Replace(StrContent,reText)
	Set objRe=Nothing
	ReplaceText = strReturn
End Function

'ƥ���ַ�
Function MatchText(PatternStr,StrContent)
	Dim objRe,blnReturn
	Set objRe = New RegExp
	objRe.Pattern = PatternStr
	objRe.Global = True
	objRe.IgnoreCase = True
	blnReturn = objRe.Test(StrContent)
	Set objRe=nothing
End Function

'************************************************
'iBook Core Lib
'2005��9��27�� ���ڶ�
'--------------------------------------------------------------
Class iBook
	
	Private strISBN,strBaseURI,strMainURI,strTOCURI
	Private blnBinary

	Private Sub Class_Initialize()
		blnBinary = False
	End Sub

	Private Sub Class_Terminate()

	End Sub
	
	Public Property Let BaseURI(ByVal strURI)
		strBaseURI = strURI
	End Property

	Public Property Let MainURI(ByVal strURI)
		strMainURI = strURI
	End Property

	Public Property Let TocURI(ByVal strURI)
		strTOCURI = strURI
	End Property

	Public Property Let Binary(ByVal bValue)
		blnBinary = bValue
	End Property

	Public Property Let ISBN(ByVal BookISBN)
		strISBN = BookISBN
	End Property

	Public Sub Write(ByVal FilePath)
		If Trim(FilePath) = "" Then
			Response.Write GetPage(strMainURI)
		Else
			If (blnBinary = True) Then
				Response.BinaryWrite GetPage(FilePath)
			Else
				Response.Write GetPage(FilePath)
			End If
		End If
	End Sub

	Private Function GetPage(ByVal path)
		Dim strHtml
		strHtml = GetHtmlContent(strBaseURI & path)
		'-------- �滻�鼮����
		'GetPage = strHtml
		GetPage = SetBookContent(strHtml,Path)
	End Function

	Public Function GetHtmlContent(ByVal URI)
		On Error Resume Next
		Dim http,strResponseText
		Set http=Server.CreateObject("Msxml2.XMLHTTP")
		    Http.open "GET",URI,false,"",""
		    Http.send()
		   If (http.status <> 200) AND (Http.readystate<>4) Then
				 GetHtmlContent = "Error Read Data."
				 Set http = nothing
				 Exit Function
			End If

		   If (blnBinary = True) Then
				strResponseText = Http.responseBody
		   Else
				'strResponseText = Http.responseText
				strResponseText = RSBinaryToString(Http.responseBody)
				'strResponseText = bytes2BSTR(Http.responseBody)
		   End If
		Set http=nothing
		If Err.number<>0 then Err.Clear
		GetHtmlContent  = strResponseText
	End Function

	Private Function GetURL(url)  '�ַ���ʽ��xmlhttp
		Set Retrieval = Server.CreateObject("Msxml2.XMLHTTP")
			  With Retrieval
			  .Open "GET", url, False, "", ""
			  .Send
			  'GetURL = .ResponseText
			  GetURL = bytes2BSTR(.Responsebody)
			  End With
		Set Retrieval = Nothing
	End Function

	Private Function SetAbsolute(ByVal Matches, ByVal Group, ByVal BaseURI)
		Dim strReturn, i, length, strTemp
		If (VarType(Matches) = 9) Then
			length = Matches.Count-1
			For i=0 To length
				If (i <> CInt(Group)) Then
					strReturn = strReturn & Matches(i)
				Else
					strTemp = GetRootURI("?ISBN="&strISBN&"&URI="&BaseURI, Matches(i))
					strReturn = strReturn & strTemp
				End If
			Next
		ElseIf (VarType(Matches)=8) Then
			strReturn = strReturn & GetRootURI("?ISBN="&strISBN&"&URI="&BaseURI, Matches)
		End If

		If (Right(strReturn,3) = "\''") Then
			strReturn = Replace(Mid(strReturn,1 , Len(strReturn)-3),"=?ISBN", "=\'?ISBN") & "'"
		End If
		If (Right(strReturn,2) = Chr(39) & Chr(34)) Then
			strReturn = Replace(Mid(strReturn,1 , Len(strReturn)-2),"=?ISBN", "='?ISBN") & Chr(34)
		End If
		strReturn = Replace(strReturn, chr(34)&Chr(34), chr(34))

		SetAbsolute = strReturn
	End Function

	'----------------------------------
	'������ʽ�滻��������ƥ���ַ�
	' 2006��8��16�� 22:08:48
	' Ridge Wong
	'-----------------------------------
	Public Function MatchReplace(ByVal SourceText, ByVal Pattern, ByVal EvalMatches)

		Dim regEx, Match, Matches
		Dim strReturn,strTemp
		Dim idxStart, idxEnd

		idxStart = 1 : idxEnd = 1
		strReturn = "" : strTemp = ""

		Set regEx = New RegExp
			regEx.Pattern = Pattern
			regEx.Global = True
			regEx.IgnoreCase = True

		Set Matches = regEx.Execute(SourceText)
			For Each Match in Matches
			
				If (Match.FirstIndex > 0) Then
					idxEnd = Match.FirstIndex
					strReturn = strReturn & Mid(SourceText, idxStart, idxEnd-idxStart+1)

					'����ƥ����Ŀ Match.Value/Match.SubMatches(6)
					EvalMatches = Replace(EvalMatches,"$", "Match.SubMatches")
					strTemp = Eval(EvalMatches)
					strReturn = strReturn & strTemp

					idxStart = idxEnd + Match.Length+1
				End If
			Next
		Set Matches = Nothing
		Set regEx = Nothing

		If (idxStart <= Len(SourceText)) Then
			strReturn = strReturn & Mid(SourceText, idxStart)
		End If

		MatchReplace = strReturn
	End Function


	'*******************
	'�滻ԭ·��Ϊ��·�����ַ�����ʽ
	'		link -> href
	'		a -> href
	'		script -> src
	'		img -> src
	'		table -> background
	'		td -> background
	'		body -> background
	'	.jpg|.jpeg|.png|.bmp|.html|.htm|.css|.js|.htc
	'-------------------------------
	Public Function SetBookContent(ByRef strHtml,ByRef strCurrentURI)

		If (blnBinary = True) Then 
			SetBookContent = strHtml
			Exit Function
		End If

		Dim strReturn,strBaseURI
		strBaseURI = GetBaseURI(strCurrentURI)
		strReturn =  MatchReplace(strHtml,""&_
		"(src|href|url|action|background)(=)(('|"")?)([^\s\>]+)(\3)", ""&_
		"SetAbsolute($,4, """&strBaseURI&""")")
		
		strReturn =  MatchReplace(strReturn, ""&_
		"(url\()([^\?\)=&]+)(\))", ""&_
		"SetAbsolute($,1, """&strBaseURI&""")")
		SetBookContent = strReturn
	End Function

	'*******************
	'�滻ԭ·��Ϊ��·�����ַ�����ʽ
	'		link -> href
	'		a -> href
	'		script -> src
	'		img -> src
	'		table -> background
	'		td -> background
	'		body -> background
	'	.jpg|.jpeg|.png|.bmp|.html|.htm|.css|.js|.htc
	'-------------------------------
	Public Function SetBookContent_Obsolute(ByRef strHtml,ByRef strCurrentURI)
		Dim objRegExp,Matches,Match,strBaseURI
		Dim strTemp,strPath,strChar,strFile
		strBaseURI = GetBaseURI(strCurrentURI)
		Set objRegExp = New Regexp
			objRegExp.IgnoreCase = True
			objRegExp.Global = True
			objRegExp.Pattern = "(\<)(link|a|img|body|table|td|script)(\s+)([^\<]*?)(href|src|background)(=)(.[^\>\s]*)"
			Set Matches =objRegExp.Execute(strHtml)
			For Each Match in Matches
				strTemp = Match.SubMatches(6)
				If (Left(strTemp,2) <> "''") And (Left(strTemp,2)<> chr(34)&chr(34) And (Left(strTemp,1)<> "\")) Then

					strChar = GetAttributeChar(strTemp)
					If InStr(1,strTemp,"?ISBN=",1) <= 0 Then
						strPath = GetRootURI("?ISBN="&strISBN&"&URI="&strBaseURI,Replace(strTemp,strChar,""))
						strHtml = Replace(strHtml,strTemp,strChar&strPath&strChar,1,-1,1)
					End If
			   End If
			Next
			Set Matches = Nothing

			'2005��10��6�� ������ 19:49:15 fix style url()
			objRegExp.Global = False
			objRegExp.Pattern = "(url\()([^\?\)=&]+)(\))"
			Set Matches =objRegExp.Execute(strHtml)
			For Each Match in Matches
				strTemp = Match.SubMatches(1)
				If (Left(strTemp,2) <> "''") And (Left(strTemp,2)<> chr(34)&chr(34) And (Left(strTemp,1)<> "\")) Then

					strChar = GetAttributeChar(strTemp)	
					strPath = GetRootURI("?ISBN="&strISBN&"&URI="&strBaseURI,Replace(strTemp,strChar,""))
					strHtml = Replace(strHtml,strTemp,strChar&strPath&strChar,1,-1,1)
			   End If
			Next
			Set Matches = Nothing

		Set objRegExp = Nothing
		SetBookContent = strHtml
	End Function

	'GetAttributeChar("'images/tpe.jpg'") = '
	Private Function GetAttributeChar(ByRef strAttribute)
		If Left(strAttribute,1) = Chr(34) Then
			GetAttributeChar = Chr(34)
		ElseIf Left(strAttribute,1) = Chr(39) Then
			GetAttributeChar = Chr(39)
		Else
			GetAttributeChar = ""
		End If
	End Function

	'GetBaseURI("http://ssss.net/asp/111.html") = http://ssss.net/asp/
	Public Function GetBaseURI(ByRef CurrentURI)
		GetBaseURI = Mid(CurrentURI,1,InstrRev(CurrentURI,"/",-1,1))
	End Function



	'***************************************************
	'���ڻ��ڹ������(û�в���)��·��ֻ��Ҫ����Httpͷ��Base����
	'---------------------
	'2005��9��27�� ���ڶ� 23:55:49 st:todo
	Private Function SetBaseURI(ByVal strURI,ByRef strHtml)
		Response.Write StrURI
		Response.Write("<hr>")
		Response.Write strBaseURI
		Response.Write("<hr>")
		SetBaseURI = strHtml
		'SetBaseURI = "<base href=""?ISBN="&strISBN&"&URI="&strURI&""">" & VbCrLf & strHtml
	End Function

	Private Function bytes2BSTR(vIn) '�ַ�������
		Dim i, ThischrCode, NextchrCode
		strReturn = ""
		For i = 1 To LenB(vIn)
			ThischrCode = AscB(MidB(vIn, i, 1))
			If ThischrCode < &H80 Then
				strReturn = strReturn & Chr(ThischrCode)
			Else
				NextchrCode = AscB(MidB(vIn, i + 1, 1))
				strReturn = strReturn & Chr(CLng(ThischrCode) * &H100 + CInt(NextchrCode))
				i = i + 1
			End If
		Next
		bytes2BSTR = strReturn
	End Function

	Private Function RSBinaryToString(xBinary)
		Dim Binary
		'MultiByte data must be converted To VT_UI1 | VT_ARRAY first.
		If vartype(xBinary)=8 Then Binary = MultiByteToBinary(xBinary) Else Binary = xBinary
	  
		Dim RS, LBinary
		Const adLongVarChar = 201
		Set RS = Server.CreateObject("ADODB.Recordset")
		LBinary = LenB(Binary)
	  
		If LBinary>0 Then
			RS.Fields.Append "mBinary", adLongVarChar, LBinary
			RS.Open
			RS.AddNew
			RS("mBinary").AppendChunk Binary 
			RS.Update
			RSBinaryToString = RS("mBinary")
		Else
			RSBinaryToString = ""
		End If
	End Function

	Private Function MultiByteToBinary(MultiByte)
	  ' 2000 Antonin Foller, http://www.motobit.com
	  Dim RS, LMultiByte, Binary
	  Const adLongVarBinary = 205
	  Set RS = CreateObject("ADODB.Recordset")
	  LMultiByte = LenB(MultiByte)
	  If LMultiByte>0 Then
		RS.Fields.Append "mBinary", adLongVarBinary, LMultiByte
		RS.Open
		RS.AddNew
		  RS("mBinary").AppendChunk MultiByte & ChrB(0)
		RS.Update
		Binary = RS("mBinary").GetChunk(LMultiByte)
	  End If
	  MultiByteToBinary = Binary
	End Function

	'******************
	'��ȡ��ǰ�����ļ�����Ը�·����ʽ
	'2005��9��27�� ���ڶ� 0:49:44
	'GetRootURI("http://www.vbyte.com/shop/idx/","../my/dss.jpg") = http://www.vbyte.com/shop/my/dss.jpg
	'GetRootURI("http://www.vbyte.com/shop/","./my/dss.jpg") = http://www.vbyte.com/shop/my/dss.jpg
	'GetRootURI("http://www.vbyte.com/shop/","ftp://vbyte.com") = ftp://vbyte.com
	'-------------------------------------
	Public Function GetRootURI(ByVal strBase,ByVal URI)
		If Right(strBase,1) <> "/" Then
			GetRootURI = URI
			Exit Function
		End If
		If InStr(1,URI,":",1)>0 Then
			GetRootURI = URI
		Else
			If Left(URI,1) = "/" Then
				GetRootURI = GetURIRoot(strBase) & Mid(URI,2)
			ElseIf Left(URI,1) = "#" Then			'�������ê���� 2005��9��28�� ������ 19:03:34
				GetRootURI = URI
			ElseIf Left(URI,2) = "./" Then
				GetRootURI = strBase & Mid(URI,3)
			ElseIf Left(URI,3) = "../" Then
				Dim idx,i,parentPathArray,pCount
				Dim getPathArray,getPath,dCount
				getPath = Replace(Replace(strBase,GetURIRoot(strBase),""),"//","/")
				getPathArray = Split(getPath,"/")
				dCount = UBound(getPathArray)		'��ַ·��������
				parentPathArray = Split(URI,"..")
				pCount = UBound(parentPathArray)	'�л�����·���Ĵ���
				If (dCount>=pCount) Then
					idx = InStrRev(strBase,"/",-1,1)
					For i=1 To pCount
						idx = InStrRev(strBase,"/",idx-1,1) '����λ��
					Next
					GetRootURI = Mid(strBase,1,idx-1) + parentPathArray(pCount)
				Else
					GetRootURI = URI
				End If
			Else
				GetRootURI = strBase & URI
			End If
		End If
	End Function

	'*****************
	'GetURIRoot("http://www.vbyte.com:80/shop/idx/") = "http://www.vbyte.com:80/"
	Public Function GetURIRoot(ByVal strURI)
		If Left(strURI,1) = "/" Then
			GetURIRoot = "/"
		Else
			Dim idx
			idx = InStr(1,strURI,"/",1)
			While (Mid(strURI,idx+1,1) = "/")
				idx = InStr(idx+2,strURI,"/",1)
			Wend
			GetURIRoot = Mid(strURI,1,idx)
		End If
	End Function

End Class


'******************************************
Class iBDataBase
     Private iBConn,iBDataArray,iBDebug,bUpdateRs
	 Private strConn,strSQL,strErrMsg
	 Public ExecCount,ExecuteDetail(4,1)

	 Public Property Let ConnString(ByVal connStr)
	    strConn = connStr
	 End Property

	 Public Property Set ConnObject(ByRef obj)
	   Set iBConn = obj
	 End Property
	 Public Property Get ConnObject()
	    if Not IsObject(iBConn) then
		  DatabaseOpen()
		  Set ConnObject = iBConn
		end if
	 End Property

	 Public Property Let SQL(ByVal sqlStr)
	    strSQL = sqlStr
	 End Property
	 Public Property Get SQL()
	    SQL = strSQL
	 End Property

	 Public Property Let RsUpdate(ByVal bValue)
	    bUpdateRs = bValue
	 End Property
	 Public Property Get RsUpdate()
		RsUpdate = bUpdateRs
	 End Property

	 Public Property Get ErrMsg()
	    ErrMsg = strErrMsg
	 End Property

	 Public Property Get ExecDetail()
	    ExecCount	= ExecuteDetail(0,1) + ExecuteDetail(1,1) + ExecuteDetail(2,1) + ExecuteDetail(3,1)
	    ExecDetail	= ExecuteDetail(0,0)&":"&ExecuteDetail(0,1)&" "&ExecuteDetail(1,0)&":"&ExecuteDetail(1,1)&" "&_ 
					ExecuteDetail(2,0)&":"&ExecuteDetail(2,1)&" "&ExecuteDetail(3,0)&":"&ExecuteDetail(3,1)&" "&_
					ExecuteDetail(4,0)&":"&ExecuteDetail(4,1)
	 End Property

	 Private Sub Class_Initialize()
	     iBDebug = true				'Ĭ�ϵ���ģʽ
		 bUpdateRs = false				'Ĭ��Ϊ����Ҫ���µļ�¼��
		 '��ϸ����:select,update,delete,exec
		 ExecCount = 0					'�������ݿ����
		 ExecuteDetail(0,0)="��ѯ" : ExecuteDetail(0,1) = 0
		 ExecuteDetail(1,0)="����" : ExecuteDetail(1,1) = 0
		 ExecuteDetail(2,0)="ɾ��" : ExecuteDetail(2,1) = 0
		 ExecuteDetail(3,0)="�洢����" : ExecuteDetail(3,1) = 0
		 ExecuteDetail(4,0)="׷��" : ExecuteDetail(4,1) = 0
		 strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath("Data.mdb")&";"
	 End Sub 
	 Private Sub Class_Terminate()
	    if isObject(iBConn) then
			if (strErrMsg<>"") then
			    iBConn.RollbackTrans()
			else
			    iBConn.CommitTrans()
			end if
		end if
	    DatabaseClose()
	    if (iBDebug and strErrMsg<>"") then
		    Dim ErrArray,i
			 ErrArray = Split(strErrMsg,"|")
			Response.Write("<fieldset style=""font-size:12px;font-family:verdana;width:75%""><legend><font color=""red"">���ִ���</font></legend><div style=""padding:18px;"">")
			for i=0 to UBound(ErrArray)
			   Response.Write("<li>"&ErrArray(i))
			next
			Response.Write("</div></fieldset>")
		 end if
	 End Sub

	 Private Sub DatabaseOpen()
	     if (not iBDebug) then On Error Resume Next
	     set iBConn = Server.CreateObject("ADODB.Connection")
		 iBConn.Open strConn
		 iBConn.BeginTrans()
		 if Err.Number<>0 then
		    iBError("�����ݿ���ִ���:iBDataBase.DatabaseOpen()")
			Err.Clear()
		 end if
	 End Sub
	 Private Sub DatabaseClose()
	     if IsObject(iBConn) then
		     if (iBConn.State=1) then iBConn.Close()
		   set iBConn = nothing
		 end if
	 End Sub

	 Private Sub iBError(ByVal strMsg)
		if (strErrMsg<>"") then 
		  strErrMsg = strErrMsg & "|" & strMsg
		else
		  strErrMsg = strMsg
		end if
	 End Sub

	 Public Function GetRs()
		 if (Not IsObject(iBConn)) then DatabaseOpen()
		 if (not iBDebug) then On Error Resume Next
		 Dim Rs
		 Set Rs = Server.CreateObject("ADODB.Recordset")
		  if (bUpdateRs=true) then
		    Rs.CursorLocation = 3 'adUseClient
			Rs.Open strSQL,iBConn,2,3
		  else
		    Rs.CursorLocation = 3 'adUseClient
			Rs.Open strSQL,iBConn,3,1
		  end if
		  'if (not Rs.eof) then
		     Set GetRs = Rs
		  'end if
		 ExecuteDetail(0,1) = ExecuteDetail(0,1) + 1
		 if Err.Number<>0 then
			iBError("�ڽ������ݿ����ʱ���ִ���:iBDataBase.GetRs()")
			Err.Clear()
		 end if
	 End Function

	 Public Function Execute()
	     Dim iCount
		     iCount = 0
		 if (Not IsObject(iBConn)) then DatabaseOpen()
		 if (not iBDebug) then On Error Resume Next

		 if InStr(1,LCase(strSQL),"select ",1)>0 then
		     Dim Rs
		     Set Rs = iBConn.Execute(strSQL)
			  ExecuteDetail(0,1) = ExecuteDetail(0,1) + 1
		      if (not Rs.eof) then iBDataArray = Rs.GetRows()
			     iCount = UBound(iBDataArray,2)
			     Rs.Close()
		     Set Rs = nothing
		 else
		     if InStr(1,LCase(strSQL),"update ",1)>0 then ExecuteDetail(1,1) = ExecuteDetail(1,1) + 1
			 if InStr(1,LCase(strSQL),"delete ",1)>0 then ExecuteDetail(2,1) = ExecuteDetail(2,1) + 1
			 if InStr(1,LCase(strSQL),"exec ",1)>0 then ExecuteDetail(3,1) = ExecuteDetail(3,1) + 1
			 if InStr(1,LCase(strSQL),"insert ",1)>0 then ExecuteDetail(4,1) = ExecuteDetail(4,1) + 1
			 iBConn.Execute strSQL,iCount
		 end if
		 if (Err.Number<>0) then
				iBError("�ڽ������ݿ����ʱ���ִ���:iBDataBase.Execute()")
				Err.Clear()
				Execute = -1   'ʧ��
			else
				Execute = iCount   '�ɹ�����
		 end if
	 End Function

	 ''''''''''''''''''''''''''''''''
	 ''' ����ֻ��һ�е�1�е�����
	 ''''''''''''''''''''''''''''''''
	 Public Function GetScalar(ByVal sqlQuery)
	      if (Not IsObject(iBConn)) then DatabaseOpen()
		  Dim Rs
		  Set Rs = iBConn.Execute(sqlQuery)
		  if Not (Rs.eof) then
	       GetScalar = Rs(0)
		  else
		   GetScalar = ""
		  end if
		  Set Rs = nothing
	 End Function

	 Public Function GetSQLStr(ByVal StrSQL)
	     GetSQLStr = Replace(CStr(StrSQL),"'","''")
	 End Function

End Class


'******************************************
'General DataPage with Templet Class 1.2.1 for ASP 
'Create Time: 2005-5-24 [R.W.]
'Modified Time:2005-7-4
'---Funtions:
'	Replace VarItem with DataIndex
'	Support Simple Function with DataIndex Synatax
'	Auto Form Pager Html
'---1.1 Add
'   Support Multi Form Pager Identity with PagerID
'---1.2 Add	@2005-5-27
'	Support Mixed VarItem with or without DataIndex
'	Support Mixed DataIndex Synatax Function
'---1.2.1 Fix 1.2 Bugs @2005��7��4�� ����һ
'---1.2.2 ��ӱ������� @2005��7��13�� ������
'**********************************************************

Class iBDataTable
     Private iTotalCount,iPagesize,iCurrentPage
	 Private objReplaceArray(),objRepArray,objRepItem,objRepIdx,objRepFun
	 Private strErrMsg,strTemplet,iBDebug,objFSO,blnPageData
	 Private strCycleCont,strPager,strPagerID,strCycleData
	 Private strFormItems,strFormItemHtml

	 Public Property Let PageSize(ByVal iNum)
	     iPagesize = iNum
	 End Property
	 Public Property Get PageSize()
	     PageSize = iPagesize
	 End Property
	 Public Property Let PageData(ByVal blnValue)
	     blnPageData = blnValue
	 End Property
	 Public Property Let CurrentPage(ByVal iNum)
	     if IsNull(iNum) or ""=iNum then
		    iNum = 1
		 else
		    iNum = CLng(iNum)
		 end if
	     iCurrentPage = iNum
	 End Property
	 Public Property Get CurrentPage()
	     CurrentPage = iCurrentPage
	 End Property
	 Public Property Let TotalCount(ByVal iNum)
	     iTotalCount  = iNum
	 End Property

	 Public Property Let PagerID(ByVal strPager)
	     strPagerID = strPager
	 End Property

	 '---------���ݱ��ģ���滻����
	  Public Property Let dtRepArray(ByVal itemArray)
	     If IsArray(itemArray) then
		   objRepArray = itemArray
		 End If
	  End Property
	  '---------���ݱ��ģ�����
	  Public Property Let dtRepItem(ByVal itemArray)
	     objRepItem = itemArray
	  End Property
	  '---------���ݱ��ģ�����ݼ�������
	  Public Property Let dtRepIdx(ByVal itemArray)
	     objRepIdx = itemArray
	  End Property
	  '---------���ݱ��ģ���������
	  Public Property Let dtRepFun(ByVal itemArray)
	     objRepFun = itemArray
	  End Property
	  Public Property Let dtTemplet(ByVal templetStr)
	     strTemplet = templetStr
	  End Property
	  Public Property Let CycleTpt(ByVal cycleStr)
	     strCycleCont = cycleStr
	  End Property
	  Public Property Get PagerData()
	     PagerData = strPager
	  End Property
	  Public Property Get CycleData()
	     CycleData = strCycleData
	  End Property

	 Private Sub Class_Initialize()
	    objFSO = "Scripting.FileSystemObject"
		blnPageData = false '�Ƿ����Ѿ���ҳ������
		iBDebug  = true     '�Ƿ��ǵ���״̬
		iTotalCount=0 : iCurrentPage = 1 : iPageSize = 20
		strFormItems = "p|pagerTotal|pagerCurrent|PagerID"

		'�ؽ�strFormItemHtml
		Dim objItem
		if (IsPostBack=true) then
		   for each objItem in Request.Form
			 AddHiddenItem objItem,Request.Form(objItem)
	       next
	    end if
	 End Sub
	 
	 Private Sub Class_Terminate()
	    
	    if (iBDebug and strErrMsg<>"") then
		    Dim ErrArray,i
			 ErrArray = Split(strErrMsg,"|")
			Response.Write("<fieldset style=""font-size:12px;font-family:verdana;width:75%""><legend><font color=""red"">���ִ���</font></legend><div style=""padding:18px;"">")
			for i=0 to UBound(ErrArray)
			   Response.Write("<li>"&ErrArray(i))
			next
			Response.Write("</div></fieldset>")
		 end if
	 End Sub

	 '--------------�����滻����
	 Private Function setRepArray()
	    'Process Objects:objRepItem,objRepIdx,objRepFun
		Dim v,b
		If IsArray(objRepItem) and IsArray(objRepIdx) and IsArray(objRepFun) then
		   b = UBound(objRepItem)
		   If (UBound(objRepIdx)<>b or UBound(objRepFun)<>b) then  
		     iBError("�滻����Ԫ�س��Ȳ�һ��:iBDataTable.setRepArray()")
			 Exit Function
		   End If
		   ReDim objReplaceArray(b,2)
		   for v=0 to b
		      objReplaceArray(v,0) = objRepItem(v)
			  objReplaceArray(v,1) = objRepIdx(v)
			  objReplaceArray(v,2) = objRepFun(v)
		   next
		   objRepArray = objReplaceArray
		End If
	 End Function

	 Private Sub iBError(ByVal strMsg)
		if (strErrMsg<>"") then 
		  strErrMsg = strErrMsg & "|" & strMsg
		else
		  strErrMsg = strMsg
		end if
	 End Sub

	 '*************************************
	 '��ӱ�������(2005-7-13)
	 '***************************************
	 Public Sub AddHiddenItem(ByVal sItemName,ByVal sItemValue)
		 if InStr(1,"|"&strFormItems&"|","|"&sItemName&"|",1)<=0 then
			strFormItems = strFormItems & "|" & sItemName
			strFormItemHtml = strFormItemHtml & "<input type=""hidden"" name="""&sItemName&""" value="""&Replace(sItemValue,chr(34),"&quot;")&""" />"
		 end if
	 End Sub

	 Public Function GetHiddenItemValue(ByVal sItemName)
		 GetHiddenItemValue = Request.Form(sItemName)
	 End Function
	 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	 Public Function Execute(ByRef objRs)
	    Dim strTable,blnModuleTable
		  blnModuleTable = false
		If IsArray(objRepItem) and IsArray(objRepIdx) and IsArray(objRepFun) then
		    if (strTemplet<>"" and InStr(1,strTemplet,"{$",1)>0) then
			   setRepArray()
			   blnModuleTable = true
			else
			   if IsArray(strCycleCont) then
				 setRepArray()
				 blnModuleTable = true
			   else
				   if (strCycleCont<>""  and InStr(1,strCycleCont,"{$",1)>0) then
					 setRepArray()
					 blnModuleTable = true
				   end if
			   end if
			end if
		end if
		
		if (blnModuleTable = false) then
		     strTable = DataTabAuto(objRs)
			else
			 strTable = DataTabModule(objRs)
		end if
		Execute = strTable
	 End Function

	 Private function DataTabAuto(ByRef objRs)
	    Dim strDataTable
        Dim strPagerColor,strTHColor,strTRColor,strTRColor2
			strTHColor = "#BFE8FB" : strPagerColor = "#f3f3f3"
			strTRColor = "#E1F4FD" : strTRColor2 = "#FFFFFF"
		if (Request.Form("pagerTotal") <> "") then
		   iTotalCount = CLng(Request.Form("PagerTotal"))
		else
		   if (not blnPageData) then iTotalCount = objRs.RecordCount
		end if
		if (Request.Form <> "") then
			if IsEmpty(Request.Form("p")) then
				  iCurrentPage = 1
			elseif IsNumeric(Request.Form("p")) then
				  iCurrentPage = CLng(Request.Form("p"))
			end if
		end if
		strDataTable = "<table width=""100%"" border=""1"" align=""center"" cellpadding=""2"" cellspacing=""0"" style=""Border-Collapse:collapse;word-break:normal;"">"&vbCrLF
		dim thArray,ColCount,k,thStr,i
		    i = 1                             '��ʼ��������
			ColCount = objRs.Fields.Count        '��ȡ������
			for k = 0 to (objRs.Fields.Count-1)
				thStr = thStr & objRs.Fields(k).name&","
			next
			 thArray =  Split((Mid(thStr,1,len(thStr)-1)),",")
		    strDataTable = strDataTable & "<tr bgcolor="""&strTHColor&""">"
				for k=0 to (ColCount-1)
				   strDataTable = strDataTable & "<th>"&thArray(k)&"</th>"
				next
			 strDataTable =  strDataTable & "</tr>"&vbCrLF

		if Clng(iTotalCount) <> 0 then
		   '---------����ѭ����ʼ------------'
		   if (not blnPageData) then '���Ѿ���ҳ����
			   objRs.PageSize		= iPageSize
			   objRs.AbsolutePage	= iCurrentPage
		   end if
		   while (not objRs.eof and i<=iPageSize)
				 if (i mod 2 =0 ) then
					 strDataTable = strDataTable & "<tr bgcolor="""&strTRColor&""">"
				  else
					 strDataTable = strDataTable & "<tr bgcolor="""&strTRColor2&""">"
				 end if
				 for k=0 to (ColCount-1)
					strDataTable = strDataTable & "<td>"&objRs(k)&"</td>"
				 next
				 strDataTable = strDataTable & "</tr>"
			 i=i+1
			 objRs.movenext
		  wend
			  objRs.close()
		  set objRs = nothing
		  '----------����ѭ������-----------'
		  strDataTable = strDataTable &vbCrLF& "<tr bgcolor="""&strPagerColor&"""><td colspan="""&(ColCount+1)&""" align=""left"" height=""22"" valign=""middle"">"&FormPostPager()&"</td></tr>"&vbCrLF
		else
		   strDataTable = strDataTable &vbCrLF& "<tr bgcolor="""&strPagerColor&"""><td colspan="""&(ColCount+1)&""" align=""center"" height=""120"" valign=""middle"">û�з���Ҫ������</td></tr>"&vbCrLF
		end if
		   strDataTable = strDataTable & "</table>"
		DataTabAuto = strDataTable
	 End Function
	 
	 '----------------Common String Operation Added : 2005-5-25
	 Public Function str_Replace(StrContent,PatternStr,reText)	
		Dim objRe
		Set objRe=New RegExp
		objRe.Pattern=PatternStr
		objRe.Global=True
		objRe.IgnoreCase=True
	   'objRe.MultiLine=True
		str_Replace=objRe.Replace(StrContent,reText)
		Set objRe=nothing
	 End Function

	'----------------Common Logical Operation Added : 2005-5-25
	Public Function IsEmptyStr(strChk)
		if IsNull(strChk) then
		   IsEmptyStr = true
		   Exit Function
		else
		   IsEmptyStr = (Trim(CStr(strChk))="")
		end if
	End Function

	Public Function IsNumber(str)
	   if not IsEmptyStr(str) then
		   isNumber = isNumeric(str)
	   else
		   isNumber = false
	   end if
	End Function

	'*****************************************
	'�Ƿ���Post����
	'*****************************************
	Public Function IsPostBack()
	   IsPostBack = (UCase(Request.ServerVariables("REQUEST_METHOD")) = "POST")
	End Function

	Public Function getValue(blnJudge,yesValue,noValue)
		if (blnJudge = true) then
			getValue = yesValue
		else
			getValue = noValue
		end if
	End Function

	 Private function FormPostPager()
	    '''''''''''''''''''''''''''''''''''
		'Random form id,for multi formPager in one Page
		'2005��5��24�� ���ڶ� [R.W.]
		'''''''''''''''''''''''''''''''''''
		Dim frmId
		    if (strPagerID<>Empty) then
			    frmId = strPagerID
			else
				Randomize timer
				frmId = "Pager$" & cint(8999*Rnd+1000)
			end if
		'''''''''''''''''''''''''''''''
		dim JSGoFunction
		JSGoFunction = "<script language=""javascript"">"&_
		"function PostPager(objFrm,n){var obj = document.getElementById(objFrm);obj.p.value = n;obj.pagerCurrent.value = n;obj.submit();}</script>"
		'''''''''''''''''''''''''''''''''''''''''''''
		dim pstr,jumpstr,totalpage
		dim prePage,nextPage
			jumpstr = "<input type='text' name='p' style='width:30px;hight:12px' value='"&iCurrentPage&"' class='entxt' onkeydown=""if(event.keyCode==13){if(doCheck(this)){event.returnValue=false;PostPager('"&frmId&"',this.value);}else{event.returnValue=false;}}"" >"
			if (iTotalCount mod iPageSize > 0) then
			   totalpage = Fix(iTotalCount/iPageSize) + 1
			else
			   totalpage = iTotalCount/iPageSize
			end if
			if (iCurrentPage>totalpage) then iCurrentPage=totalpage
			if (iCurrentPage<1) then iCurrentPage = 1

		   if (iCurrentPage=1) then
			  prePage = "��һҳ"
		   else
			  prePage = "<a href=""javascript:PostPager('"&frmId&"'," &(iCurrentPage-1)& ");"">��һҳ</a>"
		   end if

		   if (iCurrentPage = totalpage) then
			  nextPage = "��һҳ"
		   else
			  nextPage = "<a href=""javascript:PostPager('"&frmId&"'," &(iCurrentPage+1)& ");"">��һҳ</a>"
		   end if
		   pstr = pstr & "<form name="""&frmId&""" id="""&frmId&""" method=""post"" style=""padding:0px;margin:0px;"">"
		   pstr = pstr & "<style type=""text/css"">.entxt  {font-size:10px;font-family:'verdana'}</style>"&JSGoFunction &"<script language=""Javascript"">function doCheck(el){var r=new RegExp(""^\\s*(\\d+)\\s*$"");if(r.test(el.value)){if(RegExp.$1<1||RegExp.$1>"&totalpage&"){alert(""ҳ��������Χ��"");"&frmId&".p.select();return false;}return true;}alert(""ҳ������Ч��"");"&frmId&".p.select();return false;}</script>"
		   FormPostPager = pstr & "�� <span class='entxt'>"&iTotalCount&"</span> �� ÿҳ <span class='entxt'>"&iPageSize&"</span> �� ��ǰ <span class='entxt'><font color=red class='entxt'>"&iCurrentPage&"</font>/"&totalpage&"</span> ҳ <a href=""javascript:PostPager('"&frmId&"',1);"">��ҳ</a> "&prePage&" "& nextPage &" <a href=""javascript:PostPager('"&frmId&"',"&totalpage&");"">βҳ</a>  ����"&jumpstr&"ҳ<input type=""hidden"" value="""&iTotalCount&""" name=""pagerTotal""><input type=""hidden"" value="""&iCurrentPage&""" name=""pagerCurrent""><input type=""hidden"" value="""&frmId&""" name=""PagerID"">"&strFormItemHtml&"</form>"
	end function

	'***********************************
	'2005��5��8�� ������ need to be finished
	'************************************
	Private function DataTabModule(ByRef objRs)
	    if (Request.Form("pagerTotal") <> "") then
		   iTotalCount = CLng(Request.Form("PagerTotal"))
		else
		   if (not blnPageData) then iTotalCount = objRs.RecordCount
		end if
		if (Request.Form <> "") then
			if IsEmpty(Request.Form("p")) then
				  iCurrentPage = 1
			elseif IsNumeric(Request.Form("p")) then
				  iCurrentPage = CLng(Request.Form("p"))
			end if
			if (Request.Form("PagerID")<>"") then
			     strPagerID = Trim(Request.Form("PagerID"))
		    end if
		end if

	    Dim RsArray
		    'strCycleCont = strTemplet 'ѭ����־���ݿ�
		 if (not objRs.Eof) then
		     if (not blnPageData) then objRs.Move (iCurrentPage-1)*iPageSize
			 RsArray = objRs.GetRows(iPageSize,0) 'array = recordset.GetRows(Rows, Start, Fields )
			 strCycleData = tpt_Cycle(RsArray,objRepArray,strCycleCont)
			 strPager = FormPostPager()
			 'strTemplet = Replace(strTemplet,strCycleCont,strCycleData)
			 'strTemplet = Replace(strTemplet,"{$pager}",FormPostPager())
			 'DataTabModule = strTemplet
		 else
		     strPager = "���ݿ��л�û������"
			 DataTabModule = "���ݿ��л�û������"
		 end if
	End Function

	'********************************
	'����ģ���еļ�¼������ 2005��7��4�� ����һ
	'RsRowIdx --> ��¼���е�������
	'RsColIdx --> ��¼���е�������
	'***********************************
	Private Function getLegendData(ByVal MixedSynatax,ByRef RsArray,ByVal RsRowIdx,ByVal RsColIdx)
		  Dim MixedSynataxValue,iRsCount
		  Dim iStart,iMove,iStep,strIdx
		  iStart = 1
		  iRsCount = UBound(RsArray,2)				'GetRows�������ɵ�2ά����

		  if (InStr(iStart,MixedSynatax,"$",1)>0) then
			  
			  '��1�����滻����
			  while (InStr(iStart,MixedSynatax,"$",1)>0)	     
				 iStart = InStr(iStart,MixedSynatax,"$",1)
				 strIdx = Mid(MixedSynatax,iStart+1,1)
				 if IsNumeric(strIdx) then
					 iStep = 0 : iMove = ""
					 while isNumeric(strIdx)
						iMove = iMove & strIdx
						iStart = iStart + 1
						iStep = iStep + 1
						strIdx = Mid(MixedSynatax,iStart+1,1)
					 wend
					 iMove = CInt(iMove)
					 if (iMove<=iRsCount and iMove>=0) then
					     if IsNull(RsArray(iMove,RsColIdx)) then
						   MixedSynatax = Replace(MixedSynatax,"$"&iMove,"")
						 else
						   MixedSynatax = Replace(MixedSynatax,"$"&iMove,RsArray(iMove,RsColIdx))
						 end if
					 end if
				  else
					 iStart = iStart + 1
				  end if
				  iStart = iStart + iStep
			  wend
			  MixedSynatax = Replace(MixedSynatax,"\$","\/")
			  if IsNull(RsArray(RsRowIdx,RsColIdx)) then
				 MixedSynatax = Replace(MixedSynatax,"$","")
				else
				 MixedSynatax = Replace(MixedSynatax,"$",RsArray(RsRowIdx,RsColIdx))
			  end if
			  MixedSynatax = Replace(MixedSynatax,"\/","$")
			  
			  '��2�����к�����������Ӧ����
			  iStart = 1
			  while (InStr(iStart,MixedSynatax,"{FUN:",1)>0)
			    Dim strTemp
				iStart = InStr(iStart,MixedSynatax,"{FUN:",1)
				strIdx = InStr(iStart,MixedSynatax,"}",1)
				strTemp = Mid(MixedSynatax,iStart,(strIdx-iStart+1))
				strTemp = Replace(strTemp,"{FUN:","")
				strTemp = Replace(strTemp,"}","")
				if len(strTemp)>1 then strTemp = Eval(CStr(strTemp))
				MixedSynatax = Mid(MixedSynatax,1,iStart-1) & strTemp & Mid(MixedSynatax,strIdx+1)
				iStart = InStr(1,MixedSynatax,strTemp,1) + len(strTemp)
			  wend

			  MixedSynataxValue = MixedSynatax
			else
			  MixedSynataxValue = MixedSynatax
		   end if

		   getLegendData = MixedSynataxValue
	End Function
	

	Private Function tpt_Cycle(ByVal RsArray,ByVal ReplaceArray,ByVal CycleCont)
	        if (not iBDebug) then On Error Resume Next
			if (Not IsArray(RsArray)) or (Not IsArray(ReplaceArray)) then
			    iBError("û�����ݼ��ϻ��滻���ݲ���һ������:iBDataTable.tpt_Cycle()")
				Exit Function
			End if
			dim i,k,v,RsCount,RsRowIdx,RpCount,strFinal
			dim MidStr,RetStrings,rCycleCont,tempCycle
				RsCount = UBound(RsArray,2)
				RpCount = UBound(ReplaceArray)

				for i=0 to RsCount

				    if IsArray(CycleCont) then
					     k = UBound(CycleCont)
						 v = (i+1) MOD (k+1)
						 if v = 0 then
							tempCycle = CycleCont(k)
						 else
							tempCycle = CycleCont(v-1)
						 end if
					  else
					     tempCycle = CycleCont
				    end if

					''''''''''''''''�õ�ǰ�����滻ģ������
					for k=0 to RpCount
					      RsRowIdx  = CInt(ReplaceArray(k,1))		'������
						  if (RsRowIdx <>-1) then					'Update for Version 1.2
						     MidStr = RsArray(CInt(ReplaceArray(k,1)),i)
						  else
						     MidStr = ""						'ɾ���ñ���
						  end if
						  if IsNull(MidStr) then MidStr =""
						  if k=0 then rCycleCont = tempCycle

					   if len(ReplaceArray(k,2))<1 then		'�������滻
						  rCycleCont = Replace(rCycleCont,ReplaceArray(k,0),MidStr)
					   else
						  '�������������ı�����ֵ {FUN:YourFunction($1,$2,$3,$,$10)}
						  '-------------------------------------------
						  strFinal = getLegendData(ReplaceArray(k,2),RsArray,RsRowIdx,i)
						  rCycleCont = Replace(rCycleCont,ReplaceArray(k,0),strFinal)
					   end if
					next
					''''''''''''''''''''''''''''''''''''''''
					RetStrings = RetStrings & rCycleCont
				next

			tpt_Cycle = RetStrings
			if Err.Number<>0 then
		        iBError("����ѭ�������ݳ���:iBDataTable.tpt_Cycle()")
			    Err.Clear()
		    end if
	End Function

End Class



'******************************************
Class iBFileIO

	 Private objFSO,objStream

	 Private Sub Class_Initialize()
	    Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		Set objStream = Server.CreateObject("ADODB.Stream")
	 End Sub
	 
	 Private Sub Class_Terminate()
		Set objFSO = nothing
		Set objStream = nothing
	 End Sub


	Public Function FileDelete(ByVal strFilePath)
	  On Error Resume Next
	    objFSO.DeleteFile strFilePath,true
	  if (Err.Number<>0) then
		 Err.Clear
		 FileDelete = false
	  else
	     FileDelete = true
	  end if
	End Function


	Public Function FolderDelete(ByVal strFolderPath)
	  On Error Resume Next
	  objFSO.DeleteFolder strFolderPath,true
	  if (Err.Number<>0) then
		 Err.Clear
		 FolderDelete = false
	  else
	     FolderDelete = true
	  end if
	End Function


	Public Function GetFileContent(ByVal strFilePath,ByVal bUnicode)
		Dim txtStream,format
		if (bUnicode=true) then
		   format = -1 '��Unicode ��ʽ���ļ�,0=��ASCII ��ʽ���ļ�
		else
		   format = -2 'ʹ�����Լ������������е�Ĭ��ֵ��
		end if
		Set txtStream = objFSO.OpenTextFile(strFilePath,1,false,format)
		GetFileContent = txtStream.ReadAll
		Set txtStream = nothing
	End Function

	
	Public Function SaveFileContent(ByVal strFilePath,ByVal strFileContent,ByVal bUnicode)
	    Dim txtStream,format
		On Error Resume Next
		if (bUnicode=true) then
		   format = -1 '��Unicode ��ʽ���ļ�,0=��ASCII ��ʽ���ļ�
		else
		   format = -2 'ʹ�����Լ������������е�Ĭ��ֵ��
		end if
		Set txtStream = objFSO.OpenTextFile(strFilePath,2,true,format)
		    txtStream.Write(strFileContent)
		Set txtStream = nothing

	   if (Err.Number<>0) then
		  Err.Clear
		  SaveFileContent = false
	   else
		  SaveFileContent = true
	   end if
	End Function

	Public Function CreateUTF8(strFileContent,strFilePath)
	    On Error Resume Next
		objStream.Type=2
		objStream.Mode=3
		objStream.Charset="utf-8"
		objStream.Open()
		objStream.WriteText strFileContent
		objStream.SaveToFile strFilePath,2
		objStream.Close()
	   if (Err.Number<>0) then
		  Err.Clear
		  CreateUTF8 = false
	   else
		  CreateUTF8 = true
	   end if
	End Function

	Public Function StringConvert(ByRef strContent,ByVal oldCharSet,ByVal newCharset)
		objStream.Type=2
		objStream.Mode=0
		objStream.Open()
		objStream.Charset = newCharset
		objStream.WriteText strContent
		objStream.Position = 0
		objStream.Type = 2
		objStream.Charset = oldCharSet
		StringConvert = objStream.ReadText()
		objStream.Close()
	End Function

	Public Function GetFileAccessInfo(ByVal strFilePath)
	   If FileExists(strFilePath) Then
		   Dim file
		   Set File = objFSO.GetFile(strFilePath)
			  GetFileAccessInfo = Array(File.DateCreated,File.DateLastModified,File.DateLastAccessed)
		   Set File = nothing
	   Else
		    GetFileAccessInfo = Array("�ļ�������","�ļ�������","�ļ�������")
		End If
	End Function


	Public Function GetFileSize(ByVal strFilePath)
	    If FileExists(strFilePath) Then
			Dim file
			Set File = objFSO.GetFile(strFilePath)
				GetFileSize = formatSize(File.Size)
			Set File = nothing
		Else
		    GetFileSize = "�ļ�������"
		End If
	End Function


	Public Function GetFolderSize(ByVal strFolderPath)
		Dim Folder
		Set Folder = objFSO.GetFolder(strFolderPath)
		    GetFolderSize = formatSize(Folder.Size)
		Set Folder = nothing
	End Function

	Public Function FileExists(ByVal strFilePath)
		FileExists = objFSO.FileExists(strFilePath)
	End Function

	Public Function FolderExists(ByVal strFolderPath)
	    FolderExists = objFSO.FolderExists(strFolderPath)
    End Function


	Public Function CreateFolder(ByVal strFolderPath)
		Dim tPath,i,fPath
		tPath=Server.Mappath("/")
		fPath=strFolderPath
		If Right(tPath,1)="\" Then tPath=left(tPath,Len(tPath)-1) 'ȥ������"\"
		If Right(fPath,1)="\" Then Path=left(fPath,Len(Path)-1)   'ȥ������"\"
		if Instr(fPath,tpath)=0 then
			CreateFolder=False
			Exit Function '�������web���Է��ʵ�·��������False
		end if

		fPath=split(Replace(fPath,tPath,""),"\")
		for i=1 to Ubound(fPath)
			tPath=tPath&"\"&fPath(i)
			If not objFSO.FolderExists(tPath) Then
				objFSO.CreateFolder(tPath)
			End If
		next
		CreateFolder=True
	End Function

	'*********************************
	'��ʾ�ļ���С
	'***********************************
	Public Function formatSize(sFileSize)
	   if IsNumeric(sFileSize) then
		  Dim iSize
		  iSize = CDbl(sFileSize)
		  if iSize>1099511627776 then
			 formatSize = FormatNumber((iSize/1099511627776),2) & " TB"
		  elseif iSize>1073741824 then
			 formatSize = FormatNumber((iSize/1073741824),2) & " GB"
		  elseif iSize>1048576 then
			 formatSize = FormatNumber(iSize/1048576,2) & " MB"
		  elseif iSize>1024 then
			 formatSize = FormatNumber(iSize/1024,2) & " KB"
		  else
			 formatSize = sFileSize & " bytes"
		  end if
	   else
		  formatSize = sFileSize
	   end if
	End Function

	'/**
	'2007��12��10�� ���Ӷϵ������ͳ���4M�ļ���С�ķ���
	'**/
	Public Sub WriteFile(ByVal strFilePath,ByVal strSaveName)

			Dim objStream,lTotalSize
			Dim bytes,strRange
			strRange = Request.ServerVariables("HTTP_RANGE")
			Set objStream = CreateObject("ADODB.Stream")
				objStream.Open
				objStream.Type = 1
				objStream.LoadFromFile strFilePath
				lTotalSize = objStream.Size
				If strRange = "" Then

					'the Default ASP maximum buffer size is 4MB
					If lTotalSize > 4194304 Then
						objStream.Position = 0
						Response.AddHeader "Content-Length", lTotalSize
						Response.ContentType = "application/octet-stream"
						Response.AddHeader "Content-Disposition", "attachment;filename="&strSaveName
						 Do while (Not objStream.EOS And Response.IsClientConnected)
							Response.Binarywrite objStream.Read(20480)
							Response.Flush
						 Loop
					 Else
						bytes = objStream.Read
					 End If

				Else

					Dim strReqRange,idxB,objRange
					Dim idxPosition,lRangLen
						idxPosition = 0 : lRangLen = 0
						strReqRange = LCase(strRange)

					idxB = InStr(1,strReqRange,"=",1)
					If idxB > 0 Then
						'Get the range str of bytes
						strReqRange = Right(strReqRange, Len(strReqRange)-idxB)
						'Not Support serveral range
						If InStr(1,strReqRange,",",1) <= 0 Then
							objRange = Split(strReqRange,"-")
							
							'	Rangeͷ�� - "HTTP_RANGE"
							'	-----------
							'	Rangeͷ���������ʵ���һ�����߶���ӷ�Χ�����磬 
							'	��ʾͷ500���ֽڣ�bytes=0-499 
							'	��ʾ�ڶ���500�ֽڣ�bytes=500-999 
							'	��ʾ���500���ֽڣ�bytes=-500 
							'	��ʾ500�ֽ��Ժ�ķ�Χ��bytes=500- 
							'	��һ�������һ���ֽڣ�bytes=0-0,-1 
							'	ͬʱָ��������Χ��bytes=500-600,601-999
							'   ------------------------------------------------
							'bytes=0-499; bytes=500-999; bytes=-500; bytes=500-
							If UBound(objRange) = 1 Then
								'bytes=-500
								If objRange(0) = "" Then
									lRangLen = CLng(objRange(0))
									idxPosition = lTotalSize - lRangLen
								End If

								If objRange(0) <> "" And objRange(1) <> "" Then
									lRangLen = CLng(objRange(1)) - CLng(objRange(0)) + 1
									idxPosition = CLng(objRange(0))
								Else
									'bytes=500-
									lRangLen = lTotalSize - CLng(objRange(0))
									idxPosition = CLng(objRange(0))
								End If
							End If
						End if
					End If
			
					objStream.Position = idxPosition
					If lRangLen < 4194304 Then
						bytes = objStream.Read(lRangLen)
					Else
						Response.Status="206 Partial Content"
						Response.AddHeader "Content-Length", lRangLen
						Response.ContentType = "application/octet-stream"
						Response.AddHeader "Content-Range", "bytes "&CStr(idxPosition)&"-"&CStr(idxPosition+lRangLen-1)&"/"&lTotalSize
						Response.AddHeader "Content-Disposition", "attachment;filename="&strSaveName
						 Do while (Not objStream.EOS And Response.IsClientConnected)
							Response.Binarywrite objStream.Read(20480)
							Response.Flush
						 Loop
						Response.End
					End If

				End If

				objStream.Close
			Set objStream = Nothing

			'--Partial Content for Bytes
			If VarType(bytes) = 8209 Then
					Response.Clear()
					If lTotalSize > LenB(bytes) Then
						Response.Status="206 Partial Content"
					End If
					Response.AddHeader "Content-Length", LenB(bytes)
					Response.ContentType = "application/octet-stream"
					Response.AddHeader "Content-Disposition", "attachment;filename="&strSaveName
					Response.BinaryWrite(bytes)
					response.Flush()
					Response.End
			End If

	End Sub

End Class


'*********************************************
'CHM�ļ���HHC�ļ����ݽ�����
'version: 1.0 
'2005��9��29�� ������ 22:56:52
'****************************************************************
Class HHCParse
	
	Public TOCURI
	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

	'*********************************
	'��ȡHHC�ļ���Body����������
	'2005��9��29�� ������ 22:12:17 st:OK
	'--------------------------------------------------
	Function GetHHCBodyContent(ByRef strHHContent)
		Dim oRe, oMatch, oMatches, strRet
		Set oRe = New RegExp
			oRe.Pattern = "(\<body\s*[^\>]*\>)([^\0]*)(<\/body>)"
			oRe.IgnoreCase = True
			oRe.Global = True
		Set oMatches = oRe.Execute(strHHContent)
		Set oMatch = oMatches(0)
			strRet = oMatch.SubMatches(1)
		Set oMatch = Nothing
		Set oMatches = Nothing
		Set oRe = Nothing
		GetHHCBodyContent = strRet
	End Function

	'*********************************
	'ת��HHObjectΪ��Ӧ������HTML����
	'2005��9��29�� 15:20:32 st:OK
	'sup:ParamToAnchor() + FindNextTag() + FindNextCharPos()
	'--------------------------------------------------
	Function HHCObjecToLink(ByRef strURIBase,ByRef strHHContent)
		Dim RegEx, Matches, Match, strHHContentResult
		Dim retHTML, strTag, strAnchor, strParam
		TOCURI = strURIBase
		'--- to Executable Content (2005��10��8�� 15:22:12)
		strHHContent = ReplaceText(strHHContent,"(\r\n\s*)","")
		strHHContent = ReplaceText(strHHContent,"(\<object\s[^\>]*\>)",vbCrLf&"$1")
		'--------------------------------------------------------
		strHHContentResult = strHHContent		'A Copy of strHHContent

		Set RegEx = New RegExp
			'(\r\n\s*) -> (\<object\s[^\>]*\>)(.+)(\<\/object\>)
			'RegEx.Pattern = "(\<object\s[^\>]*\>)(\r\n\s+.+){0,3}(\r\n\s+)(\<\/object\>)"
			RegEx.Pattern = "(\<object\s[^\>]*\>)(.+)(\<\/object\>)"
			RegEx.IgnoreCase = True
			RegEx.Global = True

			Set Matches = RegEx.Execute(strHHContent)   ' ִ��������
			  
			  For Each Match in Matches					' �� Matches ���Ͻ��е�����
				strTag = FindNextTag(Match.FirstIndex + 1 + Match.Length,strHHContent)

				'strParam = Replace(Match.Value,Match.SubMatches(0),"",1,-1,1)
				'strParam = Replace(strParam,"</object>","",1,-1,1)
				strParam = Match.SubMatches(1)

				If (Left(strTag,4) = "<ul " Or Left(strTag,4) = "<ul>") Then
					'alert "������Ŀ" & vbCRlf & Match.Value
					strAnchor = ParamToAnchor(strParam,1,True)
				Else
					strAnchor = ParamToAnchor(strParam,11,false)		
				End If
				strHHContentResult = Replace(strHHContentResult,Match.Value,strAnchor,1,-1,1)
			  Next

			Set Matches = Nothing

		Set RegEx = Nothing
		
		Dim x, y
		x = FindNextCharPos(strHHContentResult,1,"<li>")
		y = FindNextCharPos(strHHContentResult,1,"<ul>")
		If (x > 0 And x > y) Then
			strHHContentResult = Left(strHHContentResult,y+4) & Replace(strHHContentResult,"<ul>","<ul class=""folder"">",y+5,-1,1)
		Else
			strHHContentResult = Replace(strHHContentResult,"<ul>","<ul class=""folder"">",1,-1,1)
		End If

		'---------------- to Readable HTML (2005��10��8�� 15:22:12)
		strHHContentResult = ReplaceText(strHHContentResult,"(\r\n\s*)","")
		strHHContentResult = ReplaceText(strHHContentResult,"(\<ul\s*[^\>]*\>)",vbCrLf&"$1")
		strHHContentResult = ReplaceText(strHHContentResult,"(\<li\s*[^\>]*\>)",vbCrLf&"$1")
		'-----------------------------------------------------------
		HHCObjecToLink = strHHContentResult
	End Function

	'***************************************
	'������һ��������Tag������
	'2005��9��29�� st:ok
	'----------------------------------------------------------
	Public Function FindNextTag(ByVal iPos,ByRef strHTML)
		Dim idx,iEnd,retTag
			retTag = ""
			idx = InStr(iPos,strHTML,"<",1)
		If idx > 0 Then
			iEnd = InStr(idx+1,strHTML,">",1)
			If (iEnd > 0) Then
				retTag = LCase(Mid(strHTML,idx,iEnd-idx+1))
			End  If
		End If
		FindNextTag = retTag
	End Function

	'******************************
	'ת��HHC Param Ϊ Html Anchor���
	'2005��9��29�� 11:10:02
	'sup:GetParamAttribute() st:OK
	'-----------------------------------------------------
	Public Function	ParamToAnchor(ByRef HHCParamString,ByVal repIdx,ByVal blnFolder)
		Dim RegEx,Matches,Match,Book
		Dim idx,strParam,strParamName,strRootURI
		Dim AnchorTemplet,AnchorTempletFolder
		AnchorTemplet = "<img src=""images/icon/{$ImageNumber}.gif"" border=""0""> <a href=""{$Local}"">{$Name}</a>"
		AnchorTempletFolder = "<span class=""box"" onclick=""iBTree(this,this.parentNode);""><img src=""images/plus.gif"" border=""0""> <img src=""images/folder.gif"" border=""0""></span> <a href=""{$Local}"">{$Name}</a>"
		'----�л��Ƿ�������б�  Added: 2005��9��29�� 16:49:13
		If (blnFolder = True) Then AnchorTemplet = AnchorTempletFolder

		Set RegEx = New RegExp
			RegEx.Pattern = "(\<param\s)([^\>]+)(\>)"
			RegEx.IgnoreCase = True
			RegEx.Global = True

		Set Matches = RegEx.Execute(HHCParamString)   ' ִ��������
			  For Each Match in Matches      ' �� Matches ���Ͻ��е�����
				strParam = Match.Value
				strParamName = GetParamAttribute(strParam,"Name")
				Select Case LCase(strParamName)
					Case "name"
						AnchorTemplet = Replace(AnchorTemplet,"{$Name}",GetParamAttribute(strParam,"value"))
					Case "local"
						Set Book = New iBook
							strRootURI = Book.GetRootURI(Book.GetBaseURI(TOCURI),GetParamAttribute(strParam,"value"))
						Set Book = Nothing
						AnchorTemplet = Replace(AnchorTemplet,"{$Local}",strRootURI)
					Case "imagenumber"
						idx = GetParamAttribute(strParam,"value")
						If (idx>=1 And idx<=42) Then
							AnchorTemplet = Replace(AnchorTemplet,"{$ImageNumber}",idx)
						End If
				End Select
			  Next
		Set Matches = Nothing
		Set RegEx = Nothing	

		 '--------------> Set default Icon
		 If InStr(1,AnchorTemplet,"{$ImageNumber}",1)>0 Then
			AnchorTemplet  = Replace(AnchorTemplet,"{$ImageNumber}",repIdx)
		 End If
		 '--------------> Set Empty Anchor
		 If InStr(1,AnchorTemplet,"{$Local}",1)>0 Then
			AnchorTemplet  = Replace(AnchorTemplet,"{$Local}","###"" target=""_self")
		 End If
		 '--------------> Set Match Empty
		 If InStr(1,AnchorTemplet,"{$Name}",1)>0 Then
			AnchorTemplet  = ""
		 End If
		ParamToAnchor = AnchorTemplet
	End Function

	'********************************
	'��ȡParamĳ�����Ե�ֵ 2005��9��29��
	'sup:FindNextNbChar() st:OK
	'-------------------------------
	'ParamString = "<param name=""Name"" value=""Index Z"">"
	'GetParamAttribute(ParamString,"name") = "Name"
	'GetParamAttribute(ParamString,"value") = "Index Z"
	Public Function GetParamAttribute(ByRef ParamString,ByVal AtrrName)
		Dim retValue,idx,strChar,iEnd,blnDirectFlag
		blnDirectFlag = False 'ֱ�����Ա��
		idx = InStr(1,ParamString,AtrrName,1)
		If (idx>0) Then
			idx = idx+Len(AtrrName)
			strChar = FindNextNbChar(ParamString,idx)
			'--�����������
			If strChar = "=" Then
				idx = idx + 1	'λ�Ƶ�"="��λ��
				strChar = Mid(ParamString,idx,1)	'��λ����һ�ַ�
				'�����������
				If (strChar = "'" Or strChar = Chr(34)) Then
					iEnd = FindNextCharPos(ParamString,idx+1,strChar)
				'�ո��������
				ElseIf (strChar = Chr(32)) Then
					iEnd = FindNextCharPos(ParamString,idx+1,strChar)
					If (iEnd = 0) Then
						iEnd = FindNextCharPos(ParamString,idx+1,Chr(62))
					End If
				'ֱ���������
				Else
					'�Կո����
					iEnd = FindNextCharPos(ParamString,idx+1,Chr(32))
					'��>����
					If (iEnd = 0) Then
						iEnd = FindNextCharPos(ParamString,idx+1,Chr(62))
					End If
					blnDirectFlag  =  True	'ֱ�����Ա�Ǵ�
				End If
				If (iEnd > idx) Then
					'���÷���ֵ
					If (blnDirectFlag = True) Then
						retValue = Mid(ParamString,idx,iEnd-idx)
					Else
						retValue = Mid(ParamString,idx+1,iEnd-idx-1)
					End If
				Else
					retValue = ""
				End If
			Else
				retValue = ""
			End If
		Else
			retValue = ""
		End If
		GetParamAttribute = retValue
	End Function

	'***************************
	'����һ������ʾ�ַ�,��ո�\�س�\����\�Ʊ���ŵ�
	'2005��9��29�� 11:26:01 st:ok
	'--------------------------------------
	Public Function FindNextNbChar(ByRef strSearch,ByRef iPos)
		Dim strChar,k
		 strChar = Mid(strSearch,iPos,1)
		 If Len(strChar) = 1 Then
			k = Asc(strChar)
			If (k=32 Or k=13 Or k=10 Or k=9) Then
				strChar = FindNextNbChar(strSearch,iPos+1)
			End If
		 End If
		FindNextNbChar = strChar
	End Function

	'*****************************8
	'��strSearch�д�iPos��ʼ�ı�������һstrChar��λ�ã���������Ϊ0��
	'2005��9��29�� 11:29:24 st:ok
	'-------------------------------------
	Public Function FindNextCharPos(ByRef strSearch,ByRef iPos,ByRef strChar)
		Dim idx
		idx = InStr(iPos,strSearch,strChar,1)
		If idx > 0 Then
			FindNextCharPos = idx
		Else
			FindNextCharPos = 0
		End If
	End Function

End Class
%>