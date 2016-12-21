<!--#include file="Const.asp"-->
<!--#include file="iBCore.asp"-->
<%
Dim Rs,Conn,sql,strMsg
Call DBOpen(Conn,ConnectionString)
	
'--See also Rights.asp
Sub SetManageRights(name)
	Session("Manage") = name
End Sub

Sub ClearManageRights()
	Session("Manage") = Empty
End Sub

If Request.QueryString("s") <> "p" Then
	If Session("Manage") = "" Then
	    Response.Redirect("?s=p")
	End If
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta http-equiv="Content-Language" content="zh-cn" />
<title>&#105;Reader ϵͳ����</title>
<link rel="stylesheet" href="Reader.css" type="text/css" />
<style type="text/css">
<!--
.style1 {color: #FFFFFF;font-weight:bold;}
.style2 {color: #FF0000}
a {text-decoration:none;}
a:link {text-decoration:none;}
a:hover {text-decoration:underline;color:#990000;}
-->
</style>
</head>
<body bgcolor="#666666">
<%
Select Case Request.QueryString("s")
	Case "l"
		Call ShowMenu(1)
		Call ReaderBookList()
	Case "v"
		Call ShowMenu(2)
		Call BookVoteList()
	Case "h"
		Call ShowMenu(5)
		Call BookHelpList()
	Case "s"
		Call ShowMenu(4)
		If Not IsEmptyStr(Request.QueryString("ISBN")) Then
			Call BookDataSet(Request.QueryString("ISBN"))
		Else
			Call BookDataSet(0)
		End If
	Case "x"
		Call ClearManageRights()
		ShowMsg "�ɹ��˳�����״̬��",5
		Call iReaderPass()
	Case Else
		Call iReaderPass()
End Select


'********************************
'�鼮��������
'**************************************************
Sub BookDataSet(ISBN)
	Dim BookName,publishdata,bookPath,mainPage,tocPage,sysToc,onlineDate
	Dim onService,DownHits, TotalHits, ysVstoday, lastweekHits, thisweekHits, onlineHits
	Dim blnUpdate, BookID, coverpic, PressName
		blnUpdate = False
		onlineHits = 0
		onlineDate = Now()
		BookId = 0

	If Request.QueryString("Goto") = "Preview" Then
		sql = "select top 1 * from [iReaderBooks] where BookID < (select top 1 BookID from [iReaderBooks]  where ISBN = '"&checkStr(ISBN)&"') order by BookID desc"
	ElseIf Request.QueryString("Goto") = "Next" Then
		sql = "select top 1 * from [iReaderBooks] where BookID > (select top 1 BookID from [iReaderBooks]  where ISBN = '"&checkStr(ISBN)&"') order by BookID Asc"
	Else
		sql = "select top 1 * from [iReaderBooks] where ISBN='"&checkStr(ISBN)&"'"
	End If

	'showMsg sql,100
	
	Set Rs = Server.CreateObject("ADODB.Recordset")
	Rs.Open sql,Conn,3,3

	If Not Rs.Eof Then
			If IsPostBack() Then
				blnUpdate = True
			Else
				ISBN = Rs("ISBN")
				BookName = Rs("BookName")
				coverpic = Rs("coverPath")
				publishdata = Rs("PublishData")
				bookPath = Rs("BookPath")
				mainPage = Rs("MainURI")
				tocPage = Rs("TOContent")
				sysToc = Rs("TOCSYS")
				onlineDate = Rs("OnlineDate")
				onService = Rs("STOK")
				DownHits = Rs("DownHits")
				TotalHits = Rs("Hits")
				ysVstoday = Rs("YesterdayHits") & "/" & Rs("TodayHits")
				lastweekHits = Rs("LastWeekHits")
				thisweekHits = Rs("ThisWeekHits")
				onlineHits = Rs("onlineHits")
				PressName = Rs("PressName")
			End If
		Else
			If (CStr(ISBN) <> "0" ) And Len(Request.QueryString("Goto"))>1 Then
				strMsg = strMsg & "�Ѿ�û�и���������ݵ��������ˣ�<br>"
				ISBN = 0
			ElseIf (CStr(ISBN) <> "0" ) Then
				strMsg = strMsg & "��ISBN�����ݿ��в����ڣ����Ḳ���������ݣ�<br>"
			End If
			ShowMsg strMsg,3

			DownHits = 0
			TotalHits = 0
			ysVstoday = "0/0"
			lastweekHits = 0
			thisweekHits = 0
			onlineHits = 0
	End If

	If IsPostBack() Then

			If (blnUpdate = False) Then 
				Rs.AddNew
				ISBN = Left(Request.Form("ISBN"),10)
				strMsg = strMsg & "�Ѿ����һ�¼�¼��<br>"
			End If

			BookName = Left(Request.Form("bookname"),255)
			Coverpic = Left(Request.Form("Coverpic"),255)
			publishdata = Left(Request.Form("publishdata"),50)
			PressName = Left(Trim(Request.Form("PressName")),50)
			bookPath = Left(Request.Form("bookpath"),255)
			mainPage = Left(Request.Form("mainpage"),150)
			tocPage = Left(Request.Form("tocpage"),150)
			sysToc = Left(Request.Form("systoc"),150)
			onlineDate = GetValue(IsDate(Request.Form("onlineDate")),CDate(Request.Form("onlineDate")),Now())
			onService = GetValue(Request.Form("onService")="1",True,False)
			DownHits = GetValue(IsNumeric(Request.Form("downHits")),CLng(Request.Form("downHits")),0)
			TotalHits = GetValue(IsNumeric(Request.Form("Hits")),CLng(Request.Form("Hits")),0)
			lastweekHits = GetValue(IsNumeric(Request.Form("lastweekhits")),CLng(Request.Form("lastweekhits")),0)
			thisweekHits = GetValue(IsNumeric(Request.Form("thisweekhits")),CLng(Request.Form("thisweekhits")),0)
			onlineHits = 1
			'-------------
			If (blnUpdate = False) Then 
				Rs("ISBN") = ISBN
			End If
			
			Rs("BookName") = BookName
			Rs("PressName") = PressName
			Rs("coverPath") = Coverpic
			Rs("BookPath") = bookPath
			Rs("PublishData") = publishdata
			Rs("OnlineDate") = onlineDate
			If Rs("STOK") <> onService Then
				Rs("STOK") = onService
				If (onService) Then Rs("onlineHits") = Rs("onlineHits") + 1
				onlineHits = Rs("onlineHits")
			End If
			Rs("MainURI") = mainPage
			Rs("TOContent") = tocPage
			Rs("TOCSYS") = sysToc
			Rs("Hits") = TotalHits
			Rs("DownHits") = DownHits	

			ysVstoday = Split(Request.Form("ysVstoday"),"/")
			Rs("TodayHits") = CLng(ysVstoday(0))
			Rs("YesterdayHits") = CLng(ysVstoday(1))
			ysVstoday = Request.Form("ysVstoday")

			Rs("ThisWeekHits") = thisweekHits
			Rs("LastWeekHits") = lastweekHits
			Rs.Update()
			ISBN = Rs("ISBN")

			'-------------------------
			strMsg = strMsg & "�ɹ����ø������ݣ�"
			ShowMsg strMsg,3
		End If

	Rs.Close()
	Set Rs = Nothing
%>
<form method="post" enctype="application/x-www-form-urlencoded">
<table width="95%"  border="1" align="center" cellpadding="5" cellspacing="0" bordercolor="#000000" bgcolor="#CCE7FF">
  <tr align="center" bgcolor="#000000">
    <td height="25" colspan="8" valign="middle"><span class="style1">���������Ķ�������(CHM)����</span></td>
  </tr>
  <tr>
    <td width="0" align="right" bgcolor="#f3f3f3">�鼮����</td>
    <td colspan="7"><input name="bookname" type="text" id="bookname" value="<%=bookname%>" size="70"></td>
  </tr>
  <tr>
    <td width="0" align="right" bgcolor="#f3f3f3">ISBN</td>
    <td colspan="3"><input name="ISBN" type="text" id="ISBN" value="<%=ISBN%>" size="20" maxlength="10" onChange="location.href='?s=s&ISBN='+ISBN.value;">
    (10���ַ�, ����ʹ��IB05093001��ʽ. ) <a href="#" class="style2" onClick="location.href='?s=s&ISBN='+ISBN.value;">����Ƿ����</a></td>
    <td width="0" align="center" bgcolor="#F3F3F3">��������</td>
    <td colspan="3"><input name="publishdata" type="text" id="publishdata" value="<%=publishdata%>" size="35"></td>
  </tr>
  <tr>
    <td width="0" align="right" bgcolor="#f3f3f3">���·��</td>
    <td colspan="3"><input name="bookpath" type="text" id="bookpath" value="<%=bookPath%>" size="35"> <%=GetValue(FileExist(GetValue(InStr(1,bookPath,":",1)>0,bookPath,bookRootPath & bookPath)),"<font color=""green"">(���ļ�����)</font>","<font color=""red"">(���ļ��Ѿ�������)</font>")%></td>
    <td align="center" bgcolor="#FFFFFF">����ʵ��</td>
    <td colspan="3"><input name="PressName" type="text" id="PressName" value="<%=PressName%>" size="28" /></td>
    </tr>
  <tr>
    <td width="0" align="right" bgcolor="#f3f3f3">��ҳ�ļ�</td>
    <td colspan="3"><input name="mainpage" type="text" id="mainpage" value="<%=mainPage%>" size="28"></td>
    <td width="0" align="center" bgcolor="#F3F3F3">Ŀ¼�ļ�</td>
    <td colspan="3"><input name="tocpage" type="text" id="tocpage" value="<%=tocPage%>" size="28"></td>
  </tr>
 <tr>
    <td width="0" align="right" bgcolor="#f3f3f3">ϵͳTOC�ļ�</td>
    <td colspan="7"><input name="systoc" type="text" id="systoc" value="<%=sysToc%>" size="55" /> (HHC�ļ�·��)</td>
  </tr>
   <tr>
    <td align="right" bgcolor="#f3f3f3">����ͼƬ·��</td>
    <td colspan="7"><input name="coverpic" type="text" id="coverpic" value="<%=coverpic%>" size="55" /></td>
  </tr>
  <tr>
    <td width="0" align="right" bgcolor="#f3f3f3">����ʱ��</td>
    <td colspan="3"><input name="onlineDate" type="text" id="onlineDate" value="<%=onlineDate%>" size="20">
    (���� <%=onlineHits%> ��)</td>
    <td width="0" align="center" bgcolor="#F3F3F3">����״̬</td>
    <td width="0"><input name="onService" type="checkbox" id="onService" value="1" <%=isThisValue(onService,"checked")%>>
    ���߷�����</td>
    <td width="0" align="center" bgcolor="#F3F3F3">���ش���</td>
    <td width="0"><input name="downHits" type="text" id="downHits" value="<%=DownHits%>" size="10"></td>
  </tr>
  <tr>
    <td width="0" align="right" bgcolor="#f3f3f3">�������</td>
    <td width="0"><input name="Hits" type="text" id="Hits" value="<%=TotalHits%>" size="20"></td>
    <td width="0" align="center" bgcolor="#F3F3F3">����/���� �����</td>
    <td width="0"><input name="ysVstoday" type="text" id="ysVstoday" value="<%=ysVstoday%>" size="20"></td>
    <td width="0" align="center" bgcolor="#F3F3F3">���������</td>
    <td width="0"><input name="lastweekhits" type="text" id="lastweekhits" value="<%=lastweekHits%>" size="10"></td>
    <td width="0" align="center" nowrap bgcolor="#F3F3F3">���������</td>
    <td width="0"><input name="thisweekhits" type="text" id="thisweekhits" value="<%=thisweekHits%>" size="10"></td>
  </tr>
  <tr>
    <td colspan="8" align="right" bgcolor="#F3F3F3"><a href="?s=s&ISBN=<%=ISBN%>&Goto=Preview" class="style2">��һ��</a>&nbsp;&nbsp;  <a href="?s=s&ISBN=<%=ISBN%>&Goto=Next" class="style2">��һ��</a>&nbsp;&nbsp;</td>
  </tr>
  <tr>
    <td colspan="8" align="center" bgcolor="#F3F3F3"> <input name="btnSet" type="submit" id="btnSet" value="��������">
&nbsp;&nbsp;&nbsp;&nbsp;
<input name="btnReset" type="reset" id="btnReset" value="��������"></td>
  </tr>
</table>
</form>
<%
End Sub

'********************************
'�����б�
'**************************************************
Sub ReaderBookList()
	Dim iBDb,Rs,strMsg,IO
	Set iBDb = new iBDataBase
		'iBDb.ConnString = ConnectionString
		Set iBDb.ConnObject = Conn
		Conn.BeginTrans()
		
	Dim DataTab,Pager,strCycleData
    Dim iPageSize,CurPage,CycleCont,itemCycle 
	
	iPageSize = 20 : CurPage = 1
	if (Request.Form <> "") then
	   if IsEmpty(Request.Form("p")) then
		  CurPage = 1
		elseif IsNumeric(Request.Form("p")) then
		  CurPage = CLng(Request.Form("p"))
	   end if
	end if
	
	itemCycle = "<tr {$BGCOLOR}>"&vbCrlf&_
				"    <td align=""center"">{$ISBN}</td>"&vbCrlf&_
				"    <td><a href=""?s=s&amp;ISBN={$ISBN}"" title=""BookID = {$BookID}"">{$BookName}</a></td>"&vbCrlf&_
				"    <td>{$BookPath}</td>"&vbCrlf&_
				"    <td align=""center"" nowrap>{$TimeFlag}</td>"&vbCrlf&_
				"    <td>{$OnService}</td>"&vbCrlf&_
				"    <td>{$Hits}</td>"&vbCrlf&_
				"  </tr>"
			
   CycleCont = Array(Replace(itemCycle,"{$BGCOLOR}", ""),Replace(itemCycle,"{$BGCOLOR}", "bgcolor=""#E3EBF9"""))
	iBDb.SQL = "select ISBN,BookName,BookPath,OnlineDate,Hits,STOK,BookID from [iReaderBooks] order by OnlineDate desc"
	  'Response.Write(iBDb.SQL)
	  
	  Set Rs = iBDb.GetRs()
	  Set DataTab = new iBDataTable
		  DataTab.PageSize = iPageSize
		  DataTab.PagerID = "sList"
		  DataTab.CurrentPage = CurPage
		  if (not IsPostBack) then
			 DataTab.TotalCount = iBDb.GetScalar("select count(BookId) as total from [iReaderBooks]")
		  end if
		  DataTab.PageData = false 
		  
		  DataTab.dtRepItem = Array("{$ISBN}","{$BookName}","{$BookPath}","{$TimeFlag}","{$Hits}","{$OnService}","{$BookID}")
		  DataTab.dtRepIdx = Array(0,1,2,3,4,5,6)
		  DataTab.dtRepFun = Array("","","","","","{FUN:GetValue($,""��"",""��"")}","")
		  
		  DataTab.CycleTpt = CycleCont 
		  DataTab.Execute(Rs)
		  strCycleData = DataTab.CycleData
		  Pager = DataTab.PagerData
	Set DataTab = nothing
	Set Rs = nothing
%>
<table width="95%"  border="1" align="center" cellpadding="5" cellspacing="0" bgcolor="#CCE7FF">
  <tr bgcolor="#000000">
    <td height="25" colspan="6" align="center" valign="middle"><span class="style1">���ߵ�����(CHM)�����б�</span></td>
  </tr>
  <tr align="center" bgcolor="#F3F3F3">
    <td width="11%">ISBN</td>
	<td width="40%">�鼮����</td>
    <td width="11%">�鼮·��</td>
    <td width="16%">ʱ���</td>
	<td width="5%">״̬</td>
    <td width="9%" nowrap>�������</td>
  </tr>
  <%=strCycleData%>
  <tr>
    <td colspan="6" bgcolor="#F3F3F3"><%=Pager%></td>
  </tr>
</table>
<%
End Sub


'********************************
'ͶƱ�б�
'**************************************************
Sub BookVoteList()
	
%>
<table width="95%"  border="1" align="center" cellpadding="5" cellspacing="0" bgcolor="#CCE7FF">
  <tr bgcolor="#000000">
    <td height="25" colspan="6" align="center" valign="middle"><span class="style1">������(CHM)����ͶƱ�����б�</span></td>
  </tr>
  <tr align="center" bgcolor="#F3F3F3">
    <td width="15%">ISBN</td>
	<td width="18%">�鼮���</td>
    <td width="25%">��Ʊ����</td>
    <td width="13%">�ϴ�ͶƱ��IP</td>
    <td width="18%">����ʱ��</td>
    <td width="11%">�Ự��־</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td colspan="6" bgcolor="#F3F3F3">&nbsp;</td>
  </tr>
</table>
<%
End Sub

'********************************
'�����б�
'**************************************************
Sub BookHelpList()

If Request.QueryString("id") <> "" Then
	Call HelpEdit(Request.QueryString("id"),0)
ElseIf Request.QueryString("pid") <> "" Then
	Call HelpEdit(Request.QueryString("pid"),1)
Else

	Dim iBDb,Rs,strMsg,IO
	Set iBDb = new iBDataBase
		'iBDb.ConnString = ConnectionString
		Set iBDb.ConnObject = Conn
		Conn.BeginTrans()

	'************ Update System Help HHC 2005��10��7�� ������ 23:06:26
	If Request.QueryString("do") = "hhc" Then
		Dim HelpHHC
		iBDb.RsUpdate = True
		HelpHHC = "<UL>" & BuildHelpHHC(iBDb,0) & "</UL>"
		iBDb.SQL = "select top 1 * from [iReaderConst]"
		Set Rs = iBDb.GetRs()
			Rs("HelpHHCHtml") = HelpHHC
			Rs.Update()
		Set Rs = Nothing
		iBDb.RsUpdate = False
		ShowMsg " �ɹ�����ϵͳ����HHC���� ",3
	End If
		
	Dim DataTab,Pager,strCycleData
    Dim iPageSize,CurPage,CycleCont,itemCycle 
	
	iPageSize = 20 : CurPage = 1
	if (Request.Form <> "") then
	   if IsEmpty(Request.Form("p")) then
		  CurPage = 1
		elseif IsNumeric(Request.Form("p")) then
		  CurPage = CLng(Request.Form("p"))
	   end if
	end if
	
	itemCycle = "<tr {$BGCOLOR}>"&vbCrlf&_
				"    <td align=""center""><img src=""images/icon/{$icon}.gif"" border=""0"" /></td>"&vbCrlf&_
				"    <td>{$title}</td>"&vbCrlf&_
				"    <td nowrap><a href=""?s=h&id={$ID}"">�༭����</a> <a href=""?s=h&pid={$ID}"">�Ӱ���</a></td>"&vbCrlf&_
				"  </tr>"
			
   CycleCont = Array(Replace(itemCycle,"{$BGCOLOR}", ""),Replace(itemCycle,"{$BGCOLOR}", "bgcolor=""#E3EBF9"""))
	iBDb.SQL = "select iConIndex,Title,hId from [iReaderHelp] order by hId desc"
	  'Response.Write(iBDb.SQL)
	  
	  Set Rs = iBDb.GetRs()
	  Set DataTab = new iBDataTable
		  DataTab.PageSize = iPageSize
		  DataTab.PagerID = "sList"
		  DataTab.CurrentPage = CurPage
		  if (not IsPostBack) then
			 DataTab.TotalCount = iBDb.GetScalar("select count(hId) as total from [iReaderHelp]")
		  end if
		  DataTab.PageData = false 

		  DataTab.dtRepItem = Array("{$icon}","{$title}","{$ID}")
		  DataTab.dtRepIdx = Array(0,1,2)
		  DataTab.dtRepFun = Array("","","")
  
		  DataTab.CycleTpt = CycleCont 
		  DataTab.Execute(Rs)
		  strCycleData = DataTab.CycleData
		  Pager = DataTab.PagerData
	Set DataTab = nothing
	Set Rs = nothing
	%>
<table width="95%"  border="1" align="center" cellpadding="5" cellspacing="0" bgcolor="#CCE7FF">
  <tr bgcolor="#000000">
    <td height="25" colspan="3" align="center" valign="middle"><span class="style1">iReader ���������б�</span></td>
  </tr>
  <tr align="center" bgcolor="#F3F3F3">
    <td width="30" nowrap="nowrap">ͼ��</td>
	<td width="82%">����</td>
    <td width="60">�� ��</td>
  </tr>
  <%=strCycleData%>
  <tr>
    <td colspan="2" bgcolor="#F3F3F3"><%=Pager%></td><td align="center" bgcolor="#F3F3F3"><input name="btnHtmlHelp" type="button" id="btnHtmlHelp" value="���� HelpTOC" onclick="location.href='?s=h&amp;do=hhc';" /></td>
  </tr>
</table>
	<%
End If

End Sub

'******************************
'�༭��������
'*************************************
Sub HelpEdit(Id,NewHelp)

	Dim parentId,title,idx,Content
	idx = 11

	If (NewHelp=1) Then
		parentId = Id
		Id = 0
	Else
		parentId = 0
	End If

	Set Rs = Server.CreateObject("ADODB.Recordset")
	sql = "select top 1 * from [iReaderHelp] where hId="&CheckStr(Id)
	Rs.Open sql,Conn,3,3

	If IsPostBack() Then
		If (NewHelp=1) Then Rs.AddNew
		Rs("iConIndex")	= Request.Form("iconIdx")
		Rs("Title") = Request.Form("title")
		Rs("ParentId") = Request.Form("parentId")
		Rs("Content") = Request.Form("Content")
		Rs.Update()
	End If

	If Not Rs.Eof Then
		idx = CInt(Rs("iConIndex"))
		title = Rs("Title")
		Content = Rs("Content")
		parentId = Rs("ParentId")
	End If

	Rs.Close()
	Set Rs = Nothing
%>
<script language="JavaScript">
<!--
 function setIcon(idx)
 {
	var obj = document.getElementById("icon");
	obj.src = "images/icon/"+idx+".GIF";
 }
//-->
</script>
<form method="post" name="frmExample" id="frmExample">
<table width="95%"  border="1" align="center" cellpadding="5" cellspacing="0" bgcolor="#CCE7FF">
  <tr bgcolor="#000000">
    <td height="25" colspan="4" align="center" valign="middle"><span class="style1">iReader �������ݱ༭</span></td>
  </tr>
  <tr bgcolor="#F3F3F3">
    <td width="15%" align="center">�� ��</td>
    <td colspan="3" bgcolor="#CCE7FF"><input name="title" type="text" size="32" maxlength="100" value="<%=title%>" /></td>
  </tr>
  <tr bgcolor="#F3F3F3">
    <td align="center">ͼ ��</td>
    <td width="42%" bgcolor="#CCE7FF"><img src="images/icon/<%=idx%>.GIF" name="icon" width="16" height="16" border="0" id="icon" />
	<select name="iconIdx" onchange="setIcon(this.value)"><%
			Dim i
			For i=1 To 42
			Response.Write "<option value="""&i&""""
			If i=idx Then
				Response.Write(" selected")
			End If
			Response.Write ">"&i
			Response.Write "</option>"
			Next
	%></select>(1-42֮�������)</td>
    <td width="8%" align="left" bgcolor="#F3F3F3">�������</td>
    <td width="35%" bgcolor="#CCE7FF"><input name="parentId" type="text" size="5" value="<%=parentId%>" /></td>
  </tr>
  <tr bgcolor="#F3F3F3">
    <td align="center">��������</td>
    <td colspan="3" bgcolor="#CCE7FF">
	<div style="width:500px">
	<textarea name="content" cols="85" rows="15" id="content" style="width:500px;height:200px;">
	 <%=content%>
	</textarea>
	</div>
	</td>
  </tr>
  <tr bgcolor="#F3F3F3">
    <td colspan="4" align="center"> <input name="btnSet" type="submit" id="btnSet" value="��������">
&nbsp;&nbsp;&nbsp;&nbsp;
<input name="btnReset" type="reset" id="btnReset" value="��������"></td>
  </tr>
  <tr>
    <td colspan="4" bgcolor="#F3F3F3"></td>
  </tr>
</table>
</form>
<script language="javascript">
var _formName = 'frmExample';
var _textName = 'content';
var _toolBarIconPath = '/iEditor/Icons';
var _debug = false;
var _maxCount = 64000;
//var _postAction = 'about:blank';
//����
var _a_lang = new Array();
_a_lang['pic'] = 'ͼƬ';
_a_lang['url'] = '��ַ';
_a_lang['viewe'] = '��ʾЧ��';
_a_lang['border'] = '�߿��ϸ';
_a_lang['align'] = '���뷽ʽ';
_a_lang['absmiddle'] = '���Ծ���';
_a_lang['aleft'] = '����';
_a_lang['aright'] = '����';
_a_lang['atop'] = '����';
_a_lang['amiddle'] = '�в�';
_a_lang['abottom'] = '�ײ�';
_a_lang['absbottom'] = '���Եײ�';
_a_lang['baseline'] = '����';
_a_lang['submit'] = 'ȷ��';
_a_lang['cancle'] = 'ȡ��';
_a_lang['hlink'] = '��������';
_a_lang['other'] = '����ѡ��';
_a_lang['newwindow'] = '���´��ڴ�';
_a_lang['ttop'] = '�ı�����';
_a_lang['copy'] = '����';
_a_lang['cut'] = '����';
_a_lang['pau'] = 'ճ��';
_a_lang['del'] = 'ɾ��';
_a_lang['bold'] = '����';
_a_lang['italic'] = 'б��';
_a_lang['underline'] = '�»���';
_a_lang['st'] = '�л���';
_a_lang['jl'] = '�����';
_a_lang['jc'] = '���ж���';
_a_lang['jr'] = '�Ҷ���';
_a_lang['fcolor'] = '������ɫ';
_a_lang['bcolor'] = '���ֱ�����ɫ';
_a_lang['ilist'] = '���';
_a_lang['itlist'] = '��Ŀ����';
_a_lang['sup'] = '�ϱ�';
_a_lang['sub'] = '�±�';
_a_lang['createlink'] = '��������';
_a_lang['unlink'] = 'ȡ������';
_a_lang['inserthr'] = '����ˮƽ��';
_a_lang['insertimg'] = '����/�޸�ͼƬ';
_a_lang['editsource'] = '�༭Դ�ļ�';
_a_lang['preview'] = 'Ԥ��';
_a_lang['usehtml'] = 'ʹ�ñ༭���༭';
_a_lang['font'] = '����';
_a_lang['simsun'] = '����';
_a_lang['simhei'] = '����';
_a_lang['simkai'] = '����';
_a_lang['fangsong'] = '����';
_a_lang['lishu'] = '����';
_a_lang['youyuan'] = '��Բ';
_a_lang['fontsize'] = '�ֺ�';
_a_lang['fontsize_1'] = 'һ��';
_a_lang['fontsize_2'] = '����';
_a_lang['fontsize_3'] = '����';
_a_lang['fontsize_4'] = '�ĺ�';
_a_lang['fontsize_5'] = '���';
_a_lang['fontsize_6'] = '����';
_a_lang['fontsize_7'] = '�ߺ�';
_a_lang['current'] = '��ǰ';
_a_lang['word'] = '��';
_a_lang['maxword'] = '���';
_a_lang['modify'] = '�޸�';
_a_lang['insert'] = '����';
</script>
<script language="javascript" src="/iEditor/editor_multi_lang.js"></script>
<%
End Sub


'******************************
'��ʾ����˵�
'*************************************
Sub ShowMenu(k)
 Dim str(5,1),i
 for i=0 to 5
 	 if i=k Then
		 str(i,0) = " bgcolor=""#F3F3F3"""
		 str(i,1) = "style2"
	 Else
		 str(i,0) = ""
		 str(i,1) = "style1"
	 End If
 next
%>
<table width="95%"  border="0" align="center" cellpadding="5" cellspacing="0" bgcolor="#AA5B00">
  <tr>
    <td align="center"<%=str(0,0)%>><a href="?s=p"><span class="<%=str(0,1)%>">������Ϣ</span></a></td>
	<td height="22" align="center" valign="middle"<%=str(4,0)%>><a href="?s=s"><span class="<%=str(4,1)%>">�Ǽ��鼮</span></a></td>
	<td align="center"<%=str(1,0)%>><a href="?s=l"><span class="<%=str(1,1)%>">�����б�</span></a></td>
    <td align="center"<%=str(2,0)%>><a href="?s=v"><span class="<%=str(2,1)%>">ͶƱ�б�</span></a></td>
	<td align="center"<%=str(5,0)%>><a href="?s=h"><span class="<%=str(5,1)%>">�����б�</span></a></td>
    <td align="center"<%=str(3,0)%>><a href="?s=x"><span class="<%=str(3,1)%>">�˳���¼</span></a></td>
  </tr>
  <tr><td colspan="6" bgcolor="#f3f3f3" height="25"></td></tr>
  <tr><td colspan="6" bgcolor="#DEE3F7" align="center"><div style="padding:5px;text-align:left;">����(TOC.asp?ISBN=1234567890)��TOC���ݸ�������Ӳ���do=reload��TOC����ԭʼ��������Ӳ���show=hhc��</div></td></tr>
</table>
<%
End Sub


'*******************************
'��֤����
'****************************************
Sub iReaderPass()

	Dim name,pass,np,np2
	If IsPostBack() Then
		name = Trim(Request.Form("name"))
		pass = Trim(Request.Form("pass"))
		np = Trim(Request.Form("np"))
		np2 = Trim(Request.Form("nprepeat"))
	   Else
		If Session("Manage") <> "" Then
			name = Session("Manage")
		End If
	End If
	
	If Session("Manage") <> "" Then
		Call ShowMenu(0)
	End If
	%><div style="height:120px;">&nbsp;</div>
	<form method="post" enctype="application/x-www-form-urlencoded">
	  <table width="450" border="1" align="center" cellpadding="5" cellspacing="0" bgcolor="#CCE7FF">
		<tr align="center" bgcolor="#005BAA">
		  <td height="25" colspan="2" valign="middle" bgcolor="#000000"><span class="style1">������֤��Ϣ</span></td>
		</tr>
		<tr>
		  <td align="right">�û���</td>
		  <td><input name="name" type="text" id="name" size="25" maxlength="32" value="<%=name%>" /></td>
		</tr>
		<tr>
		  <td align="right">�� ��</td>
		  <td><input name="pass" type="password" id="pass" size="25" maxlength="32" value="<%=pass%>" /></td>
		</tr>
	  <%
	  If Session("Manage") <> "" Then
	  		dim filename
			filename = Request.ServerVariables("SCRIPT_NAME")
			filename = Replace(Mid(filename,InstrRev(filename,"/",-1,1)+1),".asp","")
	  %>
		<tr>
		  <td align="right">������</td>
		  <td><input name="np" type="password" id="np" size="25" maxlength="32"  value="<%=np%>" /></td>
		</tr>
		<tr>
		  <td align="right">������ȷ��</td>
		  <td><input name="nprepeat" type="password" id="nprepeat" size="25" maxlength="32"  value="<%=np2%>" /></td>
		</tr>
		<tr>
		  <td align="right">�����ļ���</td>
		  <td><input name="filename" type="text" id="filename" value="<%=filename%>" size="15" />
	      (��д*.asp�е�*��ʾ���ļ���)</td>
	    </tr>
		<%
	  End If
	
	  If IsPostBack() Then
		Dim UserName,Password
	
		UserName = Checkstr(name)
		Password = Checkstr(pass)
	
		Set Rs = Server.CreateObject("ADODB.Recordset")
		sql = "select top 1 * from [iReaderConst] where PassName='"&UserName&"' and passkey='"&password&"' "
		Rs.Open sql,Conn,3,3
		If Rs.Eof Then
			strMsg = strMsg & "��֤��Ϣ����"
		Else
			Dim strCurrentPath
			strCurrentPath  = Server.MapPath("./") & "\"

			'*****************
			'Rename ManageFile Name
			If filename <> Request.Form("filename") Then
				If RenameFile(strCurrentPath & filename & ".asp",strCurrentPath & Request.Form("filename") & ".asp") Then
					strMsg = strMsg & "�ɹ����������ļ���Ϊ" & Request.Form("filename") & ".asp <br/>"
					filename = Request.Form("filename") & ".asp"
				Else
					strMsg = strMsg & "���������ļ���ʧ�ܣ� <br/>"
				End If
			End If
			
			If np = np2 And Len(np)>0 Then
				Rs("PassName") = UserName
				Rs("passKey") = np
				Rs.Update()
				strMsg = strMsg & "�ɹ��޸����ĵ�¼��Ϣ��"
			Else
				Rs("LastIp") = Web_GetClientIP()
				Rs("loginCount") = Rs("loginCount")+1
				Rs("LastTime") = Now()
				Rs.Update()
				SetManageRights(name)
				strMsg = strMsg & "<a href="""&filename&"?s=p"">�ɹ���¼����ѡ��˵�����</a>"
			End If
		End If
		Rs.Close()
		Set Rs = Nothing
	  %>
		<tr>
		  <td colspan="2" align="center"><span class="style2"><%=strMsg%></span></td>
		</tr>
		<%
	  End If
	  %>
		<tr>
		  <td colspan="2" align="center"><input type="submit" name="Submit" value="Ok,�㶨��" />
	      &nbsp;&nbsp;<a href="?s=x">�˳���¼</a></td>
		</tr>
	  </table>
	</form>
	<%
End Sub


Call DbClose(Conn)
%>
</body>
</html>