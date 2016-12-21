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
<title>&#105;Reader 系统管理</title>
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
		ShowMsg "成功退出管理状态！",5
		Call iReaderPass()
	Case Else
		Call iReaderPass()
End Select


'********************************
'书籍数据设置
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
				strMsg = strMsg & "已经没有该项相关数据的上下条了！<br>"
				ISBN = 0
			ElseIf (CStr(ISBN) <> "0" ) Then
				strMsg = strMsg & "该ISBN在数据库中不存在，不会覆盖已有数据！<br>"
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
				strMsg = strMsg & "已经添加一新纪录。<br>"
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
			strMsg = strMsg & "成功设置该项数据！"
			ShowMsg strMsg,3
		End If

	Rs.Close()
	Set Rs = Nothing
%>
<form method="post" enctype="application/x-www-form-urlencoded">
<table width="95%"  border="1" align="center" cellpadding="5" cellspacing="0" bordercolor="#000000" bgcolor="#CCE7FF">
  <tr align="center" bgcolor="#000000">
    <td height="25" colspan="8" valign="middle"><span class="style1">设置在线阅读电子书(CHM)资料</span></td>
  </tr>
  <tr>
    <td width="0" align="right" bgcolor="#f3f3f3">书籍名称</td>
    <td colspan="7"><input name="bookname" type="text" id="bookname" value="<%=bookname%>" size="70"></td>
  </tr>
  <tr>
    <td width="0" align="right" bgcolor="#f3f3f3">ISBN</td>
    <td colspan="3"><input name="ISBN" type="text" id="ISBN" value="<%=ISBN%>" size="20" maxlength="10" onChange="location.href='?s=s&ISBN='+ISBN.value;">
    (10个字符, 可以使用IB05093001形式. ) <a href="#" class="style2" onClick="location.href='?s=s&ISBN='+ISBN.value;">检查是否存在</a></td>
    <td width="0" align="center" bgcolor="#F3F3F3">出版数据</td>
    <td colspan="3"><input name="publishdata" type="text" id="publishdata" value="<%=publishdata%>" size="35"></td>
  </tr>
  <tr>
    <td width="0" align="right" bgcolor="#f3f3f3">相对路径</td>
    <td colspan="3"><input name="bookpath" type="text" id="bookpath" value="<%=bookPath%>" size="35"> <%=GetValue(FileExist(GetValue(InStr(1,bookPath,":",1)>0,bookPath,bookRootPath & bookPath)),"<font color=""green"">(该文件存在)</font>","<font color=""red"">(该文件已经不存在)</font>")%></td>
    <td align="center" bgcolor="#FFFFFF">出版实体</td>
    <td colspan="3"><input name="PressName" type="text" id="PressName" value="<%=PressName%>" size="28" /></td>
    </tr>
  <tr>
    <td width="0" align="right" bgcolor="#f3f3f3">首页文件</td>
    <td colspan="3"><input name="mainpage" type="text" id="mainpage" value="<%=mainPage%>" size="28"></td>
    <td width="0" align="center" bgcolor="#F3F3F3">目录文件</td>
    <td colspan="3"><input name="tocpage" type="text" id="tocpage" value="<%=tocPage%>" size="28"></td>
  </tr>
 <tr>
    <td width="0" align="right" bgcolor="#f3f3f3">系统TOC文件</td>
    <td colspan="7"><input name="systoc" type="text" id="systoc" value="<%=sysToc%>" size="55" /> (HHC文件路径)</td>
  </tr>
   <tr>
    <td align="right" bgcolor="#f3f3f3">封面图片路径</td>
    <td colspan="7"><input name="coverpic" type="text" id="coverpic" value="<%=coverpic%>" size="55" /></td>
  </tr>
  <tr>
    <td width="0" align="right" bgcolor="#f3f3f3">发布时间</td>
    <td colspan="3"><input name="onlineDate" type="text" id="onlineDate" value="<%=onlineDate%>" size="20">
    (上线 <%=onlineHits%> 次)</td>
    <td width="0" align="center" bgcolor="#F3F3F3">服务状态</td>
    <td width="0"><input name="onService" type="checkbox" id="onService" value="1" <%=isThisValue(onService,"checked")%>>
    上线服务中</td>
    <td width="0" align="center" bgcolor="#F3F3F3">下载次数</td>
    <td width="0"><input name="downHits" type="text" id="downHits" value="<%=DownHits%>" size="10"></td>
  </tr>
  <tr>
    <td width="0" align="right" bgcolor="#f3f3f3">总浏览量</td>
    <td width="0"><input name="Hits" type="text" id="Hits" value="<%=TotalHits%>" size="20"></td>
    <td width="0" align="center" bgcolor="#F3F3F3">昨天/今天 浏览量</td>
    <td width="0"><input name="ysVstoday" type="text" id="ysVstoday" value="<%=ysVstoday%>" size="20"></td>
    <td width="0" align="center" bgcolor="#F3F3F3">上周浏览量</td>
    <td width="0"><input name="lastweekhits" type="text" id="lastweekhits" value="<%=lastweekHits%>" size="10"></td>
    <td width="0" align="center" nowrap bgcolor="#F3F3F3">本周浏览量</td>
    <td width="0"><input name="thisweekhits" type="text" id="thisweekhits" value="<%=thisweekHits%>" size="10"></td>
  </tr>
  <tr>
    <td colspan="8" align="right" bgcolor="#F3F3F3"><a href="?s=s&ISBN=<%=ISBN%>&Goto=Preview" class="style2">上一本</a>&nbsp;&nbsp;  <a href="?s=s&ISBN=<%=ISBN%>&Goto=Next" class="style2">下一本</a>&nbsp;&nbsp;</td>
  </tr>
  <tr>
    <td colspan="8" align="center" bgcolor="#F3F3F3"> <input name="btnSet" type="submit" id="btnSet" value="设置数据">
&nbsp;&nbsp;&nbsp;&nbsp;
<input name="btnReset" type="reset" id="btnReset" value="重置数据"></td>
  </tr>
</table>
</form>
<%
End Sub

'********************************
'数据列表
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
		  DataTab.dtRepFun = Array("","","","","","{FUN:GetValue($,""√"",""×"")}","")
		  
		  DataTab.CycleTpt = CycleCont 
		  DataTab.Execute(Rs)
		  strCycleData = DataTab.CycleData
		  Pager = DataTab.PagerData
	Set DataTab = nothing
	Set Rs = nothing
%>
<table width="95%"  border="1" align="center" cellpadding="5" cellspacing="0" bgcolor="#CCE7FF">
  <tr bgcolor="#000000">
    <td height="25" colspan="6" align="center" valign="middle"><span class="style1">在线电子书(CHM)资料列表</span></td>
  </tr>
  <tr align="center" bgcolor="#F3F3F3">
    <td width="11%">ISBN</td>
	<td width="40%">书籍名称</td>
    <td width="11%">书籍路径</td>
    <td width="16%">时间戳</td>
	<td width="5%">状态</td>
    <td width="9%" nowrap>浏览总量</td>
  </tr>
  <%=strCycleData%>
  <tr>
    <td colspan="6" bgcolor="#F3F3F3"><%=Pager%></td>
  </tr>
</table>
<%
End Sub


'********************************
'投票列表
'**************************************************
Sub BookVoteList()
	
%>
<table width="95%"  border="1" align="center" cellpadding="5" cellspacing="0" bgcolor="#CCE7FF">
  <tr bgcolor="#000000">
    <td height="25" colspan="6" align="center" valign="middle"><span class="style1">电子书(CHM)所有投票数据列表</span></td>
  </tr>
  <tr align="center" bgcolor="#F3F3F3">
    <td width="15%">ISBN</td>
	<td width="18%">书籍类别</td>
    <td width="25%">得票总数</td>
    <td width="13%">上次投票人IP</td>
    <td width="18%">更新时间</td>
    <td width="11%">会话标志</td>
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
'帮助列表
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

	'************ Update System Help HHC 2005年10月7日 星期五 23:06:26
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
		ShowMsg " 成功更新系统帮助HHC内容 ",3
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
				"    <td nowrap><a href=""?s=h&id={$ID}"">编辑帮助</a> <a href=""?s=h&pid={$ID}"">子帮助</a></td>"&vbCrlf&_
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
    <td height="25" colspan="3" align="center" valign="middle"><span class="style1">iReader 帮助数据列表</span></td>
  </tr>
  <tr align="center" bgcolor="#F3F3F3">
    <td width="30" nowrap="nowrap">图标</td>
	<td width="82%">标题</td>
    <td width="60">编 辑</td>
  </tr>
  <%=strCycleData%>
  <tr>
    <td colspan="2" bgcolor="#F3F3F3"><%=Pager%></td><td align="center" bgcolor="#F3F3F3"><input name="btnHtmlHelp" type="button" id="btnHtmlHelp" value="更新 HelpTOC" onclick="location.href='?s=h&amp;do=hhc';" /></td>
  </tr>
</table>
	<%
End If

End Sub

'******************************
'编辑帮助内容
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
    <td height="25" colspan="4" align="center" valign="middle"><span class="style1">iReader 帮助数据编辑</span></td>
  </tr>
  <tr bgcolor="#F3F3F3">
    <td width="15%" align="center">标 题</td>
    <td colspan="3" bgcolor="#CCE7FF"><input name="title" type="text" size="32" maxlength="100" value="<%=title%>" /></td>
  </tr>
  <tr bgcolor="#F3F3F3">
    <td align="center">图 标</td>
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
	%></select>(1-42之间的数字)</td>
    <td width="8%" align="left" bgcolor="#F3F3F3">父级编号</td>
    <td width="35%" bgcolor="#CCE7FF"><input name="parentId" type="text" size="5" value="<%=parentId%>" /></td>
  </tr>
  <tr bgcolor="#F3F3F3">
    <td align="center">帮助内容</td>
    <td colspan="3" bgcolor="#CCE7FF">
	<div style="width:500px">
	<textarea name="content" cols="85" rows="15" id="content" style="width:500px;height:200px;">
	 <%=content%>
	</textarea>
	</div>
	</td>
  </tr>
  <tr bgcolor="#F3F3F3">
    <td colspan="4" align="center"> <input name="btnSet" type="submit" id="btnSet" value="设置数据">
&nbsp;&nbsp;&nbsp;&nbsp;
<input name="btnReset" type="reset" id="btnReset" value="重置数据"></td>
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
//语言
var _a_lang = new Array();
_a_lang['pic'] = '图片';
_a_lang['url'] = '地址';
_a_lang['viewe'] = '显示效果';
_a_lang['border'] = '边框粗细';
_a_lang['align'] = '对齐方式';
_a_lang['absmiddle'] = '绝对居中';
_a_lang['aleft'] = '居左';
_a_lang['aright'] = '居右';
_a_lang['atop'] = '顶部';
_a_lang['amiddle'] = '中部';
_a_lang['abottom'] = '底部';
_a_lang['absbottom'] = '绝对底部';
_a_lang['baseline'] = '基线';
_a_lang['submit'] = '确定';
_a_lang['cancle'] = '取消';
_a_lang['hlink'] = '超级链接';
_a_lang['other'] = '其他选项';
_a_lang['newwindow'] = '在新窗口打开';
_a_lang['ttop'] = '文本顶部';
_a_lang['copy'] = '复制';
_a_lang['cut'] = '剪切';
_a_lang['pau'] = '粘贴';
_a_lang['del'] = '删除';
_a_lang['bold'] = '粗体';
_a_lang['italic'] = '斜体';
_a_lang['underline'] = '下划线';
_a_lang['st'] = '中划线';
_a_lang['jl'] = '左对齐';
_a_lang['jc'] = '居中对齐';
_a_lang['jr'] = '右对齐';
_a_lang['fcolor'] = '文字颜色';
_a_lang['bcolor'] = '文字背景颜色';
_a_lang['ilist'] = '编号';
_a_lang['itlist'] = '项目符号';
_a_lang['sup'] = '上标';
_a_lang['sub'] = '下标';
_a_lang['createlink'] = '插入链接';
_a_lang['unlink'] = '取消链接';
_a_lang['inserthr'] = '插入水平线';
_a_lang['insertimg'] = '插入/修改图片';
_a_lang['editsource'] = '编辑源文件';
_a_lang['preview'] = '预览';
_a_lang['usehtml'] = '使用编辑器编辑';
_a_lang['font'] = '字体';
_a_lang['simsun'] = '宋体';
_a_lang['simhei'] = '黑体';
_a_lang['simkai'] = '楷体';
_a_lang['fangsong'] = '仿宋';
_a_lang['lishu'] = '隶书';
_a_lang['youyuan'] = '幼圆';
_a_lang['fontsize'] = '字号';
_a_lang['fontsize_1'] = '一号';
_a_lang['fontsize_2'] = '二号';
_a_lang['fontsize_3'] = '三号';
_a_lang['fontsize_4'] = '四号';
_a_lang['fontsize_5'] = '五号';
_a_lang['fontsize_6'] = '六号';
_a_lang['fontsize_7'] = '七号';
_a_lang['current'] = '当前';
_a_lang['word'] = '字';
_a_lang['maxword'] = '最多';
_a_lang['modify'] = '修改';
_a_lang['insert'] = '插入';
</script>
<script language="javascript" src="/iEditor/editor_multi_lang.js"></script>
<%
End Sub


'******************************
'显示管理菜单
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
    <td align="center"<%=str(0,0)%>><a href="?s=p"><span class="<%=str(0,1)%>">管理信息</span></a></td>
	<td height="22" align="center" valign="middle"<%=str(4,0)%>><a href="?s=s"><span class="<%=str(4,1)%>">登记书籍</span></a></td>
	<td align="center"<%=str(1,0)%>><a href="?s=l"><span class="<%=str(1,1)%>">资料列表</span></a></td>
    <td align="center"<%=str(2,0)%>><a href="?s=v"><span class="<%=str(2,1)%>">投票列表</span></a></td>
	<td align="center"<%=str(5,0)%>><a href="?s=h"><span class="<%=str(5,1)%>">帮助列表</span></a></td>
    <td align="center"<%=str(3,0)%>><a href="?s=x"><span class="<%=str(3,1)%>">退出登录</span></a></td>
  </tr>
  <tr><td colspan="6" bgcolor="#f3f3f3" height="25"></td></tr>
  <tr><td colspan="6" bgcolor="#DEE3F7" align="center"><div style="padding:5px;text-align:left;">备忘(TOC.asp?ISBN=1234567890)：TOC数据更新命令，加参数do=reload；TOC数据原始数据命令，加参数show=hhc。</div></td></tr>
</table>
<%
End Sub


'*******************************
'认证过程
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
		  <td height="25" colspan="2" valign="middle" bgcolor="#000000"><span class="style1">输入认证信息</span></td>
		</tr>
		<tr>
		  <td align="right">用户名</td>
		  <td><input name="name" type="text" id="name" size="25" maxlength="32" value="<%=name%>" /></td>
		</tr>
		<tr>
		  <td align="right">密 码</td>
		  <td><input name="pass" type="password" id="pass" size="25" maxlength="32" value="<%=pass%>" /></td>
		</tr>
	  <%
	  If Session("Manage") <> "" Then
	  		dim filename
			filename = Request.ServerVariables("SCRIPT_NAME")
			filename = Replace(Mid(filename,InstrRev(filename,"/",-1,1)+1),".asp","")
	  %>
		<tr>
		  <td align="right">新密码</td>
		  <td><input name="np" type="password" id="np" size="25" maxlength="32"  value="<%=np%>" /></td>
		</tr>
		<tr>
		  <td align="right">新密码确认</td>
		  <td><input name="nprepeat" type="password" id="nprepeat" size="25" maxlength="32"  value="<%=np2%>" /></td>
		</tr>
		<tr>
		  <td align="right">管理文件名</td>
		  <td><input name="filename" type="text" id="filename" value="<%=filename%>" size="15" />
	      (填写*.asp中的*表示的文件名)</td>
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
			strMsg = strMsg & "认证信息错误！"
		Else
			Dim strCurrentPath
			strCurrentPath  = Server.MapPath("./") & "\"

			'*****************
			'Rename ManageFile Name
			If filename <> Request.Form("filename") Then
				If RenameFile(strCurrentPath & filename & ".asp",strCurrentPath & Request.Form("filename") & ".asp") Then
					strMsg = strMsg & "成功更名管理文件名为" & Request.Form("filename") & ".asp <br/>"
					filename = Request.Form("filename") & ".asp"
				Else
					strMsg = strMsg & "更名管理文件名失败！ <br/>"
				End If
			End If
			
			If np = np2 And Len(np)>0 Then
				Rs("PassName") = UserName
				Rs("passKey") = np
				Rs.Update()
				strMsg = strMsg & "成功修改您的登录信息！"
			Else
				Rs("LastIp") = Web_GetClientIP()
				Rs("loginCount") = Rs("loginCount")+1
				Rs("LastTime") = Now()
				Rs.Update()
				SetManageRights(name)
				strMsg = strMsg & "<a href="""&filename&"?s=p"">成功登录，请选择菜单管理！</a>"
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
		  <td colspan="2" align="center"><input type="submit" name="Submit" value="Ok,搞定！" />
	      &nbsp;&nbsp;<a href="?s=x">退出登录</a></td>
		</tr>
	  </table>
	</form>
	<%
End Sub


Call DbClose(Conn)
%>
</body>
</html>