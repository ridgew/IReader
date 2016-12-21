<!--#include file="Const.asp"-->
<!--#include file="iBCore.asp"-->
<%
Dim ISBN,Book,URI,Mark
Dim strParameter,bList
Dim iBDB,Rs

bList = False
ISBN = Request.QueryString("ISBN")
URI = Request.QueryString("URI")
Mark = Request.QueryString("Mark")

If Len(ISBN)=10 Then
		strParameter = strParameter & "?ISBN=" & ISBN
	Else
		'strParameter = strParameter & "?ISBN=0735712018"
		bList = True
		Call ShowReadList()
End If

If Len(URI)>1 Then
	strParameter = strParameter & "&amp;URI=" & URI
End If

If Len(Mark)>1 Then
	strParameter = strParameter & "&amp;Mark=" & Mark
End If


If (bList = False) Then
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<title>&#105;Reader &#105;&#223;虚数传播网络(Beta)</title>
<link rel="icon" href="/favicon.ico" type="image/x-icon" />
<script language="JavaScript">
<!--
function toc(obj)
{
  try { obj.cols = (obj.cols.indexOf("300,")!=-1)? "1,*":"300,*"; } catch (e) { }
}
//-->
</script>
</head>
<frameset rows="25,*,25" cols="*" frameborder="NO" border="0" framespacing="0">
  <frame src="Menu.asp<%=strParameter%>" name="menuWin" scrolling="NO" noresize>
  <frameset cols="300,*" border="1" frameborder="1" FRAMESPACING="1"  TOPMARGIN="0"  LEFTMARGIN="0" MARGINHEIGHT="0" MARGINWIDTH="0" bordercolor="#DEE3F7" ondblclick="toc(this)" title="双击隐藏或显示TOC" name="tocPanel">
    <frame src="TOC.asp" name="tocWin" scrolling="AUTO" TOPMARGIN="0" LEFTMARGIN="0" MARGINHEIGHT="0" MARGINWIDTH="0" FRAMEBORDER="0" BORDER="0">
    <frame src="Reader.asp<%=strParameter%>" name="mainWin">
  </frameset>
  <frame src="Panel.asp<%=strParameter%>" name="panelWin" scrolling="NO">
</frameset>
<noframes><body>
</body></noframes>
</html>
<%
End If


Sub ShowReadList()
	Set iBDb = New iBDataBase
		iBDb.ConnString = ConnectionString
		
	Dim DataTab,Pager,strCycleData
    Dim iPageSize,CurPage,CycleCont,itemCycle 
	
	iPageSize = 25 : CurPage = 1
	if (Request.Form <> "") then
	   if IsEmpty(Request.Form("p")) then
		  CurPage = 1
		elseif IsNumeric(Request.Form("p")) then
		  CurPage = CLng(Request.Form("p"))
	   end if
	end if
	
	itemCycle = "<tr {$BGCOLOR}>"&vbCrlf&_
				"    <td align=""center"">{$ISBN}</td>"&vbCrlf&_
				"    <td><table cellpadding=""5""><tr><td rowspan=""4"" align=""center"" valign=""middle""><a href=""/iReader/?ISBN={$ISBN}"" target=""_blank""><img src=""Reader.asp?ISBN={$ISBN}&URI={$CoverPath}"" border=""0"" width=""90"" height=""110"" align=""left""></a></td><td><a href=""/iReader/?ISBN={$ISBN}"" target=""_blank""><strong><font color=blue>{$BookName}</font></strong></a></td></tr><tr><td>{$PressName}</td></tr><tr><td>{$PubData}</td></tr><tr><td></td></tr></table></td>"&vbCrlf&_
				"    <td nowrap>{$Expired}</td>"&vbCrlf&_
				"    <td align=""center""><a href=""/iReader/?ISBN={$ISBN}"" target=""_top""><font color=""red"">在线阅读</font></a> <a href=""Download.asp?ISBN={$ISBN}&key=greengate"" target=""_top"" title=""绿色通道：点击断点续传快速下载 ""><font color=""green"">快速下载</font></a> </td>"&vbCrlf&_
				"  </tr>"
			
   CycleCont = Array(Replace(itemCycle,"{$BGCOLOR}", ""),Replace(itemCycle,"{$BGCOLOR}", "bgcolor=""#E3EBF9"""))
	iBDb.SQL = "select ISBN,BookName,TimeFlag,CoverPath,PressName,PublishData from [iReaderBooks] where STOK=True order by BookID desc"
	  'Response.Write(iBDb.SQL)
	  
	  Set Rs = iBDb.GetRs()
	  Set DataTab = new iBDataTable
		  DataTab.PageSize = iPageSize
		  DataTab.PagerID = "sList"
		  DataTab.CurrentPage = CurPage
		  if (not IsPostBack) then
			 DataTab.TotalCount = iBDb.GetScalar("select count(BookId) as total from [iReaderBooks] where STOK=True")
		  end if
		  DataTab.PageData = false 
		  
		  DataTab.dtRepItem = Array("{$ISBN}","{$BookName}","{$Expired}","{$CoverPath}", "{$PressName}", "{$PubData}")
		  DataTab.dtRepIdx = Array(0,1,2,3,4,5)
		  DataTab.dtRepFun = Array("","","{FUN:formatDT(DateAdd(""d"",EXPIREDDAYNUM,""$""),""-1"")}","", "", "")

		  
		  DataTab.CycleTpt = CycleCont 
		  DataTab.Execute(Rs)
		  strCycleData = DataTab.CycleData
		  Pager = DataTab.PagerData
	Set DataTab = nothing
	Set Rs = nothing
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta http-equiv="Content-Language" content="zh-cn" />
<title>&#105;Reader(测试版) 虚数传播网络</title>
<link rel="stylesheet" href="Reader.css" type="text/css" />
<base target="_top" language="javascript">
<style type="text/css">
<!--
.style1 {color: #FFFFFF;font-weight:bold;}
a {text-decoration:none;}
a:link {text-decoration:none;}
a:hover {text-decoration:underline;color:#990000;}
td {font-family:"Courier New", Courier, mono;}
.style3 {color: #00FF00}
.style4 {color: #CC0000}
.style7 {color: gray}
.underline {text-decoration:underline;}
p {font-size:12px;text-indent:24px;padding:5px;}
.style8 {color: #FF0000}
-->
</style>
</head>
<body>
<table width="95%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="184" height="54" rowspan="2"><a href="http://www.vbyte.com/" target="_self"><img src="/images/logoMain.gif" alt="返回IB Networks首页" width="184" height="54" border="0" /></a></td>
    <td height="32" align="right"><a href="/">&#105; &#223;</a> | <a href="/my">会员</a> | <a href="/i">资讯</a> | <a href="/p">产品</a> | <a href="/i/astro">星相</a> | <a href="/cbs">生日榜</a> | 
<a href="/shop">e店</a> | <span style="CURSOR:pointer" onclick="this.style.behavior='url(#default#homepage)';this.sethomepage(document.location.href);return false;">设为首页</span>
<span style="cursor:pointer;" onclick="javascript:window.external.AddFavorite(document.location.href,document.title);return false"> | 加入收藏</span></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td height="50" colspan="2" valign="bottom"><strong>欢迎来到iReader -- CHM文件在线阅读器，马上读您喜欢的英文原版书。</strong></td>
  </tr>
</table>

<div align="center"><div id="read1st" style="width:95%;text-align:left;line-height:150%;border:1px dashed #666666;">
<p>【说明】：iReader尽可能地为在<a href="/shop" target="_blank">虚数E店</a>中的CHM格式电子书提供<span class="underline">在线阅读</span><span class="style3">(√)</span>或<span class="underline">自由下载</span><span class="style3">(√)</span>，因本站目前仅使用虚拟空间200M，故不能提供所有CHM格式电子书的在线阅读和下载服务，于是设置了<span class="underline">投票功能</span><span class="style4">(?)</span>，每月18日管理人员会从得票数靠前的那些电子书籍提供在线阅读和下载服务。</p>
    <p>【iReader设计目标】：(免费)在线阅读或下载原版英文书籍，初步预计阅读速度为1本/月，对应于每月重新整理iReader的在线阅读书籍。<br />
      提供书目阅读模式和个人阅读书签(?)，读于英文书籍中上下文意思不明白的英文单词可以使用&quot;<span class="underline">正常阅读模式</span>&quot;<span class="style7">(*)</span>页面底部的<span class="underline">查词功能</span><span class="style3">(√)</span>查询该词语的中文意思。<br />
      如果没有时间每天在线阅读，可以<span class="underline">直接下载</span><span class="style3">(√)</span>保存或珍藏您喜爱的电子书。</p>
    <p>【本页数据列表提示】：点击&quot;Book's Name&quot;查看该书籍的更<span class="underline">详细介绍</span><span class="style4">(?)</span>，Expire列表时该书的到期阅读时间，<span class="underline">Read</span><span class="style3">(√)</span>为在线阅读该书，<span class="underline">Vote</span><span class="style4">(?)</span>为下次能有机会继续阅读该书投票。</p>
    <p>文本中标签含义：<span class="style3">(√)</span>-&gt;已实现功能 <span class="style4">(?)</span>-&gt;暂未实现功能 <span class="style7">(*)</span>-&gt;采用类似于<span class="underline">/iReader/?ISBN=1234567890</span>的地址模式阅读书籍。</p>
    <p class="style8">对本系统和 vByte.com 网站发展感兴趣的朋友，欢迎加入本人QQ：12035729讨论(认证信息:iReader)。</p>
</div></div>
	
 <br  /> <br  />
<table width="95%"  border="1" align="center" cellpadding="5" cellspacing="0">
  <tr align="center" bgcolor="#CCCCCC">
    <td colspan="4"><strong>iReader 在线阅读书目列表</strong></td>
  </tr>
  <tr align="center" bgcolor="#666666">
    <td width="80" class="style1">ISBN</td>
    <td bgcolor="#666666" class="style1">Book's Name</td>
    <td class="style1">Expire</td>
    <td width="80" bgcolor="#666666" class="style1">Choose</td>
  </tr>
   <%=strCycleData%>
  <tr>
    <td colspan="4"><%=Pager%></td>
  </tr>
</table>

<table width="95%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><table cellspacing="0" cellpadding="5" width="100%" align="center" border="0">
          <tr align="center">
            <td height="32" colspan="9">&nbsp;</td>
          </tr>
          <tr align="center">
            <td><a href="/v">关于&#105; &#223;</a></td>
            <td><a href="/my/PassportGet.asp">用户注册</a></td>
            <td><a href="/v/privacy.asp">隐私条款</a></td>
            <td><a href="/v/copyright.asp">版权声明</a></td>
            <td><a href="/v/VisnNote.asp">&#105; &#223;记事</a></td>
            <td><a href="/p">本网作品</a></td>
            <td><a href="/shop/list.asp?type=bk" target="_blank">备用书库</a></td>
            <td><a href="/v/support.asp">在线反馈</a></td>
            <td><img height="16" alt="QQ:12035729" src="http://icon.tencent.com/12035729/s/180/" 
            width="16" border="0" /> <a href="/v/support.asp" target="_blank">服务支持</a></td>
          </tr>
		  <tr align="center">
            <td colspan="9"><hr align="center" size="1" noshade width="100%" /></td>
          </tr>
          <tr>
            <td align="center" colspan="9" height="20">Copyright &copy; 2003-2007 &#105; &#223; Networks, All Rights Reserved 版权所有 <br />
            虚数传播网络 <strong>WWW.<font color="#ff0000">V</font>BYTE.COM</strong> 版权所有&nbsp;&nbsp;<a href="http://www.miibeian.gov.cn" target="_blank">蜀ICP备05002307号</a></td>
          </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
<%
End Sub
%>