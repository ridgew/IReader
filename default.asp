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
<title>&#105;Reader &#105;&#223;������������(Beta)</title>
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
  <frameset cols="300,*" border="1" frameborder="1" FRAMESPACING="1"  TOPMARGIN="0"  LEFTMARGIN="0" MARGINHEIGHT="0" MARGINWIDTH="0" bordercolor="#DEE3F7" ondblclick="toc(this)" title="˫�����ػ���ʾTOC" name="tocPanel">
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
				"    <td align=""center""><a href=""/iReader/?ISBN={$ISBN}"" target=""_top""><font color=""red"">�����Ķ�</font></a> <a href=""Download.asp?ISBN={$ISBN}&key=greengate"" target=""_top"" title=""��ɫͨ��������ϵ������������� ""><font color=""green"">��������</font></a> </td>"&vbCrlf&_
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
<title>&#105;Reader(���԰�) ������������</title>
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
    <td width="184" height="54" rowspan="2"><a href="http://www.vbyte.com/" target="_self"><img src="/images/logoMain.gif" alt="����IB Networks��ҳ" width="184" height="54" border="0" /></a></td>
    <td height="32" align="right"><a href="/">&#105; &#223;</a> | <a href="/my">��Ա</a> | <a href="/i">��Ѷ</a> | <a href="/p">��Ʒ</a> | <a href="/i/astro">����</a> | <a href="/cbs">���հ�</a> | 
<a href="/shop">e��</a> | <span style="CURSOR:pointer" onclick="this.style.behavior='url(#default#homepage)';this.sethomepage(document.location.href);return false;">��Ϊ��ҳ</span>
<span style="cursor:pointer;" onclick="javascript:window.external.AddFavorite(document.location.href,document.title);return false"> | �����ղ�</span></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td height="50" colspan="2" valign="bottom"><strong>��ӭ����iReader -- CHM�ļ������Ķ��������϶���ϲ����Ӣ��ԭ���顣</strong></td>
  </tr>
</table>

<div align="center"><div id="read1st" style="width:95%;text-align:left;line-height:150%;border:1px dashed #666666;">
<p>��˵������iReader�����ܵ�Ϊ��<a href="/shop" target="_blank">����E��</a>�е�CHM��ʽ�������ṩ<span class="underline">�����Ķ�</span><span class="style3">(��)</span>��<span class="underline">��������</span><span class="style3">(��)</span>����վĿǰ��ʹ������ռ�200M���ʲ����ṩ����CHM��ʽ������������Ķ������ط�������������<span class="underline">ͶƱ����</span><span class="style4">(?)</span>��ÿ��18�չ�����Ա��ӵ�Ʊ����ǰ����Щ�����鼮�ṩ�����Ķ������ط���</p>
    <p>��iReader���Ŀ�꡿��(���)�����Ķ�������ԭ��Ӣ���鼮������Ԥ���Ķ��ٶ�Ϊ1��/�£���Ӧ��ÿ����������iReader�������Ķ��鼮��<br />
      �ṩ��Ŀ�Ķ�ģʽ�͸����Ķ���ǩ(?)������Ӣ���鼮����������˼�����׵�Ӣ�ĵ��ʿ���ʹ��&quot;<span class="underline">�����Ķ�ģʽ</span>&quot;<span class="style7">(*)</span>ҳ��ײ���<span class="underline">��ʹ���</span><span class="style3">(��)</span>��ѯ�ô����������˼��<br />
      ���û��ʱ��ÿ�������Ķ�������<span class="underline">ֱ������</span><span class="style3">(��)</span>����������ϲ���ĵ����顣</p>
    <p>����ҳ�����б���ʾ�������&quot;Book's Name&quot;�鿴���鼮�ĸ�<span class="underline">��ϸ����</span><span class="style4">(?)</span>��Expire�б�ʱ����ĵ����Ķ�ʱ�䣬<span class="underline">Read</span><span class="style3">(��)</span>Ϊ�����Ķ����飬<span class="underline">Vote</span><span class="style4">(?)</span>Ϊ�´����л�������Ķ�����ͶƱ��</p>
    <p>�ı��б�ǩ���壺<span class="style3">(��)</span>-&gt;��ʵ�ֹ��� <span class="style4">(?)</span>-&gt;��δʵ�ֹ��� <span class="style7">(*)</span>-&gt;����������<span class="underline">/iReader/?ISBN=1234567890</span>�ĵ�ַģʽ�Ķ��鼮��</p>
    <p class="style8">�Ա�ϵͳ�� vByte.com ��վ��չ����Ȥ�����ѣ���ӭ���뱾��QQ��12035729����(��֤��Ϣ:iReader)��</p>
</div></div>
	
 <br  /> <br  />
<table width="95%"  border="1" align="center" cellpadding="5" cellspacing="0">
  <tr align="center" bgcolor="#CCCCCC">
    <td colspan="4"><strong>iReader �����Ķ���Ŀ�б�</strong></td>
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
            <td><a href="/v">����&#105; &#223;</a></td>
            <td><a href="/my/PassportGet.asp">�û�ע��</a></td>
            <td><a href="/v/privacy.asp">��˽����</a></td>
            <td><a href="/v/copyright.asp">��Ȩ����</a></td>
            <td><a href="/v/VisnNote.asp">&#105; &#223;����</a></td>
            <td><a href="/p">������Ʒ</a></td>
            <td><a href="/shop/list.asp?type=bk" target="_blank">�������</a></td>
            <td><a href="/v/support.asp">���߷���</a></td>
            <td><img height="16" alt="QQ:12035729" src="http://icon.tencent.com/12035729/s/180/" 
            width="16" border="0" /> <a href="/v/support.asp" target="_blank">����֧��</a></td>
          </tr>
		  <tr align="center">
            <td colspan="9"><hr align="center" size="1" noshade width="100%" /></td>
          </tr>
          <tr>
            <td align="center" colspan="9" height="20">Copyright &copy; 2003-2007 &#105; &#223; Networks, All Rights Reserved ��Ȩ���� <br />
            ������������ <strong>WWW.<font color="#ff0000">V</font>BYTE.COM</strong> ��Ȩ����&nbsp;&nbsp;<a href="http://www.miibeian.gov.cn" target="_blank">��ICP��05002307��</a></td>
          </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
<%
End Sub
%>