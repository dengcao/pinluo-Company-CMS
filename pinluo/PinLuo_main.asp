<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Pinluo_Main/Conn.asp"-->
<!--#include file="PinLuo_Class.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>首页</title>
<meta content="品络科技,网站管理系统,企业网站管理系统,内容管理系统(CMS),网上商店管理系统,网站建设" name="Keywords" />
<meta content="品络科技成立于2005年6月，是一家集互联网基础服务、互联网应用软件开发、业务解决方案销售及服务于一体的高新技术企业。开发者：www.caozha.com，品络互联：www.pinluo.com" name="Description"/>
<script>
function timeoutlogout()
{
	alert('由于长时间未操作，为安全起见自动退出系统，如需要管理请重新登录！');
	top.location.href='Pl_Logout.asp?Type=TimeOut';
}
var k=setInterval('timeoutlogout()',7200000);
</script>
<link rel="stylesheet" href="images/style.css" type="text/css">
<script language="javascript" src="js/Pinluo.js" type="text/javascript"></script>
<%
act=Trim(Request("act"))
UserID=Request.Cookies("pinluo")("UserID")
dim PinLuo
Set PinLuo = New PinLuo_Class
    PinLuo.DBConnBegin
	PinLuo.CheckPurview
if act="UserWarningSave" then
	UserWarning = Trim(Request.Form("UserWarning"))
    PinLuo.PinLuo_UserWarningSave UserWarning,UserID
    'response.Write("<script language=""javascript"">location.href='PinLuo_main.asp?now="&now&"';<script>")
    PinLuo.DBConnEnd:Set PinLuo=Nothing
    response.Redirect("PinLuo_main.asp?"&now)
	response.End()
end if
    PinLuo.PinLuo_ViewAdminItem UserID
	PinLuo.PinLuo_ViewSiteConfig
	
Dim SYS_SERVER_NAME,SYS_LOCAL_ADDR
SYS_SERVER_NAME = Request.ServerVariables("SERVER_NAME")
SYS_LOCAL_ADDR = Request.ServerVariables("LOCAL_ADDR")
%>
</head>
<body class="mainbg">
	<div id="mainhearder"><span>您的位置：企业网站管理系统 >> 系统首页</span></div>
	<div id="defaulthearder"><img src="images/icon_xtaq.gif" align="absbottom" style="float:left;"><a style="float:right;color:#FFF;margin-right:20px;" href="#" onclick="document.getElementById('UserWarning2').style.display='';document.getElementById('UserWarning1').style.display='none';">修改备忘录</a></div>
	<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
	<tr>
	<td class="tableleft">
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="table1">
		<tr>
		<td valign="middle" class="rz222">
        <span style="display:" id="UserWarning1"><%=PinLuo.UserWarning%></span>
        <span style="display:none" id="UserWarning2">
          <form id="form1" name="form1" method="post" action="PinLuo_main.asp">
            <textarea name="UserWarning" style="display:none"><%if trim(PinLuo.UserWarning)<>"" then response.Write(server.HTMLEncode(PinLuo.UserWarning))%></textarea><iframe id="eWebEditor1" src="../editor/ewebeditor.htm?id=UserWarning&style=Pinluo_blue" frameBorder="0" width="100%" scrolling="no" height="350"></iframe>
            <label>
              <input name="act" type="hidden" id="act" value="UserWarningSave" />
<input type="submit" name="button" id="button" value="修改备忘录" class="button" />&nbsp;
<input type="reset" name="button2" id="button2" value="恢复重写" class="button" />
            </label>
          </form>
          </span></td>
		</tr>
		</table>
	</td>
	</tr>
	</table>
    
    <br>
	<div id="defaulthearder"><img src="images/icon_tjgg.gif" align="absbottom" /></div>
	<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
	<tr>
	<td class="tableleft">
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="table1">
		<tr>
		<td class="tjgg">
		<div id="gg">
		<ul>
		<!--script language="JavaScript" src="http://www.pinluo.com/news/getnews/qycms/?v=1&n=8"></script-->
		<li><a href="http://www.pinluo.com" target=_blank>品络企业网站系统正式发布！</a><span>2010-12-08</span></li>
		<li><a href="http://www.pinluo.com" target=_blank>品络互联提供域名注册，网站空间，服务器等！</a><span>2010-12-08</span></li>
		<li><a href="http://www.caozha.com" target=_blank>草札博客上线，欢迎访问！</a><span>2020-4-22</span></li>
		<li><a href="http://diannao.wang/tool/" target=_blank>在线工具大全发布！</a><span>2010-12-08</span></li>
		</ul>
		</div>
		</td>
		</tr>
		</table>
	</td>
	</tr>
	</table>
    
	<br>	
	<div id="defaulthearder"><img src="images/icon_dlrz.gif" align="absbottom" /></div>
	<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
	<tr>
	<td class="tableleft">
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="table1">
		<tr>
		<td class="rz222">
        
		  <table width="99%" border="0" cellpadding="1" cellspacing="0" class="rz222_table" style="margin:20px 8px;">
		    <tr>
		      <td colspan="3">本程序由品络科技（www.pinluo.com）授权给 <%=Pinluo.Pinluo_SiteName%> 使用，当前使用版本为 品络企业网站管理系统(PinLuo Qiye CMS)&nbsp;<%=Pinluo.Pinluo_Version%>&nbsp;<%=Pinluo.Pinluo_Empower%></td>
		      </tr>
		    <tr>
		      <td width="293"><strong>服务器名：</strong><%=SYS_SERVER_NAME%> (IP:<%=SYS_LOCAL_ADDR%>)　</td>
		      <td width="293"><strong>数据库使用：</strong><%If DataBaseType<>"ACCESS" Then
		Response.Write "Microsoft SQL Server"
	Else
		Response.Write "Microsoft Access"
	End If%></td>
		      <td width="270"><strong>IIS 版本：</strong><%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
		      </tr>
		    <tr>
		      <td><strong>脚本解释引擎：</strong><%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %> </td>
		      <td><strong>FSO文本读写：</strong><%=Pinluo.Pinluo_CheckObj("Scripting.FileSystemObject")%></td>
		      <td><strong>Jmail组件支持：</strong><%=Pinluo.Pinluo_CheckObj("JMail.SMTPMail")%>(4.2)&nbsp;&nbsp;<%=Pinluo.Pinluo_CheckObj("JMail.Message")%>(4.3+)</td>
		      </tr>
		    <tr>
		      <td><strong>远程采集组件：</strong><%=Pinluo.Pinluo_CheckObj("msxml2.XMLHTTP")%></td>
		      <td><strong>AspJpeg生成预览图片：</strong><%=Pinluo.Pinluo_CheckObj("Persits.Jpeg")%> </td>
		      <td><strong>CDONTS组件支持：</strong><%=Pinluo.Pinluo_CheckObj("CDONTS.NewMail")%></td>
		      </tr>
		    <tr>
		      <td><strong>服务器时间：</strong><%=now%></td>
		      <td colspan="2"><strong>站点物理路径：</strong><%=Request.ServerVariables("APPL_PHYSICAL_PATH")%></td>
		      </tr>
		    </table>
            
            </td>
		</tr>
		</table>
	</td>
	</tr>
	</table>
	
	
	<br><br />
<br />
&nbsp;
<span style="display:none">
<!--
统计代码
品络科技用于统计本程序的使用量。
此代码不会影响您网站的正常运行和使用，请您保留。
请支持品络程序的发展！
-->
<script src="http://s16.cnzz.com/stat.php?id=396649&web_id=396649" language="JavaScript"></script>
<script language="javascript" type="text/javascript" src="http://js.users.51.la/3876074.js"></script>
<noscript><a href="http://www.51.la/?3876074" target="_blank"><img alt="&#x6211;&#x8981;&#x5566;&#x514D;&#x8D39;&#x7EDF;&#x8BA1;" src="http://img.users.51.la/3876074.asp" style="border:none" /></a></noscript>
</span>
</body>
</html>
<%
PinLuo.DBConnEnd
Set PinLuo = Nothing
%>	