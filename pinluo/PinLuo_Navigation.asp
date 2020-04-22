<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Pinluo_Main/Conn.asp"-->
<!--#include file="PinLuo_Class.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<%
act=Trim(Request("act"))
dim PinLuo
Set PinLuo = New PinLuo_Class
    PinLuo.DBConnBegin
	PinLuo.CheckPurview
	PinLuo.Pinluo_CheckPurviewAdmin(0)
if act="editsave" then
    if Trim(Request("do"))="default" then
	TxtContent="<UL><LI><A href=""../"" target=main>首页</A></LI>"&vbcrlf&"<LI><A href=""PinLuo_ProductList.asp"" target=main>产品管理</A></LI>"&vbcrlf&"<LI><A href=""PinLuo_InfoList.asp"" target=main>信息管理</A></LI>"&vbcrlf&"<LI><A href=""PinLuo_ItemClass.asp"" target=main>栏目管理</A></LI>"&vbcrlf&"<LI><A href=""http://www.pinluo.com"" target=_blank>域名空间</A></LI>"&vbcrlf&"<LI><A href=""http://www.caozha.com"" target=_blank>草札</A></LI></UL>"&vbcrlf&"<!--对应顶部快捷导航-->"
	else
	TxtContent = Trim(Request.Form("TxtContent"))
	end if
    PinLuo.WriteSaveFile TxtContent,"include/Pinluo_Navigation_config.inc"
    PinLuo.DBConnEnd:Set PinLuo=Nothing
	If FoundErr = True Then
        PinLuo.PinLuo_WriteMsg ErrMsg,"PinLuo_Navigation.asp"
   		Response.End()
	End If
	response.Redirect("PinLuo_Navigation.asp")
	response.End()
end if
%>
<title>修改后台导航菜单</title>
<meta content="品络科技,网站管理系统,企业网站管理系统,内容管理系统(CMS),网上商店管理系统,网站建设" name="Keywords" />
<meta content="品络科技成立于2005年6月，是一家集互联网基础服务、互联网应用软件开发、业务解决方案销售及服务于一体的高新技术企业。开发者：www.caozha.com，品络互联：www.pinluo.com" name="Description"/>
<link rel="stylesheet" href="images/style.css" type="text/css">
<script language="JavaScript">
<!--
function submit_to(url){
	
document.PinLuo_Navigation.action=url;
document.PinLuo_Navigation.submit();
}
function Dodefault(){
if(confirm("您确定恢复默认吗？\n\n此操作将无法恢复，请慎重操作。")==true){
	submit_to('PinLuo_Navigation.asp?do=default');
	}else{
		}
	}
//-->
</script>
</head>

<body class="mainbg">
	<div id="mainhearder"><span>您的位置：企业网站管理系统 >> 修改后台快捷导航菜单</span></div>
	<div id="hearder" class="hearder1"><span>修改后台快捷菜单</span></div>
	<div class="main5" id="main5">

	<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
	<tr>
	<td class="tableleft">
    
       <form action="PinLuo_Navigation.asp" method="post" name="PinLuo_Navigation" id="PinLuo_Navigation">
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="table1">
        <tr>
		<td class="tableleft1"></td>
		<td class="tableright1">
        <%TxtContent=PinLuo.WriteReadFile("include/Pinluo_Navigation_config.inc")
		If FoundErr = True Then
        PinLuo.PinLuo_WriteMsg ErrMsg,"PinLuo_Navigation.asp"
   		Response.End()
	    End If
		%>
        <textarea name="TxtContent" style="display:none"><%if trim(TxtContent)<>"" then response.Write(replace(server.HTMLEncode(TxtContent),"&#65279;&lt;UL&gt;","&lt;UL&gt;"))%></textarea><iframe id="eWebEditor1" src="../editor/ewebeditor.htm?id=TxtContent&style=Pinluo_blue" frameBorder="0" width="100%" scrolling="no" height="350"></iframe>
        </td>
		</tr>
		<tr>
		  <td height="55" align="center" class="tableleft1" style="height:39px;">&nbsp;</td>
		  <td height="55" class="tableright1">
          
          <input name="act" type="hidden" value="editsave" />
          <input name="button2" type="button" class="button" value="完成修改" onClick="submit_to('PinLuo_Navigation.asp');">
			<input name="button2" type="reset" class="button" value="重新填写">
            <input name="button2" type="button" class="button" value="恢复默认" onClick="Dodefault();">
            <input name="button1" type="button" class="button" value="刷新本页" onClick="location.reload();">
			</td>
		  </tr>

		</table>
        </form>
       		
	</td>
	</tr>
	</table>
	
	</div>
            <br />
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