<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Pinluo_Main/Conn.asp"-->
<!--#include file="PinLuo_Class.asp"-->
<%
act=Trim(Request("act"))
SearchKeyword2 = Trim(Request("SearchKeyword"))
SearchSelect2 = Trim(Request("SearchSelect"))

OnepageNum=12
dim PinLuo
Set PinLuo = New PinLuo_Class
    PinLuo.DBConnBegin
	PinLuo.CheckPurview
	PinLuo.Pinluo_CheckPurviewAdmin(3)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="images/style.css" type="text/css">
<title>管理用户管理</title>
<meta content="品络科技,网站管理系统,企业网站管理系统,内容管理系统(CMS),网上商店管理系统,网站建设" name="Keywords" />
<meta content="品络科技成立于2005年6月，是一家集互联网基础服务、互联网应用软件开发、业务解决方案销售及服务于一体的高新技术企业。公司网址：www.5300.cn，品络互联：www.pinluo.com" name="Description"/>
<script language="JavaScript">
<!--
function CheckOthers(form)
{
	for (var i=0;i<form.elements.length;i++)
	{
		var e = form.elements[i];
			if (e.checked==false)
			{
				e.checked = true;
			}
			else
			{
				e.checked = false;
			}
	}
}

function CheckAll(form)
{
	for (var i=0;i<form.elements.length;i++)
	{
		var e = form.elements[i];
			e.checked = true
	}
}

function submit_to(url){
document.PinLuo_Adminlist.action=url;
document.PinLuo_Adminlist.submit();
}
//-->
</script>
</head>
<body class="mainbg">
	<div id="mainhearder"><span>您的位置：企业网站管理系统 >> 管理员管理</span></div>
	
	<div id="hearder" class="hearder1"><span>管理员列表</span></div>
	
   <form action="PinLuo_AdminList.asp" method="post" name="PinLuo_Adminlist" id="PinLuo_Adminlist">
   <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
   <tr>
	<td class="tableleft">
        <table width="100%" border="0" cellpadding="0" cellspacing="0" class="table">
		<tr class="stline_one">
		<td width="5%" class="heardertop1">&nbsp;</td>
		<td width="10%" class="heardertop1">编号</td>
		<td width="17%" class="heardertop1">用户帐号</td>
		<td width="13%" class="heardertop1">真实姓名</td>
		<td width="19%" class="heardertop1">登陆IP</td>
		<td width="14%" class="heardertop1">登陆时间</td>
		<td width="8%" class="heardertop1">登陆次数</td>
		<td width="8%" class="heardertop1">是否启用</td>
		<td width="6%" class="heardertop1">修改</td>
		</tr>
		
		<%PinLuo.PinLuo_AdminList OnepageNum,SearchKeyword,SearchSelect,"PinLuo_Admin"%>
		
		</table>
	</td>
	</tr>
	</table>
	<div id="page">
		<div id="add">
			<input name="button2" type="button" class="buttonnor" value="全选" onClick="CheckAll(this.form);">
			<input name="button2" type="button" class="buttonnor" value="反选" onClick="CheckOthers(this.form)">
            <input name="button2" type="button" class="buttonnor" value="刷新" onClick="window.location.reload();">
            <input name="button1" type="button" class="buttonadd" value="新增" onClick="location.href='Pinluo_Admin.asp?Act=add';">
			<input name="button4" type="button" class="buttondel" value="删除" onClick="submit_to('Pinluo_Admin.asp?Act=del');">
		</div>
        <div>
        <%=Pinluo.Pinluo_showpage_temp%>
        </div>
	</div>    

	</form><br />
<br />
<br />
&nbsp;
</body>
</html>
<%
PinLuo.DBConnEnd
Set PinLuo = Nothing
%>