<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Pinluo_Main/Conn.asp"-->
<!--#include file="PinLuo_Class.asp"-->
<%
dim PinLuo
Set PinLuo = New PinLuo_Class
    PinLuo.DBConnBegin
	PinLuo.CheckPurview
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>菜单</title>
<link rel="stylesheet" href="images/style.css" type="text/css">
<script language="JavaScript" type="text/JavaScript">
<!--
function toggle(targetid){
	if (document.getElementById){
		target=document.getElementById(targetid);
			if (target.style.display=="block"){
				target.style.display="none";
			} else {
				target.style.display="block";
			}
	}
}
-->
</script>
<script language="javascript" src="include/xmlhttp.js"></script>
</head>
<body class="leftbg">
<table width="180" border="0" cellspacing="0" cellpadding="0">
	
	<tr>
		<td>
			<div id="menuuserinfo" onclick="javascript:location.href='?Action=MenuOff'"></div>
			<div class="usermain" style="display:block">
				<ul>
					<li>账 号：<%=Request.Cookies("pinluo")("UserName")%></li>
					<li>姓 名：<%=Request.Cookies("pinluo")("RealName")%></li>
				</ul>
				<div id="userbottom"></div>
			</div>
		</td>
	</tr>
    
    	<tr>
		<td>
			<div id="menutop" onclick="toggle('div5')">
				<span class="zhgl">网站设置</span>
			</div>
			<div class="menumain" id="div5" style="display:;">
				<ul>
                    <li style="display:<%=Pinluo.Pinluo_CheckPurviewAdmin_Display(0)%>;"><span onMouseOver="this.className = 'leftmenubg'" onMouseOut="this.className='none'"><a href="PinLuo_SiteConfig.asp" target="main">网站设置</a></span></li>
                    <li style="display:<%=Pinluo.Pinluo_CheckPurviewAdmin_Display(2)%>;"><span onMouseOver="this.className = 'leftmenubg'" onMouseOut="this.className='none'"><a href="PinLuo_daohangList.asp" target="main">网站导航</a></span></li>
					<li style="display:<%=Pinluo.Pinluo_CheckPurviewAdmin_Display(2)%>;"><span onMouseOver="this.className = 'leftmenubg'" onMouseOut="this.className='none'"><a href="PinLuo_ItemClass.asp" target="main">栏目管理</a></span></li>
                    <li style="display:<%=Pinluo.Pinluo_CheckPurviewAdmin_Display(2)%>;"><span onMouseOver="this.className = 'leftmenubg'" onMouseOut="this.className='none'"><a href="PinLuo_BlockList.asp" target="main">网站碎片</a></span></li>
                    <!--li style="display:<%=Pinluo.Pinluo_CheckPurviewAdmin_Display(0)%>;"><span onMouseOver="this.className = 'leftmenubg'" onMouseOut="this.className='none'"><a href="PinLuo_Navigation.asp" target="main">后台快捷菜单</a></span></li-->
				</ul>
				<div id="menubottom"></div>
			</div>
		</td>
	</tr>

	
	<tr style="display:<%=Pinluo.Pinluo_CheckPurviewAdmin_Display(4)%>;">
		<td>
			<div id="menutop" onclick="toggle('div1')">
				<span class="jh">产品管理</span>
			</div>
			<div class="menumain" id="div1">
				<ul>
					<li><span onMouseOver="this.className = 'leftmenubg'" onMouseOut="this.className='none'"><a href="PinLuo_ProductClass.asp" target="main">产品分类</a></span></li>
                    <li><span onMouseOver="this.className = 'leftmenubg'" onMouseOut="this.className='none'"><a href="PinLuo_ProductList.asp" target="main">产品列表</a></span></li>
                    <li><span onMouseOver="this.className = 'leftmenubg'" onMouseOut="this.className='none'"><a href="Pinluo_Product.asp?Act=add" target="main">添加产品</a></span></li>
				</ul>
				<div id="menubottom"></div>
			</div>
		</td>
	</tr>
	
	<tr style="display:<%=Pinluo.Pinluo_CheckPurviewAdmin_Display(5)%>;">
		<td>
			<div id="menutop" onclick="toggle('div2')">
				<span class="xsjl">信息管理</span>
			</div>
			<div class="menumain" id="div2">
				<ul>
					<li><span onMouseOver="this.className = 'leftmenubg'" onMouseOut="this.className='none'"><a href="PinLuo_InfoClass.asp" target="main">信息分类</a></span></li>
                    <li><span onMouseOver="this.className = 'leftmenubg'" onMouseOut="this.className='none'"><a href="PinLuo_InfoList.asp" target="main">信息列表</a></span></li>
                    <li><span onMouseOver="this.className = 'leftmenubg'" onMouseOut="this.className='none'"><a href="Pinluo_info.asp?Act=add" target="main">添加信息</a></span></li>
				</ul>
				<div id="menubottom"></div>
			</div>
		</td>
	</tr>

	<tr style="display:<%=Pinluo.Pinluo_CheckPurviewAdmin_Display(6)%>;">
		<td>
			<div id="menutop" onclick="toggle('div3')">
				<span class="xtzh">留言反馈</span>
			</div>
			<div class="menumain" id="div3">
				<ul>
					<li><span onMouseOver="this.className = 'leftmenubg'" onMouseOut="this.className='none'"><a href="PinLuo_FeedbackClass.asp" target="main">留言分类</a></span></li>
                    <li><span onMouseOver="this.className = 'leftmenubg'" onMouseOut="this.className='none'"><a href="PinLuo_FeedbackList.asp" target="main">留言列表</a></span></li>
                    <li><span onMouseOver="this.className = 'leftmenubg'" onMouseOut="this.className='none'"><a href="Pinluo_Feedback.asp?Act=add" target="main">添加留言</a></span></li>
				</ul>
				<div id="menubottom"></div>
			</div>
		</td>
	</tr>
	
	<tr>
		<td>
			<div id="menutop" onclick="toggle('div4')">
				<span class="aqgl">用户管理</span>
			</div>
			<div class="menumain" id="div4" style="display:none">
				<ul>
					<li style="display:<%=Pinluo.Pinluo_CheckPurviewAdmin_Display(3)%>;"><span onMouseOver="this.className = 'leftmenubg'" onMouseOut="this.className='none'"><a href="PinLuo_AdminList.asp" target="main">管理员列表</a></span></li>
                    <li style="display:<%=Pinluo.Pinluo_CheckPurviewAdmin_Display(3)%>;"><span onMouseOver="this.className = 'leftmenubg'" onMouseOut="this.className='none'"><a href="Pinluo_Admin.asp?Act=add" target="main">管理员添加</a></span></li>
                    <li><span onMouseOver="this.className = 'leftmenubg'" onMouseOut="this.className='none'"><a href="PinLuo_EditMyInfo.asp" target="main">修改我的资料</a></span></li>
				</ul>
				<div id="menubottom"></div>
			</div>
		</td>
	</tr>


	<tr>
		<td>
			<div id="menutop" onclick="toggle('div6');">
				<span class="sygj">实用工具</span>
			</div>
			<div class="menumain" id="div6" style="display:none">
				<ul>					
                    <li style="display:<%=Pinluo.Pinluo_CheckPurviewAdmin_Display(1)%>;"><span onMouseOver="this.className = 'leftmenubg'" onMouseOut="this.className='none'"><a href="PinLuo_Database.asp?Action=Backup" target="main">数据库备份</a></span></li>
                    <li style="display:<%=Pinluo.Pinluo_CheckPurviewAdmin_Display(1)%>;"><span onMouseOver="this.className = 'leftmenubg'" onMouseOut="this.className='none'"><a href="PinLuo_Database.asp?Action=Compact" target="main">压缩数据库</a></span></li>
                    <li style="display:<%=Pinluo.Pinluo_CheckPurviewAdmin_Display(1)%>;"><span onMouseOver="this.className = 'leftmenubg'" onMouseOut="this.className='none'"><a href="PinLuo_Database.asp?Action=SpaceSize" target="main">空间统计</a></span></li>
                    <li><span onMouseOver="this.className = 'leftmenubg'" onMouseOut="this.className='none'"><a href="http://diannao.wang/tool/" target="_blank">工具大全</a></span></li>
                    <li><span onMouseOver="this.className = 'leftmenubg'" onMouseOut="this.className='none'"><a href="http://www.pinluo.com/" target="_blank">域名空间</a></span></li>
                    <li><span onMouseOver="this.className = 'leftmenubg'" onMouseOut="this.className='none'"><a href="http://www.caozha.com/" target="_blank">作者博客</a></span></li>
				</ul>
				<div id="menubottom"></div>
			</div>
		</td>
	</tr>
	
	<tr>
		<td>
			<div id="menutop" onclick="toggle('div7')">
				<span class="kf">服务支持</span>
			</div>
			<div class="menumain" id="div7">
				<ul>
					<li><span onMouseOver="this.className = 'leftmenubg'" onMouseOut="this.className='none'"><a href="http://www.pinluo.com/?help/?classid=26" target="_blank">在线帮助</a></span></li>
                    <li><span onMouseOver="this.className = 'leftmenubg'" onMouseOut="this.className='none'"><a href="http://www.pinluo.com/?cms/qycms/template.asp" target="_blank">模板下载</a></span></li>
                    <li><span onMouseOver="this.className = 'leftmenubg'" onMouseOut="this.className='none'"><a href="http://www.pinluo.com/?cms/qycms/update.asp" target="_blank">升级更新</a></span></li>
                    <li><span onMouseOver="this.className = 'leftmenubg'" onMouseOut="this.className='none'"><a href="http://www.pinluo.com/?cms/qycms/license.asp" target="_blank">版权声明</a></span></li>
				</ul>
				<div id="menubottom"></div>
			</div>
		</td>
	</tr>
	
</table>

<script>
if (screen.width<=1024)
{
document.getElementById("div3").style.display="none";
document.getElementById("div4").style.display="none";
}
</script>

</body>
</html>
<%
PinLuo.DBConnEnd
Set PinLuo = Nothing
%>