<!--#include file="../Pinluo_Main/Conn.asp"-->
<!--#include file="PinLuo_Class.asp"-->
<%
    If Request.Cookies("pinluo")("UserID") = "" or Request.Cookies("pinluo")("UserName") = "" then
	   response.Redirect("PL_Logout.asp")
	   response.End()
	End If
dim PinLuo
Set PinLuo = New PinLuo_Class
    PinLuo.DBConnBegin
	PinLuo.CheckPurview
	PinLuo.PinLuo_ViewSiteConfig
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="images/style.css" type="text/css">
<title><%=Pinluo.Pinluo_SiteName%>网站后台管理系统</title>
<meta content="品络科技,网站管理系统,企业网站管理系统,内容管理系统(CMS),网上商店管理系统,网站建设" name="Keywords" />
<meta content="品络科技成立于2005年6月，是一家集互联网基础服务、互联网应用软件开发、业务解决方案销售及服务于一体的高新技术企业。公司网址：www.5300.cn，品络互联：www.pinluo.com" name="Description"/>
</head>
<FRAMESET border=0 name=top_frame frameSpacing=0 rows=99,* frameBorder=0 scrolling=auto>
<FRAME name=head marginWidth=0 marginHeight=0 src="PinLuo_Top.asp" frameBorder=0 id="head">
<FRAMESET border=0 name=bottom_frame cols=180,6,*>
<FRAME name=menu marginWidth=0 style="OVERFLOW:auto;OVERFLOW-X:HIDDEN;" marginHeight=0 src="PinLuo_Menu.asp" noResize>
<FRAME name=partition marginWidth=0 marginHeight=0 src="PinLuo_Partition.asp">
<FRAME name=main marginWidth=0 marginHeight=0 src="PinLuo_main.asp">
</FRAMESET>
</FRAMESET>
<noframes></noframes>
</html>
<%PinLuo.DBConnEnd
Set PinLuo = Nothing%>