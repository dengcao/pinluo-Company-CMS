<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Pinluo_Main/Conn.asp"-->
<!--#include file="PinLuo_Class.asp"-->
<%
dim PinLuo
Set PinLuo = New PinLuo_Class
    PinLuo.DBConnBegin
	PinLuo.CheckPurview
	PinLuo.DBConnEnd
Set PinLuo = Nothing
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>头部</title>
<link rel="stylesheet" href="images/style.css" type="text/css">
</head>
<body class="bg">
	<div id="top">
		<div id="topleft"><span><a href="http://www.pinluo.com" target="_blank"><img src="images/logo.png" border="0"></a></span></div>
		<div id="topright">
			<div id="nav">
				
                <UL>
                   <LI><A href="../" target=_blank>首页</A> 
                   <LI><A href="PinLuo_ProductList.asp" target=main>产品管理</A> 
                   <LI><A href="PinLuo_InfoList.asp" target=main>信息管理</A> 
                   <LI><A href="PinLuo_ItemClass.asp" target=main>栏目管理</A> 
                   <LI><A href="http://www.pinluo.com/" target=_blank>官方网站</A> 
                   <LI><A href="http://www.pinluo.com/" target=_blank>域名主机</A> 
                   <LI><A href="http://www.caozha.com" target=_blank>作者博客</A> </LI></UL>
               
			</div>
		</div>
	</div>
	
	<div id="mebinfobg">
		<div id="mebinfo">您好，管理员：<%=Request.Cookies("pinluo")("RealName")%>　<a href="PinLuo_EditMyInfo.asp" target="main">登陆系统：<span class="ye1"><%=Request.Cookies("pinluo")("LoginTimes")%></span> 次</a>　<a href="PinLuo_main.asp" target="main"><span class="jxs">系统首页</span></a>　<a href="Pl_Logout.asp" target="_top"><span class="exit">退出</span></a></div>
	</div>
	
</body>
</html>