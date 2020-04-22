<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!--#include file="Pinluo_Main/Conn.asp"-->
<!--#include file="Pinluo/PinLuo_Class.asp"-->
<%
'☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆
'☆                                                                         ☆
'☆  系 统：品络企业网站管理系统 Version 1.5                                    ☆
'☆  日 期：2010-05                                                          ☆
'☆  开 发：草札(www.caozha.com)                                              ☆
'☆  声 明: 使用本系统必须保留此版权声明信息！本文字不会影响系统的正常运行。            ☆
'☆  Copyright (C) 2010 品络(www.pinluo.com) All Rights Reserved.            ☆
'☆                                                                         ☆
'☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆
 
act=Trim(Request("act"))
dim PinLuo
Set PinLuo = New PinLuo_Class
    PinLuo.DBConnBegin
	
if act="editsave" then
    sqlDbPath_ml = Trim(Request.Form("sqlDbPath_ml"))
	Pinluo_SiteName = Trim(Request.Form("Pinluo_SiteName"))
	Pinluo_SeoTitle = Trim(Request.Form("Pinluo_SeoTitle"))
	Pinluo_DelProImg = Trim(Request.Form("Pinluo_DelProImg"))
	Pinluo_IsFeedback = Trim(Request.Form("Pinluo_IsFeedback"))
	Pinluo_SiteUrl = Trim(Request.Form("Pinluo_SiteUrl"))
	Pinluo_SeoIndexTitle = Trim(Request.Form("Pinluo_SeoIndexTitle"))
	Pinluo_SeoIndexKeyword = Trim(Request.Form("Pinluo_SeoIndexKeyword"))
	Pinluo_SeoIndexMS = Trim(Request.Form("Pinluo_SeoIndexMS"))
	Pinluo_Logo = Trim(Request.Form("Pinluo_Logo"))
	Pinluo_Banner = Trim(Request.Form("Pinluo_Banner"))
   if Pinluo_SiteName="" then response.Write("<script language=""javascript"">alert('请填写网站名称！');window.history.back();</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
PinLuo.PinLuo_SiteConfigEdit Pinluo_SiteName,Pinluo_SeoTitle,Pinluo_DelProImg,Pinluo_IsFeedback,Pinluo_SiteUrl,Pinluo_SeoIndexTitle,Pinluo_SeoIndexKeyword,Pinluo_SeoIndexMS,Pinluo_Logo,Pinluo_Banner
Pinluo_ValidateCode=PinLuo.GetPinluo_ValidateCode(30)&".accdb"
PinLuo.DBConnEnd:Set PinLuo=Nothing
fileUpdate sqlDbPath_access,Pinluo_ValidateCode
WriteFile sqlDbPath_ml&"Pinluo_Main/database/"&Pinluo_ValidateCode
Installnow=true
else
Installnow=false
PinLuo.PinLuo_ViewSiteConfig
end if
%>
<title>安装 - 品络企业网站系统</title>
<meta content="品络科技,网站管理系统,企业网站管理系统,内容管理系统(CMS),网上商店管理系统,网站建设" name="Keywords" />
<meta content="品络科技成立于2005年6月，是一家集互联网基础服务、互联网应用软件开发、业务解决方案销售及服务于一体的高新技术企业。开发者：www.caozha.com，品络互联：www.pinluo.com" name="Description"/>
<link rel="stylesheet" href="Pinluo/images/style.css" type="text/css">
<script language="javascript" src="Pinluo/js/Pinluo.js" type="text/javascript"></script>
</head>

<body class="mainbg" style="text-align:center;">
	<div style="width:650px; text-align:left; margin:auto; margin-top:20px;">
	<div id="hearder" class="hearder2"><span>在线安装 - 品络企业网站系统</span></div>
	<div class="main5" id="main5">

<%if Installnow=false then%>
	<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
	<tr>
	<td class="tableleft1">
    
        <table width="100%" border="0" align="left" cellpadding="0" cellspacing="0" class="table1">
		<tr>
		<td style="padding:12px; text-align:left; font-weight:normal; font-size:13px;">
        
        <form name="form1" method="post" action="Install.asp">
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="table1">
		<tr>
		  <td width="19%" class="tableleft1"><strong>网站名称：</strong></td>
		  <td width="81%" class="tableright1"><input type='text' size='55' maxlength='255' name='Pinluo_SiteName' class='input' value="<%=PinLuo.Pinluo_SiteName%>"></td>
		  </tr>
        <tr>
		  <td width="19%" class="tableleft1"><strong>全站SEO标题：</strong></td>
		  <td width="81%" class="tableright1"><input type='text' size='55' maxlength='255' name='Pinluo_SeoTitle' class='input' value="<%=PinLuo.Pinluo_SeoTitle%>"></td>
		  </tr>  
         <tr>
		  <td width="19%" class="tableleft1"><strong>网站地址：</strong></td>
		  <td width="81%" class="tableright1"><input type='text' size='55' maxlength='255' name='Pinluo_SiteUrl' class='input' value="<%=PinLuo.GetServerPath(0)%>">
		  建议默认，末尾带/</td>
		  </tr>
          
          <tr style="display:none">
		  <td width="19%" class="tableleft1"><strong>网站LOGO：</strong></td>
		  <td width="81%" class="tableright1"><input type='text' size='55' maxlength='255' name='Pinluo_Logo' id="Pinluo_Logo" class='input' value="<%=PinLuo.Pinluo_Logo%>"></td>
		  </tr>  
         <tr style="display:none">
		  <td width="19%" class="tableleft1"><strong>网站Banner：</strong></td>
		  <td width="81%" class="tableright1"><input type='text' size='55' maxlength='255' name='Pinluo_Banner' class='input' value="<%=PinLuo.Pinluo_Banner%>"></td>
		  </tr>
          
        <tr>
		  <td width="19%" class="tableleft1"><strong>首页SEO标题：</strong></td>
		  <td width="81%" class="tableright1"><input type='text' size='55' maxlength='255' name='Pinluo_SeoIndexTitle' class='input' value="<%=PinLuo.Pinluo_SeoIndexTitle%>"></td>
		  </tr>  
          <tr>
		  <td width="19%" class="tableleft1"><strong>首页SEO关键词：</strong></td>
		  <td width="81%" class="tableright1"><input type='text' size='55' maxlength='255' name='Pinluo_SeoIndexKeyword' class='input' value="<%=PinLuo.Pinluo_SeoIndexKeyword%>"> 多个关键词用,分隔</td>
		  </tr>
        <tr>
		  <td width="19%" class="tableleft1"><strong>首页SEO描述：</strong></td>
		  <td width="81%" class="tableright1"><input type='text' size='55' name='Pinluo_SeoIndexMS' class='input' value="<%=PinLuo.Pinluo_SeoIndexMS%>"></td>
		  </tr>  
		<tr>
		<td class="tableleft1"><strong>删除产品时：</strong></td>
		<td class="tableright1"><label>
		  <input type="radio" name="Pinluo_DelProImg" id="radio" value="true"<%if PinLuo.Pinluo_DelProImg=true then%> checked="checked"<%end if%> />
		  删除图片&nbsp;&nbsp;&nbsp;
		</label>
		  <input name="Pinluo_DelProImg" type="radio" id="radio2" value="false"<%if PinLuo.Pinluo_DelProImg=false then%> checked="checked"<%end if%> /> 
		  保留图片&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(建议保留)
</td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>留言：</strong></td>
		<td class="tableright1"><label>
		  <input type="radio" name="Pinluo_IsFeedback" id="radio" value="true"<%if PinLuo.Pinluo_IsFeedback=true then%> checked="checked"<%end if%> />
		  开放&nbsp;&nbsp;&nbsp;
		</label>
		  <input name="Pinluo_IsFeedback" type="radio" id="radio2" value="false"<%if PinLuo.Pinluo_IsFeedback=false then%> checked="checked"<%end if%> /> 
		  关闭
</td>
		</tr>
        <tr>
          <td class="tableleft1">默认管理员账号：</td>
          <td class="tableright1">&nbsp;&nbsp;&nbsp;admin</td>
        </tr>
        <tr>
          <td class="tableleft1">默认管理员密码：</td>
          <td class="tableright1">&nbsp;&nbsp;&nbsp;pinluo  &nbsp;&nbsp;&nbsp;(安装成功后，请马上登录后台修改)</td>
        </tr>
        <tr>
          <td class="tableleft1" style="color:red;font-weight:bold;">特别注意事项：</td>
          <td class="tableright1" style="color:red;font-weight:bold;">&nbsp;&nbsp;&nbsp;网站服务器的IIS必须开启父路径，否则程序会执行错误。<a href="https://my.oschina.net/dengzhenhua/blog/3295146" target=_blank>点此查看开启方法</a></td>
        </tr>
        
		<tr>
		  <td height="55" align="center" class="tableleft1" style="height:39px;">&nbsp;</td>
		  <td height="55" class="tableright1">
          <input name="sqlDbPath_ml" type="hidden" value="<%=PinLuo.GetServerPath(0)%>" />
          <input name="act" type="hidden" value="editsave" />
          <input type="submit" name="submit" value="马上安装" class="button">
		    <input type="reset" value="重置" class="button" name="reset" tabindex="25">
</td>
		  </tr>

		</table>
        </form>
       </td>
		</tr>

		</table>
       		
	</td>
	</tr>
	</table>
    <%
PinLuo.DBConnEnd:Set PinLuo=Nothing
else%>
    <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
	<tr>
	<td class="tableleft1">
    
    <table width="100%" border="0" align="left" cellpadding="0" cellspacing="0" class="table1">
		<tr>
		<td style="padding:12px; text-align:left; font-weight:normal; font-size:13px;"><font color=red><b>系统已经安装成功！</b></font><br><br><br /><center><input name='button' onclick="location.href='<%=sqlDbPath_ml&"pinluo/PL_login.asp"%>';" class=button type=button value='登录后台'>&nbsp;&nbsp;<input name='button' onclick="location.href='<%=sqlDbPath_ml%>index.asp';" class=button type=button value='浏览首页'></center></td>
		</tr>

		</table>
        
        </td>
	</tr>
	</table>
    <%end if%>
    
	</div>
            <br />
<br />
&nbsp; </div>
</body>
</html>
<%
if Installnow=true then
   fileUpdate "index.asp.bak","index.asp"
   fileUpdate "index.html","index.html.bak"
   fileUpdate "Install.asp","Install.asp.bak"
end if

Function WriteFile(path)
  Dim st   
  Set st=Server.CreateObject("ADODB.Stream")   
  st.Type=2   
  st.Mode=3   
  st.Charset="utf-8"
  st.Open()
	st.WriteText "<" & "%" & vbcrlf
	st.WriteText "'设置数据库路径，草札 www.caozha.com" & vbcrlf
	st.WriteText "sqlDbPath_access=" & chr(34) & path & chr(34) & vbcrlf
	st.WriteText "%" & ">"
  st.SaveToFile Server.MapPath("Pinluo_Main/Site_Conn.asp"),2
  st.Close()
  Set st=Nothing
End Function

Function fileUpdate(filePath,extendName) 
Set fso = Server.CreateObject("Scripting.FileSystemObject") 
Set f = fso.GetFile(Server.MapPath(filePath))
f.name = extendName
newname=f.name
fileUpdate=newname
End Function 

%>	