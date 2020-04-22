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
response.Write("<script language=""javascript"">alert('更新网站配置成功！');location.href='PinLuo_SiteConfig.asp';</script>")
PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
else
PinLuo.PinLuo_ViewSiteConfig
end if
%>
<title>网站配置</title>
<link rel="stylesheet" href="images/style.css" type="text/css">
<script language="javascript" src="js/Pinluo.js" type="text/javascript"></script>
<script Language=javascript>
// 参数说明
// s_Type : 文件类型，可用值为"image","flash","media","file"
// s_Link : 文件上传后，用于接收上传文件路径文件名的表单名
// s_Thumbnail : 文件上传后，用于接收上传图片时所产生的缩略图文件的路径文件名的表单名，当未生成缩略图时，返回空值，原图用s_Link参数接收，此参数专用于缩略图
function showUploadDialog(s_Type, s_Link, s_Thumbnail){
	//以下style=coolblue,值可以依据实际需要修改为您的样式名,通过此样式的后台设置来达到控制允许上传文件类型及文件大小
	var arr = showModalDialog("../editor/dialog/i_upload.htm?style=coolblue&type="+s_Type+"&link="+s_Link+"&thumbnail="+s_Thumbnail, window, "dialogWidth:0px;dialogHeight:0px;help:no;scroll:no;status:no");
}
</script>
</head>

<body class="mainbg">
	<div id="mainhearder"><span>您的位置：企业网站管理系统 >> 网站配置</span></div>
	<div id="hearder" class="hearder2"><span>网站配置</span></div>
	<div class="main5" id="main5">


	<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
	<tr>
	<td class="tableleft">
    
                    <form name="form1" method="post" action="PinLuo_SiteConfig.asp">
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="table1">
		<tr>
		  <td width="15%" class="tableleft1"><strong>网站名称：</strong></td>
		  <td width="85%" class="tableright1"><input type='text' size='55' maxlength='255' name='Pinluo_SiteName' class='input' value="<%=PinLuo.Pinluo_SiteName%>"></td>
		  </tr>
        <tr>
		  <td width="15%" class="tableleft1"><strong>全站SEO标题：</strong></td>
		  <td width="85%" class="tableright1"><input type='text' size='55' maxlength='255' name='Pinluo_SeoTitle' class='input' value="<%=PinLuo.Pinluo_SeoTitle%>"></td>
		  </tr>  
         <tr>
		  <td width="15%" class="tableleft1"><strong>网站地址：</strong></td>
		  <td width="85%" class="tableright1"><input type='text' size='55' maxlength='255' name='Pinluo_SiteUrl' class='input' value="<%=PinLuo.Pinluo_SiteUrl%>">
		  末尾带/</td>
		  </tr>
          
          <tr>
		  <td width="15%" class="tableleft1"><strong>网站LOGO：</strong></td>
		  <td width="85%" class="tableright1"><input type='text' size='55' maxlength='255' name='Pinluo_Logo' id="Pinluo_Logo" class='input' value="<%=PinLuo.Pinluo_Logo%>"> <img src="Images/openimg.gif" title="查看图片" onclick="window.open(document.getElementById('Pinluo_Logo').value);" style="cursor:pointer" /> <input type=button class=btn value="上传图片" onclick="showUploadDialog('image', 'form1.Pinluo_Logo', '')"></td>
		  </tr>  
         <tr>
		  <td width="15%" class="tableleft1"><strong>网站Banner：</strong></td>
		  <td width="85%" class="tableright1"><input type='text' size='55' maxlength='255' name='Pinluo_Banner' class='input' value="<%=PinLuo.Pinluo_Banner%>"> <img src="Images/openimg.gif" title="查看图片" onclick="window.open(document.getElementById('Pinluo_Banner').value);" style="cursor:pointer" /> <input type=button class=btn value="上传图片" onclick="showUploadDialog('image', 'form1.Pinluo_Banner', '')"></td>
		  </tr>
          
        <tr>
		  <td width="15%" class="tableleft1"><strong>首页SEO标题：</strong></td>
		  <td width="85%" class="tableright1"><input type='text' size='55' maxlength='255' name='Pinluo_SeoIndexTitle' class='input' value="<%=PinLuo.Pinluo_SeoIndexTitle%>"></td>
		  </tr>  
          <tr>
		  <td width="15%" class="tableleft1"><strong>首页SEO关键词：</strong></td>
		  <td width="85%" class="tableright1"><input type='text' size='55' maxlength='255' name='Pinluo_SeoIndexKeyword' class='input' value="<%=PinLuo.Pinluo_SeoIndexKeyword%>"> 多个关键词用,分隔</td>
		  </tr>
        <tr>
		  <td width="15%" class="tableleft1"><strong>首页SEO描述：</strong></td>
		  <td width="85%" class="tableright1"><input type='text' size='55' name='Pinluo_SeoIndexMS' class='input' value="<%=PinLuo.Pinluo_SeoIndexMS%>"></td>
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
		  <td height="55" align="center" class="tableleft1" style="height:39px;">&nbsp;</td>
		  <td height="55" class="tableright1">
          <input name="act" type="hidden" value="editsave" />
          <input type="submit" name="submit" value="完成修改" class="button">
		    <input type="reset" value="重新填写" class="button" name="reset" tabindex="25">
		    <input type="button" value="刷新本页" class="button" name="button" tabindex="25" onclick="location.reload()"></td>
		  </tr>

		</table>
        </form>
       		
	</td>
	</tr>
	</table>

<%
PinLuo.DBConnEnd
Set PinLuo = Nothing
%>	
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