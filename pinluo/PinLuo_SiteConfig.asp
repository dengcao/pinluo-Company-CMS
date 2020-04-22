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
<script src="js/jquery-1.10.2.min.js"></script>
<script type="text/javascript" charset="utf-8" src="../ueditor/ueditor.config.js"></script>
<script type="text/javascript" charset="utf-8" src="../ueditor/ueditor.all.min.js"> </script>
<script type="text/javascript" charset="utf-8" src="../ueditor/lang/zh-cn/zh-cn.js"></script>
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
		  <td width="85%" class="tableright1"><input type='text' size='55' maxlength='255' name='Pinluo_Logo' id="Pinluo_Logo" class='input' value="<%=PinLuo.Pinluo_Logo%>"> <img src="Images/openimg.gif" title="查看图片" onclick="window.open(document.getElementById('Pinluo_Logo').value);" style="cursor:pointer" /> <input type=button class=btn value="上传图片" onclick="upImage_media_url();">	  
			

			<script type="text/plain" id="upload_ue_media_url"></script>                              
 <script type="text/javascript">
var _editor_media_url;

$(function() {

//重新实例化一个编辑器，防止在上面的editor编辑器中显示上传的图片或者文件
_editor_media_url = UE.getEditor('upload_ue_media_url');
_editor_media_url.ready(function () {
//设置编辑器不可用
_editor_media_url.setDisabled();

//隐藏编辑器，因为不会用到这个编辑器实例，所以要隐藏

_editor_media_url.hide();

//侦听图片上传

_editor_media_url.addListener('beforeInsertImage', function (t, arg) {

//将地址赋值给相应的input,只去第一张图片的路径

	//document.getElementById("media_url").value=arg[0].src;
$("#Pinluo_Logo").attr("value", arg[0].src);

//图片预览

//$("#preview").attr("src", arg[0].src);

})



});

}); 

//弹出图片上传的对话框

function upImage_media_url() {

var myImage_media_url = _editor_media_url.getDialog("insertimage");

myImage_media_url.open();

}

</script>
			  
			  </td>
		  </tr>  
         <tr>
		  <td width="15%" class="tableleft1"><strong>网站Banner：</strong></td>
		  <td width="85%" class="tableright1"><input type='text' size='55' maxlength='255' name='Pinluo_Banner' id="Pinluo_Banner" class='input' value="<%=PinLuo.Pinluo_Banner%>"> <img src="Images/openimg.gif" title="查看图片" onclick="window.open(document.getElementById('Pinluo_Banner').value);" style="cursor:pointer" /> <input type=button class=btn value="上传图片" onclick="upImage_media_url2();">
			 
			 <script type="text/plain" id="upload_ue_media_url2"></script>                              
 <script type="text/javascript">
var _editor_media_url2;

$(function() {

//重新实例化一个编辑器，防止在上面的editor编辑器中显示上传的图片或者文件
_editor_media_url2 = UE.getEditor('upload_ue_media_url2');
_editor_media_url2.ready(function () {
//设置编辑器不可用
_editor_media_url2.setDisabled();

//隐藏编辑器，因为不会用到这个编辑器实例，所以要隐藏

_editor_media_url2.hide();

//侦听图片上传

_editor_media_url2.addListener('beforeInsertImage', function (t, arg) {

//将地址赋值给相应的input,只去第一张图片的路径

	//document.getElementById("media_url").value=arg[0].src;
$("#Pinluo_Banner").attr("value", arg[0].src);

//图片预览

//$("#preview").attr("src", arg[0].src);

})



});

}); 

//弹出图片上传的对话框

function upImage_media_url2() {

var myImage_media_url2 = _editor_media_url2.getDialog("insertimage");

myImage_media_url2.open();

}

</script>
			 </td>
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