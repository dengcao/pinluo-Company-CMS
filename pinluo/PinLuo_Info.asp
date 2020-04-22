<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Pinluo_Main/Conn.asp"-->
<!--#include file="PinLuo_Class.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<%
act=Trim(Request("act"))
infoid=Trim(Request("infoid"))
Classid=Trim(Request("Classid"))
SearchKeyword2 = Trim(Request("SearchKeyword"))
SearchSelect2 = Trim(Request("SearchSelect"))
page2 = Trim(Request("page"))
dim PinLuo
Set PinLuo = New PinLuo_Class
    PinLuo.DBConnBegin
	PinLuo.CheckPurview
	PinLuo.Pinluo_CheckPurviewAdmin(5)
	
if act="add" then
InfoItemName="添加信息"
elseif act="edit" then
InfoItemName="修改信息"
elseif act="addsave" then
	  InfoTitle=Trim(Request("InfoTitle"))
	  SEO_Title=Trim(Request("SEO_Title"))
	  SEO_Keyword=Trim(Request("SEO_Keyword"))
	  SEO_Description=Trim(Request("SEO_Description"))
	  InfoContent=Trim(Request("InfoContent"))
	  InfoImg=Trim(Request("InfoImg"))
	  Author=Trim(Request("Author"))
	  Origin=Trim(Request("Origin"))
	  UpdateTime=Trim(Request("UpdateTime"))
	  hits=Trim(Request("hits"))
	  OrderID=Trim(Request("OrderID"))
	  Shenhe=eval(Request("Shenhe"))
if Classid="" then response.Write("<script language=""javascript"">alert('请选择信息分类！');window.history.back();</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
if InfoTitle="" then response.Write("<script language=""javascript"">alert('请填写标题！');window.history.back();</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
if InfoContent="" then response.Write("<script language=""javascript"">alert('请填写内容！');window.history.back();</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
PinLuo.PinLuo_AddInfo Classid,InfoTitle,SEO_Title,SEO_Keyword,SEO_Description,InfoContent,InfoImg,Author,Origin,UpdateTime,hits,OrderID,Shenhe,"PinLuo_InfoList"
response.Write("<script language=""javascript"">alert('添加信息成功！');location.href='PinLuo_Info.asp?act=add&classid="&Classid&"';</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
elseif act="editsave" then
	  InfoTitle=Trim(Request("InfoTitle"))
	  SEO_Title=Trim(Request("SEO_Title"))
	  SEO_Keyword=Trim(Request("SEO_Keyword"))
	  SEO_Description=Trim(Request("SEO_Description"))
	  InfoContent=Trim(Request("InfoContent"))
	  InfoImg=Trim(Request("InfoImg"))
	  Author=Trim(Request("Author"))
	  Origin=Trim(Request("Origin"))
	  UpdateTime=Trim(Request("UpdateTime"))
	  hits=Trim(Request("hits"))
	  OrderID=Trim(Request("OrderID"))
	  Shenhe=eval(Request("Shenhe"))
if isnumeric(Infoid)=false then response.Write("<script language=""javascript"">alert('必需参数丢失，修改失败！');window.history.back();</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()	  
if Classid="" then response.Write("<script language=""javascript"">alert('请选择信息分类！');window.history.back();</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
if InfoTitle="" then response.Write("<script language=""javascript"">alert('请填写标题！');window.history.back();</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
if InfoContent="" then response.Write("<script language=""javascript"">alert('请填写内容！');window.history.back();</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
PinLuo.PinLuo_EditInfo Infoid,Classid,InfoTitle,SEO_Title,SEO_Keyword,SEO_Description,InfoContent,InfoImg,Author,Origin,UpdateTime,hits,OrderID,Shenhe,"PinLuo_InfoList"
response.Write("<script language=""javascript"">alert('修改信息成功！');location.href='PinLuo_InfoList.asp?classid="&Trim(Request("BackClassid"))&"&SearchKeyword="&SearchKeyword2&"&SearchSelect="&SearchSelect2&"&page="&page2&"';</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
elseif act="del" then
strManageID = Trim(Request.Form("DelInfoID"))
if trim(replace(strManageID,",",""))="" then response.Write("<script language=""javascript"">alert('请先选择要删除信息！');window.history.back();</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
PinLuo.DelInfoAll strManageID,"PinLuo_InfoList"
response.Write("<script language=""javascript"">alert('删除信息成功！');location.href='PinLuo_InfoList.asp?classid="&Classid&"&SearchKeyword="&SearchKeyword2&"&SearchSelect="&SearchSelect2&"&page="&page2&"';</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
else
InfoItemName="添加信息"
end if
%>
<title>信息管理</title>
<link rel="stylesheet" href="images/style.css" type="text/css">
<script language="javascript" src="js/Pinluo.js" type="text/javascript"></script>
<script src="js/jquery-1.10.2.min.js"></script>
<script type="text/javascript" charset="utf-8" src="../ueditor/ueditor.config.js"></script>
<script type="text/javascript" charset="utf-8" src="../ueditor/ueditor.all.min.js"> </script>
<script type="text/javascript" charset="utf-8" src="../ueditor/lang/zh-cn/zh-cn.js"></script>
</head>

<body class="mainbg">
	<div id="mainhearder"><span>您的位置：企业网站管理系统 >> 信息管理 >> <%=InfoItemName%></span></div>
	<div id="hearder" class="hearder1"><span><%=InfoItemName%></span><a href="PinLuo_InfoList.asp" style="color:#ccc; float:right; margin-right:20px;">信息管理</a></div>
	<div class="main5" id="main5">

<%if act="add" then%>	
	<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
	<tr>
	<td class="tableleft">
    
            <form name="form1" method="post" action="PinLuo_Info.asp">
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="table1">
		<tr>
		<td width="15%"class="tableleft1"><strong>分类：</strong></td>
		<td width="85%" class="tableright1"><label>
		  <select name="Classid" id="Classid">
          <%=PinLuo.PinLuo_GetClass_Option("PinLuo_InfoClass",0,classid,-1)%>
		    </select>
		  </label></td>
		</tr>
		<tr>
		<td class="tableleft1"><strong>标题：</strong></td>
		<td class="tableright1"><input type='text' size='80' maxlength='255' name='InfoTitle' class='input'></td>
		</tr>
        
        <tr>
		<td class="tableleft1"><strong>SEO标题：</strong></td>
		<td class="tableright1"><input type='text' size='50' maxlength='255' name='SEO_Title' class='input'> 
		(为空则默认使用信息标题)</td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>SEO关键词：</strong></td>
		<td class="tableright1"><input type='text' size='50' maxlength='255' name='SEO_Keyword' class='input'>
(多个用,分隔)</td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>SEO描述：</strong></td>
		<td class="tableright1"><input type='text' size='50' maxlength='255' name='SEO_Description' class='input'></td>
		</tr>
        
        <tr>
		<td class="tableleft1"><strong>内容：</strong></td>
		<td class="tableright1">
			<textarea name="InfoContent" id="myEditor" style="width:96%;height:450px;"></textarea>
<script type="text/javascript">
    var editor = new UE.ui.Editor();
    editor.render("myEditor");
    //UE.getEditor('myEditor')
</script>
        </td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>缩略图：</strong></td>
		<td class="tableright1"><input type='text' size='25' maxlength='255' name='InfoImg' id='InfoImg' class='input'><input type=button value="上传图片" onclick="upImage_media_url();">&nbsp;&nbsp;(选填)
			
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
$("#InfoImg").attr("value", arg[0].src);

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
		<td class="tableleft1"><strong>作者：</strong></td>
		<td class="tableright1"><input type='text' size='15' maxlength='255' name='Author' class='input'></td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>出处：</strong></td>
		<td class="tableright1"><input type='text' size='25' maxlength='255' name='Origin' class='input'>&nbsp;&nbsp;(比如品络科技,XX报等)</td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>发布时间：</strong></td>
		<td class="tableright1"><input type='text' size='25' maxlength='255' name='UpdateTime' class='input' value="<%=now%>">&nbsp;&nbsp;(格式如：<%=date%>)</td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>阅读次数：</strong></td>
		<td class="tableright1"><input type='text' size='15' maxlength='255' name='hits' class='input' value="0"></td>
		</tr>
		<tr>
		<td class="tableleft1"><strong>排序：</strong></td>
		<td class="tableright1"><input name='OrderID' type='text' class='input' value="0" size='15' maxlength='10'>&nbsp;&nbsp;(数值越大越靠前)</td>
		</tr>
		<tr>
		<td class="tableleft1"><strong>审核通过：</strong></td>
		<td class="tableright1"><label>
		  <input type="radio" name="Shenhe" id="radio" value="true" checked="checked" />
		  是&nbsp;&nbsp;&nbsp;
		</label>
		  <input name="Shenhe" type="radio" id="radio2" value="false" /> 
		  否
</td>
		</tr>
        
		<tr>
		  <td height="55" align="center" class="tableleft1" style="height:39px;">&nbsp;</td>
		  <td height="55" class="tableright1">
          <input name="act" type="hidden" value="addsave" />
          <input type="submit" name="submit" value="完成添加" class="button">
		    <input type="reset" value="重新填写" class="button" name="reset" tabindex="25">
		    <input type="button" value="刷新本页" class="button" name="button" tabindex="25" onclick="location.reload()"></td>
		  </tr>

		</table>
        </form>
       		
	</td>
	</tr>
	</table>
<%elseif act="edit" then
PinLuo.PinLuo_ViewInfoItem InfoID,"PinLuo_InfoList"
%>
	<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
	<tr>
	<td class="tableleft">
    
            <form name="form1" method="post" action="PinLuo_Info.asp">
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="table1">
		<tr>
		<td width="15%"class="tableleft1"><strong>分类：</strong></td>
		<td width="85%" class="tableright1"><label>
		  <select name="Classid" id="Classid">
          <%=PinLuo.PinLuo_GetClass_Option("PinLuo_InfoClass",0,PinLuo.Classid,-1)%>
		    </select>
		  </label></td>
		</tr>
		<tr>
		<td class="tableleft1"><strong>标题：</strong></td>
		<td class="tableright1"><input type='text' size='80' maxlength='255' name='InfoTitle' class='input' value="<%=PinLuo.InfoTitle%>"></td>
		</tr>
        
        <tr>
		<td class="tableleft1"><strong>SEO标题：</strong></td>
		<td class="tableright1"><input type='text' size='50' maxlength='255' name='SEO_Title' class='input' value="<%=PinLuo.SEO_Title%>">
(为空则默认使用信息标题)</td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>SEO关键词：</strong></td>
		<td class="tableright1"><input type='text' size='50' maxlength='255' name='SEO_Keyword' class='input' value="<%=PinLuo.SEO_Keyword%>">
(多个用,分隔)</td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>SEO描述：</strong></td>
		<td class="tableright1"><input type='text' size='50' maxlength='255' name='SEO_Description' class='input' value="<%=PinLuo.SEO_Description%>"></td>
		</tr>
        
        <tr>
		<td class="tableleft1"><strong>内容：</strong></td>
		<td class="tableright1">
			<textarea name="InfoContent" id="myEditor" style="width:96%;height:450px;"><%if trim(PinLuo.InfoContent)<>"" then response.Write(server.HTMLEncode(PinLuo.InfoContent))%></textarea>
<script type="text/javascript">
    var editor = new UE.ui.Editor();
    editor.render("myEditor");
    //UE.getEditor('myEditor')
</script>        
        </td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>缩略图：</strong></td>
		<td class="tableright1"><input type='text' size='25' maxlength='255' name='InfoImg' id='InfoImg' class='input' value="<%=PinLuo.InfoImg%>"><input type=button value="上传图片" onclick="upImage_media_url();">&nbsp;&nbsp;(选填)
			
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
$("#InfoImg").attr("value", arg[0].src);

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
		<td class="tableleft1"><strong>作者：</strong></td>
		<td class="tableright1"><input type='text' size='15' maxlength='255' name='Author' class='input' value="<%=PinLuo.Author%>"></td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>出处：</strong></td>
		<td class="tableright1"><input type='text' size='25' maxlength='255' name='Origin' class='input' value="<%=PinLuo.Origin%>">&nbsp;&nbsp;(比如品络科技,XX报等)</td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>发布时间：</strong></td>
		<td class="tableright1"><input type='text' size='25' maxlength='255' name='UpdateTime' class='input' value="<%=PinLuo.UpdateTime%>">&nbsp;&nbsp;(格式如：<%=date%>)</td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>阅读次数：</strong></td>
		<td class="tableright1"><input type='text' size='15' maxlength='255' name='hits' class='input' value="<%=PinLuo.hits%>"></td>
		</tr>
		<tr>
		<td class="tableleft1"><strong>排序：</strong></td>
		<td class="tableright1"><input name='OrderID' type='text' class='input' value="<%=PinLuo.OrderID%>" size='15' maxlength='10'>&nbsp;&nbsp;(数值越大越靠前)</td>
		</tr>
		<tr>
		<td class="tableleft1"><strong>审核通过：</strong></td>
		<td class="tableright1"><label>
		  <input type="radio" name="Shenhe" id="radio" value="true"<%if PinLuo.Shenhe then%> checked="checked"<%end if%> />
		  是&nbsp;&nbsp;&nbsp;
		</label>
		  <input name="Shenhe" type="radio" id="radio2" value="false"<%if PinLuo.Shenhe=false then%> checked="checked"<%end if%> /> 
		  否
</td>
		</tr>
        
		<tr>
		  <td height="55" align="center" class="tableleft1" style="height:39px;">&nbsp;</td>
		  <td height="55" class="tableright1">
          <input name="InfoID" type="hidden" value="<%=Infoid%>" />
          <input name="SearchKeyword" type="hidden" value="<%=SearchKeyword2%>" />
          <input name="SearchSelect" type="hidden" value="<%=SearchSelect2%>" />
          <input name="page" type="hidden" value="<%=page2%>" />
          <input name="BackClassid" type="hidden" value="<%=Classid%>" />
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
<%end if%>
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