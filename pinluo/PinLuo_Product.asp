<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Pinluo_Main/Conn.asp"-->
<!--#include file="PinLuo_Class.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<%
act=Trim(Request("act"))
Proid=Trim(Request("Proid"))
Classid=Trim(Request("Classid"))
SearchKeyword2 = Trim(Request("SearchKeyword"))
SearchSelect2 = Trim(Request("SearchSelect"))
page2 = Trim(Request("page"))
dim PinLuo
Set PinLuo = New PinLuo_Class
    PinLuo.DBConnBegin
	PinLuo.CheckPurview
	PinLuo.Pinluo_CheckPurviewAdmin(4)
	
if act="add" then
ProductItemName="添加产品"
elseif act="edit" then
ProductItemName="修改产品"
elseif act="addsave" then
	  ProName=Trim(Request("ProName"))
	  SEO_Title=Trim(Request("SEO_Title"))
	  SEO_Keyword=Trim(Request("SEO_Keyword"))
	  SEO_Description=Trim(Request("SEO_Description"))
	  ProContent=Trim(Request("ProContent"))
	  ProImg1=Trim(Request("ProImg1"))
	  ProImg2=Trim(Request("ProImg2"))
	  ProPrice1=Trim(Request("ProPrice1"))
	  ProPrice2=Trim(Request("ProPrice2"))
	  Saled=Trim(Request("Saled"))
	  Jian=Trim(Request("Jian"))
	  Hot=Trim(Request("Hot"))
	  Cheap=Trim(Request("Cheap"))
	  UpdateTime=Trim(Request("UpdateTime"))
	  hits=Trim(Request("hits"))
	  OrderID=Trim(Request("OrderID"))
	  Shenhe=eval(Request("Shenhe"))
if Classid="" then response.Write("<script language=""javascript"">alert('请选择产品分类！');window.history.back();</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
if ProName="" then response.Write("<script language=""javascript"">alert('请填写名称！');window.history.back();</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
if ProContent="" then response.Write("<script language=""javascript"">alert('请填写产品介绍内容！');window.history.back();</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
PinLuo.PinLuo_AddProduct Classid,ProName,SEO_Title,SEO_Keyword,SEO_Description,ProContent,ProImg1,ProImg2,ProPrice1,ProPrice2,Saled,Jian,Hot,Cheap,UpdateTime,hits,OrderID,Shenhe,"PinLuo_ProductList"
response.Write("<script language=""javascript"">alert('添加产品成功！');location.href='PinLuo_Product.asp?act=add&classid="&Classid&"';</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
elseif act="editsave" then
	  ProName=Trim(Request("ProName"))
	  SEO_Title=Trim(Request("SEO_Title"))
	  SEO_Keyword=Trim(Request("SEO_Keyword"))
	  SEO_Description=Trim(Request("SEO_Description"))
	  ProContent=Trim(Request("ProContent"))
	  ProImg1=Trim(Request("ProImg1"))
	  ProImg2=Trim(Request("ProImg2"))
	  ProPrice1=Trim(Request("ProPrice1"))
	  ProPrice2=Trim(Request("ProPrice2"))
	  Saled=Trim(Request("Saled"))
	  Jian=Trim(Request("Jian"))
	  Hot=Trim(Request("Hot"))
	  Cheap=Trim(Request("Cheap"))
	  UpdateTime=Trim(Request("UpdateTime"))
	  hits=Trim(Request("hits"))
	  OrderID=Trim(Request("OrderID"))
	  Shenhe=eval(Request("Shenhe"))
if isnumeric(Proid)=false then response.Write("<script language=""javascript"">alert('必需参数丢失，修改失败！');window.history.back();</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()	  
if Classid="" then response.Write("<script language=""javascript"">alert('请选择产品分类！');window.history.back();</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
if ProName="" then response.Write("<script language=""javascript"">alert('请填写名称！');window.history.back();</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
if ProContent="" then response.Write("<script language=""javascript"">alert('请填写产品介绍内容！');window.history.back();</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
PinLuo.PinLuo_EditProduct ProID,Classid,ProName,SEO_Title,SEO_Keyword,SEO_Description,ProContent,ProImg1,ProImg2,ProPrice1,ProPrice2,Saled,Jian,Hot,Cheap,UpdateTime,hits,OrderID,Shenhe,"PinLuo_ProductList"
response.Write("<script language=""javascript"">alert('修改产品成功！');location.href='PinLuo_ProductList.asp?classid="&Trim(Request("BackClassid"))&"&SearchKeyword="&SearchKeyword2&"&SearchSelect="&SearchSelect2&"&page="&page2&"';</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
elseif act="del" then
strManageID = Trim(Request.Form("DelProID"))
if trim(replace(strManageID,",",""))="" then response.Write("<script language=""javascript"">alert('请先选择要删除产品！');window.history.back();</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
PinLuo.DelProductAll strManageID,"PinLuo_ProductList"
response.Write("<script language=""javascript"">alert('删除产品成功！');location.href='PinLuo_ProductList.asp?classid="&Classid&"&SearchKeyword="&SearchKeyword2&"&SearchSelect="&SearchSelect2&"&page="&page2&"';</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
else
ProductItemName="添加产品"
end if
%>
<title>产品管理</title>
<link rel="stylesheet" href="images/style.css" type="text/css">
<script language="javascript" src="js/Pinluo.js" type="text/javascript"></script>
<script src="js/jquery-1.10.2.min.js"></script>
<script type="text/javascript" charset="utf-8" src="../ueditor/ueditor.config.js"></script>
<script type="text/javascript" charset="utf-8" src="../ueditor/ueditor.all.min.js"> </script>
<script type="text/javascript" charset="utf-8" src="../ueditor/lang/zh-cn/zh-cn.js"></script>
</head>

<body class="mainbg">
	<div id="mainhearder"><span>您的位置：企业网站管理系统 >> 产品管理 >> <%=ProductItemName%></span></div>
	<div id="hearder" class="hearder2"><span><%=ProductItemName%></span><a href="PinLuo_ProductList.asp" style="color:#ccc; float:right; margin-right:20px;">产品管理</a></div>
	<div class="main5" id="main5">

<%if act="add" then%>	
	<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
	<tr>
	<td class="tableleft">
    
            <form name="form1" method="post" action="PinLuo_Product.asp">
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="table1">
		<tr>
		<td width="15%"class="tableleft1"><strong>分类：</strong></td>
		<td width="85%" class="tableright1"><label>
		  <select name="Classid" id="Classid">
          <%=PinLuo.PinLuo_GetClass_Option("PinLuo_ProductClass",0,classid,-1)%>
		    </select>
		  </label></td>
		</tr>
		<tr>
		<td class="tableleft1"><strong>产品名称：</strong></td>
		<td class="tableright1"><input type='text' size='80' maxlength='255' name='ProName' class='input'></td>
		</tr>
        
        <tr>
		<td class="tableleft1"><strong>SEO标题：</strong></td>
		<td class="tableright1"><input type='text' size='50' maxlength='255' name='SEO_Title' class='input'> (为空则默认使用产品名称)</td>
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
		<td class="tableleft1"><strong>小图：</strong></td>
		<td class="tableright1"><input type='text' size='25' maxlength='255' name='ProImg1' id='ProImg1' class='input'> <input type=button value="上传图片" onclick="upImage_media_url();">
			
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
$("#ProImg1").attr("value", arg[0].src);

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
		<td class="tableleft1"><strong>大图：</strong></td>
		<td class="tableright1"><input type='text' size='25' maxlength='255' name='ProImg2' id='ProImg2' class='input'> <input type=button value="上传图片" onclick="upImage_media_url2()">
			
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
$("#ProImg2").attr("value", arg[0].src);

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
		<td class="tableleft1"><strong>市场价：</strong></td>
		<td class="tableright1"><input name='ProPrice2' type='text' class='input' value="-1" size='15' maxlength='255'> (填写数值,-1为面议)</td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>现价：</strong></td>
		<td class="tableright1"><input name='ProPrice1' type='text' class='input' value="-1" size='15' maxlength='255'> (填写数值,-1为面议)</td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>产品介绍：</strong></td>
		<td class="tableright1">
			<textarea name="ProContent" id="myEditor" style="width:96%;height:450px;"></textarea>
<script type="text/javascript">
    var editor = new UE.ui.Editor();
    editor.render("myEditor");
    //UE.getEditor('myEditor')
</script>        
        </td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>已售出：</strong></td>
		<td class="tableright1"><input name='Saled' type='text' class='input' value="0" size='15' maxlength='255'></td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>发布时间：</strong></td>
		<td class="tableright1"><input type='text' size='25' maxlength='255' name='UpdateTime' class='input' value="<%=now%>">&nbsp;&nbsp;(格式如：<%=date%>)</td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>查看次数：</strong></td>
		<td class="tableright1"><input type='text' size='15' maxlength='255' name='hits' class='input' value="0"></td>
		</tr>
		<tr>
		<td class="tableleft1"><strong>排序：</strong></td>
		<td class="tableright1"><input name='OrderID' type='text' class='input' value="0" size='15' maxlength='10'>&nbsp;&nbsp;(数值越大越靠前)</td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>推荐：</strong></td>
		<td class="tableright1"><label>
		  <input type="radio" name="Jian" id="radio" value="true" />
		  是&nbsp;&nbsp;&nbsp;
		</label>
		  <input name="Jian" type="radio" id="radio2" value="false" checked="checked" /> 
		  否
</td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>热卖：</strong></td>
		<td class="tableright1"><label>
		  <input type="radio" name="Hot" id="radio" value="true" />
		  是&nbsp;&nbsp;&nbsp;
		</label>
		  <input name="Hot" type="radio" id="radio2" value="false" checked="checked" /> 
		  否
</td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>打折：</strong></td>
		<td class="tableright1"><label>
		  <input type="radio" name="Cheap" id="radio" value="true" />
		  是&nbsp;&nbsp;&nbsp;
		</label>
		  <input name="Cheap" type="radio" id="radio2" value="false" checked="checked" /> 
		  否
</td>
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
PinLuo.PinLuo_ViewProductItem ProID,"PinLuo_ProductList"
%>
	<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
	<tr>
	<td class="tableleft">
        
        <form name="form1" method="post" action="PinLuo_Product.asp">
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="table1">
		<tr>
		<td width="15%"class="tableleft1"><strong>分类：</strong></td>
		<td width="85%" class="tableright1"><label>
		  <select name="Classid" id="Classid">
          <%=PinLuo.PinLuo_GetClass_Option("PinLuo_ProductClass",0,Classid,-1)%>
		    </select>
		  </label></td>
		</tr>
		<tr>
		<td class="tableleft1"><strong>产品名称：</strong></td>
		<td class="tableright1"><input type='text' size='80' maxlength='255' name='ProName' class='input' value="<%=PinLuo.ProName%>"></td>
		</tr>
        
        <tr>
		<td class="tableleft1"><strong>SEO标题：</strong></td>
		<td class="tableright1"><input type='text' size='50' maxlength='255' name='SEO_Title' class='input' value="<%=PinLuo.SEO_Title%>">
(为空则默认使用产品名称)</td>
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
		<td class="tableleft1"><strong>小图：</strong></td>
		<td class="tableright1"><input type='text' size='25' maxlength='255' name='ProImg1' id='ProImg1' class='input' value="<%=PinLuo.ProImg1%>"> <input type=button value="上传图片" onclick="upImage_media_url();">
			
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
$("#ProImg1").attr("value", arg[0].src);

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
		<td class="tableleft1"><strong>大图：</strong></td>
		<td class="tableright1"><input type='text' size='25' maxlength='255' name='ProImg2' id='ProImg2' class='input' value="<%=PinLuo.ProImg2%>"> <input type=button value="上传图片" onclick="upImage_media_url2();">
			
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
$("#ProImg2").attr("value", arg[0].src);

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
		<td class="tableleft1"><strong>市场价：</strong></td>
		<td class="tableright1"><input name='ProPrice2' type='text' class='input' size='15' maxlength='255' value="<%=PinLuo.ProPrice2%>"> (填写数值,-1为面议)</td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>现价：</strong></td>
		<td class="tableright1"><input name='ProPrice1' type='text' class='input' size='15' maxlength='255' value="<%=PinLuo.ProPrice1%>"> (填写数值,-1为面议)</td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>产品介绍：</strong></td>
		<td class="tableright1">
			<textarea name="ProContent" id="myEditor" style="width:96%;height:450px;"><%if trim(PinLuo.ProContent)<>"" then response.Write(server.HTMLEncode(PinLuo.ProContent))%></textarea>
<script type="text/javascript">
    var editor = new UE.ui.Editor();
    editor.render("myEditor");
    //UE.getEditor('myEditor')
</script>
        </td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>已售出：</strong></td>
		<td class="tableright1"><input name='Saled' type='text' class='input' value="<%=PinLuo.Saled%>" size='15' maxlength='255'></td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>发布时间：</strong></td>
		<td class="tableright1"><input type='text' size='25' maxlength='255' name='UpdateTime' class='input' value="<%=PinLuo.UpdateTime%>">&nbsp;&nbsp;(格式如：<%=date%>)</td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>查看次数：</strong></td>
		<td class="tableright1"><input type='text' size='15' maxlength='255' name='hits' class='input' value="<%=PinLuo.hits%>"></td>
		</tr>
		<tr>
		<td class="tableleft1"><strong>排序：</strong></td>
		<td class="tableright1"><input name='OrderID' type='text' class='input' value="<%=PinLuo.OrderID%>" size='15' maxlength='10'>&nbsp;&nbsp;(数值越大越靠前)</td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>推荐：</strong></td>
		<td class="tableright1"><label>
		  <input type="radio" name="Jian" id="radio" value="true"<%if PinLuo.Jian then%> checked="checked"<%end if%> />
		  是&nbsp;&nbsp;&nbsp;
		</label>
		  <input name="Jian" type="radio" id="radio2" value="false"<%if PinLuo.Jian=false then%> checked="checked"<%end if%> /> 
		  否
</td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>热卖：</strong></td>
		<td class="tableright1"><label>
		  <input type="radio" name="Hot" id="radio" value="true"<%if PinLuo.Hot then%> checked="checked"<%end if%> />
		  是&nbsp;&nbsp;&nbsp;
		</label>
		  <input name="Hot" type="radio" id="radio2" value="false"<%if PinLuo.Hot=false then%> checked="checked"<%end if%> /> 
		  否
</td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>打折：</strong></td>
		<td class="tableright1"><label>
		  <input type="radio" name="Cheap" id="radio" value="true"<%if PinLuo.Cheap then%> checked="checked"<%end if%> />
		  是&nbsp;&nbsp;&nbsp;
		</label>
		  <input name="Cheap" type="radio" id="radio2" value="false"<%if PinLuo.Cheap=false then%> checked="checked"<%end if%> /> 
		  否
</td>
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
          <input name="ProID" type="hidden" value="<%=Proid%>" />
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