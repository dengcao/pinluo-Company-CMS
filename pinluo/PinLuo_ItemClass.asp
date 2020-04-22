<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Pinluo_Main/Conn.asp"-->
<!--#include file="PinLuo_Class.asp"-->
<%
act=Trim(Request("act"))
classid=Trim(Request("classid"))
dim PinLuo
Set PinLuo = New PinLuo_Class
    PinLuo.DBConnBegin
	PinLuo.CheckPurview
	PinLuo.Pinluo_CheckPurviewAdmin(2)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="images/style.css" type="text/css">
<script type="text/javascript" charset="utf-8" src="../ueditor/ueditor.config.js"></script>
<script type="text/javascript" charset="utf-8" src="../ueditor/ueditor.all.min.js"> </script>
<script type="text/javascript" charset="utf-8" src="../ueditor/lang/zh-cn/zh-cn.js"></script>
<title>栏目管理</title>
<style type="text/css">
.classlist li{
	padding:2px 0px;
	}
*+body .classlist li{padding:2px 18px;}
.classlist li span{
	padding-left:20px;
	}
.classlist #class1{background:url(images/arr2.png) no-repeat;}

</style>
<script language="javascript">
function ShowLink(t){
	if(t==1){
		document.getElementById("Showlink").style.display="";
		}else{
		document.getElementById("Showlink").style.display="none";
		}
	}
function ShowMenu(id){
	document.getElementById("M"+id).style.display=""
	}
function HideMenu(id){
	document.getElementById("M"+id).style.display="none"
	}
function cf(t){
if(confirm("您确定删除【"+t+"】这个栏目吗？\n\n删除此栏目将同时删除其下级栏目。\n\n此操作将无法恢复，请慎重操作。")==true){
	return true;
	}else{
	return false;	
		}
	}
</script>
</head>

<body class="mainbg">
<%
if act="addsave" then
ClassName=Trim(Request("ClassName"))
SEO_Title=Trim(Request("SEO_Title"))
SEO_Keyword=Trim(Request("SEO_Keyword"))
SEO_Description=Trim(Request("SEO_Description"))
ParentID=int(Request("ParentID"))
IsOuter=eval(Request("IsOuter"))
PathUrl=Trim(Request("PathUrl"))
Visible=eval(Request("Visible"))
OrderID=int(Request("OrderID"))
ClassContents=Trim(Request("ClassContents"))
if ClassName="" then response.Write("<script language=""javascript"">alert('栏目名称不能为空！');window.history.back();</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
PinLuo.PinLuo_AddClass ClassName,SEO_Title,SEO_Keyword,SEO_Description,ClassContents,ParentID,IsOuter,PathUrl,Visible,OrderID,"PinLuo_ItemClass"
elseif act="editsave" then
ClassName=Trim(Request("ClassName"))
SEO_Title=Trim(Request("SEO_Title"))
SEO_Keyword=Trim(Request("SEO_Keyword"))
SEO_Description=Trim(Request("SEO_Description"))
IsOuter=eval(Request("IsOuter"))
PathUrl=Trim(Request("PathUrl"))
Visible=eval(Request("Visible"))
OrderID=int(Request("OrderID"))
ClassContents=Trim(Request("ClassContents"))
if Classid="" then response.Write("<script language=""javascript"">alert('请选择更新的栏目！');window.history.back();</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
if ClassName="" then response.Write("<script language=""javascript"">alert('栏目名称不能为空！');window.history.back();</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
PinLuo.PinLuo_EditClass ClassName,SEO_Title,SEO_Keyword,SEO_Description,ClassContents,ClassID,IsOuter,PathUrl,Visible,OrderID,"PinLuo_ItemClass"
elseif act="del" then
if classid="" then response.Write("<script language=""javascript"">alert('请选择删除的栏目！');window.history.back();</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
PinLuo.PinLuo_DeleteItemClass ClassID,"PinLuo_ItemClass"
end if
%>
	<div id="mainhearder"><span>您的位置：企业网站管理系统 >> 栏目管理</span></div>
	<div id="hearder" class="hearder2"><span><%if act="add" then%>添加栏目<%elseif act="edit" then%>修改栏目<%else%>栏目管理<%end if%></span><a href=?act=add&classid=0 style="color:#ccc; float:right; margin-right:20px;">添加栏目</a></div>
	<div class="main5" id="main5">
	
	<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
	<tr>
	<td class="tableleft">

        <%if act="add" then%>
        <form name="form1" method="post" action="PinLuo_ItemClass.asp">
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="table1">
		<tr>
		<td width="16%"class="tableleft1"><strong>上级栏目：</strong></td>
		<td width="84%" class="tableright1"><label>
		  <select name="ParentID" id="ParentID">
          <option value='0'>|--&nbsp;作为顶级栏目 </option>
          <%=PinLuo.PinLuo_GetClass_Option("PinLuo_ItemClass",0,classid,-1)%>
		    </select>
		  </label></td>
		</tr>
		<tr>
		<td class="tableleft1"><strong>栏目名称：</strong></td>
		<td class="tableright1"><input type='text' size='30' maxlength='255' name='ClassName' class='input'></td>
		</tr>
        
        <tr>
		<td class="tableleft1"><strong>SEO标题：</strong></td>
		<td class="tableright1"><input type='text' size='30' maxlength='255' name='SEO_Title' class='input'></td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>SEO关键词：</strong></td>
		<td class="tableright1"><input type='text' size='30' maxlength='255' name='SEO_Keyword' class='input'></td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>SEO描述：</strong></td>
		<td class="tableright1"><input type='text' size='30' maxlength='255' name='SEO_Description' class='input'></td>
		</tr>
        
		<tr>
		<td class="tableleft1"><strong>排序：</strong></td>
		<td class="tableright1"><input name='OrderID' type='text' class='input' value="0" size='30' maxlength='10'>&nbsp;&nbsp;(数值越大越靠前)</td>
		</tr>
		<tr>
		<td class="tableleft1"><strong>隐藏导航：</strong></td>
		<td class="tableright1"><label>
		  <input type="radio" name="Visible" id="radio" value="false" />
		  是&nbsp;&nbsp;&nbsp;
		</label>
		  <input name="Visible" type="radio" id="radio2" value="true" checked="checked" /> 
		  否
</td>
		</tr>
		<tr>
		<td class="tableleft1"><strong>外部链接：</strong></td>
		<td class="tableright1"><label>
		  <input type="radio" name="IsOuter" id="radio" value="true" onclick="ShowLink(1)" />
		  是&nbsp;&nbsp;&nbsp;
		</label>
		  <input name="IsOuter" type="radio" id="radio2" value="false" checked="checked" onclick="ShowLink(0)" /> 
		  否
</td>
		</tr>
        <tr id="Showlink" style="display:none;">
		<td class="tableleft1"><strong>链接地址：</strong></td>
		<td class="tableright1"><input type='text' size='30' maxlength='255' name='PathUrl' class='input'>&nbsp;&nbsp;(为外部链接时填写)</td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>栏目内容：</strong></td>
		<td class="tableright1">
			<textarea name="ClassContents" id="myEditor" style="width:96%;height:450px;"></textarea>
<script type="text/javascript">
    var editor = new UE.ui.Editor();
    editor.render("myEditor");
    //UE.getEditor('myEditor')
</script>
        </td>
		</tr>
		<tr>
		  <td height="55" align="center" class="tableleft1" style="height:39px;">&nbsp;</td>
		  <td height="55" class="tableright1"><input name="act" type="hidden" value="addsave" /><input type="submit" name="submit" value="完成添加" class="button">
		    <input type="reset" value="重新填写" class="button" name="reset" tabindex="25"></td>
		  </tr>

		</table>
        </form>
        <%elseif act="edit" then
		PinLuo.PinLuo_ViewClassItem ClassID,"PinLuo_ItemClass"
		%>
        <form name="form1" method="post" action="PinLuo_ItemClass.asp">
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="table1">
		<tr>
		<td width="16%"class="tableleft1"><strong>上级栏目：</strong></td>
		<td width="84%" class="tableright1">&nbsp;<%=PinLuo.PinLuo_ViewClassName(ClassID,"PinLuo_ItemClass")%>&nbsp;&nbsp;&nbsp;(ID:<%=ClassID%>)</td>
		</tr>
		<tr>
		<td class="tableleft1"><strong>栏目名称：</strong></td>
		<td class="tableright1"><input type='text' size='30' maxlength='255' name='ClassName' class='input' value="<%=PinLuo.ClassName%>"></td>
		</tr>
        
        <tr>
		<td class="tableleft1"><strong>SEO标题：</strong></td>
		<td class="tableright1"><input type='text' size='30' maxlength='255' name='SEO_Title' class='input' value="<%=PinLuo.SEO_Title%>"></td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>SEO关键词：</strong></td>
		<td class="tableright1"><input type='text' size='30' maxlength='255' name='SEO_Keyword' class='input' value="<%=PinLuo.SEO_Keyword%>"></td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>SEO描述：</strong></td>
		<td class="tableright1"><input type='text' size='30' maxlength='255' name='SEO_Description' class='input' value="<%=PinLuo.SEO_Description%>"></td>
		</tr>
        
		<tr>
		<td class="tableleft1"><strong>排序：</strong></td>
		<td class="tableright1"><input type='text' size='30' maxlength='10' name='OrderID' class='input' value="<%=PinLuo.OrderID%>">&nbsp;&nbsp;(数值越大越靠前)</td>
		</tr>
		<tr>
		<td class="tableleft1"><strong>隐藏导航：</strong></td>
		<td class="tableright1"><label>
		  <input type="radio" name="Visible" id="radio" value="false"<%if PinLuo.Visible=false then%> checked="checked"<%end if%> />
		  是&nbsp;&nbsp;&nbsp;
		</label>
		  <input name="Visible" type="radio" id="radio2" value="true"<%if PinLuo.Visible=true then%> checked="checked"<%end if%> /> 
		  否
</td>
		</tr>
		<tr>
		<td class="tableleft1"><strong>外部链接：</strong></td>
		<td class="tableright1"><label>
		  <input type="radio" name="IsOuter" id="radio" value="true"<%if PinLuo.IsOuter=true then%> checked="checked"<%end if%> onclick="ShowLink(1)" />
		  是&nbsp;&nbsp;&nbsp;
		</label>
		  <input name="IsOuter" type="radio" id="radio2" value="false"<%if PinLuo.IsOuter=false then%> checked="checked"<%end if%> onclick="ShowLink(0)" /> 
		  否
</td>
		</tr>
        <tr id="Showlink" style="display:<%if PinLuo.IsOuter=false then%>none<%end if%>;">
		<td class="tableleft1"><strong>链接地址：</strong></td>
		<td class="tableright1"><input type='text' size='30' maxlength='255' name='PathUrl' class='input' value="<%=PinLuo.PathUrl%>">&nbsp;&nbsp;(栏目为外部链接时有效)</td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>栏目内容：</strong></td>
		<td class="tableright1">
			<textarea name="ClassContents" id="myEditor" style="width:96%;height:450px;"><%if trim(PinLuo.ClassContents)<>"" then response.Write(server.HTMLEncode(PinLuo.ClassContents))%></textarea>
<script type="text/javascript">
    var editor = new UE.ui.Editor();
    editor.render("myEditor");
    //UE.getEditor('myEditor')
</script>        
        </td>
		</tr>
		<tr>
		  <td height="55" align="center" class="tableleft1" style="height:39px;">&nbsp;</td>
		  <td height="55" class="tableright1"><input name="classid" type="hidden" value="<%=classid%>" /><input name="act" type="hidden" value="editsave" /><input type="submit" name="submit" value="完成修改" class="button">
		    <input type="reset" value="重新填写" class="button" name="reset" tabindex="25"></td>
		  </tr>

		</table>
        </form>
        <%else%>
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="table1">
		<tr>
		  <td colspan="2" class="tableleft1 classlist" style="text-align:left;padding:5px 10px;">
          
<%
	PinLuo.PinLuo_ViewClass "",10,"PinLuo_ItemClass.asp?act=edit&","PinLuo_ItemClass"
%>
<br><br /><br />
          <font style="color:#aaa; font-weight:normal">说明：( )内数字代表包含下属栏目，[ ]内数字代表排序号，“隐”代表隐藏栏目，“外”代表外部链接。</font></td>
		  </tr>
		</table>
<%end if

PinLuo.DBConnEnd
Set PinLuo = Nothing
%>
	</td>
	</tr>
	</table>
		<br />
<br />	
	</div>
</body>
</html>
