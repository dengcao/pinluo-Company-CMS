<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Pinluo_Main/Conn.asp"-->
<!--#include file="../Pinluo_Main/Inc/md5.asp"-->
<!--#include file="PinLuo_Class.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<%
act=Trim(Request("act"))
Block_ID=Trim(Request("Block_ID"))
SearchKeyword2 = Trim(Request("SearchKeyword"))
SearchSelect2 = Trim(Request("SearchSelect"))
page2 = Trim(Request("page"))
dim PinLuo
Set PinLuo = New PinLuo_Class
    PinLuo.DBConnBegin
	PinLuo.CheckPurview
	PinLuo.Pinluo_CheckPurviewAdmin(2)
	
if act="add" then
InfoItemName="添加碎片"
elseif act="edit" then
InfoItemName="修改碎片"
elseif act="addsave" then
	Block_Title = Trim(Request.Form("Block_Title"))
	Block_Content = Trim(Request.Form("Block_Content"))
	Block_Time = Request.Form("Block_Time")

if Block_Title="" then response.Write("<script language=""javascript"">alert('请填写碎片标题！');window.history.back();</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()

PinLuo_AddBlock_chk=PinLuo.PinLuo_AddBlock(Block_Title,Block_Content,Block_Time)
if PinLuo_AddBlock_chk=true then
response.Write("<script language=""javascript"">alert('添加碎片成功！');location.href='PinLuo_Block.asp?act=add';</script>")
else
response.Write("<script language=""javascript"">alert('"&PinLuo.ErrMsg&"');window.history.back();</script>")
end if
PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
elseif act="editsave" then
	Block_Title = Trim(Request.Form("Block_Title"))
	Block_Content = Trim(Request.Form("Block_Content"))
	Block_Time = Trim(Request.Form("Block_Time"))

if Block_Title="" then response.Write("<script language=""javascript"">alert('请填写碎片标题！');window.history.back();</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()

PinLuo_EditBlock_chk=PinLuo.PinLuo_EditBlock(Block_ID,Block_Title,Block_Content,Block_Time)
if PinLuo_EditBlock_chk=true then
response.Write("<script language=""javascript"">alert('修改碎片成功！');location.href='PinLuo_BlockList.asp?SearchKeyword="&SearchKeyword2&"&SearchSelect="&SearchSelect2&"&page="&page2&"';</script>")
else
response.Write("<script language=""javascript"">alert('"&PinLuo.ErrMsg&"');window.history.back();</script>")
end if
PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
elseif act="del" then
strManageID = Trim(Request.Form("DelBlockID"))
if trim(replace(strManageID,",",""))="" then response.Write("<script language=""javascript"">alert('请先选择要删除的碎片！');window.history.back();</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
PinLuo.DelBlockAll strManageID
response.Write("<script language=""javascript"">alert('删除碎片成功！');location.href='PinLuo_BlockList.asp?SearchKeyword="&SearchKeyword2&"&SearchSelect="&SearchSelect2&"&page="&page2&"';</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
else
InfoItemName="添加碎片"
end if
%>
<title>碎片管理</title>
<meta content="品络科技,网站管理系统,企业网站管理系统,内容管理系统(CMS),网上商店管理系统,网站建设" name="Keywords" />
<meta content="品络科技成立于2005年6月，是一家集互联网基础服务、互联网应用软件开发、业务解决方案销售及服务于一体的高新技术企业。公司网址：www.5300.cn，品络互联：www.pinluo.com" name="Description"/>
<link rel="stylesheet" href="images/style.css" type="text/css">
<script language="javascript" src="js/Pinluo.js" type="text/javascript"></script>
</head>

<body class="mainbg">
	<div id="mainhearder"><span>您的位置：企业网站管理系统 >> 碎片管理 >> <%=InfoItemName%></span></div>
	<div id="hearder" class="hearder1"><span><%=InfoItemName%></span><a href="PinLuo_BlockList.asp" style="color:#ccc; float:right; margin-right:20px;">碎片列表</a></div>
	<div class="main5" id="main5">

<%if act="add" then%>	
	<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
	<tr>
	<td class="tableleft">
    
            <form name="form1" method="post" action="PinLuo_Block.asp">
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="table1">
		<tr>
		  <td width="15%" class="tableleft1"><strong>碎片标题：</strong></td>
		  <td width="85%" class="tableright1"><input type='text' size='55' maxlength='255' name='Block_Title' class='input'>&nbsp;&nbsp;<font color="red">*</font></td>
		  </tr>
        <tr>
		  <td width="15%" class="tableleft1"><strong>碎片内容：</strong></td>
		  <td width="85%" class="tableright1"><textarea name="Block_Content" style="display:none"></textarea><iframe id="eWebEditor1" src="../editor/ewebeditor.htm?id=Block_Content&style=pinluo" frameBorder="0" width="100%" scrolling="no" height="350"></iframe></td>
		  </tr>  
        <tr>
		<td class="tableleft1"><strong>添加时间：</strong></td>
		<td class="tableright1"><input type='text' size='20' maxlength='255' name='Block_Time' value="<%=now%>" class='input'>&nbsp; 格式如：<%=now%></td>
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
PinLuo.PinLuo_ViewBlockItem Block_ID
%>
	<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
	<tr>
	<td class="tableleft">
    
                    <form name="form1" method="post" action="PinLuo_Block.asp">
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="table1">
		<tr>
		  <td width="15%" class="tableleft1"><strong>碎片标题：</strong></td>
		  <td width="85%" class="tableright1"><input type='text' size='55' maxlength='255' name='Block_Title' class='input' value="<%=PinLuo.Block_Title%>">&nbsp;&nbsp;<font color="red">*</font></td>
		  </tr>
        <tr>
		  <td width="15%" class="tableleft1"><strong>碎片内容：</strong></td>
		  <td width="85%" class="tableright1"><textarea name="Block_Content" style="display:none"><%if trim(PinLuo.Block_Content)<>"" then response.Write(server.HTMLEncode(PinLuo.Block_Content))%></textarea><iframe id="eWebEditor1" src="../editor/ewebeditor.htm?id=Block_Content&style=pinluo" frameBorder="0" width="100%" scrolling="no" height="350"></iframe></td>
		  </tr>  
        <tr>
		<td class="tableleft1"><strong>添加时间：</strong></td>
		<td class="tableright1"><input type='text' size='20' maxlength='255' name='Block_Time' class='input' value="<%=PinLuo.Block_Time%>"> &nbsp;格式如：<%=now%></td>
		</tr>
        
        
		<tr>
		  <td height="55" align="center" class="tableleft1" style="height:39px;">&nbsp;</td>
		  <td height="55" class="tableright1">
          <input name="act" type="hidden" value="editsave" />
          <input name="Block_ID" type="hidden" value="<%=Block_ID%>" />
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