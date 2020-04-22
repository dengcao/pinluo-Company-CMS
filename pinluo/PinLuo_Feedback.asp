<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Pinluo_Main/Conn.asp"-->
<!--#include file="PinLuo_Class.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<%
act=Trim(Request("act"))
Feedbackid=Trim(Request("Feedbackid"))
Classid=Trim(Request("Classid"))
SearchKeyword2 = Trim(Request("SearchKeyword"))
SearchSelect2 = Trim(Request("SearchSelect"))
page2 = Trim(Request("page"))
dim PinLuo
Set PinLuo = New PinLuo_Class
    PinLuo.DBConnBegin
	PinLuo.CheckPurview
	PinLuo.Pinluo_CheckPurviewAdmin(6)
	
if act="add" then
FeedbackItemName="添加留言"
elseif act="edit" then
FeedbackItemName="修改留言"
elseif act="addsave" then
	  FeedbackTitle=Trim(Request("FeedbackTitle"))
	  FeedbackContent=Trim(Request("FeedbackContent"))
	  Author=Trim(Request("Author"))
	  UpdateTime=Trim(Request("UpdateTime"))
	  Shenhe=Trim(Request("Shenhe"))
	  Phone=Trim(Request("Phone"))
	  Email=Trim(Request("Email"))
	  QQ=Trim(Request("QQ"))
	  ReplyContent=Trim(Request("ReplyContent"))
	  ReplyTime=Trim(Request("ReplyTime"))
	  Hits=Trim(Request("Hits"))
	  OrderID=Trim(Request("OrderID"))
if Classid="" then response.Write("<script language=""javascript"">alert('请选择留言分类！');window.history.back();</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
if FeedbackTitle="" then response.Write("<script language=""javascript"">alert('请填写留言标题！');window.history.back();</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
if FeedbackContent="" then response.Write("<script language=""javascript"">alert('请填写留言内容！');window.history.back();</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
PinLuo.PinLuo_AddFeedback Classid,FeedbackTitle,FeedbackContent,Author,UpdateTime,Shenhe,Phone,Email,QQ,ReplyContent,ReplyTime,Hits,OrderID,"PinLuo_FeedbackList"
response.Write("<script language=""javascript"">alert('添加留言成功！');location.href='PinLuo_Feedback.asp?act=add&classid="&Classid&"';</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
elseif act="editsave" then
	  FeedbackTitle=Trim(Request("FeedbackTitle"))
	  FeedbackContent=Trim(Request("FeedbackContent"))
	  Author=Trim(Request("Author"))
	  UpdateTime=Trim(Request("UpdateTime"))
	  Shenhe=Trim(Request("Shenhe"))
	  Phone=Trim(Request("Phone"))
	  Email=Trim(Request("Email"))
	  QQ=Trim(Request("QQ"))
	  ReplyContent=Trim(Request("ReplyContent"))
	  ReplyTime=Trim(Request("ReplyTime"))
	  Hits=Trim(Request("Hits"))
	  OrderID=Trim(Request("OrderID"))
	  ReplyContent=Trim(Request("ReplyContent"))
	  ReplyTime=Trim(Request("ReplyTime"))
	  Hits=Trim(Request("Hits"))
	  OrderID=Trim(Request("OrderID"))
if isnumeric(Feedbackid)=false then response.Write("<script language=""javascript"">alert('必需参数丢失，修改失败！');window.history.back();</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()	  
if Classid="" then response.Write("<script language=""javascript"">alert('请选择留言分类！');window.history.back();</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
if FeedbackTitle="" then response.Write("<script language=""javascript"">alert('请填写留言标题！');window.history.back();</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
if FeedbackContent="" then response.Write("<script language=""javascript"">alert('请填写留言内容！');window.history.back();</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
PinLuo.PinLuo_EditFeedback Feedbackid,Classid,FeedbackTitle,FeedbackContent,Author,UpdateTime,Shenhe,Phone,Email,QQ,ReplyContent,ReplyTime,Hits,OrderID,"PinLuo_FeedbackList"
response.Write("<script language=""javascript"">alert('修改留言成功！');location.href='PinLuo_FeedbackList.asp?classid="&Trim(Request("BackClassid"))&"&SearchKeyword="&SearchKeyword2&"&SearchSelect="&SearchSelect2&"&page="&page2&"';</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
elseif act="del" then
strManageID = Trim(Request.Form("DelFeedbackID"))
if trim(replace(strManageID,",",""))="" then response.Write("<script language=""javascript"">alert('请先选择要删除留言！');window.history.back();</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
PinLuo.DelFeedbackAll strManageID,"PinLuo_FeedbackList"
response.Write("<script language=""javascript"">alert('删除留言成功！');location.href='PinLuo_FeedbackList.asp?classid="&Classid&"&SearchKeyword="&SearchKeyword2&"&SearchSelect="&SearchSelect2&"&page="&page2&"';</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
else
InfoItemName="添加留言"
end if
%>
<title>留言管理</title>
<meta content="品络科技,网站管理系统,企业网站管理系统,内容管理系统(CMS),网上商店管理系统,网站建设" name="Keywords" />
<meta content="品络科技成立于2005年6月，是一家集互联网基础服务、互联网应用软件开发、业务解决方案销售及服务于一体的高新技术企业。开发者：www.caozha.com，品络互联：www.pinluo.com" name="Description"/>
<link rel="stylesheet" href="images/style.css" type="text/css">
<script language="javascript" src="js/Pinluo.js" type="text/javascript"></script>
<script src="js/jquery-1.10.2.min.js"></script>
<script type="text/javascript" charset="utf-8" src="../ueditor/ueditor.config.js"></script>
<script type="text/javascript" charset="utf-8" src="../ueditor/ueditor.all.min.js"> </script>
<script type="text/javascript" charset="utf-8" src="../ueditor/lang/zh-cn/zh-cn.js"></script>
</head>

<body class="mainbg">
	<div id="mainhearder"><span>您的位置：企业网站管理系统 >> 留言反馈管理 >> <%=FeedbackItemName%></span></div>
	<div id="hearder" class="hearder1"><span><%=FeedbackItemName%></span><a href="PinLuo_FeedbackList.asp" style="color:#ccc; float:right; margin-right:20px;">留言管理</a></div>
	<div class="main5" id="main5">

<%if act="add" then%>	
	<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
	<tr>
	<td class="tableleft">
    
            <form name="form1" method="post" action="PinLuo_Feedback.asp">
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="table1">
		<tr>
		<td width="15%"class="tableleft1"><strong>分类：</strong></td>
		<td width="85%" class="tableright1"><label>
		  <select name="Classid" id="Classid">
          <%=PinLuo.PinLuo_GetClass_Option("PinLuo_FeedbackClass",0,classid,-1)%>
		    </select>
		  </label></td>
		</tr>
		<tr>
		<td class="tableleft1"><strong>标题：</strong></td>
		<td class="tableright1"><input type='text' size='80' maxlength='255' name='FeedbackTitle' class='input'></td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>内容：</strong></td>
		<td class="tableright1">
			<textarea name="FeedbackContent" id="myEditor" style="width:96%;height:450px;"></textarea>
<script type="text/javascript">
    var editor = new UE.ui.Editor();
    editor.render("myEditor");
    //UE.getEditor('myEditor')
</script>
        </td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>留言者：</strong></td>
		<td class="tableright1"><input type='text' size='15' maxlength='255' name='Author' value="匿名" class='input'></td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>电话：</strong></td>
		<td class="tableright1"><input type='text' size='15' maxlength='255' name='Phone' class='input'></td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>邮箱：</strong></td>
		<td class="tableright1"><input type='text' size='15' maxlength='255' name='Email' class='input'></td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>联系QQ：</strong></td>
		<td class="tableright1"><input type='text' size='15' maxlength='255' name='QQ' class='input'></td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>发布时间：</strong></td>
		<td class="tableright1"><input type='text' size='25' maxlength='255' name='UpdateTime' class='input' value="<%=now%>">&nbsp;&nbsp;(格式如：<%=date%>)</td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>点击：</strong></td>
		<td class="tableright1"><input type='text' size='15' maxlength='255' name='Hits' value="0" class='input'></td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>排序：</strong></td>
		<td class="tableright1"><input type='text' size='15' maxlength='255' name='OrderID' value="0" class='input'></td>
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
		<td class="tableleft1"><strong>回复内容：</strong></td>
		<td class="tableright1">
			<textarea name="ReplyContent" id="myEditor2" style="width:96%;height:450px;"></textarea>
<script type="text/javascript">
    var editor2 = new UE.ui.Editor();
    editor2.render("myEditor2");
    //UE.getEditor('myEditor2')
</script>
        </td>
		</tr>
        
        <tr>
		<td class="tableleft1"><strong>回复时间：</strong></td>
		<td class="tableright1"><input type='text' size='25' maxlength='255' name='ReplyTime' class='input' value="<%=now%>">&nbsp;&nbsp;(格式如：<%=date%>)</td>
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
PinLuo.PinLuo_ViewFeedbackItem FeedbackID,"PinLuo_FeedbackList"
%>
	<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
	<tr>
	<td class="tableleft">
    
          <form name="form1" method="post" action="PinLuo_Feedback.asp">
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="table1">
		<tr>
		<td width="15%"class="tableleft1"><strong>分类：</strong></td>
		<td width="85%" class="tableright1"><label>
		  <select name="Classid" id="Classid">
          <%=PinLuo.PinLuo_GetClass_Option("PinLuo_FeedbackClass",0,classid,-1)%>
		    </select>
		  </label></td>
		</tr>
		<tr>
		<td class="tableleft1"><strong>标题：</strong></td>
		<td class="tableright1"><input type='text' size='80' maxlength='255' name='FeedbackTitle' class='input' value="<%=Pinluo.FeedbackTitle%>"></td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>内容：</strong></td>
		<td class="tableright1">
			<textarea name="FeedbackContent" id="myEditor" style="width:96%;height:450px;"><%if trim(PinLuo.FeedbackContent)<>"" then response.Write(server.HTMLEncode(PinLuo.FeedbackContent))%></textarea>
<script type="text/javascript">
    var editor = new UE.ui.Editor();
    editor.render("myEditor");
    //UE.getEditor('myEditor')
</script>
        </td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>留言者：</strong></td>
		<td class="tableright1"><input type='text' size='15' maxlength='255' name='Author' class='input' value="<%=Pinluo.Author%>"></td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>电话：</strong></td>
		<td class="tableright1"><input type='text' size='15' maxlength='255' name='Phone' class='input' value="<%=Pinluo.Phone%>"></td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>邮箱：</strong></td>
		<td class="tableright1"><input type='text' size='15' maxlength='255' name='Email' class='input' value="<%=Pinluo.Email%>"></td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>联系QQ：</strong></td>
		<td class="tableright1"><input type='text' size='15' maxlength='255' name='QQ' class='input' value="<%=Pinluo.QQ%>"></td>
		</tr>
        
        <tr>
		<td class="tableleft1"><strong>点击：</strong></td>
		<td class="tableright1"><input type='text' size='15' maxlength='255' name='Hits' value="<%=Pinluo.Hits%>" class='input'></td>
		</tr>
        <tr>
		<td class="tableleft1"><strong>排序：</strong></td>
		<td class="tableright1"><input type='text' size='15' maxlength='255' name='OrderID' value="<%=Pinluo.OrderID%>" class='input'></td>
		</tr>
        
        <tr>
		<td class="tableleft1"><strong>发布时间：</strong></td>
		<td class="tableright1"><input type='text' size='25' maxlength='255' name='UpdateTime' class='input' value="<%=Pinluo.UpdateTime%>">&nbsp;&nbsp;(格式如：<%=date%>)</td>
		</tr>
		<tr>
		<td class="tableleft1"><strong>审核通过：</strong></td>
		<td class="tableright1"><label>
		  <input type="radio" name="Shenhe" id="radio" value="true"<%if PinLuo.Shenhe then%> checked="checked"<%end if%>/>
		  是&nbsp;&nbsp;&nbsp;
		</label>
		  <input name="Shenhe" type="radio" id="radio2" value="false"<%if PinLuo.Shenhe=false then%> checked="checked"<%end if%> /> 
		  否
</td>
		</tr>
        
        <tr>
		<td class="tableleft1"><strong>回复内容：</strong></td>
		<td class="tableright1">
			<textarea name="ReplyContent" id="myEditor2" style="width:96%;height:450px;"><%if trim(PinLuo.ReplyContent)<>"" then response.Write(server.HTMLEncode(PinLuo.ReplyContent))%></textarea>
<script type="text/javascript">
    var editor2 = new UE.ui.Editor();
    editor2.render("myEditor2");
    //UE.getEditor('myEditor2')
</script>
        </td>
		</tr>
        
        <tr>
		<td class="tableleft1"><strong>回复时间：</strong></td>
		<td class="tableright1"><input type='text' size='25' maxlength='255' name='ReplyTime' class='input' value="<%if isdate(Pinluo.ReplyTime)=false then response.write(now) else response.write(Pinluo.ReplyTime)%>">&nbsp;&nbsp;(格式如：<%=now%>)</td>
		</tr>
        
		<tr>
		  <td height="55" align="center" class="tableleft1" style="height:39px;">&nbsp;</td>
		  <td height="55" class="tableright1">
          <input name="FeedbackID" type="hidden" value="<%=Feedbackid%>" />
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