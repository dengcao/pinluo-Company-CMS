<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>保存留言</title>
<!--#include file="../Pinluo_Main/Config.asp"-->
<%
act=Trim(Request("act"))
dim PinLuo,gotourl
Set PinLuo = New PinLuo_Class
    PinLuo.DBConnBegin
	PinLuo.PinLuo_ViewSiteConfig
	If PinLuo.Pinluo_IsFeedback=false then	response.Write("<script language=""javascript"">alert('抱歉，留言系统已经关闭，不允许提交！请您和管理员联系。');history.back();</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
	if act="addsave" then
	  Classid=Trim(Request("Classid"))
	  FeedbackTitle=Trim(Request("FeedbackTitle"))
	  FeedbackContent=Trim(Request("FeedbackContent"))
	  Author=Trim(Request("Author"))
	  UpdateTime=now
	  Shenhe=false
	  Phone=Trim(Request("Phone"))
	  Email=Trim(Request("Email"))
	  QQ=Trim(Request("QQ"))
	  SEO_Title=""
	  SEO_Keyword=""
	  SEO_Description=""
	  ReplyContent=""
	  ReplyTime=empty	  
	  Hits=0
	  OrderID=0
if FeedbackTitle="" then response.Write("<script language=""javascript"">alert('请填写留言标题！');window.history.back();</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
if FeedbackContent="" then response.Write("<script language=""javascript"">alert('请填写留言内容！');window.history.back();</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
PinLuo.PinLuo_AddFeedback Classid,FeedbackTitle,FeedbackContent,Author,UpdateTime,Shenhe,Phone,Email,QQ,"PinLuo_FeedbackList"
response.Write("<script language=""javascript"">alert('提交留言成功！我们会及时对您的留言审核和处理。谢谢您对我们的支持！');history.back();</script>"):PinLuo.DBConnEnd:Set PinLuo=Nothing:response.End()
    end if
    PinLuo.DBConnEnd
Set PinLuo = Nothing
%>