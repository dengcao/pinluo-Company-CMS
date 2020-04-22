<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!--#include file="../Pinluo_Main/Config.asp"-->
<%
dim SearchKeyword
SearchKeyword=Trim(request("SearchKeyword"))

ClassID=Trim(Request("ClassID"))
if ClassID="" then ClassID=Trim(Request("ID"))
ClassName=""
dim PinLuo
Set PinLuo = New PinLuo_Class
    PinLuo.DBConnBegin
	PinLuo.PinLuo_ViewSiteConfig
	if PinLuo.isnumeric(ClassID) then
	  PinLuo.PinLuo_ViewItemContent ClassID,"PinLuo_ItemClass"
	  ClassName=PinLuo.ClassName
	  ClassName_title=PinLuo.SEO_Title&" - "
	  keywords=PinLuo.SEO_Keyword
	  descriptions=PinLuo.SEO_Description
	  ParentID=PinLuo.ParentID
	else
	  ClassID="" 
	  ClassName_title="网站栏目 - "
	  ClassName="网站栏目"
	  keywords="网站栏目"
	  descriptions="网站栏目"
	  ParentID=0
	end if
%>
<title><%=ClassName_title&Pinluo.Pinluo_SeoTitle%></title>
<meta name="keywords" content="<%=keywords%>">
<META NAME="description" CONTENT="<%=descriptions%>">
<!--#include file="../inc/head.asp"-->
  <div class="xia">
   <div class="zuo">
    <div class="dt">
	  <div class="dt_tou gs1_tou"><span><%=PinLuo.PinLuo_ViewClassName(ParentID,"PinLuo_ItemClass","关于我们")%></span></div>
	  <div class="xian"></div>
	  <ul>
	    <%PinLuo.PinLuo_Classlist_View ParentID,3,"../item/?","PinLuo_ItemClass","",8%>
	  </ul>
	</div>
	<div class="lx1">
	  <div class="lx1_tou"><h3><%=Pinluo.PinLuo_ViewBlockItem(1,1)%></h3></div>
	  <div class="xian"></div>
	  <%=Pinluo.PinLuo_ViewBlockItem(1,0)%>
	</div>
   </div>
   <div class="you">
    <div class="gs1">
	  <div class="gs1_tou"><span><%=ClassName%></span></div>
	  <div class="xian"></div>
	  <div class="lx2_lb ItemContent">

             <%=Pinluo.ClassContents%> 

	  </div>
	</div>
   </div>
  </div>
  <div class="clear"></div>
  <div class="bq">
    <!--#include file="../inc/foot.asp"-->
</div>
</body>
</html>
<%PinLuo.DBConnEnd
Set PinLuo = Nothing%>