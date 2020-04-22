<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!--#include file="../Pinluo_Main/Config.asp"-->
<%
dim PinLuo
Set PinLuo = New PinLuo_Class
    PinLuo.DBConnBegin
	PinLuo.PinLuo_ViewSiteConfig
	Infoid=PinLuo.Pinluo_GetUrlID
	PinLuo.PinLuo_ViewInfoItem infoid,"PinLuo_InfoList","PinLuo_InfoClass"
	classid=Pinluo.classid
	classname=Pinluo.classname
	infotitle=Pinluo.SEO_Title&"_"
	keywords=PinLuo.SEO_Keyword
	descriptions=PinLuo.SEO_Description
%>
<title><%=infotitle&Pinluo.Pinluo_SeoTitle%></title>
<meta name="keywords" content="<%=keywords%>">
<META NAME="description" CONTENT="<%=descriptions%>">

<!--#include file="../inc/head.asp"-->
  <div class="xia">
   <div class="zuo">
    <div class="dt">
	  <div class="dt_tou"><h3>新闻动态</h3></div>
	  <div class="xian"></div>
	  <ul>
      
      <%PinLuo.PinLuo_Classlist_View 0,3,"../info/?","PinLuo_InfoClass","",8%>

	  </ul>
	</div>
    <div class="zx">
	  <div class="zx_tou"><h3>最新资讯</h3></div>
	  <div class="xian"></div>
	  <ul>
      <%HtmlStr="<li><a href=""../info/s.asp?{$infoid}.html"">· {$title}</a></li>"
		Pinluo.PinLuo_GetInfolist HtmlStr,18,5,6,"","PinLuo_InfoList","PinLuo_InfoClass","new"%>
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

	  <div class="gs1_lb">
	    <!--新闻内容-->
        <script type="text/javascript" language="javascript">
<!--
function ContentSize(size)
{
	var obj=document.getElementById("artview_content");
	obj.style.fontSize=size+"px";
}
-->
</script>

<div class="artview">
	<div class="artview_intr">&nbsp;&nbsp;新闻动态 > <%=classname%> > 详细信息</div>
	<div class="artview_title"><%=Pinluo.InfoTitle%></div>
	<div class="artview_info">发布时间: <%=Pinluo.UpdateTime%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;浏览次数：<%=Pinluo.hits%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;大小:&nbsp;&nbsp;<a href="javascript:ContentSize(16)">16px</a>&nbsp;&nbsp;<a href="javascript:ContentSize(14)">14px</a>&nbsp;&nbsp;<a href="javascript:ContentSize(12)">12px</a></div>

	<div id="artview_content"><br><%=Pinluo.InfoContent%><br /></div>
	<div class="blankbar1"></div>
</div>
        
	  </div>
	  <div class="clear"></div>
	  <div class="ji"></div>
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