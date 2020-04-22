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
	PinLuo.PinLuo_ViewProductItem infoid,"PinLuo_ProductList","PinLuo_ProductClass"
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
	  <div class="dt_tou"><h3>产品分类</h3></div>
	  <div class="xian"></div>
	  <ul>
      
         <%PinLuo.PinLuo_Classlist_View 0,3,"../product/?","PinLuo_ProductClass","",8%>

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

   </div>
   <div class="you">
    <div class="gs1">

	  <div class="gs1_lb">
	    <!--产品内容-->


<div class="artview">
	
	<div class="artview_intr">&nbsp;&nbsp;产品展示 > <%=classname%> > 详细信息</div>
    <div class="prodview_info" style="text-align:right;color:#999">&nbsp;&nbsp;发布时间: <%=Pinluo.UpdateTime%>&nbsp;&nbsp;&nbsp;&nbsp;查看次数: <%=Pinluo.hits%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</div>
	<div style="font-weight:bold;color:#555">
	<a href="<%=Pinluo.ProImg2%>" target="_blank"><img src="<%=Pinluo.ProImg2%>" width="350" border="0" /></a><br><br>产品名称：<%=Pinluo.ProName%><br />
市场价：<%=Pinluo.ProPrice1%><br />
    优惠价：<%=Pinluo.ProPrice2%>
</div>
	<div class="prodview_prices">    
	</div>
	<div class="prodview_content" style="color:#555"><br /><%=Pinluo.ProContent%><br /></div>
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