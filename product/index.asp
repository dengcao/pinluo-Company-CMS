<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!--#include file="../Pinluo_Main/Config.asp"-->
<%
dim SearchKeyword
SearchKeyword=Trim(request("SearchKeyword"))

ClassID=Trim(Request("ClassID"))
ClassName=""
dim PinLuo
Set PinLuo = New PinLuo_Class
    PinLuo.DBConnBegin
	PinLuo.PinLuo_ViewSiteConfig
	if PinLuo.isnumeric(ClassID) then
	  PinLuo.PinLuo_ViewClassItem ClassID,"PinLuo_ProductClass"
	  ClassName=PinLuo.ClassName
	  ClassName_title=PinLuo.SEO_Title&" - "
	  keywords=PinLuo.SEO_Keyword
	  descriptions=PinLuo.SEO_Description
	else
	  ClassID="" 
	  ClassName_title="产品展示 - "
	  ClassName="产品展示"
	  keywords="产品展示"
	  descriptions="产品展示"
	end if
%>
<title><%=ClassName_title&Pinluo.Pinluo_SeoTitle%></title>
<meta name="keywords" content="<%=keywords%>">
<META NAME="description" CONTENT="<%=descriptions%>">
<!--#include file="../inc/head.asp"-->
  <div class="xia">
   <div class="zuo">
    <div class="dt">
	  <div class="zs_tou">
	    <h3>产品分类</h3>
	  </div>
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
    <div class="zs1">
	  <div class="gs1_tou">
	    <span><%=ClassName%></span>
	  </div>
	  <div class="xian"></div>
	  <ul>
             <%
PagelistHtmlSt="<li><a href=""s.asp?{$proid}.html""><img src=""{$proimg1}"" border=0 width=124 height=93 /></a><p><a href=""s.asp?{$proid}.html"">{$proname}</a></p></li>"&vbcrlf
Pinluo.PinLuo_productList PagelistHtmlSt,8,5,12,"productlist","index.asp",SearchKeyword,"Shenhe2",ClassID,"PinLuo_ProductList","PinLuo_ProductClass",""%>

	  </ul>
      
           <div class="clear"></div>
          <div class="ji"><form action="index.asp" method="post" name="productlist" id="productlist">
<%=pinluo.Pinluo_showpage_temp%>
</form></div>
      
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