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
	  PinLuo.PinLuo_ViewClassItem ClassID,"PinLuo_FeedbackClass"
	  ClassName=PinLuo.ClassName
	  ClassName_title=PinLuo.SEO_Title&" - "
	  keywords=PinLuo.SEO_Keyword
	  descriptions=PinLuo.SEO_Description
	  ParentID=PinLuo.ParentID
	else
	  ClassID="" 
	  ClassName_title="留言反馈 - "
	  ClassName="留言反馈"
	  keywords="留言反馈"
	  descriptions="留言反馈"
	  ParentID=0
	end if
%>
<title><%=ClassName_title&Pinluo.Pinluo_SeoTitle%></title>
<meta name="keywords" content="<%=keywords%>">
<META NAME="description" CONTENT="<%=descriptions%>">

<!--#include file="../inc/head.asp"-->
  <div class="xia">
   <div class="zuo">
    <div class="zx">
	  <div class="zx_tou"><h3><%=PinLuo.PinLuo_ViewClassName(ParentID,"PinLuo_FeedbackClass","留言反馈")%></h3></div>
	  <div class="xian"></div>
	  <ul>
	    <%PinLuo.PinLuo_Classlist_View ParentID,3,"../feedback/?","PinLuo_FeedbackClass","",8%>
	  </ul>
	</div>
	<div class="lx1">
	  <div class="lx1_tou"><h3><%=Pinluo.PinLuo_ViewBlockItem(1,1)%></h3></div>
	  <div class="xian"></div>
	  <%=Pinluo.PinLuo_ViewBlockItem(1,0)%>
	</div>
   </div>
   <div class="you">
    <div class="ly">
	  <div class="ly_tou">
	    <h3>给我留言</h3>
	  </div>
	  <div class="xian"></div>
	  <div class="ly1 ly100"><TABLE width="100%" 
            border=0 cellPadding=0 cellSpacing=0 style="Z-INDEX: 100">
              <TBODY>
              <TR>
                <TD 
                style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px">
                  <FORM id=List name=List action=index.asp method=post>
                  <SCRIPT language=JavaScript>
<!--
///////////////////////////////////////////////////////////////////////////////////////////
//	函数名：WinOpenSmall
//	作  用：打开绝对居中窗口
//	作  者：Builder
//	参  数：url		目标文件
//	参  数：width    打开窗口宽度
//	参  数：height   打开窗口高度
///////////////////////////////////////////////////////////////////////////////////////////
	function WinOpenSmall(url,width,height)
	{
		var left = (screen.width/2) - width/2;
		var top = (screen.height/2) - height/2;
		var styleStr = 'toolbar=no,location=no,directories=auto,status=no,menubar=no,scrollbars=yes resizable=yes,z-lock=yes,width='+width+',height='+height+',left='+left+',top='+top+',screenX='+left+',screenY='+top;
		window.open(url,"", styleStr);
	}
	
//-->
</SCRIPT>

                  <TABLE style="TABLE-LAYOUT: fixed" height=28 cellSpacing=0 
                  cellPadding=0 width="100%" border=0>
                    <TBODY>
                    <TR height=3 width="100%">
                      <TD>
                        <TABLE style="TABLE-LAYOUT: fixed" height=3 
                        cellSpacing=0 cellPadding=0 width="100%" border=0>
                          <TBODY>
                          <TR height=1>
                            <TD width=1></TD>
                            <TD width=1></TD>
                            <TD width=1></TD>
                            <TD bgColor=#cccccc></TD>
                            <TD width=1></TD>
                            <TD width=1></TD>
                            <TD width=1></TD></TR>
                          <TR height=1>
                            <TD></TD>
                            <TD bgColor=#cccccc colSpan=2></TD>
                            <TD bgColor=#f7f7f7></TD>
                            <TD bgColor=#cccccc colSpan=2></TD>
                            <TD></TD></TR>
                          <TR height=1>
                            <TD></TD>
                            <TD bgColor=#cccccc></TD>
                            <TD bgColor=#f7f7f7 colSpan=3></TD>
                            <TD bgColor=#cccccc></TD>
                            <TD></TD></TR></TBODY></TABLE></TD></TR>
                    <TR>
                      <TD>
                        <TABLE style="TABLE-LAYOUT: fixed" height="100%" 
                        cellSpacing=0 cellPadding=0 border=0>
                          <TBODY>
                          <TR>
                            <TD width=1 bgColor=#cccccc></TD>
                            <TD id=oINNER 
                            style="PADDING-RIGHT: 1px; PADDING-LEFT: 3px; PADDING-TOP: 1px" 
                            align=middle bgColor=#f7f7f7>
                              <TABLE width="100%" 
                                border=1 align="left" 
                              cellPadding=1 cellSpacing=1 borderColorLight=#bbbbbb borderColorDark=#ffffff 
                              bgColor=#ffffff>
                                <TBODY>
                                <TR style="FONT-WEIGHT: bold" align=middle 
                                height=30>
                                <TD width="8%" bgColor=#f1f1f1><div align="left">编号</div></TD>
                                <TD width="39%" bgColor=#f1f1f1><div align="left">留言主题</div></TD>
                                <TD width="13%" bgColor=#f1f1f1><div align="left">留言时间</div></TD>
                                <TD width="8%" bgColor=#f1f1f1><div align="left">联系人</div></TD></TR>
                                
                                <%
PagelistHtmlSt="<TR bgColor=#ffffff height=27><TD align=middle height=27><div align=""left"">{$feedbackid}</div></TD><TD style=""PADDING-LEFT: 4px""><div align=""left""><A style=""COLOR: #333"" href=""s.asp?{$feedbackid}.html"" target=_blank>{$feedbacktitle}</A></div></TD><TD style=""PADDING-LEFT: 4px"" align=middle><div align=""left"">{$time}</div></TD><TD align=middle><div align=""left"">{$author}</div></TD></TR>"&vbcrlf
Pinluo.PinLuo_FeedbackList PagelistHtmlSt,43,6,10,"Feedbacklist","index.asp",SearchKeyword,"Shenhe2",ClassID,"PinLuo_FeedbackList","PinLuo_FeedbackClass",""%>
                                
                                <TR>
                                <TD align=middle bgColor=#f1f1f1 colSpan=4 
                                height=36>
                                <form action="" method="post" name="Feedbacklist" id="Feedbacklist">
<%=pinluo.Pinluo_showpage_temp%>
</form>
                                </TD></TR></TBODY>
                              </TABLE></TD>
                            <TD width=1 
                    bgColor=#cccccc></TD></TR></TBODY></TABLE></TD></TR>
                    <TR height=3 width="100%">
                      <TD>
                        <TABLE style="TABLE-LAYOUT: fixed" height=3 
                        cellSpacing=0 cellPadding=0 width="100%" border=0>
                          <TBODY>
                          <TR height=1>
                            <TD width=1></TD>
                            <TD width=1 bgColor=#cccccc></TD>
                            <TD width=1 bgColor=#f7f7f7></TD>
                            <TD bgColor=#f7f7f7></TD>
                            <TD width=1 bgColor=#f7f7f7></TD>
                            <TD width=1 bgColor=#cccccc></TD>
                            <TD width=1></TD></TR>
                          <TR height=1>
                            <TD></TD>
                            <TD bgColor=#cccccc colSpan=2></TD>
                            <TD bgColor=#f7f7f7></TD>
                            <TD bgColor=#cccccc colSpan=2></TD>
                            <TD></TD></TR>
                          <TR height=1>
                            <TD colSpan=3></TD>
                            <TD bgColor=#cccccc></TD>
                    <TD 
                  colSpan=3></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></FORM>
                  <FORM id=Feedback name=Feedback action=post.asp method=post>
                  <SCRIPT language=JavaScript 
                   src="ValiDate.js"></SCRIPT>

                  <TABLE style="TABLE-LAYOUT: fixed" height=28 cellSpacing=0 
                  cellPadding=0 width="100%" border=0>
                    <TBODY>
                    <TR height=3 width="100%">
                      <TD>
                        <TABLE style="TABLE-LAYOUT: fixed" height=3 
                        cellSpacing=0 cellPadding=0 width="100%" border=0>
                          <TBODY>
                          <TR height=1>
                            <TD width=1></TD>
                            <TD width=1></TD>
                            <TD width=1></TD>
                            <TD bgColor=#cccccc></TD>
                            <TD width=1></TD>
                            <TD width=1></TD>
                            <TD width=1></TD></TR>
                          <TR height=1>
                            <TD></TD>
                            <TD bgColor=#cccccc colSpan=2></TD>
                            <TD bgColor=#f7f7f7></TD>
                            <TD bgColor=#cccccc colSpan=2></TD>
                            <TD></TD></TR>
                          <TR height=1>
                            <TD></TD>
                            <TD bgColor=#cccccc></TD>
                            <TD bgColor=#f7f7f7 colSpan=3></TD>
                            <TD bgColor=#cccccc></TD>
                            <TD></TD></TR></TBODY></TABLE></TD></TR>
                    <TR>
                      <TD>
                        <TABLE width="100%" height="100%" border=0 cellPadding=0 
                        cellSpacing=0 style="TABLE-LAYOUT: fixed">
                          <TBODY>
                          <TR>
                            <TD width=1 bgColor=#cccccc></TD>
                            <TD id=oINNER 
                            style="PADDING-RIGHT: 1px; PADDING-LEFT: 3px; PADDING-TOP: 1px" 
                            align=middle bgColor=#f7f7f7>
                              <TABLE width="100%" 
                                border=1 
                              cellPadding=1 cellSpacing=1 borderColorLight=#bbbbbb borderColorDark=#ffffff 
                              bgColor=#ffffff>
                                <TBODY>
                                <TR bgColor=#fafafa>
                                <TD height=30 colSpan=4 align="left" bgColor=#e1e1e1><B>[ 
                                在线留言 ]</B></TD>
                                </TR>
                                <TR align="left" bgColor=#fafafa>
                                <TD class=td_padding1166 noWrap 
                                bgColor=#f7f7f7 height=36>留言类型：</TD>
                                <TD height=27 
                                colSpan=3 bgColor=#f7f7f7 class=td_padding1166><LABEL 
                                for=Classid></LABEL><SELECT id=Classid 
                                name=Classid>
                                <%=PinLuo.PinLuo_GetClass_Option("PinLuo_FeedbackClass",0,classid,-1)%>
                                </SELECT> <FONT 
                                color=#ff0000>* </FONT></TD></TR>
                                <TR align="left">
                                <TD class=td_padding1166 
                                bgColor=#f1f1f1 height=36>您的姓名：</TD>
                                <TD bgColor=#f1f1f1 class=td_padding1166><INPUT 
                                class=inputText id=Author maxLength=60 
                                name=Author> <FONT color=#ff0000>* 
                                </FONT></TD>
                                <TD class=td_padding1166 
                                bgColor=#f1f1f1>电子邮件：</TD>
                                <TD bgColor=#f1f1f1 class=td_padding1166><INPUT 
                                class=inputText id=Email maxLength=60 
                                name=Email> <FONT color=#ff0000>* 
                                </FONT></TD></TR>
                                <TR align="left" bgColor=#fafafa>
                                <TD class=td_padding1166 
                                bgColor=#f7f7f7 height=36>联系电话：</TD>
                                <TD bgColor=#f7f7f7 class=td_padding1166><INPUT 
                                class=inputText id=Phone maxLength=60 
                                name=Phone> <FONT color=#e4edf9>&nbsp; 
                                </FONT></TD>
                                <TD class=td_padding1166 
                                bgColor=#f7f7f7>联系QQ：</TD>
                                <TD bgColor=#f7f7f7 class=td_padding1166><INPUT 
                                class=inputText id=QQ maxLength=60 
                                name=QQ></TD></TR>
                                <TR align="left" bgColor=#ffffff>
                                <TD class=td_padding1166 
                                bgColor=#f1f1f1 height=36>留言主题：</TD>
                                <TD 
                                colSpan=3 bgColor=#f1f1f1 class=td_padding1166><INPUT class=inputText id=FeedbackTitle 
                                maxLength=60 size=60 name=FeedbackTitle> <FONT 
                                color=#ff0000>* </FONT></TD></TR>
                                <TR align="left" bgColor=#fafafa>
                                <TD class=td_padding1166 
                                bgColor=#f7f7f7 height=36>留言内容：</TD>
                                <TD 
                                colSpan=3 bgColor=#f7f7f7 class=td_padding1166><TEXTAREA class=inputText id=FeedbackContent onkeydown=gbcount(this.form.FB_Content,this.form.total,this.form.used,this.form.remain); onkeyup=gbcount(this.form.FB_Content,this.form.total,this.form.used,this.form.remain); name=FeedbackContent rows=6 cols=60></TEXTAREA> 
                                <FONT color=#ff0000>* </FONT><A 
                                href="javascript:admin_Size(-3,'FB_Content')"><IMG 
                                height=20 
                                src="../images/minus.gif" 
                                width=20 border=0 unselectable="on"></A> <A 
                                href="javascript:admin_Size(3,'FB_Content')"><IMG 
                                height=20 
                                src="../images/plus.gif" 
                                width=20 border=0 unselectable="on"></A> 
                                <BR>
                                最多字数：<INPUT class=inputText disabled 
                                maxLength=4 size=3 value=1000 
                                name=total>已用字数：<INPUT class=inputText disabled 
                                maxLength=4 size=3 value=0 name=used>剩余字数：<INPUT 
                                class=inputText disabled maxLength=4 size=3 
                                value=1000 name=remain> </TD></TR>
                                <TR bgColor=#fafafa>
                                <TD class=td_padding1166 align=right 
                                bgColor=#f1f1f1 height=36>&nbsp;</TD>
                                <TD class=td_padding1166 bgColor=#f1f1f1 
                                colSpan=3><INPUT type=hidden 
                                value="addsave" name=act> <INPUT class=inputButton onclick=CheckFeedback(this.form) type=button value="提 交" name=Submit> 
<INPUT class=inputButton type=reset value="重 置" name=Submit> 
<INPUT class=inputButton id=reload onclick=window.location.reload() type=button value="刷 新" name=reload> 
                                &nbsp;&nbsp; </TD></TR></TBODY></TABLE></TD>
                          <TD width=1 
                    bgColor=#cccccc></TD></TR></TBODY></TABLE></TD></TR>
                    <TR height=3 width="100%">
                      <TD>
                        <TABLE style="TABLE-LAYOUT: fixed" height=3 
                        cellSpacing=0 cellPadding=0 width="100%" border=0>
                          <TBODY>
                          <TR height=1>
                            <TD width=1></TD>
                            <TD width=1 bgColor=#cccccc></TD>
                            <TD width=1 bgColor=#f7f7f7></TD>
                            <TD bgColor=#f7f7f7></TD>
                            <TD width=1 bgColor=#f7f7f7></TD>
                            <TD width=1 bgColor=#cccccc></TD>
                            <TD width=1></TD></TR>
                          <TR height=1>
                            <TD></TD>
                            <TD bgColor=#cccccc colSpan=2></TD>
                            <TD bgColor=#f7f7f7></TD>
                            <TD bgColor=#cccccc colSpan=2></TD>
                            <TD></TD></TR>
                          <TR height=1>
                            <TD colSpan=3></TD>
                            <TD bgColor=#cccccc></TD>
                    <TD 
                  colSpan=3></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></FORM></TD></TR></TBODY></TABLE></div>
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