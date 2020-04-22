<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
<!--#include file="../Pinluo_Main/Config.asp"-->
<%
dim PinLuo
Set PinLuo = New PinLuo_Class
    PinLuo.DBConnBegin
	PinLuo.PinLuo_ViewSiteConfig
	Feedbackid=PinLuo.Pinluo_GetUrlID
	PinLuo.PinLuo_ViewFeedbackItem Feedbackid,"PinLuo_FeedbackList","PinLuo_FeedbackClass"
	classid=Pinluo.classid
	classname=Pinluo.classname
	Feedbacktitle=Pinluo.SEO_Title&"_"
	keywords=PinLuo.SEO_Keyword
	descriptions=PinLuo.SEO_Description
%>
<title><%=Feedbacktitle&Pinluo.Pinluo_SeoTitle%></title>
<meta name="keywords" content="<%=keywords%>">
<META NAME="description" CONTENT="<%=descriptions%>">
<link href="<%=Pinluo.Pinluo_SiteUrl%>images/style.css" rel="stylesheet" type="text/css" />
<style type="text/css">
body,td,th {
	font-size: 14px;
}
body {
	background-color: #FFF;
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
</style>
</head>
<body>
<br>
<table width="98%" height=28 border=0 align="center" cellPadding=0 cellSpacing=0 style="table-LAYOUT: fixed">
  <tbody>
    <tr height=3 width="100%">
      <td><table style="table-LAYOUT: fixed" height=3 cellSpacing=0 cellPadding=0 width="100%" border=0>
          <tbody>
            <tr height=1>
              <td width=1></td>
              <td width=1></td>
              <td width=1></td>
              <td bgColor="#CCCCCC"></td>
              <td width=1></td>
              <td width=1></td>
              <td width=1></td>
            </tr>
            <tr height=1>
              <td></td>
              <td bgColor="#CCCCCC" colSpan=2></td>
              <td bgColor="#F7F7F7"></td>
              <td bgColor="#CCCCCC" colSpan=2></td>
              <td></td>
            </tr>
            <tr height=1>
              <td></td>
              <td bgColor="#CCCCCC"></td>
              <td bgColor="#F7F7F7" colSpan=3></td>
              <td bgColor="#CCCCCC"></td>
              <td></td>
            </tr>
          </tbody>
      </table></td>
    </tr>
    <tr>
      <td><table style="table-LAYOUT: fixed" height="100%" cellSpacing=0 cellPadding=0 border=0>
          <tbody>
            <tr>
              <td width=1 bgColor="#CCCCCC"></td>
              <td id=oINNER bgColor="#F7F7F7" align="center" style="padding-left:3px; padding-right:1px; padding-top:1px;">



	<table width="100%" border="1" align="center" cellpadding="1" cellspacing="1" bordercolorlight="#BBBBBB" bordercolordark="#FFFFFF" bgcolor="#FFFFFF">
          <tr bgcolor="#666666">
            <td height="30" colspan="2" bgcolor="#E1E1E1"><strong style="float:left">[ 查看留言详情 ]</strong><div style="float:right;padding-right:20px;"><a href="<%=Pinluo.Pinluo_Siteurl%>">返回首页</a></div></td>
            </tr>
          <tr bgcolor="#E2E2E2">
            <td width="15%" height="27" align="right" nowrap bgcolor="#F7F7F7" class="td_padding1166">留言类型：</td>
            <td bgcolor="#F7F7F7" class="td_padding1166"><%=PinLuo.classname%>
              &nbsp;</td>
          </tr>
          <tr bgcolor="#F1F1F1">
            <td height="27" align="right" class="td_padding1166">留言主题：</td>
            <td class="td_padding1166"><%=Pinluo.FeedbackTitle%>
              &nbsp;</td>
          </tr>
          <tr bgcolor="#E2E2E2">
            <td height="27" align="right" bgcolor="#F7F7F7" class="td_padding1166">留言内容：</td>
            <td bgcolor="#F7F7F7" class="td_padding1166"><%=Pinluo.FeedbackContent%></td>
          </tr>
          <tr bgcolor="#F1F1F1">
            <td height="27" align="right" class="td_padding1166">留言时间：</td>
            <td class="td_padding1166"><%=Pinluo.UpdateTime%>
              &nbsp;</td>
          </tr>
          <tr bgcolor="#F1F1F1">
            <td height="27" align="right" class="td_padding1166">浏览次数：</td>
            <td class="td_padding1166"><%=Pinluo.hits%>
              &nbsp;</td>
          </tr>
          <tr bgcolor="#F9F9F9">
            <td height="27" align="right" bgcolor="#F7F7F7" class="td_padding1166" style="color:#F00">回复留言：</td>
            <td bgcolor="#F7F7F7" class="td_padding1166" style="table-layout:fixed;word-break:break-all; line-height:23px; padding-left:4px; padding-right:4px;"><%=Pinluo.ReplyContent%>
              &nbsp;</td>
          </tr>
          <tr bgcolor="#F1F1F1">
            <td height="27" align="right" class="td_padding1166">回复时间：</td>
            <td class="td_padding1166"><%=Pinluo.ReplyTime%>
              &nbsp;</td>
          </tr>
          <tr bgcolor="#F9F9F9">
            <td height="35" colspan="2" align="right" bgcolor="#F7F7F7">						  <a href="javascript:window.print();">『打印本页』</a> &nbsp;<a href="javascript:window.close();">『关闭窗口』</a></td>
            </tr>
      </table>
 

</td>
              <td width=1 bgColor="#CCCCCC"></td>
            </tr>
          </tbody>
      </table></td>
    </tr>
    <tr height=3 width="100%">
      <td><table style="table-LAYOUT: fixed" height=3 cellSpacing=0 cellPadding=0 width="100%" border=0>
          <tbody>
            <tr height=1>
              <td width=1></td>
              <td width=1 bgColor="#CCCCCC"></td>
              <td width=1 bgColor="#F7F7F7"></td>
              <td bgColor="#F7F7F7"></td>
              <td width=1 bgColor="#F7F7F7"></td>
              <td width=1 bgColor="#CCCCCC"></td>
              <td width=1></td>
            </tr>
            <tr height=1>
              <td></td>
              <td bgColor="#CCCCCC" colSpan=2></td>
              <td bgColor="#F7F7F7"></td>
              <td bgColor="#CCCCCC" colSpan=2></td>
              <td></td>
            </tr>
            <tr height=1>
              <td colSpan=3></td>
              <td bgColor="#CCCCCC"></td>
              <td colSpan=3></td>
            </tr>
          </tbody>
      </table></td>
    </tr>
  </tbody>
</table>
</body>
</html>	
<%PinLuo.DBConnEnd
Set PinLuo = Nothing%>