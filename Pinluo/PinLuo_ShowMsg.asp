<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>系统提示信息-品络企业网站系统</title>
<meta content="品络科技,网站管理系统,企业网站管理系统,内容管理系统(CMS),网上商店管理系统,网站建设" name="Keywords" />
<meta content="品络科技成立于2005年6月，是一家集互联网基础服务、互联网应用软件开发、业务解决方案销售及服务于一体的高新技术企业。公司网址：www.5300.cn，品络互联：www.pinluo.com" name="Description"/>
<link rel="stylesheet" href="images/style.css" type="text/css">
<script language="javascript" src="js/Pinluo.js" type="text/javascript"></script>
</head>

<body class="mainbg" style="text-align:center;">
	<div style="width:500px; text-align:left; margin:auto; margin-top:20px;">
	<div id="hearder" class="hearder2"><span>系统提示！</span></div>
	<div class="main5" id="main5">

	<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
	<tr>
	<td class="tableleft1">
    
        <table width="100%" border="0" align="left" cellpadding="0" cellspacing="0" class="table1">
		<tr>
		<td style="padding:12px; text-align:left; font-weight:normal; font-size:13px;"><%=Request.Cookies("pinluo")("PinLuo_WriteMsg_ErrMsg")%><br><br /><center><input name='button' onclick="location.href='<%=Request.Cookies("pinluo")("PinLuo_WriteMsg_comeurl")%>';" class=button type=button value=' 确 认 '></center></td>
		</tr>

		</table>
       		
	</td>
	</tr>
	</table>
	</div>
            <br />
<br />
&nbsp; </div>
</body>
</html>