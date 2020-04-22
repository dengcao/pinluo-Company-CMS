<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
option explicit
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
%>
<!--#include file="Pinluo_Config.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<link href="../Pinluo_Main/css/User_Login.css" rel="stylesheet" type="text/css" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta content="品络科技,网站管理系统,企业网站管理系统,内容管理系统(CMS),网上商店管理系统,网站建设" name="Keywords" />
<meta content="品络科技成立于2005年6月，是一家集互联网基础服务、互联网应用软件开发、业务解决方案销售及服务于一体的高新技术企业。公司网址：www.5300.cn，品络互联：www.pinluo.com" name="Description"/>

<title>管理员登录</title></head>
<body id="userlogin_body">
<form name="Login" method="post" action="PinLuo_Chklogin.asp" id="Login">
<div>
</div>
<div>
</div>
<div id="panSiteFactory">
	
    <div id="siteFactoryLogin">
        <div id="user_login">
            <dl>
                <dd id="user_top">
                    <ul>
                        <li class="user_top_l"></li>
                        <li class="user_top_c"></li>
                        <li class="user_top_r"></li>
                    </ul>
                </dd>
                <dd id="user_main">
                    <ul>
                        <li class="user_main_l"></li>
                        <li class="user_main_c">
                            <div class="user_main_box">
                                <ul>
                                    <li class="user_main_text">
                                        管理员：
                                    </li>
                                    <li class="user_main_input">
                                        <input name="UserName" type="text" maxlength="20" id="UserName" class="TxtUserNameCssClass" />
                                    </li>
                                </ul>
                                <ul>
                                    <li class="user_main_text">
                                        密 码：
                                    </li>
                                    <li class="user_main_input">
                                        <input name="UserPassword" type="password" id="UserPassword" class="TxtPasswordCssClass" /></li>
                                </ul>
                                
                                <ul>
                                    <li class="user_main_text">验证码： </li>
                                    <li class="user_main_input">
                                        <input name="CheckCode" type="text" id="CheckCode" class="TxtValidateCodeCssClass" /><img src="../Pinluo_Main/CheckCode/checkcode<%=CheckCodeType%>.asp?u=" height="<%if CheckCodeType=3 then response.Write("20") else response.Write("15")%>" align="absmiddle" style="cursor:pointer" onClick="this.src+=parseInt(Math.random()*10)" alt="看不清楚？点击更换下一张。">
                                    </li>
                                </ul>
                            </div>
                        </li>
                        <li class="user_main_r">
                            <input type="image" name="IbtnEnter" id="IbtnEnter" class="IbtnEnterCssClass" src="../Pinluo_Main/images/user_botton.gif" style="border-width:0px;" /></li>
                    </ul>
                </dd>
                <dd id="user_bottom">
                    <ul>
                        <li class="user_bottom_l"></li>
                        <li class="user_bottom_c admin_bottom_c"></li>
                        <li class="user_bottom_r"></li>
                    </ul>
                </dd>
            </dl>
        </div>


	</div>
       </div>

</div>
</form>
</body>
</html>