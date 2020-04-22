<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
<link href="<%=Pinluo.Pinluo_SiteUrl%>images/style.css" rel="stylesheet" type="text/css" />
</head>

<body>
  <div class="logo"><a href="<%=Pinluo.Pinluo_Siteurl%>"><img src="<%=Pinluo.Pinluo_Logo%>" border="0" /></a></div>
  <div class="dh">
    <ul>
      <%HtmlStr="<li><a class=""d"" href=""{$url}"" target=""{$Blank}"">{$title}</a></li>"
		PinLuo.PinLuo_GetDaohanglist HtmlStr,10,100,"PinLuo_Daohang"%>
	</ul>
  </div>
  <div class="banner"><img src="<%=Pinluo.Pinluo_Banner%>" /></div>