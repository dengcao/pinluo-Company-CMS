<!--#include file="Site_Conn.asp"-->
<%
'☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆
'☆                                                                         ☆
'☆  系 统：品络企业网站管理系统 Version 1.5                                    ☆
'☆  日 期：2010-05                                                          ☆
'☆  开 发：草札(www.caozha.com)                                              ☆
'☆  声 明: 使用本系统必须保留此版权声明信息！本文字不会影响系统的正常运行。            ☆
'☆  Copyright (C) 2010 品络(www.pinluo.com) All Rights Reserved.            ☆
'☆                                                                         ☆
'☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆
 
 
    Dim sqlLocalName, sqlUsername, sqlPassword, sqlDatabaseName,sqlDbPath,DataBaseType,InstallDir
	
	InstallDir="/" '系统安装目录，如根目录请填/
	DataBaseType="ACCESS" '设置数据库类型，ACCESS或SQL	
	
	'//如果是ACCESS设置以下参数
	sqlDbPath = sqlDbPath_access
	
	'//如果是SQL Server设置以下参数
	sqlLocalName = "127.0.0.1"   '连接IP,本机默认用local或127.0.0.1
	sqlUsername = "pinluo.com"          '用户名
	sqlPassword = "pinluo123"           '用户密码
	sqlDatabaseName = "pinluo"       '数据库名
%>