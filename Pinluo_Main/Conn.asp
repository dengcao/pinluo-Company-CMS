<!--#include file="Site_Conn.asp"-->
<%
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