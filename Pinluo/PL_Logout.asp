<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Pinluo_Main/Conn.asp"-->
<!--#include file="PinLuo_Class.asp"-->
<%
dim PinLuo
Set PinLuo = New PinLuo_Class
    PinLuo.DBConnBegin
	PinLuo.PinLuo_LogOut
	PinLuo.DBConnEnd
Set PinLuo = Nothing
%>