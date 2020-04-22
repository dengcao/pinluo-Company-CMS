<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Pinluo_Main/Conn.asp"-->
<!--#include file="../Pinluo_Main/Inc/md5.asp"-->
<!--#include file="PinLuo_Class.asp"-->
<%
dim PinLuo
Set PinLuo = New PinLuo_Class
    PinLuo.DBConnBegin
	PinLuo.ValidateLoginData
If FoundErr = True Then
    ComeUrl="PL_login.asp"
	PinLuo.PinLuo_WriteMsg ErrMsg,ComeUrl
   	Response.End()
End If
	PinLuo.DBConnEnd
Set PinLuo = Nothing
%>