<!--#include file="../Pinluo_Main/Conn.asp"-->
<!--#include file="PinLuo_Class.asp"-->
<%
dim PinLuo
Set PinLuo = New PinLuo_Class
    PinLuo.DBConnBegin
	PinLuo.CheckPurview
	PinLuo.Pinluo_CheckPurviewAdmin(1)
	
Dim dbpath, Barwidth
Action=trim(request("Action"))

ObjInstalled_FSO=PinLuo.IsObjInstalled("Scripting.FileSystemObject")

dbpath = Server.MapPath(sqlDbPath)

Barwidth = 500
Response.Write "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">" & vbCrLf
Response.Write "<html xmlns=""http://www.w3.org/1999/xhtml"">"& vbCrLf
Response.Write "<head>"& vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>" & vbCrLf
Response.Write "<title>数据库管理</title>"& vbCrLf
Response.Write "<link href='images/style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<body class=""mainbg"">" & vbCrLf
Response.Write "<div id=""mainhearder""><span>您的位置：企业网站管理系统 >> 数据库管理</span></div>" & vbCrLf
Response.Write "<div id=hearder class=hearder2><span>数据库管理</span></div>" & vbCrLf
Response.Write "<div class=main5 id=main5>" & vbCrLf
Response.Write "<table width=98% border=0 align=center cellpadding=0 cellspacing=0>" & vbCrLf
Response.Write "<tr><td class=tableleft>" & vbCrLf

Select Case Action
Case "Backup"
    Call ShowBackup
Case "BackupData"
    Call BackupData
Case "Compact"
    Call ShowCompact
Case "CompactData"
    Call CompactData
Case "Restore"
    Call ShowRestore
Case "RestoreData"
    Call RestoreData
Case "Init"
    Call ShowInit
Case "Clear"
    Call ShowInit
Case "SpaceSize"
    Call SpaceSize
Case Else
    FoundErr = True
    ErrMsg = ErrMsg & "错误参数！"
End Select
If FoundErr = True Then
    PinLuo.DBConnEnd
    PinLuo.PinLuo_WriteMsg ErrMsg,ComeUrl
End If
Response.Write "</td>	</tr>	</table>" & vbCrLf
Response.Write "</div><br /><br />&nbsp; " & vbCrLf
Response.Write "</body></html>"
PinLuo.DBConnEnd
Set PinLuo = Nothing

Sub ShowBackup()
    Response.Write "<form method='post' action='Pinluo_Database.asp?action=BackupData'>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='0' class='table1'>"
    Response.Write "  <tr class='title'>"
    Response.Write "      <td align='center' height='22' valign='middle' class='tableleft1' style='text-align:center;'><b>备 份 数 据 库</b></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr>"
    Response.Write "    <td height='150' align='left' valign='middle'>"
    Response.Write "<table cellpadding='0' cellspacing='0' border='0' width='100%'>"
    Response.Write "  <tr>"
    Response.Write " <td width='200' height='33' align='right' class='tableleft1'>备份目录：</td>"
    Response.Write " <td colspan='2'><input type=text size=36 name=bkfolder class='input' value=Databackup>"
    Response.Write " &nbsp;相对路径目录，如目录不存在，将自动创建</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr>"
    Response.Write " <td width='200' height='34' align='right' class='tableleft1'>备份名称：</td>"
    Response.Write " <td height='34' colspan='2'><input type=text size=36 class='input' name=bkDBname value='" & replace(Date,"/","-") & "'>"
    Response.Write " &nbsp;不用输入文件名后缀（默认为“.asa”）。如有同名文件，将覆盖</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr align='center'>"
    Response.Write " <td height='40' colspan='3'><br><input name='submit' class=button type=submit value=' 开始备份 '"
    If DataBaseType<>"ACCESS" Or ObjInstalled_FSO = False Then
        Response.Write " disabled"
    End If
    Response.Write "><br>&nbsp;</td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    If ObjInstalled_FSO = False Then
        Response.Write "<b><font color=red>你的服务器不支持 FSO(Scripting.FileSystemObject)! 不能使用本功能</font></b>"
    End If
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "</form>"
    If DataBaseType<>"ACCESS" Then
        Response.Write "<br><b>说明：</b><br>&nbsp;&nbsp;&nbsp;&nbsp;您使用的是SQL版，请直接使用SQL2000提供的数据库备份功能进行备份！<br><br>"
    End If
End Sub

Sub ShowCompact()
    Response.Write "<form method='post' action='Pinluo_Database.asp?action=CompactData'>"
    Response.Write "<table class='table1' width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>"
    Response.Write " <tr class='title'>"
    Response.Write "     <td align='center' height='22' valign='middle' class='tableleft1' style='text-align:center;'><b>数据库在线压缩</b></td>"
    Response.Write " </tr>"
    Response.Write " <tr>"
    Response.Write "     <td align='center' height='150' valign='middle'>"
    Response.Write "      <br>"
    Response.Write "      <br>"
    Response.Write "      压缩前，建议先备份数据库，以免发生意外错误。 <br>"
    Response.Write "      <br>"
    Response.Write "      <br>"
    Response.Write " <input name='submit' type=submit class=button value=' 压缩数据库 '"
    If DataBaseType<>"ACCESS" Then
        Response.Write " disabled"
    End If
    Response.Write "><br><br>"
    If ObjInstalled_FSO = False Then
        Response.Write "<b><font color=red>你的服务器不支持 FSO(Scripting.FileSystemObject)! 不能使用本功能</font></b>"
    End If
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "</form>"
    If DataBaseType<>"ACCESS" Then
        Response.Write "<br><b>说明：</b><br>&nbsp;&nbsp;&nbsp;&nbsp;您使用的是SQL版，无需进行压缩操作！<br><br>"
    End If
End Sub

Sub ShowRestore()
    Response.Write "<form method='post' action='Pinluo_Database.asp?action=RestoreData'>"
    Response.Write "<table width='100%' class='table1' border='0' align='center' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr class='title'>"
    Response.Write "    <td align='center' height='22' valign='middle'><b>数据库恢复</b></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr>"
    Response.Write "    <td align='center' height='150' valign='middle'>"
    Response.Write "      <table width='100%' border='0' cellspacing='0' cellpadding='0'>"
    Response.Write "        <tr>"
    Response.Write "          <td width='200' height='30' align='right' class='tableleft1'>原备份数据库路径（相对）：</td>"
    Response.Write "          <td height='30'><input name=backpath class='input' type=text id='backpath' value='Databackup\data.asa' size=50 maxlength='200'></td>"
    Response.Write "        </tr>"
    Response.Write "        <tr align='center'>"
    Response.Write "          <td height='40' colspan='2'><input class=button name='submit' type=submit value=' 恢复数据 '"
    If DataBaseType<>"ACCESS" Or ObjInstalled_FSO = False Then
        Response.Write " disabled"
    End If
    Response.Write ">"
    Response.Write "          </td>"
    Response.Write "        </tr>"
    Response.Write "      </table>"
    If ObjInstalled_FSO = False Then
        Response.Write "<b><font color=red>你的服务器不支持 FSO(Scripting.FileSystemObject)! 不能使用本功能</font></b>"
    End If
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "</form>"
    If DataBaseType<>"ACCESS" Then
        Response.Write "<br><b>说明：</b><br>&nbsp;&nbsp;&nbsp;&nbsp;您使用的是SQL版，请直接使用SQL2000提供的数据库恢复功能进行恢复！<br><br>"
    Else
        Response.Write "<br><b>说明：</b><br>&nbsp;&nbsp;&nbsp;&nbsp;原备份数据库的扩展名必须为：asa或者asp<br><br>"
    End If
End Sub

Sub SpaceSize()
    On Error Resume Next
    Response.Write "<table class='table1' width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td align='center' height='22' valign='middle' class='tableleft1' style='text-align:center'><b>系统文件占用空间情况</b></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr>"
    Response.Write "    <td width='100%' height='150' valign='middle' class='tableleft1' style='text-align:left;padding-left:80px;'>"
    Response.Write "    <blockquote><br>"
    Response.Write "      上传文件占用空间：" & ShowSpace("editor/uploadfile")
	Response.Write "      <br>"
    Response.Write "      主程序占用空间：" & ShowSpace("Pinluo_Main/")
    Response.Write "      <br>"
    Response.Write "      数据库占用空间：" & ShowSpace("Pinluo_Main/Database")
    Response.Write "      <br>"
	Response.Write "      备份数据库占用空间：" & ShowSpace("Pinluo/Databackup")
    Response.Write "      <br>"
    Response.Write "      后台占用空间：" & ShowSpace("Pinluo")
    Response.Write "      <br>"
    Response.Write "      网站占用空间总计：" & ShowSpace(" ")
    Response.Write "    </blockquote>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
End Sub

Sub BackupData()
    Dim bkfolder, bkdbname
    bkfolder = Trim(Request("bkfolder"))
    bkdbname = Trim(Request("bkdbname"))
    If bkfolder = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定备份目录！</li>"
		ComeUrl = "PinLuo_Database.asp?Action=Backup"
    End If
    If bkdbname = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>请指定备份文件名</li>"
		ComeUrl = "PinLuo_Database.asp?Action=Backup"
    End If
    If FoundErr = True Then Exit Sub
    bkfolder = Server.MapPath(bkfolder)
	
	SET fso=server.CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(dbpath) Then
        If fso.FolderExists(bkfolder) = False Then
            fso.CreateFolder (bkfolder)
        End If
        fso.copyfile dbpath, bkfolder & "\" & bkdbname & ".asa"
		FoundErr = True
        ErrMsg = ErrMsg & "<font color=red><b>备份数据库成功，备份的数据库为：</b></font><br>" & bkfolder & "\" & bkdbname & ".asa"
		ComeUrl = "PinLuo_Database.asp?Action=Backup"
    Else
        FoundErr = True
        ErrMsg = ErrMsg & "找不到源数据库文件，请检查Conn.asp中的配置。"
		ComeUrl = "PinLuo_Database.asp?Action=Backup"
    End If
	set fso=nothing
End Sub

Sub CompactData()
    On Error Resume Next

    Dim Engine, strDBPath
    PinLuo.DBConnEnd

    strDBPath = Left(dbpath, InStrRev(dbpath, "\"))
	SET fso=server.CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(dbpath) Then
        Set Engine = Server.CreateObject("JRO.JetEngine")
        Engine.CompactDatabase "Provider=Microsoft.ACE.OLEDB.12.0;Data Source = " & dbpath, " Provider=Microsoft.ACE.OLEDB.12.0;Data Source = " & strDBPath & "temp.accdb;Jet OLEDB:Engine Type=5"
        fso.copyfile strDBPath & "temp.accdb", dbpath
        fso.DeleteFile (strDBPath & "temp.accdb")
        Set Engine = Nothing
		FoundErr = True
        ErrMsg = ErrMsg & "<font color=red><b>数据库压缩成功!</b></font><br>"
		ComeUrl = "PinLuo_Database.asp?Action=Compact"
    Else
		FoundErr = True
        ErrMsg = ErrMsg & "数据库没有找到!"
		ComeUrl = "PinLuo_Database.asp?Action=Compact"
    End If
	set fso=nothing
    If Err.Number <> 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & Err.Description
        Err.Clear
        Exit Sub
    End If
End Sub

Sub RestoreData()
    Dim backpath
    backpath = Trim(Request.Form("backpath"))
    If backpath = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "请指定原备份的数据库文件名！"
        Exit Sub
    End If
    If GetFileExt(backpath) <> "asa" And GetFileExt(backpath) <> "asp" Then
        FoundErr = True
        ErrMsg = ErrMsg & "原备份数据库文件的扩展名必须为asa或asp！"
        Exit Sub
    End If
    backpath = Server.MapPath(backpath)
	SET fso=server.CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(backpath) Then
        fso.copyfile backpath, dbpath
		FoundErr = True
        ErrMsg = ErrMsg & "成功恢复数据！"
    Else
        FoundErr = True
        ErrMsg = ErrMsg & "找不到指定的备份文件！"
    End If
	set fso=nothing
End Sub


Function ShowSpace(FolderPath)
    Dim ft, fd, fs, TotalSize, SpaceSize, FolderBarWidth, arrPath, strSize, i
	SET fso=server.CreateObject("Scripting.FileSystemObject")
    Set ft = fso.GetFolder(Server.MapPath(InstallDir))
    TotalSize = ft.size
    If TotalSize = 0 Then TotalSize = 1

    SpaceSize = 0
    arrPath = Split(FolderPath, "|")
    For i = 0 To UBound(arrPath)
        If arrPath(i) = "SiteRoot" Then
            Set fd = fso.GetFolder(Server.MapPath(InstallDir))
            For Each fs In fd.Files
                SpaceSize = SpaceSize + fs.size
            Next
        Else
            If fso.FolderExists(Server.MapPath(InstallDir & arrPath(i))) Then
                Set fd = fso.GetFolder(Server.MapPath(InstallDir & arrPath(i)))
                SpaceSize = SpaceSize + fd.size
            End If
        End If
    Next
    FolderBarWidth = CLng((SpaceSize / TotalSize) * Barwidth)

    strSize = SpaceSize & "&nbsp;Byte"
    If SpaceSize > 1024 Then
       SpaceSize = (SpaceSize / 1024)
       strSize = FormatNumber(SpaceSize, 2, vbTrue, vbFalse, vbTrue) & "&nbsp;KB"
    End If
    If SpaceSize > 1024 Then
       SpaceSize = (SpaceSize / 1024)
       strSize = FormatNumber(SpaceSize, 2, vbTrue, vbFalse, vbTrue) & "&nbsp;MB"
    End If
    If SpaceSize > 1024 Then
       SpaceSize = (SpaceSize / 1024)
       strSize = FormatNumber(SpaceSize, 2, vbTrue, vbFalse, vbTrue) & "&nbsp;GB"
    End If
    strSize = "<font face=verdana>" & strSize & "</font>"
    ShowSpace = "&nbsp;<img src='images/bar.gif' width='" & FolderBarWidth & "' height='10' title='目录：" & InstallDir&FolderPath & "'>&nbsp;" & strSize
	set fso=nothing
End Function

Function GetOtherFolder()
    Dim ft, fd, strOther, strSystem, arrPath
	SET fso=server.CreateObject("Scripting.FileSystemObject")
    strSystem = "AD|Admin|AuthorPic|BlogPic|CopyFromPic|Count|Database|Editor|FriendSite|Images|Inc|JS|Language|Reg|Sdms|SiteMap|Skin|Temp|User|xml"

    Set ft = fso.GetFolder(Server.MapPath(InstallDir))
    For Each fd In ft.SubFolders
        If InStr("|" & strSystem & "|", "|" & fd.name & "|") = 0 Then
            If strOther = "" Then
                strOther = fd.name
            Else
                strOther = strOther & "|" & fd.name
            End If
        End If
    Next
    GetOtherFolder = strOther
	set fso=nothing
End Function

%>