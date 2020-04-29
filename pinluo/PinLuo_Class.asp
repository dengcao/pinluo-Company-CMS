<!--#include file="Pinluo_Config.asp"-->
<%  

'☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆
'☆                                                                         ☆
'☆  系 统：品络企业网站管理系统 Version 1.5                                    ☆
'☆  日 期：2010-05                                                          ☆
'☆  开 发：草札(www.caozha.com)                                              ☆
'☆  鸣 谢：穷店(www.qiongdian.com) 品络(www.pinluo.com)                      ☆
'☆  声 明: 使用本程序源码必须保留此版权声明等相关信息！                            ☆
'☆  Copyright ©2010 www.caozha.com All Rights Reserved.                    ☆
'☆                                                                         ☆
'☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆

Class PinLuo_Class
Public objConn,objCmd
Public Pinluo_SiteName,Pinluo_SeoTitle,Pinluo_DelProImg,Pinluo_IsFeedback,Pinluo_Version,Pinluo_Empower,Pinluo_SiteUrl,Pinluo_SeoIndexTitle,Pinluo_SeoIndexKeyword,Pinluo_SeoIndexMS,Pinluo_Logo,Pinluo_Banner
Public ClassID,ClassName,SEO_Title,SEO_Keyword,SEO_Description,ClassContents,ParentID,Depth,IsOuter,PathUrl,Visible,OrderID,ChildID
Public InfoID,InfoTitle,InfoContent,InfoImg,Author,Origin
Public UpdateTime,hits,Shenhe
Public ProID,ProName,ProContent,ProImg1,ProImg2,ProPrice1,ProPrice2,Saled,Jian,Hot,Cheap
Public UserName,UserPassword,UserPassed,RealName,Mobile,Email,LastLoginIP,LastLoginTime,LastLogoutTime,LoginTimes,UserPopedom,UserWarning
Public Feedbackid,FeedbackTitle,FeedbackContent,Phone,QQ,ReplyContent,ReplyTime
Public Block_ID,Block_Title,Block_Content,Block_Time
Public Daohang_ID,Daohang_Title,Daohang_Url,Daohang_Blank,Daohang_order
Public StrUserPopedom,UserPopedomCheck
Public page,Pinluo_showpage_temp
Public RsList_i,ErrMsg

Private Sub Class_Initialize()
    StrUserPopedom="网站设置|数据库操作|栏目管理|管理员管理|产品管理|信息管理|留言反馈"
    UserPopedomCheck="0|0|0|0|0|0|0"
End Sub

Private Sub Class_Terminate()  
End Sub

Public Sub DBConnBegin()
  On Error Resume Next
  If IsObject(objConn) = True Then Exit Sub
  If DataBaseType="ACCESS" Then
	strConn = "Provider=Microsoft.Jet.OleDB.4.0; Data Source = " & Server.MapPath(sqlDbPath)
  Else
	strConn = "Provider=Sqloledb; User ID=" & SqlUsername & "; Password=" & SqlPassword & "; Initial Catalog = " & SqlDatabaseName & "; Data Source=" & SqlLocalName & ";"
  End If
	Set objConn = Server.CreateObject("ADODB.Connection")
	objConn.Open strConn
	If Err Then
		err.Clear
		Set objConn = Nothing
		FoundErr = True
        ErrMsg = ErrMsg & "<font color=red><b>数据库连接出错，请检查数据库参数设置。</b></font><br>"
		ComeUrl = ""
	End If
	If FoundErr = True Then
    PinLuo.DBConnEnd
    PinLuo.PinLuo_WriteMsg ErrMsg,ComeUrl
    End If
	Set objCmd = Server.CreateObject("ADODB.Command")
End Sub

Public Sub DBConnEnd()
	On Error Resume Next
	If IsObject(objCmd) Then
		objCmd.Close
		Set objCmd = Nothing
	End if
	If IsObject(objConn) Then
		objConn.Close
		Set objConn = Nothing
	End if
End Sub

Public Sub PinLuo_ViewSiteConfig()
	set RsList = server.CreateObject("adodb.recordset")
	SqlList = "SELECT * FROM PinLuo_SiteConfig"
	RsList.open SqlList, objConn, 1, 1
	if not(RsList.eof) then
	  Pinluo_SiteName=RsList("Pinluo_SiteName")
	  Pinluo_SeoTitle=RsList("Pinluo_SeoTitle")
	  Pinluo_DelProImg=RsList("Pinluo_DelProImg")
	  Pinluo_IsFeedback=RsList("Pinluo_IsFeedback")
	  Pinluo_Version=RsList("Pinluo_Version")
	  Pinluo_Empower=RsList("Pinluo_Empower")
	  Pinluo_SiteUrl=RsList("Pinluo_SiteUrl")
	  Pinluo_SeoIndexTitle=RsList("Pinluo_SeoIndexTitle")
	  Pinluo_SeoIndexKeyword=RsList("Pinluo_SeoIndexKeyword")
	  Pinluo_SeoIndexMS=RsList("Pinluo_SeoIndexMS")
	  Pinluo_Logo=RsList("Pinluo_Logo")
	  Pinluo_Banner=RsList("Pinluo_Banner")
	end if
	RsList.close
	Set RsList = Nothing
End Sub

Public Sub PinLuo_SiteConfigEdit(Pinluo_SiteName,Pinluo_SeoTitle,Pinluo_DelProImg,Pinluo_IsFeedback,Pinluo_SiteUrl,Pinluo_SeoIndexTitle,Pinluo_SeoIndexKeyword,Pinluo_SeoIndexMS,Pinluo_Logo,Pinluo_Banner)
    if Pinluo_SiteName="" then PinLuo_SiteConfigEdit=false:exit Sub
	set RsList = server.CreateObject("adodb.recordset")
	SqlList = "SELECT * FROM PinLuo_SiteConfig"
	RsList.open SqlList, objConn, 1, 3
	if not(RsList.eof) then
	  RsList("Pinluo_SiteName")=Pinluo_SiteName
	  RsList("Pinluo_SeoTitle")=Pinluo_SeoTitle
	  RsList("Pinluo_DelProImg")=Pinluo_DelProImg
	  RsList("Pinluo_IsFeedback")=Pinluo_IsFeedback
	  RsList("Pinluo_SiteUrl")=Pinluo_SiteUrl
	  RsList("Pinluo_SeoIndexTitle")=Pinluo_SeoIndexTitle
	  RsList("Pinluo_SeoIndexKeyword")=Pinluo_SeoIndexKeyword
	  RsList("Pinluo_SeoIndexMS")=Pinluo_SeoIndexMS
	  RsList("Pinluo_Logo")=Pinluo_Logo
	  RsList("Pinluo_Banner")=Pinluo_Banner
	  RsList.update
	end if
	RsList.close
	Set RsList = Nothing
End Sub

Public Sub PinLuo_ViewClass(ParentID,Depth,GotoUrl,Datatable)
	set RsList = server.CreateObject("adodb.recordset")
	SqlList = "SELECT A.*, ( SELECT Count(B.ClassID) FROM "&Datatable&" AS B WHERE A.ClassID = B.ParentID) AS CountList FROM "&Datatable&" AS A where A.ClassID>0 "
	if ParentID<>"" then
	SqlList = SqlList&"and ParentID="&int(ParentID)&" "
	else
	SqlList = SqlList&"and ParentID=0 "
	end if
	if Depth<>"" then SqlList = SqlList&"and Depth<="&Depth&" "
	SqlList = SqlList&"ORDER BY A.OrderID Desc,A.ClassID "
	RsList.open SqlList, objConn, 1, 1
	RsList_i=RsList_i+1
	if (RsList.eof and RsList_i=1) then response.Write("<a href='?act=add&classid=0'><span style=font-weight:normal>还没有任何分类，点击添加分类</span></a>")
	do while not (RsList.eof)
	   response.Write("<li id=""class"&RsList("Depth")&""" onmouseover=""ShowMenu("&RsList("ClassID")&")"" onmouseout=""HideMenu("&RsList("ClassID")&")""")
	   if RsList("Depth")>1 then
	   response.Write(" style='padding-left:"&((RsList("Depth")-1)*15)&"px;'>|--- ")
	   else
	   response.Write(">")
	   end if
	   if GotoUrl="" then response.Write(RsList("ClassName")) else	response.Write("<a href='"&GotoUrl&"classid="&RsList("ClassID")&"'>"&RsList("ClassName")&"</a>")
	   response.Write("&nbsp;&nbsp;<font style='color:#999;font-weight:normal'>("&RsList("CountList")&")&nbsp;["&RsList("OrderID")&"]</font>")
	   if RsList("Visible")=false then response.Write("&nbsp;<font color=#888888>隐</font>")
	   if RsList("IsOuter")=true then response.Write("&nbsp;<font color=red>外</font>")
	   response.Write("<span id=""M"&RsList("ClassID")&""" style=display:none;font-weight:normal>&nbsp;&nbsp;&nbsp;&nbsp;<a href=?act=add&classid="&RsList("ClassID")&">添加小类</a> <a href=?act=edit&classid="&RsList("ClassID")&">修改</a> <a href=?act=del&classid="&RsList("ClassID")&" onclick=""return cf('"&RsList("ClassName")&"')"">删除</a></span></li>")
	   PinLuo_ViewClass RsList("ClassID"),Depth,GotoUrl,Datatable
	   RsList.movenext
	loop
	RsList.close
	Set RsList = Nothing
End Sub

Public Sub PinLuo_AddClass(ClassName,SEO_Title,SEO_Keyword,SEO_Description,ClassContents,ParentID,IsOuter,PathUrl,Visible,OrderID,Datatable)
    if isnumeric(ParentID)=false then ParentID=0
	if isnumeric(OrderID)=false then OrderID=0
	set RsList = server.CreateObject("adodb.recordset")
	SqlList = "SELECT * FROM "&Datatable&" where ClassID="&int(ParentID)
	RsList.open SqlList, objConn, 1, 3
	  if not RsList.eof then Depth=RsList("Depth")+1 else Depth=1
	RsList.addnew
	  RsList("ClassName")=ClassName
	  RsList("SEO_Title")=SEO_Title
	  RsList("SEO_Keyword")=SEO_Keyword
	  RsList("SEO_Description")=SEO_Description
	  RsList("ClassContents")=ClassContents
	  RsList("ParentID")=ParentID
	  RsList("Depth")=Depth
	  RsList("IsOuter")=eval(IsOuter)
	  RsList("PathUrl")=PathUrl
	  RsList("Visible")=eval(Visible)
	  RsList("OrderID")=OrderID
	  RsList("ChildID")=""
	RsList.update
	  ClassID=RsList("ClassID")
	RsList.close
	Set RsList = Nothing
	'更新上级孩子数
	PinLuo_UpdateAddClassChild ClassID,ParentID,Datatable
End Sub

Public Sub PinLuo_UpdateAddClassChild(NowClassID,ParentID,Datatable) '更新父目录的所有子分类，形式如1,2,3...
	set RsList = server.CreateObject("adodb.recordset")
	SqlList = "SELECT * FROM "&Datatable&" where ClassID="&int(ParentID)
	RsList.open SqlList, objConn, 1, 3
	  if not RsList.eof then
	     if trim(RsList("ChildID"))<>"" then
		 RsList("ChildID")=RsList("ChildID")&","&NowClassID
		 else
		 RsList("ChildID")=NowClassID
		 end if
		 ParentID=RsList("ParentID")
		 RsList.update
      end if
	RsList.close
	Set RsList = Nothing
	if ParentID>0 then PinLuo_UpdateAddClassChild NowClassID,ParentID,Datatable
End Sub

Public Sub PinLuo_EditClass(ClassName,SEO_Title,SEO_Keyword,SEO_Description,ClassContents,ClassID,IsOuter,PathUrl,Visible,OrderID,Datatable)
    if isnumeric(ClassID)=false then exit Sub
	if isnumeric(OrderID)=false then OrderID=0
	set RsList = server.CreateObject("adodb.recordset")
	SqlList = "SELECT * FROM "&Datatable&" where ClassID="&int(ClassID)
	RsList.open SqlList, objConn, 1, 3
	if not(RsList.eof) then
	  RsList("ClassName")=ClassName
	  RsList("SEO_Title")=SEO_Title
	  RsList("SEO_Keyword")=SEO_Keyword
	  RsList("SEO_Description")=SEO_Description
	  RsList("ClassContents")=ClassContents
	  RsList("IsOuter")=eval(IsOuter)
	  RsList("PathUrl")=PathUrl
	  RsList("Visible")=eval(Visible)
	  RsList("OrderID")=OrderID
	RsList.update
	end if
	RsList.close
	Set RsList = Nothing	
End Sub

Public Sub PinLuo_DeleteInfoClass(ClassID,Datatable,ListDatatable)
    if isnumeric(ClassID)=false then exit Sub
    set RsList = server.CreateObject("adodb.recordset")
	SqlList = "SELECT * FROM "&Datatable&" where ClassID="&int(ClassID)
	RsList.open SqlList, objConn, 1, 3
	if not RsList.eof then
	  ChildID=trim(RsList("ChildID"))
	  if ChildID<>"" then 
	    objconn.execute("delete from "&Datatable&" where classid in("&ChildID&")")'删除下属分类
		objconn.execute("delete from "&ListDatatable&" where classid in("&ChildID&")")
	  end if
	  ParentID=int(RsList("ParentID"))
	  objconn.execute("delete from "&ListDatatable&" where ClassID="&int(ClassID))'删除文章
	end if
	RsList.delete
	RsList.close
	Set RsList = Nothing
	'更新上级栏目孩子数
	ChildID=ClassID&","&ChildID
	PinLuo_UpdateDeleteInfoClassChild ChildID,ParentID,Datatable
End Sub

Public Sub PinLuo_UpdateDeleteInfoClassChild(ChildID,ParentID,Datatable) '更新父目录的所有子分类，形式如1,2,3...
	set RsList = server.CreateObject("adodb.recordset")
	SqlList = "SELECT * FROM "&Datatable&" where ClassID="&int(ParentID)
	RsList.open SqlList, objConn, 1, 3
	  if not RsList.eof then
	     if trim(RsList("ChildID"))<>"" then
		    ChildIDArr=Split(ChildID,",")
			ChildIDsmtp="|,"&RsList("ChildID")&",|"
            N=UBound(ChildIDArr)
            For i=0 To N
			  if trim(ChildIDArr(i))<>"" then
                 ChildIDsmtp=replace(ChildIDsmtp,","&ChildIDArr(i)&",",",")
			  end if
            Next
			ChildIDsmtp=trim(replace(ChildIDsmtp,"|,",""))
			ChildIDsmtp=trim(replace(ChildIDsmtp,",|",""))
			ChildIDsmtp=trim(replace(ChildIDsmtp,"|",""))
		 RsList("ChildID")=ChildIDsmtp
		 end if
		 ParentID=RsList("ParentID")
		 RsList.update
      end if
	RsList.close
	Set RsList = Nothing
	if ParentID>0 then PinLuo_UpdateDeleteInfoClassChild ChildID,ParentID,Datatable
End Sub

Public Sub PinLuo_ViewClassItem(ClassID,Datatable)
    if isnumeric(ClassID)=false then exit Sub
	set RsList = server.CreateObject("adodb.recordset")
	SqlList = "SELECT * FROM "&Datatable&" where ClassID="&int(ClassID)
	RsList.open SqlList, objConn, 1, 1
	if not(RsList.eof) then
	  ClassName=RsList("ClassName")
	  SEO_Title=Trim(RsList("SEO_Title"))
	  SEO_Keyword=Trim(RsList("SEO_Keyword"))
	  SEO_Description=Trim(RsList("SEO_Description"))
	  ClassContents=RsList("ClassContents")
	  ParentID=RsList("ParentID")
	  Depth=RsList("Depth")
	  IsOuter=RsList("IsOuter")
	  PathUrl=RsList("PathUrl")
	  Visible=RsList("Visible")
	  OrderID=RsList("OrderID")
	end if
	RsList.close
	Set RsList = Nothing
End Sub

Public Function PinLuo_ViewClassName(ClassID,Datatable)
    PinLuo_ViewClassName="未知"
    if isnumeric(ClassID)=false then exit Function
	set RsList = server.CreateObject("adodb.recordset")
	SqlList = "SELECT ClassName FROM "&Datatable&" where ClassID="&int(ClassID)
	RsList.open SqlList, objConn, 1, 1
	 if not RsList.eof then
	  PinLuo_ViewClassName=RsList("ClassName")
	 end if
	RsList.close
	Set RsList = Nothing	
End Function

Public Function PinLuo_GetClassChildID(ClassID,Datatable)
    PinLuo_GetClassChildID=""
    if isnumeric(ClassID)=false then exit Function
	set RsList = server.CreateObject("adodb.recordset")
	SqlList = "SELECT ChildID FROM "&Datatable&" where ClassID="&int(ClassID)
	RsList.open SqlList, objConn, 1, 1
	 if not RsList.eof then
	  PinLuo_GetClassChildID=trim(RsList("ChildID"))
	 end if
	RsList.close
	Set RsList = Nothing	
End Function

Public Function PinLuo_ViewClass_Select(Datatable,ParentID2,NowClassid,k2) '列表框显示分类名
    dim allnum
    Set ShowClass = objConn.Execute("select count(*) as allnum from "&Datatable)
	    Allnum=ShowClass("allnum")
		ShowClass.close
	set ShowClass = nothing
	SelectTemp=SelectTemp&PinLuo_GetClass_Option(Datatable,ParentID2,NowClassid,k2)
	PinLuo_ViewClass_Select = SelectTemp
End Function

Public Function PinLuo_GetClass_Option(Datatable,ParentID,NowClassid,k)   '树型列表显示分类名
    '如调用信息类所有栏目：PinLuo_GetClass_Option("PinLuo_InfoClass",0,0,-1)
    if trim(ParentID)="" then ParentID=0
	Set rsClass = objConn.Execute("select * from "&Datatable&" where ParentID="&int(ParentID)&" order by OrderID desc,ClassID asc")
	Do While Not rsClass.EOF
	strTemp=strTemp&"<option value='"&rsClass("ClassID")&"' "
	   if trim(NowClassid)=trim(rsClass("ClassID")) then
	     strTemp=strTemp&" selected='selected'"
	   end if
	   if rsClass("Visible")=false then
	      Visible_temp="(隐)"
	   else
	      Visible_temp=""
	   end if
	   if rsClass("IsOuter")=true then
	      IsOuter_temp="(外)"
	   else
	      IsOuter_temp=""
	   end if
	strTemp=strTemp&">" & PinLuo_tmp(k,"|&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;") &"|--&nbsp;"&rsClass("ClassName")&" "&Visible_temp&IsOuter_temp&"</option>"	
	strTemp=strTemp&PinLuo_GetClass_Option(Datatable,rsClass("ClassID"),NowClassid,k+1)
	rsClass.MoveNext
	Loop
	rsClass.close
	set rsClass = nothing
	PinLuo_GetClass_Option = strTemp
End Function

Public Function PinLuo_tmp(n,text)   '显示缩进符号
	For i = 0 To n
		PinLuo_tmp = PinLuo_tmp & text
	Next
End Function

Public Sub PinLuo_InfoListView(OnepageNum,SearchKeyword,SearchSelect,SearchClassID,Datatable,ClassDatatable)
  	set RsList = server.CreateObject("adodb.recordset")
	SqlList = " select * from "&Datatable&" "
	SqlList = SqlList + " where InfoID > 0 "
	if isnumeric(SearchClassID) and trim(SearchClassID)<>"" then
	    SearchClassChildID=PinLuo_GetClassChildID(SearchClassID,ClassDatatable)
	    if SearchClassChildID="" then
	       SqlList = SqlList + " and Classid="&int(SearchClassID)
		else
		   SqlList = SqlList + " and Classid in("&SearchClassID&","&SearchClassChildID&") "
		end if
	end if
	if SearchKeyword <> "" then
		Select Case SearchSelect
			Case "InfoTitle"
				SqlList = SqlList + " and InfoTitle like '%"&SearchKeyword&"%' " 
			Case "InfoContent"
				SqlList = SqlList + " and InfoContent like '%"&SearchKeyword&"%' " 
			Case "Author"
				SqlList = SqlList + " and Author like '%"&SearchKeyword&"%' " 
			Case "Origin"
				SqlList = SqlList + " and Origin like '%"&SearchKeyword&"%' " 
			Case Else
				SqlList = SqlList + " and InfoTitle&InfoContent&Author&Origin like '%"&SearchKeyword&"%' " 
		End Select
	End If
	    Select Case SearchSelect
	        Case "Shenhe1"
				SqlList = SqlList + " and Shenhe=false " 
			Case "Shenhe2"
				SqlList = SqlList + " and Shenhe=true " 
			Case "InfoImg"
				SqlList = SqlList + " and InfoImg<>'' " 
		End Select
	SqlList = SqlList + " ORDER BY OrderID desc, UpdateTime DESC , InfoID DESC "
	RsList.open SqlList, objConn, 1, 1
	if RsList.eof or RsList.bof then
		response.Write("<tr class=""stline""><td colspan=""9"">没有找到任何信息记录！</td></tr>")
	else 
		
	RsList.pagesize = OnepageNum
	page = trim(request("page"))
	If IsNumeric(page) = False Then 
		page = 1 
	Else
		page = cint(page)
	End If 
	if page < 1 then page = 1
	if page > RsList.pagecount then page = RsList.pagecount
	
	RsList.absolutepage = page
	for i = 1 to RsList.pagesize 
	j = RsList.pagesize * ( page - 1) + i
	
	response.Write("<tr class=""stline"" onMouseOver=""this.className='nd'"" onMouseOut=""this.className='stline'"">"&vbcr)
	response.Write("<td><input type=""checkbox"" name=""DelInfoID"" value="""&RsList("InfoID")&"""></td>"&vbcr)
	response.Write("<td>"&RsList("InfoID")&"</td>"&vbcr)
	response.Write("<td align=left><a href=""Pinluo_info.asp?Act=edit&Infoid="&RsList("InfoID")&"&classid="&SearchClassID&"&SearchKeyword="&SearchKeyword&"&SearchSelect="&SearchSelect&"&page="&page&""">"&gotTopic(RsList("InfoTitle"),200)&"</a>")
	if trim(RsList("InfoImg"))<>"" then response.Write("<font color=red>(图)</font>"&vbcr)
	response.Write("</td>"&vbcr)
	response.Write("<td>"&PinLuo_ViewClassName(RsList("Classid"),"PinLuo_InfoClass")&"</td>"&vbcr)
	response.Write("<td>"&GetCurrentDate(RsList("UpdateTime"),5)&"</td>"&vbcr)
	response.Write("<td>"&RsList("hits")&"</td>"&vbcr)
	response.Write("<td>"&ChkShenHe(RsList("shenhe"))&"</td>"&vbcr)
	response.Write("<td><input name=""InfoRank_"& RsList("InfoID") &""" type=""text"" id=""InfoRank_"& RsList("InfoID") &""" title="" 显示顺序，数值越大越靠前，默认为0 ""  value="""& RsList("OrderID")&""" size=""2"" maxlength=""12"" onFocus=""if(this.value=='"& RsList("OrderID")&"')this.value='';"" onBlur=""if(this.value=='')this.value='"&RsList("OrderID")&"';"" ></td>"&vbcr)
	response.Write("<td><a href=""Pinluo_info.asp?Act=edit&Infoid="&RsList("InfoID")&"&classid="&SearchClassID&"&SearchKeyword="&SearchKeyword&"&SearchSelect="&SearchSelect&"&page="&page&"""><img src=""images/icon_xg.gif"" width=16 height=16 border=0></a></td>"&vbcr)
	response.Write("</tr>"&vbcr)
	
			RsList.movenext
		if RsList.eof then exit for
	Next
	Pinluo_showpage_temp=Pinluo_showpage("PinLuo_Infolist","PinLuo_InfoList.asp",page,RsList, "SearchKeyword#" & SearchKeyword & "#SearchSelect#" & SearchSelect & "#classid#" & SearchClassID)
		
		
	end if 
	RsList.close
	set RsList = nothing
End Sub

Public Sub PinLuo_ViewInfoItem(InfoID,Datatable)
    if isnumeric(InfoID)=false then exit Sub
	set RsList = server.CreateObject("adodb.recordset")
	SqlList = "SELECT * FROM "&Datatable&" where InfoID="&InfoID
	RsList.open SqlList, objConn, 1, 1
	if not(RsList.eof) then
	  Classid=RsList("Classid")
	  InfoTitle=RsList("InfoTitle")
	  SEO_Title=RsList("SEO_Title")
	  SEO_Keyword=RsList("SEO_Keyword")
	  SEO_Description=RsList("SEO_Description")
	  InfoContent=RsList("InfoContent")
	  InfoImg=RsList("InfoImg")
	  Author=RsList("Author")
	  Origin=RsList("Origin")
	  UpdateTime=RsList("UpdateTime")
	  hits=RsList("hits")
	  OrderID=RsList("OrderID")
	  Shenhe=RsList("Shenhe")
	end if
	RsList.close
	Set RsList = Nothing
End Sub

Public Function PinLuo_AddInfo(Classid,InfoTitle,SEO_Title,SEO_Keyword,SEO_Description,InfoContent,InfoImg,Author,Origin,UpdateTime,hits,OrderID,Shenhe,Datatable)
    if isnumeric(Classid)=false then PinLuo_AddInfo=false:exit Function
	if isdate(UpdateTime)=false then UpdateTime=now
	if isnumeric(hits)=false then hits=0
	if isnumeric(OrderID)=false then OrderID=0
	set RsList = server.CreateObject("adodb.recordset")
	SqlList = "SELECT * FROM "&Datatable
	RsList.open SqlList, objConn, 1, 3
	RsList.addnew
      RsList("Classid")=Classid
	  RsList("InfoTitle")=InfoTitle
	  RsList("SEO_Title")=SEO_Title
	  RsList("SEO_Keyword")=SEO_Keyword
	  RsList("SEO_Description")=SEO_Description
	  RsList("InfoContent")=InfoContent
	  RsList("InfoImg")=InfoImg
	  RsList("Author")=Author
	  RsList("Origin")=Origin
	  RsList("UpdateTime")=UpdateTime
	  RsList("hits")=hits
	  RsList("OrderID")=OrderID
	  RsList("Shenhe")=eval(Shenhe)
	RsList.update
	  InfoID=RsList("InfoID")
	RsList.close
	Set RsList = Nothing
	PinLuo_AddInfo=true
End Function

Public Function PinLuo_EditInfo(Infoid,Classid,InfoTitle,SEO_Title,SEO_Keyword,SEO_Description,InfoContent,InfoImg,Author,Origin,UpdateTime,hits,OrderID,Shenhe,Datatable)
    if isnumeric(Infoid)=false then PinLuo_EditInfo=false:exit Function
	if isnumeric(Classid)=false then PinLuo_EditInfo=false:exit Function
	if isdate(UpdateTime)=false then UpdateTime=now
	if isnumeric(hits)=false then hits=0
	if isnumeric(OrderID)=false then OrderID=0
	set RsList = server.CreateObject("adodb.recordset")
	SqlList = "SELECT * FROM "&Datatable&" where Infoid="&int(Infoid)
	RsList.open SqlList, objConn, 1, 3
      RsList("Classid")=Classid
	  RsList("InfoTitle")=InfoTitle
	  RsList("SEO_Title")=SEO_Title
	  RsList("SEO_Keyword")=SEO_Keyword
	  RsList("SEO_Description")=SEO_Description
	  RsList("InfoContent")=InfoContent
	  RsList("InfoImg")=InfoImg
	  RsList("Author")=Author
	  RsList("Origin")=Origin
	  RsList("UpdateTime")=UpdateTime
	  RsList("hits")=hits
	  RsList("OrderID")=OrderID
	  RsList("Shenhe")=eval(Shenhe)
	RsList.update
	  InfoID=RsList("InfoID")
	RsList.close
	Set RsList = Nothing
	PinLuo_EditInfo=true
End Function

Public Sub DelInfoAll(DelInfoID,Datatable)
	If trim(DelInfoID)="" Then Exit Sub
	Dim RsGetList, sqlGetList
	Set RsGetList = Server.CreateObject("adodb.recordset")
	sqlGetList = " select * from "&Datatable&" where InfoID in ( "&DelInfoID&" ) "
	RsGetList.open sqlGetList, objConn, 1, 2
	While Not( RsGetList.bof or RsGetList.eof )
		Dim aSaveFileName
		Dim sSaveFileName
		sSaveFileName = RsGetList("InfoImg")
		If sSaveFileName <> "" Then 
			aSaveFileName = Split(sSaveFileName, "|")
			Dim i
			For i = 0 To UBound(aSaveFileName)
				if isnull( aSaveFileName(i) ) = false or aSaveFileName(i) <> "" then 
					DelFileUploadfile "",aSaveFileName(i)
				end if 
			Next
		End If
		RsGetList.delete
		RsGetList.movenext
	wend
	RsGetList.Close
	Set RsGetList = Nothing
End Sub

Public Sub PinLuo_ProductListView(OnepageNum,SearchKeyword,SearchSelect,SearchClassID,Datatable,ClassDatatable)
  	set RsList = server.CreateObject("adodb.recordset")
	SqlList = " select * from "&Datatable&" "
	SqlList = SqlList + " where ProID > 0 "
	if isnumeric(SearchClassID) and trim(SearchClassID)<>"" then
	    SearchClassChildID=PinLuo_GetClassChildID(SearchClassID,ClassDatatable)
	    if SearchClassChildID="" then
	       SqlList = SqlList + " and Classid="&int(SearchClassID)
		else
		   SqlList = SqlList + " and Classid in("&SearchClassID&","&SearchClassChildID&") "
		end if
	end if
	if SearchKeyword <> "" then
		Select Case SearchSelect
			Case "ProName"
				SqlList = SqlList + " and ProName like '%"&SearchKeyword&"%' " 
			Case "ProContent"
				SqlList = SqlList + " and ProContent like '%"&SearchKeyword&"%' " 
			Case Else
				SqlList = SqlList + " and ProName&ProContent like '%"&SearchKeyword&"%' " 
		End Select
	End If
	    Select Case SearchSelect
		    Case "Jian"
				SqlList = SqlList + " and Jian=true " 
			Case "Hot"
				SqlList = SqlList + " and Hot=true " 
			Case "Cheap"
				SqlList = SqlList + " and Cheap=true " 	
	        Case "Shenhe1"
				SqlList = SqlList + " and Shenhe=false " 
			Case "Shenhe2"
				SqlList = SqlList + " and Shenhe=true " 
		End Select
	SqlList = SqlList + " ORDER BY OrderID desc, UpdateTime DESC , ProID DESC "
	RsList.open SqlList, objConn, 1, 1
	if RsList.eof or RsList.bof then
		response.Write("<tr class=""stline""><td colspan=""9"">没有找到任何产品记录！</td></tr>")
	else 
		
	RsList.pagesize = OnepageNum
	page = trim(request("page"))
	If IsNumeric(page) = False Then 
		page = 1 
	Else
		page = cint(page)
	End If 
	if page < 1 then page = 1
	if page > RsList.pagecount then page = RsList.pagecount
	
	RsList.absolutepage = page
	for i = 1 to RsList.pagesize 
	j = RsList.pagesize * ( page - 1) + i
	
	response.Write("<tr class=""stline"" onMouseOver=""this.className='nd'"" onMouseOut=""this.className='stline'"">"&vbcr)
	response.Write("<td><input type=""checkbox"" name=""DelProID"" value="""&RsList("ProID")&"""></td>"&vbcr)
	response.Write("<td>"&RsList("ProID")&"</td>"&vbcr)
	response.Write("<td align=left><a href=""Pinluo_Product.asp?Act=edit&ProID="&RsList("ProID")&"&classid="&SearchClassID&"&SearchKeyword="&SearchKeyword&"&SearchSelect="&SearchSelect&"&page="&page&""">"&gotTopic(RsList("ProName"),200)&"</a> ")
	if RsList("Jian")=true then response.Write("<font color=#FF6600>(荐)</font>"&vbcr)
	if RsList("Hot")=true then response.Write("<font color=red>(热)</font>"&vbcr)
	if RsList("Cheap")=true then response.Write("<font color=green>(折)</font>"&vbcr)
	response.Write("</td>"&vbcr)
	response.Write("<td>"&PinLuo_ViewClassName(RsList("Classid"),"PinLuo_ProductClass")&"</td>"&vbcr)
	response.Write("<td>"&GetCurrentDate(RsList("UpdateTime"),5)&"</td>"&vbcr)
	response.Write("<td>"&RsList("hits")&"</td>"&vbcr)
	response.Write("<td>"&ChkShenHe(RsList("shenhe"))&"</td>"&vbcr)
	response.Write("<td><input name=""ProductRank_"& RsList("ProID") &""" type=""text"" id=""ProductRank_"& RsList("ProID") &""" title="" 显示顺序，数值越大越靠前，默认为0 ""  value="""& RsList("OrderID")&""" size=""2"" maxlength=""12"" onFocus=""if(this.value=='"& RsList("OrderID")&"')this.value='';"" onBlur=""if(this.value=='')this.value='"&RsList("OrderID")&"';"" ></td>"&vbcr)
	response.Write("<td><a href=""Pinluo_Product.asp?Act=edit&ProID="&RsList("ProID")&"&classid="&SearchClassID&"&SearchKeyword="&SearchKeyword&"&SearchSelect="&SearchSelect&"&page="&page&"""><img src=""images/icon_xg.gif"" width=16 height=16 border=0></a></td>"&vbcr)
	response.Write("</tr>"&vbcr)
	
			RsList.movenext
		if RsList.eof then exit for
	Next
	Pinluo_showpage_temp=Pinluo_showpage("PinLuo_Productlist","PinLuo_ProductList.asp",page,RsList, "SearchKeyword#" & SearchKeyword & "#SearchSelect#" & SearchSelect & "#classid#" & SearchClassID)
		
		
	end if 
	RsList.close
	set RsList = nothing
End Sub

Public Function PinLuo_AddProduct(Classid,ProName,SEO_Title,SEO_Keyword,SEO_Description,ProContent,ProImg1,ProImg2,ProPrice1,ProPrice2,Saled,Jian,Hot,Cheap,UpdateTime,hits,OrderID,Shenhe,Datatable)
    if isnumeric(Classid)=false then PinLuo_AddProduct=false:exit Function
	if isdate(UpdateTime)=false then UpdateTime=now
	if isnumeric(hits)=false then hits=0
	if isnumeric(OrderID)=false then OrderID=0
	if isnumeric(Saled)=false then Saled=0
	if isnumeric(ProPrice1)=false then ProPrice1=0
	if isnumeric(ProPrice2)=false then ProPrice2=0
	set RsList = server.CreateObject("adodb.recordset")
	SqlList = "SELECT * FROM "&Datatable
	RsList.open SqlList, objConn, 1, 3
	RsList.addnew
      RsList("Classid")=Classid
	  RsList("ProName")=ProName
	  RsList("SEO_Title")=SEO_Title
	  RsList("SEO_Keyword")=SEO_Keyword
	  RsList("SEO_Description")=SEO_Description
	  RsList("ProContent")=ProContent
	  RsList("ProImg1")=ProImg1
	  RsList("ProImg2")=ProImg2
	  RsList("ProPrice1")=ProPrice1
	  RsList("ProPrice2")=ProPrice2
	  RsList("Saled")=Saled
	  RsList("Jian")=Jian
	  RsList("Hot")=Hot
	  RsList("Cheap")=Cheap
	  RsList("UpdateTime")=UpdateTime
	  RsList("hits")=hits
	  RsList("OrderID")=OrderID
	  RsList("Shenhe")=eval(Shenhe)
	RsList.update
	  ProID=RsList("ProID")
	RsList.close
	Set RsList = Nothing
	PinLuo_AddProduct=true
End Function

Public Function PinLuo_EditProduct(ProID,Classid,ProName,SEO_Title,SEO_Keyword,SEO_Description,ProContent,ProImg1,ProImg2,ProPrice1,ProPrice2,Saled,Jian,Hot,Cheap,UpdateTime,hits,OrderID,Shenhe,Datatable)
    if isnumeric(ProID)=false then PinLuo_EditProduct=false:exit Function
    if isnumeric(Classid)=false then PinLuo_EditProduct=false:exit Function
	if isdate(UpdateTime)=false then UpdateTime=now
	if isnumeric(hits)=false then hits=0
	if isnumeric(OrderID)=false then OrderID=0
	if isnumeric(Saled)=false then Saled=0
	if isnumeric(ProPrice1)=false then ProPrice1=0
	if isnumeric(ProPrice2)=false then ProPrice2=0
	set RsList = server.CreateObject("adodb.recordset")
	SqlList = "SELECT * FROM "&Datatable&" where ProID="&int(ProID)
	RsList.open SqlList, objConn, 1, 3
      RsList("Classid")=Classid
	  RsList("ProName")=ProName
	  RsList("SEO_Title")=SEO_Title
	  RsList("SEO_Keyword")=SEO_Keyword
	  RsList("SEO_Description")=SEO_Description
	  RsList("ProContent")=ProContent
	  RsList("ProImg1")=ProImg1
	  RsList("ProImg2")=ProImg2
	  RsList("ProPrice1")=ProPrice1
	  RsList("ProPrice2")=ProPrice2
	  RsList("Saled")=Saled
	  RsList("Jian")=Jian
	  RsList("Hot")=Hot
	  RsList("Cheap")=Cheap
	  RsList("UpdateTime")=UpdateTime
	  RsList("hits")=hits
	  RsList("OrderID")=OrderID
	  RsList("Shenhe")=eval(Shenhe)
	RsList.update
	  ProID=RsList("ProID")
	RsList.close
	Set RsList = Nothing
	PinLuo_EditProduct=true
End Function

Public Sub PinLuo_ViewProductItem(ProID,Datatable)
    if isnumeric(ProID)=false then exit Sub
	set RsList = server.CreateObject("adodb.recordset")
	SqlList = "SELECT * FROM "&Datatable&" where ProID="&ProID
	RsList.open SqlList, objConn, 1, 1
	if not(RsList.eof) then
	  ProID=RsList("ProID")
	  Classid=RsList("Classid")
	  ProName=RsList("ProName")
	  SEO_Title=RsList("SEO_Title")
	  SEO_Keyword=RsList("SEO_Keyword")
	  SEO_Description=RsList("SEO_Description")
	  ProContent=RsList("ProContent")
	  ProImg1=RsList("ProImg1")
	  ProImg2=RsList("ProImg2")
	  ProPrice1=RsList("ProPrice1")
	  ProPrice2=RsList("ProPrice2")
	  Saled=RsList("Saled")
	  Jian=RsList("Jian")
	  Hot=RsList("Hot")
	  Cheap=RsList("Cheap")
	  UpdateTime=RsList("UpdateTime")
	  hits=RsList("hits")
	  OrderID=RsList("OrderID")
	  Shenhe=RsList("Shenhe")
	end if
	RsList.close
	Set RsList = Nothing
End Sub

Public Sub DelProductAll(DelProID,Datatable)
	If trim(DelProID)="" Then Exit Sub
	Dim RsGetList, sqlGetList
	Set RsGetList = Server.CreateObject("adodb.recordset")
	sqlGetList = " select * from "&Datatable&" where ProID in ( "&DelProID&" ) "
	RsGetList.open sqlGetList, objConn, 1, 2
	call PinLuo_ViewSiteConfig() '取得网站配置
	While Not( RsGetList.bof or RsGetList.eof )
	if Pinluo_DelProImg=true then '设置删除产品图片
		Dim aSaveFileName,aSaveFileName2
		Dim sSaveFileName,sSaveFileName2
		sSaveFileName = RsGetList("ProImg1")
		sSaveFileName2 = RsGetList("ProImg2")
		If sSaveFileName <> "" Then 
			aSaveFileName = Split(sSaveFileName, "|")
			Dim i
			For i = 0 To UBound(aSaveFileName)
				if isnull( aSaveFileName(i) ) = false or aSaveFileName(i) <> "" then 
					DelFileUploadfile "",aSaveFileName(i)
				end if 
			Next
		End If
		If sSaveFileName2 <> "" Then 
			aSaveFileName2 = Split(sSaveFileName2, "|")
			Dim j
			For j = 0 To UBound(aSaveFileName2)
				if isnull( aSaveFileName2(j) ) = false or aSaveFileName2(j) <> "" then 
					DelFileUploadfile "",aSaveFileName2(j)
				end if 
			Next
		End If
	end if
		RsGetList.delete
		RsGetList.movenext
	wend
	RsGetList.Close
	Set RsGetList = Nothing
End Sub

Public Sub PinLuo_AdminList(OnepageNum,SearchKeyword,SearchSelect,Datatable)
  	set RsList = server.CreateObject("adodb.recordset")
	SqlList = " select * from "&Datatable&" "
	SqlList = SqlList + " where UserID > 0 "
	if SearchKeyword <> "" then
		Select Case SearchSelect
			Case "UserName"
				SqlList = SqlList + " and UserName like '%"&SearchKeyword&"%' " 
			Case "RealName"
				SqlList = SqlList + " and RealName like '%"&SearchKeyword&"%' " 
			Case "Mobile"
				SqlList = SqlList + " and Mobile like '%"&SearchKeyword&"%' " 
			Case "Email"
				SqlList = SqlList + " and Email like '%"&SearchKeyword&"%' " 	
			Case Else
				SqlList = SqlList + " and UserName&RealName&Mobile&Email like '%"&SearchKeyword&"%' " 
		End Select
	End If
	    Select Case SearchSelect
		    Case "UserPassed1"
				SqlList = SqlList + " and UserPassed=true " 
			Case "UserPassed2"
				SqlList = SqlList + " and UserPassed=false " 
		End Select
	SqlList = SqlList + " ORDER BY UserID asc "
	RsList.open SqlList, objConn, 1, 1
	if RsList.eof or RsList.bof then
		response.Write("<tr class=""stline""><td colspan=""9"">没有找到任何记录！</td></tr>")
	else 
		
	RsList.pagesize = OnepageNum
	page = trim(request("page"))
	If IsNumeric(page) = False Then 
		page = 1 
	Else
		page = cint(page)
	End If 
	if page < 1 then page = 1
	if page > RsList.pagecount then page = RsList.pagecount
	
	RsList.absolutepage = page
	for i = 1 to RsList.pagesize 
	j = RsList.pagesize * ( page - 1) + i
	
	response.Write("<tr class=""stline"" onMouseOver=""this.className='nd'"" onMouseOut=""this.className='stline'"">"&vbcr)
	response.Write("<td><input type=""checkbox"" name=""DelUserID"" value="""&RsList("UserID")&"""></td>"&vbcr)
	response.Write("<td>"&RsList("UserID")&"</td>"&vbcr)
	response.Write("<td align=left><a href=""Pinluo_Admin.asp?Act=edit&UserID="&RsList("UserID")&"&SearchKeyword="&SearchKeyword&"&SearchSelect="&SearchSelect&"&page="&page&""">"&RsList("UserName")&"")
	response.Write("</a></td>"&vbcr)
	response.Write("<td>"&RsList("RealName")&"</td>"&vbcr)
	response.Write("<td><a href=""http://www.pinluo.com/tool/ip/?domain="&trim(RsList("LastLoginIP"))&""" target=_blank>"&RsList("LastLoginIP")&"</td>"&vbcr)
	response.Write("<td>"&GetCurrentDate(RsList("LastLoginTime"),5)&"</td>"&vbcr)
	response.Write("<td>"&RsList("LoginTimes")&"</td>"&vbcr)
	response.Write("<td>")
	if RsList("UserPassed")=true then response.Write("<font color=green><b>√</b></font>"&vbcr) else response.Write("<font color=red><b>×</b></font>"&vbcr)
	response.Write("</td>"&vbcr)
	response.Write("<td><a href=""Pinluo_Admin.asp?Act=edit&UserID="&RsList("UserID")&"&SearchKeyword="&SearchKeyword&"&SearchSelect="&SearchSelect&"&page="&page&"""><img src=""images/icon_xg.gif"" width=16 height=16 border=0></a></td>"&vbcr)
	response.Write("</tr>"&vbcr)
	
			RsList.movenext
		if RsList.eof then exit for
	Next
	Pinluo_showpage_temp=Pinluo_showpage("PinLuo_Adminlist","PinLuo_Adminlist.asp",page,RsList, "SearchKeyword#" & SearchKeyword & "#SearchSelect#" & SearchSelect)
		
		
	end if 
	RsList.close
	set RsList = nothing
End Sub

Public Function PinLuo_AddAdmin(UserName,UserPassword,UserPassed,RealName,Mobile,Email,UserPopedom)
    if trim(UserName)="" then PinLuo_AddAdmin=false:exit Function
	if trim(UserPassword)="" then PinLuo_AddAdmin=false:exit Function
	set RsList = server.CreateObject("adodb.recordset")
	SqlList = "SELECT * FROM PinLuo_Admin where UserName = '"&UserName&"' "
	RsList.open SqlList, objConn, 1, 3
	If not( RsList.bof or RsList.eof ) Then
		ErrMsg = "帐号："&UserName&"已经存在！请您重新输入！"
		PinLuo_AddAdmin=false
	Else
	RsList.addnew
      RsList("UserName")=UserName
	  RsList("UserPassword")=md5(UserPassword,32)
	  RsList("UserPassed")=UserPassed
	  RsList("RealName")=RealName
	  RsList("Mobile")=Mobile
	  RsList("Email")=Email
	  RsList("LastLoginIP") = "127.0.0.1"
	  RsList("LastLoginTime") = Now()
	  RsList("LastLogoutTime") = Now()
	  RsList("LoginTimes") = 1
	  RsList("UserPopedom") = UserPopedom
	  RsList.update
	  UserID=RsList("UserID")
	  PinLuo_AddAdmin=true
	End If
	RsList.close
	Set RsList = Nothing
End Function

Public Function PinLuo_EditAdmin(UserID,UserName,UserPassword,UserPassed,RealName,Mobile,Email,UserPopedom)
    if trim(UserID)="" then PinLuo_EditAdmin=false:exit Function
	set RsList = server.CreateObject("adodb.recordset")
	SqlList = "SELECT * FROM PinLuo_Admin where UserID = "&int(UserID)
	RsList.open SqlList, objConn, 1, 3
	If not( RsList.eof ) Then
      RsList("UserName")=UserName
	  if trim(UserPassword)<>"" then
	  RsList("UserPassword")=md5(UserPassword,32)
	  end if
	  RsList("UserPassed")=UserPassed
	  RsList("RealName")=RealName
	  RsList("Mobile")=Mobile
	  RsList("Email")=Email
	  if trim(UserPopedom)<>"" then
	  RsList("UserPopedom")=UserPopedom
	  end if
	  RsList.update
	  UserID=RsList("UserID")
	  PinLuo_EditAdmin=true
	Else
	  ErrMsg = "更新失败，帐号ID："&UserID&"不存在！请检查！"
	  PinLuo_EditAdmin=false
	End If
	RsList.close
	Set RsList = Nothing
End Function

Public Sub PinLuo_ViewAdminItem(UserID)
    if isnumeric(UserID)=false or trim(UserID)="" then exit Sub
	set RsList = server.CreateObject("adodb.recordset")
	SqlList = "SELECT * FROM PinLuo_Admin where UserID="&UserID
	RsList.open SqlList, objConn, 1, 1
	if not(RsList.eof) then
	  UserName=RsList("UserName")
	  UserPassword=RsList("UserPassword")
	  UserPassed=RsList("UserPassed")
	  RealName=RsList("RealName")
	  Mobile=RsList("Mobile")
	  Email=RsList("Email")
	  LastLoginIP=RsList("LastLoginIP")
	  LastLoginTime=RsList("LastLoginTime")
	  LastLogoutTime=RsList("LastLogoutTime")
	  LoginTimes=RsList("LoginTimes")
	  UserPopedom=RsList("UserPopedom")
	  UserWarning=RsList("UserWarning")
	end if
	RsList.close
	Set RsList = Nothing
End Sub

Public Sub DelAdminAll(DelUserID)
	If trim(DelUserID)="" Then Exit Sub
	Dim RsGetList, sqlGetList
	Set RsGetList = Server.CreateObject("adodb.recordset")
	sqlGetList = " select * from PinLuo_Admin where UserID in ( "&DelUserID&" ) "
	RsGetList.open sqlGetList, objConn, 1, 2
	While Not( RsGetList.bof or RsGetList.eof )
		RsGetList.delete
		RsGetList.movenext
	wend
	RsGetList.Close
	Set RsGetList = Nothing
End Sub

Public Sub PinLuo_BlockList(OnepageNum,SearchKeyword,SearchSelect,Datatable)
  	set RsList = server.CreateObject("adodb.recordset")
	SqlList = " select * from "&Datatable&" "
	SqlList = SqlList + " where Block_ID > 0 "
	if SearchKeyword <> "" then
		Select Case SearchSelect
			Case "Block_Title"
				SqlList = SqlList + " and Block_Title like '%"&SearchKeyword&"%' " 
			Case "Block_Content"
				SqlList = SqlList + " and Block_Content like '%"&SearchKeyword&"%' " 
			Case Else
				SqlList = SqlList + " and Block_Title&Block_Content like '%"&SearchKeyword&"%' " 
		End Select
	End If

	SqlList = SqlList + " ORDER BY Block_Time asc,Block_ID asc "
	RsList.open SqlList, objConn, 1, 1
	if RsList.eof or RsList.bof then
		response.Write("<tr class=""stline""><td colspan=""5"">没有找到任何记录！</td></tr>")
	else 
		
	RsList.pagesize = OnepageNum
	page = trim(request("page"))
	If IsNumeric(page) = False Then 
		page = 1 
	Else
		page = cint(page)
	End If 
	if page < 1 then page = 1
	if page > RsList.pagecount then page = RsList.pagecount
	
	RsList.absolutepage = page
	for i = 1 to RsList.pagesize 
	j = RsList.pagesize * ( page - 1) + i
	
	response.Write("<tr class=""stline"" onMouseOver=""this.className='nd'"" onMouseOut=""this.className='stline'"">"&vbcr)
	response.Write("<td><input type=""checkbox"" name=""DelBlockID"" value="""&RsList("Block_ID")&"""></td>"&vbcr)
	response.Write("<td>"&RsList("Block_ID")&"</td>"&vbcr)
	response.Write("<td align=left><a href=""Pinluo_Block.asp?Act=edit&Block_ID="&RsList("Block_ID")&"&SearchKeyword="&SearchKeyword&"&SearchSelect="&SearchSelect&"&page="&page&""">"&RsList("Block_Title")&"")
	response.Write("</a></td>"&vbcr)
	response.Write("<td>"&RsList("Block_Time")&"</td>"&vbcr)
	response.Write("<td><a href=""Pinluo_Block.asp?Act=edit&Block_ID="&RsList("Block_ID")&"&SearchKeyword="&SearchKeyword&"&SearchSelect="&SearchSelect&"&page="&page&"""><img src=""images/icon_xg.gif"" width=16 height=16 border=0></a></td>"&vbcr)
	response.Write("</tr>"&vbcr)
	
			RsList.movenext
		if RsList.eof then exit for
	Next
	Pinluo_showpage_temp=Pinluo_showpage("PinLuo_Blocklist","PinLuo_Blocklist.asp",page,RsList, "SearchKeyword#" & SearchKeyword & "#SearchSelect#" & SearchSelect)
		
		
	end if 
	RsList.close
	set RsList = nothing
End Sub

Public Function PinLuo_AddBlock(Block_Title,Block_Content,Block_Time)
    if trim(Block_Title)="" then PinLuo_AddBlock=false:exit Function
	if isdate(Block_Time)=false then Block_Time=now
	set RsList = server.CreateObject("adodb.recordset")
	SqlList = "SELECT * FROM PinLuo_Block"
	RsList.open SqlList, objConn, 1, 3

	  RsList.addnew
      RsList("Block_Title")=Block_Title
	  RsList("Block_Content")=Block_Content
	  RsList("Block_Time")=Block_Time
	  RsList.update
	  Block_ID=RsList("Block_ID")
	  PinLuo_AddBlock=true

	RsList.close
	Set RsList = Nothing
End Function

Public Function PinLuo_EditBlock(Block_ID,Block_Title,Block_Content,Block_Time)
    if trim(Block_ID)="" then PinLuo_EditBlock=false:exit Function	
	set RsList = server.CreateObject("adodb.recordset")
	SqlList = "SELECT * FROM PinLuo_Block where Block_ID = "&int(Block_ID)
	RsList.open SqlList, objConn, 1, 3
	If not( RsList.eof ) Then
      RsList("Block_Title")=Block_Title
	  RsList("Block_Content")=Block_Content
	  if isdate(Block_Time)=true then
	  RsList("Block_Time")=Block_Time
	  end if
	  RsList.update
	  Block_ID=RsList("Block_ID")
	  PinLuo_EditBlock=true
	Else
	  ErrMsg = "更新失败，ID："&Block_ID&"不存在！请检查！"
	  PinLuo_EditBlock=false
	End If
	RsList.close
	Set RsList = Nothing
End Function

Public Sub PinLuo_ViewBlockItem(Block_ID)
    if isnumeric(Block_ID)=false or trim(Block_ID)="" then exit Sub
	set RsList = server.CreateObject("adodb.recordset")
	SqlList = "SELECT * FROM PinLuo_Block where Block_ID="&Block_ID
	RsList.open SqlList, objConn, 1, 1
	if not(RsList.eof) then
	  Block_Title=RsList("Block_Title")
	  Block_Content=RsList("Block_Content")
	  Block_Time=RsList("Block_Time")
	end if
	RsList.close
	Set RsList = Nothing
End Sub

Public Sub DelBlockAll(DelBlock_ID)
	If trim(DelBlock_ID)="" Then Exit Sub
	Dim RsGetList, sqlGetList
	Set RsGetList = Server.CreateObject("adodb.recordset")
	sqlGetList = " select * from PinLuo_Block where Block_ID in ( "&DelBlock_ID&" ) "
	RsGetList.open sqlGetList, objConn, 1, 2
	While Not( RsGetList.bof or RsGetList.eof )
		RsGetList.delete
		RsGetList.movenext
	wend
	RsGetList.Close
	Set RsGetList = Nothing
End Sub

Public Sub PinLuo_DaohangList(OnepageNum,SearchKeyword,SearchSelect,Datatable)
  	set RsList = server.CreateObject("adodb.recordset")
	SqlList = " select * from "&Datatable&" "
	SqlList = SqlList + " where Daohang_ID > 0 "
	if SearchKeyword <> "" then
		Select Case SearchSelect
			Case "Daohang_Title"
				SqlList = SqlList + " and Daohang_Title like '%"&SearchKeyword&"%' " 
			Case "Daohang_Url"
				SqlList = SqlList + " and Daohang_Url like '%"&SearchKeyword&"%' " 
			Case Else
				SqlList = SqlList + " and Daohang_Title&Daohang_Url like '%"&SearchKeyword&"%' " 
		End Select
	End If

	SqlList = SqlList + " ORDER BY Daohang_order Desc "
	RsList.open SqlList, objConn, 1, 1
	if RsList.eof or RsList.bof then
		response.Write("<tr class=""stline""><td colspan=""5"">没有找到任何记录！</td></tr>")
	else 
		
	RsList.pagesize = OnepageNum
	page = trim(request("page"))
	If IsNumeric(page) = False Then 
		page = 1 
	Else
		page = cint(page)
	End If 
	if page < 1 then page = 1
	if page > RsList.pagecount then page = RsList.pagecount
	
	RsList.absolutepage = page
	for i = 1 to RsList.pagesize 
	j = RsList.pagesize * ( page - 1) + i
	
	response.Write("<tr class=""stline"" onMouseOver=""this.className='nd'"" onMouseOut=""this.className='stline'"">"&vbcr)
	response.Write("<td><input type=""checkbox"" name=""DelDaohangID"" value="""&RsList("Daohang_ID")&"""></td>"&vbcr)
	response.Write("<td>"&RsList("Daohang_ID")&"</td>"&vbcr)
	response.Write("<td align=left><a href=""Pinluo_Daohang.asp?Act=edit&Daohang_ID="&RsList("Daohang_ID")&"&SearchKeyword="&SearchKeyword&"&SearchSelect="&SearchSelect&"&page="&page&""">"&RsList("Daohang_Title")&"")
	response.Write("</a></td>"&vbcr)
	response.Write("<td align=left>"&RsList("Daohang_Url")&"</td>"&vbcr)
	response.Write("<td align=left>"&RsList("Daohang_order")&"</td>"&vbcr)
	response.Write("<td><a href=""Pinluo_Daohang.asp?Act=edit&Daohang_ID="&RsList("Daohang_ID")&"&SearchKeyword="&SearchKeyword&"&SearchSelect="&SearchSelect&"&page="&page&"""><img src=""images/icon_xg.gif"" width=16 height=16 border=0></a></td>"&vbcr)
	response.Write("</tr>"&vbcr)
	
			RsList.movenext
		if RsList.eof then exit for
	Next
	Pinluo_showpage_temp=Pinluo_showpage("PinLuo_Daohanglist","PinLuo_Daohanglist.asp",page,RsList, "SearchKeyword#" & SearchKeyword & "#SearchSelect#" & SearchSelect)
		
		
	end if 
	RsList.close
	set RsList = nothing
End Sub

Public Function PinLuo_AddDaohang(Daohang_Title,Daohang_Url,Daohang_Blank,Daohang_order)
    if trim(Daohang_Title)="" then PinLuo_AddDaohang=false:exit Function
	set RsList = server.CreateObject("adodb.recordset")
	SqlList = "SELECT * FROM PinLuo_Daohang"
	RsList.open SqlList, objConn, 1, 3

	  RsList.addnew
      RsList("Daohang_Title")=Daohang_Title
	  RsList("Daohang_Url")=Daohang_Url
	  RsList("Daohang_Blank")=Daohang_Blank
	  RsList("Daohang_order")=Daohang_order
	  RsList.update
	  Daohang_ID=RsList("Daohang_ID")
	  PinLuo_AddDaohang=true

	RsList.close
	Set RsList = Nothing
End Function

Public Function PinLuo_EditDaohang(Daohang_ID,Daohang_Title,Daohang_Url,Daohang_Blank,Daohang_order)
    if trim(Daohang_ID)="" then PinLuo_EditDaohang=false:exit Function	
	set RsList = server.CreateObject("adodb.recordset")
	SqlList = "SELECT * FROM PinLuo_Daohang where Daohang_ID = "&int(Daohang_ID)
	RsList.open SqlList, objConn, 1, 3
	If not( RsList.eof ) Then
      RsList("Daohang_Title")=Daohang_Title
	  RsList("Daohang_Url")=Daohang_Url
	  RsList("Daohang_Blank")=Daohang_Blank
	  RsList("Daohang_order")=Daohang_order
	  RsList.update
	  Daohang_ID=RsList("Daohang_ID")
	  PinLuo_EditDaohang=true
	Else
	  ErrMsg = "更新失败，ID："&Daohang_ID&"不存在！请检查！"
	  PinLuo_EditDaohang=false
	End If
	RsList.close
	Set RsList = Nothing
End Function

Public Sub PinLuo_ViewDaohangItem(Daohang_ID)
    if isnumeric(Daohang_ID)=false or trim(Daohang_ID)="" then exit Sub
	set RsList = server.CreateObject("adodb.recordset")
	SqlList = "SELECT * FROM PinLuo_Daohang where Daohang_ID="&Daohang_ID
	RsList.open SqlList, objConn, 1, 1
	if not(RsList.eof) then
	  Daohang_Title=RsList("Daohang_Title")
	  Daohang_Url=RsList("Daohang_Url")
	  Daohang_Blank=RsList("Daohang_Blank")
	  Daohang_order=RsList("Daohang_order")
	end if
	RsList.close
	Set RsList = Nothing
End Sub

Public Sub DelDaohangAll(DelDaohang_ID)
	If trim(DelDaohang_ID)="" Then Exit Sub
	Dim RsGetList, sqlGetList
	Set RsGetList = Server.CreateObject("adodb.recordset")
	sqlGetList = " select * from PinLuo_Daohang where Daohang_ID in ( "&DelDaohang_ID&" ) "
	RsGetList.open sqlGetList, objConn, 1, 2
	While Not( RsGetList.bof or RsGetList.eof )
		RsGetList.delete
		RsGetList.movenext
	wend
	RsGetList.Close
	Set RsGetList = Nothing
End Sub

Public Sub PinLuo_DeleteItemClass(ClassID,Datatable)
    if isnumeric(ClassID)=false then exit Sub
    set RsList = server.CreateObject("adodb.recordset")
	SqlList = "SELECT * FROM "&Datatable&" where ClassID="&int(ClassID)
	RsList.open SqlList, objConn, 1, 3
	if not RsList.eof then
	  ChildID=trim(RsList("ChildID"))
	  if ChildID<>"" then 
	    objconn.execute("delete from "&Datatable&" where classid in("&ChildID&")")'删除下属分类
	  end if
	  ParentID=int(RsList("ParentID"))
	end if
	RsList.delete
	RsList.close
	Set RsList = Nothing
	'更新上级栏目孩子数
	ChildID=ClassID&","&ChildID
	PinLuo_UpdateDeleteInfoClassChild ChildID,ParentID,Datatable
End Sub

Public Sub PinLuo_FeedbackListView(OnepageNum,SearchKeyword,SearchSelect,SearchClassID,Datatable,ClassDatatable)
  	set RsList = server.CreateObject("adodb.recordset")
	SqlList = " select * from "&Datatable&" "
	SqlList = SqlList + " where FeedbackID > 0 "
	if isnumeric(SearchClassID) and trim(SearchClassID)<>"" then
	    SearchClassChildID=PinLuo_GetClassChildID(SearchClassID,ClassDatatable)
	    if SearchClassChildID="" then
	       SqlList = SqlList + " and Classid="&int(SearchClassID)
		else
		   SqlList = SqlList + " and Classid in("&SearchClassID&","&SearchClassChildID&") "
		end if
	end if
	if SearchKeyword <> "" then
		Select Case SearchSelect
			Case "FeedbackTitle"
				SqlList = SqlList + " and FeedbackTitle like '%"&SearchKeyword&"%' " 
			Case "FeedbackContent"
				SqlList = SqlList + " and FeedbackContent like '%"&SearchKeyword&"%' " 
			Case "Author"
				SqlList = SqlList + " and Author like '%"&SearchKeyword&"%' " 
			Case Else
				SqlList = SqlList + " and FeedbackTitle&FeedbackContent&Author like '%"&SearchKeyword&"%' " 
		End Select
	End If
	    Select Case SearchSelect
	        Case "Shenhe1"
				SqlList = SqlList + " and Shenhe=false " 
			Case "Shenhe2"
				SqlList = SqlList + " and Shenhe=true " 
		End Select
	SqlList = SqlList + " ORDER BY UpdateTime DESC , FeedbackID DESC "
	RsList.open SqlList, objConn, 1, 1
	if RsList.eof or RsList.bof then
		response.Write("<tr class=""stline""><td colspan=""9"">没有找到任何信息记录！</td></tr>")
	else 
		
	RsList.pagesize = OnepageNum
	page = trim(request("page"))
	If IsNumeric(page) = False Then 
		page = 1 
	Else
		page = cint(page)
	End If 
	if page < 1 then page = 1
	if page > RsList.pagecount then page = RsList.pagecount
	
	RsList.absolutepage = page
	for i = 1 to RsList.pagesize 
	j = RsList.pagesize * ( page - 1) + i
	
	response.Write("<tr class=""stline"" onMouseOver=""this.className='nd'"" onMouseOut=""this.className='stline'"">"&vbcr)
	response.Write("<td><input type=""checkbox"" name=""DelFeedbackID"" value="""&RsList("FeedbackID")&"""></td>"&vbcr)
	response.Write("<td>"&RsList("FeedbackID")&"</td>"&vbcr)
	response.Write("<td align=left><a href=""Pinluo_Feedback.asp?Act=edit&Feedbackid="&RsList("FeedbackID")&"&classid="&SearchClassID&"&SearchKeyword="&SearchKeyword&"&SearchSelect="&SearchSelect&"&page="&page&""">"&gotTopic(RsList("FeedbackTitle"),200))
	response.Write("</a></td>"&vbcr)
	response.Write("<td>"&PinLuo_ViewClassName(RsList("Classid"),"PinLuo_FeedbackClass")&"</td>"&vbcr)
	response.Write("<td>"&GetCurrentDate(RsList("UpdateTime"),5)&"</td>"&vbcr)
	response.Write("<td>"&RsList("Author")&"</td>"&vbcr)
	response.Write("<td>"&ChkShenHe(RsList("shenhe"))&"</td>"&vbcr)
	response.Write("<td><a href=""Pinluo_Feedback.asp?Act=edit&Feedbackid="&RsList("FeedbackID")&"&classid="&SearchClassID&"&SearchKeyword="&SearchKeyword&"&SearchSelect="&SearchSelect&"&page="&page&"""><img src=""images/icon_xg.gif"" width=16 height=16 border=0></a></td>"&vbcr)
	response.Write("</tr>"&vbcr)
	
			RsList.movenext
		if RsList.eof then exit for
	Next
	Pinluo_showpage_temp=Pinluo_showpage("PinLuo_Feedbacklist","PinLuo_FeedbackList.asp",page,RsList, "SearchKeyword#" & SearchKeyword & "#SearchSelect#" & SearchSelect & "#classid#" & SearchClassID)
		
		
	end if 
	RsList.close
	set RsList = nothing
End Sub

Public Function PinLuo_AddFeedback(Classid,FeedbackTitle,FeedbackContent,Author,UpdateTime,Shenhe,Phone,Email,QQ,ReplyContent,ReplyTime,Hits,OrderID,Datatable)
    if isnumeric(Classid)=false then PinLuo_AddFeedback=false:exit Function
	if trim(FeedbackTitle)="" then PinLuo_AddFeedback=false:exit Function
	if trim(FeedbackContent)="" then PinLuo_AddFeedback=false:exit Function
	if isdate(UpdateTime)=false then UpdateTime=now
	if isnumeric(OrderID)=false then OrderID=0
	if isnumeric(Hits)=false then Hits=0
	if isdate(ReplyTime)=false then ReplyTime=empty
	set RsList = server.CreateObject("adodb.recordset")
	SqlList = "SELECT * FROM "&Datatable
	RsList.open SqlList, objConn, 1, 3
	RsList.addnew
      RsList("Classid")=Classid
	  RsList("FeedbackTitle")=FeedbackTitle
	  RsList("FeedbackContent")=FeedbackContent
	  RsList("Author")=Author
	  RsList("UpdateTime")=UpdateTime
	  RsList("Shenhe")=Shenhe
	  RsList("Phone") = Phone
	  RsList("Email") = Email
	  RsList("QQ") = QQ
	  RsList("ReplyContent") = ReplyContent
	  RsList("ReplyTime") = ReplyTime
	  RsList("Hits") = Hits
	  RsList("OrderID") = OrderID
	  RsList.update
	  FeedbackID=RsList("FeedbackID")
	  PinLuo_AddFeedback=true
	RsList.close
	Set RsList = Nothing
End Function

Public Function PinLuo_EditFeedback(Feedbackid,Classid,FeedbackTitle,FeedbackContent,Author,UpdateTime,Shenhe,Phone,Email,QQ,ReplyContent,ReplyTime,Hits,OrderID,Datatable)
    if isnumeric(Feedbackid)=false then PinLuo_EditFeedback=false:exit Function
    if isnumeric(Classid)=false then PinLuo_EditFeedback=false:exit Function
	if trim(FeedbackTitle)="" then PinLuo_EditFeedback=false:exit Function
	if trim(FeedbackContent)="" then PinLuo_EditFeedback=false:exit Function
	if isdate(UpdateTime)=false then UpdateTime=now
	if isnumeric(OrderID)=false then OrderID=0
	if isnumeric(Hits)=false then Hits=0
	if isdate(ReplyTime)=false then ReplyTime=empty
	set RsList = server.CreateObject("adodb.recordset")
	SqlList = "SELECT * FROM "&Datatable&" where Feedbackid="&int(Feedbackid)
	RsList.open SqlList, objConn, 1, 3
	if not RsList.eof then
      RsList("Classid")=Classid
	  RsList("FeedbackTitle")=FeedbackTitle
	  RsList("FeedbackContent")=FeedbackContent
	  RsList("Author")=Author
	  RsList("UpdateTime")=UpdateTime
	  RsList("Shenhe")=Shenhe
	  RsList("Phone") = Phone
	  RsList("Email") = Email
	  RsList("QQ") = QQ
	  RsList("ReplyContent") = ReplyContent
	  RsList("ReplyTime") = ReplyTime
	  RsList("Hits") = Hits
	  RsList("OrderID") = OrderID
	  RsList.update
	  FeedbackID=RsList("FeedbackID")
	  PinLuo_EditFeedback=true
	else
	  PinLuo_EditFeedback=false
	end if
	RsList.close
	Set RsList = Nothing
End Function

Public Sub DelFeedbackAll(DelFeedbackID,Datatable)
	If trim(DelFeedbackID)="" Then Exit Sub
	Dim RsGetList, sqlGetList
	Set RsGetList = Server.CreateObject("adodb.recordset")
	sqlGetList = " select * from "&Datatable&" where FeedbackID in ( "&DelFeedbackID&" ) "
	RsGetList.open sqlGetList, objConn, 1, 2
	While Not( RsGetList.bof or RsGetList.eof )
		RsGetList.delete
		RsGetList.movenext
	wend
	RsGetList.Close
	Set RsGetList = Nothing
End Sub

Public Sub PinLuo_ViewFeedbackItem(FeedbackID,Datatable)
    if isnumeric(FeedbackID)=false then exit Sub
	set RsList = server.CreateObject("adodb.recordset")
	SqlList = "SELECT * FROM "&Datatable&" where FeedbackID="&FeedbackID
	RsList.open SqlList, objConn, 1, 1
	if not(RsList.eof) then
	  FeedbackID=RsList("FeedbackID")
	  Classid=RsList("Classid")
	  FeedbackTitle=RsList("FeedbackTitle")
	  FeedbackContent=RsList("FeedbackContent")
	  Author=RsList("Author")
	  UpdateTime=RsList("UpdateTime")
	  Shenhe=RsList("Shenhe")
	  Phone=RsList("Phone")
	  Email=RsList("Email")
	  QQ=RsList("QQ")
	  ReplyContent=RsList("ReplyContent")
	  ReplyTime=RsList("ReplyTime")
	  Hits=RsList("Hits")
	  OrderID=RsList("OrderID")
	end if
	RsList.close
	Set RsList = Nothing
End Sub

Public Sub PinLuo_WriteMsg(ErrMsg,ComeUrl)
  On Error Resume Next
  If IsObject(PinLuo) Then Set PinLuo = Nothing
  if trim(comeurl)="" then comeurl="#"
  Response.Cookies("pinluo")("PinLuo_WriteMsg_ErrMsg")=ErrMsg
  Response.Cookies("pinluo")("PinLuo_WriteMsg_comeurl")=ComeUrl
  response.Redirect("PinLuo_ShowMsg.asp")
  response.End()
End Sub


'/////////////////////////////////////////////////////////////////////////////////////////
'	过程名：CheckComeUrl
'	作  用：禁止直接输入地址访问后台
'	参  数：无
'/////////////////////////////////////////////////////////////////////////////////////////
Public Sub CheckComeUrl()

	Dim strComeUrl,strAdminUrl
	strComeUrl = trim(request.ServerVariables("HTTP_REFERER"))

	If strComeUrl = "" Then
		FoundErr = True
		ErrMsg = ErrMsg & "<br><li>抱歉，为了系统安全，不允许直接输入地址访问本系统的后台管理页面。</li>"
	Else
		strAdminUrl = trim("http://" & Request.ServerVariables("SERVER_NAME"))
		
		If ubound(split(strComeUrl,":"))>1 Then
			strAdminUrl=strAdminUrl & ":" & Request.ServerVariables("SERVER_PORT")
		End If
		
		strAdminUrl=strAdminUrl & request.ServerVariables("SCRIPT_NAME")
	
		If lcase(left(strComeUrl,instrrev(strComeUrl,"pinluo/"))) <> lcase(left(strAdminUrl,instrrev(strAdminUrl,"pinluo/"))) Then
		FoundErr = True
		ErrMsg = ErrMsg & "<br><li>抱歉，为了系统安全，不允许直接输入地址访问本系统的后台管理页面。</li>"
		End If
	End If
	
End Sub

Public Sub CheckPurview()
        ComeUrl="PL_login.asp"
        Call DBConnBegin()
        Call CheckComeUrl()
	If FoundErr = True Then
		call DBConnEnd()
        PinLuo_WriteMsg ErrMsg,ComeUrl
   		Response.End()
	End If
	Call CheckSessionData()
	If FoundErr = True Then
		call DBConnEnd()
        PinLuo_WriteMsg ErrMsg,ComeUrl
   		Response.End()
	End If
End Sub

Public Sub ValidateLoginData()

	Dim strUserName
	strUserName = Trim(Request.Form("UserName"))
	strUserName = Replace( strUserName, "'", "" )
	Dim strUserPassword
	strUserPassword = Trim(Request.Form("UserPassword"))
	strUserPassword = Replace( strUserPassword, "'", "" )
	Dim strCheckCode
	strCheckCode = Trim(Request.Form("CheckCode"))
	strCheckCode = Replace( strCheckCode, "'", "" )
	
	'检查提交数据的有效性
	If strUserName = "" Then
		FoundErr = True
		ErrMsg = ErrMsg & "<br>用户帐号不能为空！"
		Exit Sub
	End If
	If strUserPassword = "" Then
		FoundErr = True
		ErrMsg = ErrMsg & "<br>用户密码不能为空！"
		Exit Sub
	End If
	If strCheckCode = "" Then
		FoundErr = True
		ErrMsg = ErrMsg & "<br>验证码不能为空！"
		Exit Sub
	End If
	If Trim(Request.Cookies("pinluo")("CheckCode")) = "" Then
		FoundErr = True
		ErrMsg = ErrMsg & "<br>您停留在登录页面的时间过长，请重新返回登录页面进行登录。"
		Exit Sub
	End If
	if strCheckCode <> trim(Request.Cookies("pinluo")("CheckCode")) then
		FoundErr = True
		ErrMsg = ErrMsg & "<br>您输入的确认码和系统产生的不一致，请重新输入。"
		Exit Sub
	end if

	'检查数据的正确性
	strUserPassword = MD5( strUserPassword,32 )
	Dim RsGetAdmin, sqlGetAdmin
	Set RsGetAdmin = Server.CreateObject("adodb.recordset")
	sqlGetAdmin = " select * from PinLuo_Admin where UserPassword = '"&strUserPassword&"' and UserName = '"&strUserName&"' "
	RsGetAdmin.open sqlGetAdmin, objConn, 1, 3
	If RsGetAdmin.bof and RsGetAdmin.eof Then
		FoundErr = True
		ErrMsg = ErrMsg & "用户帐号或密码错误！请您重新输入！<br>"
	Else
		If strUserPassword <> RsGetAdmin("UserPassword") Then
			FoundErr = True
			ErrMsg=ErrMsg & "用户帐号或密码错误！请您重新输入！<br>"
		End If
		If RsGetAdmin("UserPassed")=false Then
			FoundErr = True
			ErrMsg=ErrMsg & "您的帐号"&strUserName&"已被停用，如有疑问请联系管理员。<br>"
		End If
	End If
	If FoundErr = True Then
	    Response.Cookies("pinluo")("UserID")=""
		Response.Cookies("pinluo")("UserName")=""
		Response.Cookies("pinluo")("RealName")=""
		Response.Cookies("pinluo")("LoginTimes")=""
		Response.Cookies("pinluo")("UserPassword")=""
		Response.Cookies("pinluo")("RndPassword")=""
		Response.Cookies("pinluo")("UserPopedom")=""
		RsGetAdmin.close
		Set RsGetAdmin = Nothing
		Exit Sub
	End If
	
	Dim RndPassword
	RndPassword = GetRndPassword(32)

	RsGetAdmin("LastLoginIP") = getIP() 'Request.ServerVariables("REMOTE_ADDR")
	RsGetAdmin("LastLoginTime") = now()
	RsGetAdmin("LoginTimes") = RsGetAdmin("LoginTimes")+1
	RsGetAdmin("RndPassword") = RndPassword
	if Len(Trim(UserPopedomCheck)) <> Len(Trim(RsGetAdmin("UserPopedom"))) or IsNull(RsGetAdmin("UserPopedom")) then 
	RsGetAdmin("UserPopedom") = UserPopedomCheck
	end if 

	RsGetAdmin.update
	Response.Cookies("pinluo")("UserID") = RsGetAdmin("UserID")
	Response.Cookies("pinluo")("UserName") = RsGetAdmin("UserName")
	Response.Cookies("pinluo")("RealName") = RsGetAdmin("RealName")
	Response.Cookies("pinluo")("LoginTimes") = RsGetAdmin("LoginTimes")
	Response.Cookies("pinluo")("UserPassword") = RsGetAdmin("UserPassword")
	Response.Cookies("pinluo")("UserPopedom") = RsGetAdmin("UserPopedom")
	Response.Cookies("pinluo")("RndPassword") = RndPassword
	
	RsGetAdmin.Close
	Set RsGetAdmin = Nothing
	Call DBConnEnd()
	Set PinLuo = Nothing
	Response.Redirect "Index.asp"
	
End Sub

Public Sub PinLuo_LogOut()
  UserID=Request.Cookies("pinluo")("UserID")
  if isnumeric(UserID) then
	set RsList = server.CreateObject("adodb.recordset")
	SqlList = "SELECT LastLogoutTime FROM PinLuo_Admin where UserID = "&int(UserID)&" "
	RsList.open SqlList, objConn, 1, 3
	if not(RsList.eof) then
	  RsList("LastLogoutTime")=now
	  RsList.update
	end if
	RsList.close
	Set RsList = Nothing
 end if
    'session.Abandon() 
    Response.Cookies("pinluo")("UserID") = ""
	Response.Cookies("pinluo")("UserName") = ""
	Response.Cookies("pinluo")("RealName") = ""
	Response.Cookies("pinluo")("LoginTimes") = ""
	Response.Cookies("pinluo")("UserPassword") = ""
	Response.Cookies("pinluo")("RndPassword") = ""
	Response.Cookies("pinluo")("UserPopedom") = ""
	Response.Cookies("pinluo")("CheckCode") = ""
    call DBConnEnd()
	If IsObject(PinLuo) Then
		Set PinLuo = Nothing
	End if
	Response.Redirect("PL_login.asp") 
End Sub

'/////////////////////////////////////////////////////////////////////////////////////////
'	过程名：CheckSessionData
'	作  用：检查超时用户及未登陆用户
'	参  数：无
'/////////////////////////////////////////////////////////////////////////////////////////
Public Sub CheckSessionData()

	Dim rsGetAdmin,sqlGetAdmin
	Dim UserName,UserPassword,RndPassword

	UserName = replace(trim(Request.Cookies("pinluo")("UserName")),"'","")
	UserPassword = replace(trim(Request.Cookies("pinluo")("UserPassword")),"'","")
	RndPassword = replace(trim(Request.Cookies("pinluo")("RndPassword")),"'","")
	if UserName = "" or UserPassword = "" or RndPassword = "" then
		call DBConnEnd()
		FoundErr = True
		ErrMsg = ErrMsg & "<br><p><font color='red'>抱歉，管理员登陆会话超时！</font></p><p>您不能继续进行网站后台管理操作，请重新登陆系统！</p><p>您可以<a href='PL_login.asp' target='_top'>点此重新登录</a>。</p>"
		Exit Sub
	end if

	sqlGetAdmin="select * from PinLuo_Admin where UserName = '" & UserName & "' and UserPassword = '" & UserPassword & "'"
	set rsGetAdmin = server.CreateObject("adodb.recordset")
	rsGetAdmin.open sqlGetAdmin, objConn, 1, 1
	if rsGetAdmin.bof and rsGetAdmin.eof then
		rsGetAdmin.close
		set rsGetAdmin = nothing
		call DBConnEnd()
		FoundErr = True
		ErrMsg = ErrMsg & "<br><p><font color='red'>抱歉，管理员登陆会话超时！</font></p><p>您不能继续进行网站后台管理操作，请重新登陆系统！</p><p>您可以<a href='PL_login.asp' target='_top'>点此重新登录</a>。</p>"
		Exit Sub
	else
		if trim(rsGetAdmin("RndPassword")) <> RndPassword then
			FoundErr = True
			ErrMsg = ErrMsg & "<br><p><font color='red'>抱歉，为了系统安全，本系统不允许两个人使用同一个管理员帐号进行登录！</font></p><p>因为现在有人已经在其他地方使用此管理员帐号进行登录了，所以您将不能继续进行网站后台管理操作。</p><p>您可以<a href='PL_login.asp' target='_top'>点此重新登录</a>。</p>"
			rsGetAdmin.close
			set rsGetAdmin=nothing
			call DBConnEnd()
			Exit Sub
		end if
	end if
End Sub


Public Sub PinLuo_UserWarningSave(UserWarning,UserID)
    if isnumeric(UserID)=false then exit Sub
	set RsList = server.CreateObject("adodb.recordset")
	SqlList = "SELECT UserWarning FROM PinLuo_Admin where UserID="&int(UserID)
	RsList.open SqlList, objConn, 1, 3
	if not(RsList.eof) then
	  RsList("UserWarning")=UserWarning
	RsList.update
	end if
	RsList.close
	Set RsList = Nothing	
End Sub

Public Sub WriteSaveFile(TxtContent,FilePath)
  On Error Resume Next
  if trim(FilePath)="" then Exit Sub
  if IsObjInstalled("ADODB.Stream")=false then 
    FoundErr = True
    ErrMsg = ErrMsg & "<br><p><font color='red'>抱歉，您的主机不支持ADODB.Stream,不能保存文件。</p>"
    Exit Sub
  end if
  Dim st   
  Set st=Server.CreateObject("ADODB.Stream")
  st.Type=2   
  st.Mode=3   
  st.Charset="utf-8"
  st.Open()
	'st.WriteText "<!-- 该文件生成于"&now&"，由品络科技(www.pinluo.com)自主开发的企业网站管理系统自动生成。-->"& vbcrlf
  st.WriteText trim(TxtContent)
  st.SaveToFile Server.MapPath(FilePath),2
  st.Close()
  Set st=Nothing
End Sub

Public Function WriteReadFile(FilePath)
  On Error Resume Next
  WriteReadFile=""
  if trim(FilePath)="" then Exit Function
  if IsObjInstalled("ADODB.Stream")=False then 
    FoundErr = True
    ErrMsg = ErrMsg & "<br><p><font color='red'>抱歉，您的主机不支持ADODB.Stream,不能读取文件。</p>"
    Exit Function
  end if
  Dim st   
  Set st=Server.CreateObject("ADODB.Stream")
  st.Type=2   
  st.Mode=3   
  st.Open()
  st.LoadFromFile Server.MapPath(FilePath)
  If Err.Number<>0 Then
    FoundErr = True
    ErrMsg = ErrMsg & "<br><p><font color='red'>抱歉，文件"&FilePath&"无法打开，请检查是否存在！</p>"
    Err.Clear
    Exit Function
  End if
  st.Charset="utf-8"
  st.Position = 2
  WriteReadFile = trim(st.ReadText)
  st.Close()
  Set st=Nothing
End Function
 
Public Function UserPopedomShow(UserPopedomarr,UserPopedomCheckarr)
	StrUserPopedom2 = Split(UserPopedomarr, "|")
	UserPopedomCheck2 = Split(UserPopedomCheckarr, "|")
	For i = 0 To UBound(StrUserPopedom2)
	response.Write("<input name=""UserPopedom"&i&""" type=""checkbox"" value=""1"" ")
	if Trim(UserPopedomCheck2(i)) = 1 then Response.Write(" checked")
	response.Write(">"&StrUserPopedom2(i)&"&nbsp;&nbsp;")
    if i mod 8 = 7 then Response.Write("<br>")
	Next
End Function

Public Sub Pinluo_CheckPurviewAdmin(i)
    if isnumeric(i)=false then Exit Sub
    StrCheckPurviewAdmin = Split(Trim(Request.Cookies("pinluo")("UserPopedom")), "|")
	If Trim(StrCheckPurviewAdmin(i)) = 0 Then 
		FoundErr = True
		ErrMsg = ErrMsg & "<font color=red>抱歉，您没有本操作的权限，如有疑问请与管理员联系！</font>"
	End If 
	If FoundErr = True Then
		call DBConnEnd()
        PinLuo_WriteMsg ErrMsg,"#"
   		Response.End()
	End If
End Sub

Public Function Pinluo_CheckPurviewAdmin_Display(i)
    if isnumeric(i)=false then Exit Function
    StrCheckPurviewAdmin = Split(Trim(Request.Cookies("pinluo")("UserPopedom")), "|")
	If Trim(StrCheckPurviewAdmin(i)) = 0 Then 
		Pinluo_CheckPurviewAdmin_Display="none"
	Else
	    Pinluo_CheckPurviewAdmin_Display=""
	End If 
End Function

'/////////////////////////////////////////////////////////////////////////////////////////
'函数名：DelFile
'参  数：FilePath   ----文件夹名
'        FileName ----文件名
'/////////////////////////////////////////////////////////////////////////////////////////
Public Sub DelFile(FilePath,FileName)
	if FilePath<>"" and trim(left(FileName,6))<>"http://" then
		FilePath2=server.MapPath(FilePath)
	else
		exit sub
	end if
	SET fs=server.CreateObject("Scripting.FileSystemObject")
	if fs.FileExists(FilePath2 & "/" & FileName) then
		fs.DeleteFile FilePath2 & "/" & FileName,true
	end if
	set fs=nothing
end sub

'/////////////////////////////////////////////////////////////////////////////////////////
'函数名：DelFileUploadfile
'参  数：FilePath   ----文件夹名
'        FileName ----文件名
'/////////////////////////////////////////////////////////////////////////////////////////
Public Sub DelFileUploadfile(FilePath,FileName)
	if FileName<>"" and instr(FileName,"ditor/uploadfile")>0 then
		FilePath2=server.MapPath(FilePath & FileName)
	else
		exit sub
	end if
	SET fs=server.CreateObject("Scripting.FileSystemObject")
	if fs.FileExists(FilePath2) then
		fs.DeleteFile FilePath2,true
	end if
	set fs=nothing
end sub

Public Function Pinluo_showpage(form_name,action_name,page,rs,str)
'***********************************************************************************
'	函数名：Pinluo_showpage
'	作  用：通用分页函数
'	调  用：show_page "formname","index.asp",page,rs,"name#"name"#title#"title"#date#"date
'	参  数：窗体名称、页面的名称、页面的当前页数、SQL 数据指针记录集、变量名称、变量的值
'	参  数：formname		窗体名称
'	参  数：index.asp    页面的名称
'	参  数：page   页面的当前页数
'	参  数：rs      SQL 数据指针记录集
'	参  数："name#"  与  name      "name#" 为第一个变量名称，name 为该变量的值 
'	参  数："#title#" 与  title      "#title#" 为第二个变量名称，title 为该变量的值
'	参  数："#date#"  与  date     "#date#" 为第三个变量名称，date 为该变量的值
'	注  意：如果还有更多变量继续以此类推： #"变量4"#变量4;
'	注  意：除了第一个变量不用在其前面加"#" 外，其他的变量和值分别在其前面加 "#".
'***********************************************************************************

	Pinluo_showpage=Pinluo_showpage&"共找到<b>"&rs.recordcount&"</b>条记录&nbsp;&nbsp;"
	Pinluo_showpage=Pinluo_showpage&"每页显示<b>"&rs.pagesize&"</b>条&nbsp;"
	Pinluo_showpage=Pinluo_showpage&"<b>"&page&"/"&rs.pagecount&"</b>&nbsp;页&nbsp;&nbsp;"	
	
	strs=split(str,"#")
	
    if rs.pagecount=1 then Pinluo_showpage=Pinluo_showpage&"首页&nbsp;上页&nbsp;下页&nbsp;尾页&nbsp;"
	if page<>1 and rs.pagecount<>0 then
		constr="<a href=" & action_name & "?page=1"
		constr2="<a href=" & action_name & "?page=" & (page-1)
		for i=0 to ubound(strs)-1 step 2
			constr=constr & "&" & strs(i) & "=" & strs(i+1)
			constr2=constr2  & "&" & strs(i) & "=" & strs(i+1)
		next	   
		constr=constr & ">首页</a>&nbsp;"
		constr2=constr2 & ">上页</a>&nbsp;"
		Pinluo_showpage=Pinluo_showpage&constr
		Pinluo_showpage=Pinluo_showpage&constr2	   
		if page=rs.pagecount then Pinluo_showpage=Pinluo_showpage&"下页&nbsp;尾页&nbsp;"
	end if

	if page<>rs.pagecount then
		if page=1 then Pinluo_showpage=Pinluo_showpage&"首页&nbsp;上页&nbsp;"
		constr3="<a href=" & action_name & "?page=" & (page+1)
		constr4="<a href=" & action_name & "?page=" & rs.pagecount
		for i=0 to ubound(strs)-1 step 2
			constr3=constr3  & "&" & strs(i) & "=" & strs(i+1)
			constr4=constr4  & "&" & strs(i) & "=" & strs(i+1)
		next
		constr3=constr3 & ">下页</a>&nbsp;"
		constr4=constr4 & ">尾页</a>&nbsp;"
		Pinluo_showpage=Pinluo_showpage&constr3
		Pinluo_showpage=Pinluo_showpage&constr4
	end if
	
	if rs.pagecount>=1 then
		Pinluo_showpage=Pinluo_showpage& vbcrlf &"<select name='change_page'  id='in'>"& vbcrlf '说明 id='in' 是跳转页的着色
		for i=1 to rs.pagecount
			if page=i then
				Pinluo_showpage=Pinluo_showpage&"<option value=" &i& " selected>第" &i& "页</option>"& vbcrlf
			else
				Pinluo_showpage=Pinluo_showpage&"<option value=" & i & ">第" &i& "页</option>"& vbcrlf
			end if
		next
		Pinluo_showpage=Pinluo_showpage&"</select>"& vbcrlf
	end if

	for i=0 to ubound(strs)-1 step 2
		Pinluo_showpage=Pinluo_showpage&"<input type='hidden' name=" & strs(i) & " value=" & strs(i+1) &">"& vbcrlf
	next
	Pinluo_showpage=Pinluo_showpage&"<input type='hidden' name='str' value=" & str &">"& vbcrlf

	Pinluo_showpage=Pinluo_showpage&vbcrlf
	Pinluo_showpage=Pinluo_showpage&"<script language=""VBScript"">"& vbcrlf
	Pinluo_showpage=Pinluo_showpage&"	sub change_page_onchange()"& vbcrlf
	Pinluo_showpage=Pinluo_showpage&"		str="&form_name&".str.value"& vbcrlf
	Pinluo_showpage=Pinluo_showpage&"		strs=split(str,""#"")"& vbcrlf
	Pinluo_showpage=Pinluo_showpage&"		page22="&form_name&".change_page.value"& vbcrlf
	Pinluo_showpage=Pinluo_showpage&"		strh="""&action_name&"?page=""&page22"& vbcrlf
	Pinluo_showpage=Pinluo_showpage&"		for i_page = 0 to ubound(strs)-1 step 2"& vbcrlf
	Pinluo_showpage=Pinluo_showpage&"			strh=strh & ""&"" & strs(i_page) & ""="" & strs(i_page+1)"& vbcrlf
	Pinluo_showpage=Pinluo_showpage&"		next"& vbcrlf
	Pinluo_showpage=Pinluo_showpage&"		location.href=strh"& vbcrlf
	Pinluo_showpage=Pinluo_showpage&"	end sub"& vbcrlf
	Pinluo_showpage=Pinluo_showpage&"</"&"script>"& vbcrlf
End Function

'**************************************************
'函数名：IsObjInstalled
'作  用：检查组件是否已经安装
'参  数：strClassString ----组件名
'返回值：True  ----已经安装
'        False ----没有安装
'**************************************************
Function IsObjInstalled(strClassString)
    On Error Resume Next
    IsObjInstalled = False
    Err = 0
    Dim xTestObj
    Set xTestObj = CreateObject(strClassString)
    If Err.Number = 0 Then IsObjInstalled = True
    Set xTestObj = Nothing
    Err = 0
End Function

Function Pinluo_CheckObj(objid)
	If Not IsObjInstalled(objid) Then
		Pinluo_CheckObj = "<font color=""red"">&times;</font>"
	Else
		Pinluo_CheckObj = "<font color=""green"">&radic;</font>"
	End If
End Function

Public Sub ExecuteErr()
	'记录错误事件
		Response.Write "<span style='font-size:12px;'><br />"
		Response.Write "错 误 号：" & Err.Number & "<br />"
		Response.Write "错误描述：" & Err.Description & "<br />"
		Response.Write "错误来源：" & Err.Source & "</span>"
		Err.Clear
		Response.end
End Sub

Public Function ChkShenHe(str)
   if trim(str)="" then
     chkshenhe="未知"
   elseif eval(str)=true then
     chkshenhe="已审核"
   else
     chkshenhe="未审核"
   end if   
End Function
  
Public Function GetCurrentDate(strDate,strType)
'/////////////////////////////////////////////////////////////////////////////////////////
'函数名：GetCurrentDate
'作  用：格式化输入日期
'参  数：strDate   ----日期数据
'返回值：格式化后的字符串
'/////////////////////////////////////////////////////////////////////////////////////////
	strType = CStr(Trim(strType))
	If Not IsDate(strDate) Or Not IsNumeric(strType) Then
		GetCurrentDate = "---"
		Exit Function
	End If
	Select Case strType
		Case "0"
			GetCurrentDate = FormatDateTime(strDate, 0) 
		Case "1"
			GetCurrentDate = FormatDateTime(strDate, 1) 
		Case "2"
			GetCurrentDate = FormatDateTime(strDate, 2) 
		Case "3"
			GetCurrentDate = FormatDateTime(strDate, 3) 
		Case "4"
			GetCurrentDate = FormatDateTime(strDate, 4) 
		Case "5"
			GetCurrentDate = year(strDate) & "-" & right("0" & month(strDate),2) & "-" & right("0" & day(strDate),2)
		Case Else
			GetCurrentDate = strDate
		End Select
End Function

'/////////////////////////////////////////////////////////////////////////////////////////
'函数名：gotTopic
'作  用：截字符串，汉字一个算两个字符，英文算一个字符
'参  数：str   ----原字符串
'       strlen ----截取长度
'返回值：截取后的字符串
'/////////////////////////////////////////////////////////////////////////////////////////
Public Function gotTopic(str,strlen)
	if trim(str) = "" or trim(str) = null then
		gotTopic=""
		exit function
	end if
	dim l,t,c, i
	str=replace(replace(replace(replace(str,"&nbsp;"," "),"&quot;",chr(34)),"&gt;",">"),"&lt;","<")
	l=len(str)
	t=0
	for i=1 to l
		c=Abs(Asc(Mid(str,i,1)))
		if c>255 then
			t=t+2
		else
			t=t+1
		end if
		if t>=strlen then
			gotTopic=left(str,i) & "..."
			exit for
		else
			gotTopic=str
		end if
	next
	gotTopic=replace(replace(replace(replace(gotTopic," ","&nbsp;"),chr(34),"&quot;"),">","&gt;"),"<","&lt;")
End Function

'**************************************************
'函数名：ReplaceBadChar
'作  用：过滤非法的SQL字符
'参  数：strChar-----要过滤的字符
'返回值：过滤后的字符
'**************************************************
Public Function ReplaceBadChar(strChar)
    If strChar = "" Or IsNull(strChar) Then
        ReplaceBadChar = ""
        Exit Function
    End If
    Dim strBadChar, arrBadChar, tempChar, i
    strBadChar = "+,',%,^,&,?,(,),<,>,[,],{,},/,\,;,:," & Chr(34) & "," & Chr(0) & ",--"
    arrBadChar = Split(strBadChar, ",")
    tempChar = strChar
    For i = 0 To UBound(arrBadChar)
        tempChar = Replace(tempChar, arrBadChar(i), "")
    Next
    tempChar = Replace(tempChar, "@@", "@")
    ReplaceBadChar = tempChar
End Function

'/////////////////////////////////////////////////////////////////////////////////////////
'函数名：GetRndPassword
'作  用：生成随机码
'参  数：PasswordLen 
'返  回：RndPassword = GetRndPassword(16) t85b621664s5v6HL
'/////////////////////////////////////////////////////////////////////////////////////////
Public Function GetRndPassword(PasswordLen)
	Dim Ran,i,strPassword
	strPassword = ""
	For i=1 To PasswordLen
		Randomize
		Ran = CInt(Rnd * 2)
		Randomize
		If Ran = 0 Then
			Ran = CInt(Rnd * 25) + 97
			strPassword =strPassword & UCase(Chr(Ran))
		ElseIf Ran = 1 Then
			Ran = CInt(Rnd * 9)
			strPassword = strPassword & Ran
		ElseIf Ran = 2 Then
			Ran = CInt(Rnd * 25) + 97
			strPassword = strPassword & Chr(Ran)
		End If
	Next
	GetRndPassword = strPassword
End Function

Private Function getIP() 
Dim strIPAddr 
If Request.ServerVariables("HTTP_X_FORWARDED_FOR") = "" OR InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), "unknown") > 0 Then 
strIPAddr = Request.ServerVariables("REMOTE_ADDR") 
ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",") > 0 Then 
strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",")-1) 
ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";") > 0 Then 
strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";")-1) 
Else 
strIPAddr = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
End If 
getIP = Trim(Mid(strIPAddr, 1, 30)) 
End Function 

Public Function GetServerPath(s)
Dim Path
Dim Pos
if s=0 then
Path=Request.ServerVariables("script_name")
else
Path="http://" & Request.ServerVariables("server_name") & Request.ServerVariables("script_name")
end if
Pos=InStrRev(Path,"/")
GetServerPath=Left(Path,Pos)
End Function

Function GetPinluo_ValidateCode(PinLuo_codenum) '设置要生成多少位数
      Const PinLuo_cAmount = 36 ' 文字数量
      Const PinLuo_cCode = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
      Dim PinLuo_vCode(), PinLuo_vCodes
   redim PinLuo_vCode(PinLuo_codenum)
   Randomize
      For i = 0 To PinLuo_codenum-1
            PinLuo_vCode(i) = Int(Rnd * PinLuo_cAmount)
            PinLuo_vCodes = PinLuo_vCodes & Mid(PinLuo_cCode,PinLuo_vCode(i) + 1, 1)
      Next
  
   GetPinluo_ValidateCode="PinLuo.COM__"&PinLuo_vCodes
End Function
End Class  
%>
