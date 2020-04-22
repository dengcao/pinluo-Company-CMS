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

Class PinLuo_Class
Public objConn,objCmd
Public Pinluo_SiteName,Pinluo_SeoTitle,Pinluo_DelProImg,Pinluo_IsFeedback,Pinluo_Version,Pinluo_Empower,Pinluo_SiteUrl,Pinluo_SeoIndexTitle,Pinluo_SeoIndexKeyword,Pinluo_SeoIndexMS,Pinluo_Logo,Pinluo_Banner
Public ClassID,ClassName,ClassContents,ParentID,Depth,IsOuter,PathUrl,Visible,OrderID,ChildID
Public SEO_Title,SEO_Keyword,SEO_Description
Public InfoID,InfoTitle,InfoContent,InfoImg,Author,Origin
Public UpdateTime,hits,Shenhe
Public ProID,ProName,ProContent,ProImg1,ProImg2,ProPrice1,ProPrice2,Saled,Jian,Hot,Cheap
Public Feedbackid,FeedbackTitle,FeedbackContent,Phone,QQ,ReplyContent,ReplyTime
Public Block_ID,Block_Title,Block_Content,Block_Time
Public Daohang_ID,Daohang_Title,Daohang_Url,Daohang_Blank,Daohang_order
Public page,Pinluo_showpage_temp
Public RsList_i,ErrMsg

Private Sub Class_Initialize()
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
	  Pinluo_SeoIndexTitle=trim(RsList("Pinluo_SeoIndexTitle"))
	  Pinluo_SeoIndexKeyword=RsList("Pinluo_SeoIndexKeyword")
	  Pinluo_SeoIndexMS=RsList("Pinluo_SeoIndexMS")
	  Pinluo_Logo=RsList("Pinluo_Logo")
	  Pinluo_Banner=RsList("Pinluo_Banner")
	  if Pinluo_SeoIndexTitle="" then Pinluo_SeoIndexTitle=Pinluo_SeoTitle
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

Public Sub PinLuo_Classlist_View(ParentID,Depth,GotoUrl,Datatable,ii,px)
    if isnumeric(px)=false then px=6
	set RsList = server.CreateObject("adodb.recordset")
	SqlList = "SELECT A.*, ( SELECT Count(B.ClassID) FROM "&Datatable&" AS B WHERE A.ClassID = B.ParentID) AS CountList FROM "&Datatable&" AS A where A.Visible=true "
	if ParentID<>"" then
	SqlList = SqlList&"and ParentID="&int(ParentID)&" "
	else
	SqlList = SqlList&"and ParentID=0 "
	ParentID=0
	end if
	if Depth<>"" then SqlList = SqlList&"and Depth<="&Depth&" "
	SqlList = SqlList&"ORDER BY A.OrderID Desc,A.ClassID "
	RsList.open SqlList, objConn, 1, 1
	RsList_i=RsList_i+1
	do while not (RsList.eof)
	   response.Write("<li id=""class"&RsList("Depth")&""" ")
	   if RsList("ParentID")=int(ParentID) and ii<>"" then
	   response.Write(" style='padding-left:"&((RsList("Depth")-1)*px)&"px;'>|--- ")
	   else
	   response.Write(">")
	   end if
	   if GotoUrl="" then
	     response.Write(RsList("ClassName"))
	   elseif RsList("IsOuter")=true then
	     response.Write("<a href='"&RsList("PathUrl")&"'>"&RsList("ClassName")&"</a>")
	   else
	   	 response.Write("<a href='"&GotoUrl&"classid="&RsList("ClassID")&"'>"&RsList("ClassName")&"</a>")
	   end if
	   response.Write("</li>")
	   PinLuo_Classlist_View RsList("ClassID"),Depth,GotoUrl,Datatable,"1",px
	   RsList.movenext
	loop
	RsList.close
	Set RsList = Nothing
End Sub

Public Sub PinLuo_ViewClassItem(ClassID,Datatable)
    if isnumeric(ClassID)=false then exit Sub
	set RsList = server.CreateObject("adodb.recordset")
	SqlList = "SELECT * FROM "&Datatable&" where ClassID="&int(ClassID)
	RsList.open SqlList, objConn, 1, 1
	if not(RsList.eof) then
	  ClassName=RsList("ClassName")
	  SEO_Title=RsList("SEO_Title")
	  SEO_Keyword=RsList("SEO_Keyword")
	  SEO_Description=RsList("SEO_Description")
	  ClassContents=RsList("ClassContents")
	  ParentID=RsList("ParentID")
	  Depth=RsList("Depth")
	  IsOuter=RsList("IsOuter")
	  PathUrl=RsList("PathUrl")
	  Visible=RsList("Visible")
	  OrderID=RsList("OrderID")
	  if trim(SEO_Title)="" or isnull(SEO_Title) then SEO_Title=ClassName
	  if trim(SEO_Keyword)="" or isnull(SEO_Keyword) then SEO_Keyword=ClassName
	  if trim(SEO_Description)="" or isnull(SEO_Description) then SEO_Description=ClassName
	end if
	RsList.close
	Set RsList = Nothing
End Sub

Public Function PinLuo_ViewClassName(ClassID,Datatable,DefaultName)
    PinLuo_ViewClassName=DefaultName
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

Public Sub PinLuo_InfoList(HtmlStr,TitleNum,FormatTime,OnepageNum,PageForm,PageName,SearchKeyword,SearchSelect,SearchClassID,Datatable,ClassDatatable,Orderby)
    if isnumeric(TitleNum)=false then TitleNum=15
	if isnumeric(FormatTime)=false then FormatTime=5
	if HtmlStr="" then exit Sub
  	set RsList = server.CreateObject("adodb.recordset")
	SqlList = " select * from "&Datatable&" "
	SqlList = SqlList + " where InfoID > 0 "
	if isnumeric(SearchClassID)=true and trim(SearchClassID)<>"" then
	    SearchClassChildID=PinLuo_GetClassChildID(SearchClassID,ClassDatatable)
	    if SearchClassChildID="" then
	       'SqlList = SqlList + " and Classid="&int(SearchClassID)
		   SqlList = SqlList + " and Classid in("&SearchClassID&") "
		else
		   SqlList = SqlList + " and Classid in("&SearchClassID&","&SearchClassChildID&") "
		end if
	elseif trim(SearchClassID)<>"" then
	    SqlList = SqlList + " and Classid in("&SearchClassID&") "
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
	If Orderby="asc" then
	   SqlList = SqlList + " ORDER BY OrderID desc, InfoID aSC , UpdateTime DESC "
	Elseif  Orderby="new" then
	   SqlList = SqlList + " ORDER BY UpdateTime DESC , InfoID DESC "
	Elseif  Orderby="hot" then
	   SqlList = SqlList + " ORDER BY hits DESC , InfoID DESC "
	Else
	   SqlList = SqlList + " ORDER BY OrderID desc, UpdateTime DESC , InfoID DESC "
	End If	
	RsList.open SqlList, objConn, 1, 1
	if RsList.eof or RsList.bof then
		response.Write("没有找到任何信息记录！")
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
	
	HtmlStr2=replace(HtmlStr,"{$infoid}",RsList("InfoID"))
	HtmlStr2=replace(HtmlStr2,"{$title}",gotTopic(RsList("InfoTitle"),TitleNum))
	if RsList("InfoImg")<>null then HtmlStr2=replace(HtmlStr2,"{$infoimg}",RsList("InfoImg"))
	HtmlStr2=replace(HtmlStr2,"{$author}",RsList("Author"))
	HtmlStr2=replace(HtmlStr2,"{$origin}",RsList("Origin"))
	HtmlStr2=replace(HtmlStr2,"{$classid}",RsList("classid"))
	HtmlStr2=replace(HtmlStr2,"{$classname}",PinLuo_ViewClassName(RsList("Classid"),ClassDatatable,"未知"))
	HtmlStr2=replace(HtmlStr2,"{$time}",GetCurrentDate(RsList("UpdateTime"),FormatTime))
	HtmlStr2=replace(HtmlStr2,"{$hits}",RsList("hits"))
	HtmlStr2=replace(HtmlStr2,"{$orderid}",RsList("OrderID"))
	response.Write(HtmlStr2&vbcrlf)
	
	RsList.movenext
		if RsList.eof then exit for
	Next
	if PageForm="" then PageForm="Pinluo_Infolist"
	Pinluo_showpage_temp=Pinluo_showpage(PageForm,PageName,page,RsList, "SearchKeyword#" & SearchKeyword & "#SearchSelect#" & SearchSelect & "#classid#" & SearchClassID)
		
		
	end if 
	RsList.close
	set RsList = nothing
End Sub

Public Sub PinLuo_ViewInfoItem(InfoID,Datatable,ClassDatatable)
    if isnumeric(InfoID)=false or InfoID="" then exit Sub
	set RsList = server.CreateObject("adodb.recordset")
	SqlList = "SELECT * FROM "&Datatable&" where InfoID="&InfoID
	SqlList = SqlList + " and Shenhe=true " 
	RsList.open SqlList, objConn, 1, 3
	if not(RsList.eof) then
	  RsList("hits")=RsList("hits")+1
	  RsList.update
	  Classid=RsList("Classid")
	  ClassName=PinLuo_ViewClassName(RsList("Classid"),ClassDatatable,"未知")
	  InfoTitle=RsList("InfoTitle")
	  InfoContent=RsList("InfoContent")
	  InfoImg=RsList("InfoImg")
	  Author=RsList("Author")
	  Origin=RsList("Origin")
	  UpdateTime=RsList("UpdateTime")
	  hits=RsList("hits")
	  OrderID=RsList("OrderID")
	  Shenhe=RsList("Shenhe")
	  SEO_Title=RsList("SEO_Title")
	  SEO_Keyword=RsList("SEO_Keyword")
	  SEO_Description=RsList("SEO_Description")
	  if trim(SEO_Title)="" or isnull(SEO_Title) then SEO_Title=InfoTitle
	  if trim(SEO_Keyword)="" or isnull(SEO_Keyword) then SEO_Keyword=InfoTitle
	  if trim(SEO_Description)="" or isnull(SEO_Description) then SEO_Description=InfoTitle
	end if
	RsList.close
	Set RsList = Nothing
End Sub

Public Sub PinLuo_ViewNextInfo(HtmlStr,types,TitleNum,FormatTime,InfoID,SearchClassID,Datatable,ClassDatatable)
    if isnumeric(InfoID)=false or InfoID="" then exit Sub
	set RsList = server.CreateObject("adodb.recordset")
	SqlList = "SELECT top 1 * FROM "&Datatable&" where InfoID>0 "
	if isnumeric(SearchClassID) and trim(SearchClassID)<>"" then
	    SearchClassChildID=PinLuo_GetClassChildID(SearchClassID,ClassDatatable)
	    if SearchClassChildID="" then
	       SqlList = SqlList + " and Classid="&int(SearchClassID)
		else
		   SqlList = SqlList + " and Classid in("&SearchClassID&","&SearchClassChildID&") "
		end if
	end if
	if types=1 then
	SqlList = SqlList +"and InfoID>"&InfoID&" and Shenhe=true order by InfoID"
	else
	SqlList = SqlList +"and InfoID<"&InfoID&" and Shenhe=true order by InfoID desc"
	end if
	RsList.open SqlList, objConn, 1, 1
	if not(RsList.eof) then
	HtmlStr2=replace(HtmlStr,"{$infoid}",RsList("InfoID"))
	HtmlStr2=replace(HtmlStr2,"{$title}",gotTopic(RsList("InfoTitle"),TitleNum))
	if RsList("InfoImg")<>null then HtmlStr2=replace(HtmlStr2,"{$infoimg}",RsList("InfoImg"))
	HtmlStr2=replace(HtmlStr2,"{$author}",RsList("Author"))
	HtmlStr2=replace(HtmlStr2,"{$origin}",RsList("Origin"))
	HtmlStr2=replace(HtmlStr2,"{$classid}",RsList("classid"))
	HtmlStr2=replace(HtmlStr2,"{$classname}",PinLuo_ViewClassName(RsList("Classid"),ClassDatatable,"未知"))
	HtmlStr2=replace(HtmlStr2,"{$time}",GetCurrentDate(RsList("UpdateTime"),FormatTime))
	HtmlStr2=replace(HtmlStr2,"{$hits}",RsList("hits"))
	HtmlStr2=replace(HtmlStr2,"{$orderid}",RsList("OrderID"))
	else
	HtmlStr2=replace(HtmlStr,"{$title}","没有了")
	end if
	RsList.close
	Set RsList = Nothing
	response.Write(HtmlStr2&vbcrlf)
End Sub

Public Sub PinLuo_GetInfolist(HtmlStr,TitleNum,FormatTime,InfoNum,SearchClassID,Datatable,ClassDatatable,Orderby)
    if isnumeric(InfoNum)=false then InfoNum=5
    if isnumeric(TitleNum)=false then TitleNum=15
	if isnumeric(FormatTime)=false then FormatTime=5
	if HtmlStr="" then exit Sub
	set RsList = server.CreateObject("adodb.recordset")
	SqlList = " select top "&InfoNum&" * from "&Datatable&" "
	SqlList = SqlList + " where InfoID > 0 and Shenhe=true "
	if isnumeric(SearchClassID) and trim(SearchClassID)<>"" then
	    SearchClassChildID=PinLuo_GetClassChildID(SearchClassID,ClassDatatable)
	    if SearchClassChildID="" then
	       'SqlList = SqlList + " and Classid="&int(SearchClassID)
		   SqlList = SqlList + " and Classid in("&SearchClassID&") "
		else
		   SqlList = SqlList + " and Classid in("&SearchClassID&","&SearchClassChildID&") "
		end if
	elseif trim(SearchClassID)<>"" then
	    SqlList = SqlList + " and Classid in("&SearchClassID&") "
	end if
	If Orderby="asc" then
	   SqlList = SqlList + " ORDER BY OrderID desc, InfoID aSC , UpdateTime DESC "
	Elseif  Orderby="new" then
	   SqlList = SqlList + " ORDER BY UpdateTime DESC , InfoID DESC "
	Elseif  Orderby="hot" then
	   SqlList = SqlList + " ORDER BY hits DESC , InfoID DESC "
	Else
	   SqlList = SqlList + " ORDER BY OrderID desc, UpdateTime DESC , InfoID DESC "
	End If	
	RsList.open SqlList, objConn, 1, 1
	do while not(RsList.eof)
	HtmlStr2=replace(HtmlStr,"{$infoid}",RsList("InfoID"))
	HtmlStr2=replace(HtmlStr2,"{$title}",gotTopic(RsList("InfoTitle"),TitleNum))
	if RsList("InfoImg")<>null then HtmlStr2=replace(HtmlStr2,"{$infoimg}",RsList("InfoImg"))
	HtmlStr2=replace(HtmlStr2,"{$author}",RsList("Author"))
	HtmlStr2=replace(HtmlStr2,"{$origin}",RsList("Origin"))
	HtmlStr2=replace(HtmlStr2,"{$classid}",RsList("classid"))
	HtmlStr2=replace(HtmlStr2,"{$classname}",PinLuo_ViewClassName(RsList("Classid"),ClassDatatable,"未知"))
	HtmlStr2=replace(HtmlStr2,"{$time}",GetCurrentDate(RsList("UpdateTime"),FormatTime))
	HtmlStr2=replace(HtmlStr2,"{$hits}",RsList("hits"))
	HtmlStr2=replace(HtmlStr2,"{$orderid}",RsList("OrderID"))
	response.Write(HtmlStr2&vbcrlf)
	RsList.movenext
	loop
	RsList.close
	Set RsList = Nothing
End Sub

Public Sub PinLuo_ProductList(HtmlStr,TitleNum,FormatTime,OnepageNum,PageForm,PageName,SearchKeyword,SearchSelect,SearchClassID,Datatable,ClassDatatable,Orderby)
    if isnumeric(TitleNum)=false then TitleNum=15
	if isnumeric(FormatTime)=false then FormatTime=5
	if HtmlStr="" then exit Sub
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
	If Orderby="asc" then
	   SqlList = SqlList + " ORDER BY OrderID desc, ProID aSC , UpdateTime DESC "
	Elseif  Orderby="new" then
	   SqlList = SqlList + " ORDER BY UpdateTime DESC , ProID DESC "
	Elseif  Orderby="hot" then
	   SqlList = SqlList + " ORDER BY hits DESC , ProID DESC "
	Else
	   SqlList = SqlList + " ORDER BY OrderID desc, UpdateTime DESC , ProID DESC "
	End If	
	RsList.open SqlList, objConn, 1, 1
	if RsList.eof or RsList.bof then
		response.Write("没有找到任何信息记录！")
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
	
	HtmlStr2=replace(HtmlStr,"{$proid}",RsList("ProID"))
	HtmlStr2=replace(HtmlStr2,"{$proname}",gotTopic(RsList("ProName"),TitleNum))
	HtmlStr2=replace(HtmlStr2,"{$procontent}",RsList("ProContent"))
	HtmlStr2=replace(HtmlStr2,"{$proimg1}",RsList("ProImg1"))
	HtmlStr2=replace(HtmlStr2,"{$proimg2}",RsList("ProImg2"))
	HtmlStr2=replace(HtmlStr2,"{$proprice1}",RsList("ProPrice1"))
	HtmlStr2=replace(HtmlStr2,"{$proprice2}",RsList("ProPrice2"))
	HtmlStr2=replace(HtmlStr2,"{$saled}",RsList("Saled"))
	HtmlStr2=replace(HtmlStr2,"{$classid}",RsList("classid"))
	HtmlStr2=replace(HtmlStr2,"{$classname}",PinLuo_ViewClassName(RsList("Classid"),ClassDatatable,"未知"))
	HtmlStr2=replace(HtmlStr2,"{$time}",GetCurrentDate(RsList("UpdateTime"),FormatTime))
	HtmlStr2=replace(HtmlStr2,"{$hits}",RsList("hits"))
	HtmlStr2=replace(HtmlStr2,"{$orderid}",RsList("OrderID"))
	response.Write(HtmlStr2&vbcrlf)
	
	RsList.movenext
		if RsList.eof then exit for
	Next
	if PageForm="" then PageForm="Pinluo_productlist"
	Pinluo_showpage_temp=Pinluo_showpage(PageForm,PageName,page,RsList, "SearchKeyword#" & SearchKeyword & "#SearchSelect#" & SearchSelect & "#classid#" & SearchClassID)
		
		
	end if 
	RsList.close
	set RsList = nothing
End Sub

Public Sub PinLuo_GetProductlist(HtmlStr,TitleNum,FormatTime,ProNum,SearchClassID,Datatable,ClassDatatable,Orderby)
    if isnumeric(ProNum)=false then ProNum=5
    if isnumeric(TitleNum)=false then TitleNum=15
	if isnumeric(FormatTime)=false then FormatTime=5
	if HtmlStr="" then exit Sub
	set RsList = server.CreateObject("adodb.recordset")
	SqlList = " select top "&ProNum&" * from "&Datatable&" "
	SqlList = SqlList + " where ProID > 0 and Shenhe=true "
	if isnumeric(SearchClassID) and trim(SearchClassID)<>"" then
	    SearchClassChildID=PinLuo_GetClassChildID(SearchClassID,ClassDatatable)
	    if SearchClassChildID="" then
	       'SqlList = SqlList + " and Classid="&int(SearchClassID)
		   SqlList = SqlList + " and Classid in("&SearchClassID&") "
		else
		   SqlList = SqlList + " and Classid in("&SearchClassID&","&SearchClassChildID&") "
		end if
	elseif trim(SearchClassID)<>"" then
	    SqlList = SqlList + " and Classid in("&SearchClassID&") "
	end if
	If Orderby="asc" then
	   SqlList = SqlList + " ORDER BY OrderID desc, ProID aSC , UpdateTime DESC "
	Elseif  Orderby="new" then
	   SqlList = SqlList + " ORDER BY UpdateTime DESC , ProID DESC "
	Elseif  Orderby="hot" then
	   SqlList = SqlList + " ORDER BY hits DESC , ProID DESC "
	Else
	   SqlList = SqlList + " ORDER BY OrderID desc, UpdateTime DESC , ProID DESC "
	End If	
	RsList.open SqlList, objConn, 1, 1
	do while not(RsList.eof)
	HtmlStr2=replace(HtmlStr,"{$proid}",RsList("ProID"))
	HtmlStr2=replace(HtmlStr2,"{$proname}",gotTopic(RsList("ProName"),TitleNum))
	HtmlStr2=replace(HtmlStr2,"{$proimg1}",RsList("ProImg1"))
	HtmlStr2=replace(HtmlStr2,"{$proimg2}",RsList("ProImg2"))
	HtmlStr2=replace(HtmlStr2,"{$price1}",RsList("ProPrice1"))
	HtmlStr2=replace(HtmlStr2,"{$price2}",RsList("ProPrice2"))
	HtmlStr2=replace(HtmlStr2,"{$saled}",RsList("Saled"))
	HtmlStr2=replace(HtmlStr2,"{$procontent}",RsList("ProContent"))
	HtmlStr2=replace(HtmlStr2,"{$classid}",RsList("classid"))
	HtmlStr2=replace(HtmlStr2,"{$classname}",PinLuo_ViewClassName(RsList("Classid"),ClassDatatable,"未知"))
	HtmlStr2=replace(HtmlStr2,"{$time}",GetCurrentDate(RsList("UpdateTime"),FormatTime))
	HtmlStr2=replace(HtmlStr2,"{$hits}",RsList("hits"))
	HtmlStr2=replace(HtmlStr2,"{$orderid}",RsList("OrderID"))
	response.Write(HtmlStr2&vbcrlf)
	RsList.movenext
	loop
	RsList.close
	Set RsList = Nothing
End Sub

Public Sub PinLuo_ViewProductItem(ProID,Datatable,ClassDatatable)
    if isnumeric(ProID)=false or ProID="" then exit Sub
	set RsList = server.CreateObject("adodb.recordset")
	SqlList = "SELECT * FROM "&Datatable&" where ProID="&ProID
	SqlList = SqlList + " and Shenhe=true " 
	RsList.open SqlList, objConn, 1, 3
	if not(RsList.eof) then
	  RsList("hits")=RsList("hits")+1
	  RsList.update
	  Classid=RsList("Classid")
	  ClassName=PinLuo_ViewClassName(RsList("Classid"),ClassDatatable,"未知")
	  ProName=RsList("ProName")
	  ProContent=RsList("ProContent")
	  ProImg1=RsList("ProImg1")
	  ProImg2=RsList("ProImg2")
	  ProPrice1=RsList("ProPrice1")
	  ProPrice2=RsList("ProPrice2")
	  if ProPrice1<0 then ProPrice1="面议"
	  if ProPrice2<0 then ProPrice2="面议"
	  Saled=RsList("Saled")
	  Jian=RsList("Jian")
	  Hot=RsList("Hot")
	  Cheap=RsList("Cheap")
	  UpdateTime=RsList("UpdateTime")
	  hits=RsList("hits")
	  OrderID=RsList("OrderID")
	  Shenhe=RsList("Shenhe")
	  SEO_Title=RsList("SEO_Title")
	  SEO_Keyword=RsList("SEO_Keyword")
	  SEO_Description=RsList("SEO_Description")
	  if trim(SEO_Title)="" or isnull(SEO_Title) then SEO_Title=ProName
	  if trim(SEO_Keyword)="" or isnull(SEO_Keyword) then SEO_Keyword=ProName
	  if trim(SEO_Description)="" or isnull(SEO_Description) then SEO_Description=ProName
	end if
	RsList.close
	Set RsList = Nothing
End Sub

Public Function PinLuo_AddFeedback(Classid,FeedbackTitle,FeedbackContent,Author,UpdateTime,Shenhe,Phone,Email,QQ,Datatable)
    if isnumeric(Classid)=false then PinLuo_AddFeedback=false:exit Function
	if trim(FeedbackTitle)="" then PinLuo_AddFeedback=false:exit Function
	if trim(FeedbackContent)="" then PinLuo_AddFeedback=false:exit Function
	if isdate(UpdateTime)=false then UpdateTime=now
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
	  RsList("ReplyContent") = ""
	  RsList("ReplyTime") = empty
	  RsList("Hits") = 0
	  RsList("OrderID") = 0
	  RsList.update
	  FeedbackID=RsList("FeedbackID")
	  PinLuo_AddFeedback=true
	RsList.close
	Set RsList = Nothing
End Function

Public Function PinLuo_ViewBlockItem(Block_ID,t)
    PinLuo_ViewBlockItem=""
    if isnumeric(Block_ID)=false or trim(Block_ID)="" then exit Function
	set RsList = server.CreateObject("adodb.recordset")
	SqlList = "SELECT * FROM PinLuo_Block where Block_ID="&Block_ID
	RsList.open SqlList, objConn, 1, 1
	if not(RsList.eof) then
	  if t=1 then
	     PinLuo_ViewBlockItem=RsList("Block_Title")
	  elseif t=2 then
	     PinLuo_ViewBlockItem=RsList("Block_Time")
	  else
	     PinLuo_ViewBlockItem=RsList("Block_Content")
	  end if
	end if
	RsList.close
	Set RsList = Nothing
End Function

Public Sub PinLuo_GetDaohanglist(HtmlStr,TitleNum,InfoNum,Datatable)
    if isnumeric(InfoNum)=false then InfoNum=5
    if isnumeric(TitleNum)=false then TitleNum=15
	if HtmlStr="" then exit Sub
	set RsList = server.CreateObject("adodb.recordset")
	SqlList = " select top "&InfoNum&" * from "&Datatable&" "
    SqlList = SqlList + " ORDER BY Daohang_order desc"

	RsList.open SqlList, objConn, 1, 1
	do while not(RsList.eof)
	HtmlStr2=replace(HtmlStr,"{$ID}",RsList("Daohang_ID"))
	HtmlStr2=replace(HtmlStr2,"{$title}",gotTopic(RsList("Daohang_Title"),TitleNum))
	HtmlStr2=replace(HtmlStr2,"{$url}",RsList("Daohang_Url"))
	if RsList("Daohang_Blank")=1 then Daohang_Blank_s="_blank" else  Daohang_Blank_s="_self"
	HtmlStr2=replace(HtmlStr2,"{$Blank}",Daohang_Blank_s)
	HtmlStr2=replace(HtmlStr2,"{$order}",RsList("Daohang_order"))
	if isObject(Pinluo) then HtmlStr2=replace(HtmlStr2,"{$siteurl}",Pinluo.Pinluo_SiteUrl)
	response.Write(HtmlStr2&vbcrlf)
	RsList.movenext
	loop
	RsList.close
	Set RsList = Nothing
End Sub

Public Sub PinLuo_ViewItemContent(ClassID,Datatable)
    if isnumeric(ClassID)=false then exit Sub
	set RsList = server.CreateObject("adodb.recordset")
	SqlList = "SELECT * FROM "&Datatable&" where ClassID="&int(ClassID)
	RsList.open SqlList, objConn, 1, 1
	if not(RsList.eof) then
	  ClassName=RsList("ClassName")
	  SEO_Title=RsList("SEO_Title")
	  SEO_Keyword=RsList("SEO_Keyword")
	  SEO_Description=RsList("SEO_Description")
	  ClassContents=RsList("ClassContents")
	  ParentID=RsList("ParentID")
	  Depth=RsList("Depth")
	  IsOuter=RsList("IsOuter")
	  PathUrl=RsList("PathUrl")
	  Visible=RsList("Visible")
	  OrderID=RsList("OrderID")
	  ChildID=RsList("ChildID")
	  if trim(SEO_Title)="" or isnull(SEO_Title) then SEO_Title=ClassName
	  if trim(SEO_Keyword)="" or isnull(SEO_Keyword) then SEO_Keyword=ClassName
	  if trim(SEO_Description)="" or isnull(SEO_Description) then SEO_Description=ClassName
	end if
	RsList.close
	Set RsList = Nothing
End Sub

Public Sub PinLuo_FeedbackList(HtmlStr,TitleNum,FormatTime,OnepageNum,PageForm,PageName,SearchKeyword,SearchSelect,SearchClassID,Datatable,ClassDatatable,Orderby)
    if isnumeric(TitleNum)=false then TitleNum=15
	if isnumeric(FormatTime)=false then FormatTime=5
	if HtmlStr="" then exit Sub
  	set RsList = server.CreateObject("adodb.recordset")
	SqlList = " select * from "&Datatable&" "
	SqlList = SqlList + " where FeedbackID > 0 "
	if isnumeric(SearchClassID)=true and trim(SearchClassID)<>"" then
	    SearchClassChildID=PinLuo_GetClassChildID(SearchClassID,ClassDatatable)
	    if SearchClassChildID="" then
	       'SqlList = SqlList + " and Classid="&int(SearchClassID)
		   SqlList = SqlList + " and Classid in("&SearchClassID&") "
		else
		   SqlList = SqlList + " and Classid in("&SearchClassID&","&SearchClassChildID&") "
		end if
	elseif trim(SearchClassID)<>"" then
	    SqlList = SqlList + " and Classid in("&SearchClassID&") "
	end if
	if SearchKeyword <> "" then
		Select Case SearchSelect
			Case "FeedbackTitle"
				SqlList = SqlList + " and FeedbackTitle like '%"&SearchKeyword&"%' " 
			Case "FeedbackContent"
				SqlList = SqlList + " and FeedbackContent like '%"&SearchKeyword&"%' " 
			Case "Author"
				SqlList = SqlList + " and Author like '%"&SearchKeyword&"%' " 
			Case "Phone"
				SqlList = SqlList + " and Phone like '%"&SearchKeyword&"%' " 
			Case "Email"
				SqlList = SqlList + " and Email like '%"&SearchKeyword&"%' " 
			Case "QQ"
				SqlList = SqlList + " and QQ like '%"&SearchKeyword&"%' " 	
			Case "ReplyContent"
				SqlList = SqlList + " and ReplyContent like '%"&SearchKeyword&"%' " 		
			Case Else
				SqlList = SqlList + " and FeedbackTitle&FeedbackContent&Author&Phone&Email&QQ&ReplyContent like '%"&SearchKeyword&"%' " 
		End Select
	End If
	    Select Case SearchSelect
	        Case "Shenhe1"
				SqlList = SqlList + " and Shenhe=false " 
			Case "Shenhe2"
				SqlList = SqlList + " and Shenhe=true " 
		End Select
	If Orderby="asc" then
	   SqlList = SqlList + " ORDER BY OrderID desc, FeedbackID aSC , UpdateTime DESC "
	Elseif  Orderby="new" then
	   SqlList = SqlList + " ORDER BY UpdateTime DESC , FeedbackID DESC "
	Elseif  Orderby="hot" then
	   SqlList = SqlList + " ORDER BY hits DESC , FeedbackID DESC "
	Else
	   SqlList = SqlList + " ORDER BY OrderID desc, UpdateTime DESC , FeedbackID DESC "
	End If	
	RsList.open SqlList, objConn, 1, 1
	if RsList.eof or RsList.bof then
		response.Write("没有找到任何信息记录！")
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
	
	HtmlStr2=replace(HtmlStr,"{$feedbackid}",RsList("FeedbackID"))
	HtmlStr2=replace(HtmlStr2,"{$feedbacktitle}",gotTopic(RsList("FeedbackTitle"),TitleNum))
	HtmlStr2=replace(HtmlStr2,"{$author}",RsList("Author"))
	HtmlStr2=replace(HtmlStr2,"{$phone}",RsList("Phone"))
	HtmlStr2=replace(HtmlStr2,"{$email}",RsList("Email"))
	HtmlStr2=replace(HtmlStr2,"{$qq}",RsList("QQ"))
	HtmlStr2=replace(HtmlStr2,"{$classid}",RsList("classid"))
	HtmlStr2=replace(HtmlStr2,"{$classname}",PinLuo_ViewClassName(RsList("Classid"),ClassDatatable,"未知"))
	HtmlStr2=replace(HtmlStr2,"{$time}",GetCurrentDate(RsList("UpdateTime"),FormatTime))
	HtmlStr2=replace(HtmlStr2,"{$hits}",RsList("hits"))
	HtmlStr2=replace(HtmlStr2,"{$orderid}",RsList("OrderID"))
	if trim(RsList("ReplyTime"))<>"" and isnull(RsList("ReplyTime"))=false then HtmlStr2=replace(HtmlStr2,"{$replytime}",RsList("ReplyTime"))
	response.Write(HtmlStr2&vbcrlf)
	
	RsList.movenext
		if RsList.eof then exit for
	Next
	if PageForm="" then PageForm="Pinluo_Feedbacklist"
	Pinluo_showpage_temp=Pinluo_showpage(PageForm,PageName,page,RsList, "SearchKeyword#" & SearchKeyword & "#SearchSelect#" & SearchSelect & "#classid#" & SearchClassID)
		
		
	end if 
	RsList.close
	set RsList = nothing
End Sub

Public Sub PinLuo_ViewFeedbackItem(FeedbackID,Datatable,ClassDatatable)
    if isnumeric(FeedbackID)=false or FeedbackID="" then exit Sub
	set RsList = server.CreateObject("adodb.recordset")
	SqlList = "SELECT * FROM "&Datatable&" where FeedbackID="&FeedbackID
	SqlList = SqlList + " and Shenhe=true " 
	RsList.open SqlList, objConn, 1, 3
	if not(RsList.eof) then
	  RsList("hits")=RsList("hits")+1
	  RsList.update
	  Classid=RsList("Classid")
	  ClassName=PinLuo_ViewClassName(RsList("Classid"),ClassDatatable,"未知")
	  FeedbackTitle=RsList("FeedbackTitle")
	  FeedbackContent=RsList("FeedbackContent")
	  Author=RsList("Author")
	  UpdateTime=RsList("UpdateTime")
	  hits=RsList("hits")
	  OrderID=RsList("OrderID")
	  Shenhe=RsList("Shenhe")
	  Phone=RsList("Phone")
	  Email=RsList("Email")
	  QQ=RsList("QQ")
	  ReplyContent=RsList("ReplyContent")
	  ReplyTime=RsList("ReplyTime")
	  SEO_Title=RsList("SEO_Title")
	  SEO_Keyword=RsList("SEO_Keyword")
	  SEO_Description=RsList("SEO_Description")
	  if trim(SEO_Title)="" or isnull(SEO_Title) then SEO_Title=FeedbackTitle
	  if trim(SEO_Keyword)="" or isnull(SEO_Keyword) then SEO_Keyword=FeedbackTitle
	  if trim(SEO_Description)="" or isnull(SEO_Description) then SEO_Description=FeedbackTitle
	end if
	RsList.close
	Set RsList = Nothing
End Sub

Public Function Pinluo_GetUrl()
On Error Resume Next     
Dim strTemp     
If LCase(Request.ServerVariables("HTTPS"))="off" Then     
strTemp="http://"     
Else     
strTemp="https://"     
End If     
strTemp=strTemp&Request.ServerVariables("SERVER_NAME")     
If Request.ServerVariables("SERVER_PORT") <> 80 Then strTemp=strTemp&":"&Request.ServerVariables("SERVER_PORT")     
strTemp= strTemp&Request.ServerVariables("URL")     
If Trim(Request.QueryString)<>"" Then strTemp=strTemp&"?"&Trim(Request.QueryString)     
Pinluo_GetUrl= strTemp
Pinluo_GetUrl=mid(Pinluo_GetUrl,instr(Pinluo_GetUrl,"?")+1)
Pinluo_GetUrl=trim(replace(Pinluo_GetUrl,"ss=1&",""))
End Function

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

Public Function isnumeric(str)
   if trim(str)="" or isnull(str) then isnumeric=false:exit Function
   if InStr(str,",") > 0 then isnumeric=false:exit Function
   if InStr(str,"，") > 0 then isnumeric=false:exit Function
isnumeric=true
for i =1 to len(str)
e = Lcase(Mid(str, i, 1))
if InStr("0123456789.", e) <= 0 then
isnumeric = false
exit function
end if
next
End Function

Function Pinluo_GetUrlID()
On Error Resume Next     
Dim strTemp     
If LCase(Request.ServerVariables("HTTPS"))="off" Then     
strTemp="http://"     
Else     
strTemp="https://"     
End If     
strTemp=strTemp&Request.ServerVariables("SERVER_NAME")     
If Request.ServerVariables("SERVER_PORT") <> 80 Then strTemp=strTemp&":"&Request.ServerVariables("SERVER_PORT")     
strTemp= strTemp&Request.ServerVariables("URL")     
If Trim(Request.QueryString)<>"" Then strTemp=strTemp&"?"&Trim(Request.QueryString)     
Pinluo_GetUrlID= strTemp
Pinluo_GetUrlID=mid(Pinluo_GetUrlID,instr(Pinluo_GetUrlID,"?")+1)
Pinluo_GetUrlID=trim(replace(Pinluo_GetUrlID,".html",""))
End Function

End Class  
%>