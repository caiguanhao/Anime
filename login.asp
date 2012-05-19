<!-- #include file="@universal.asp" --><%
	Set conn=Server.CreateObject("ADODB.Connection")
	conn.Provider="Microsoft.Jet.OLEDB.4.0"
	conn.Open Server.Mappath(vDatabase)
	Set RSR=Server.CreateObject("ADODB.RecordSet")
	If Request("logout")="1" Then
		Session("UserName")=""
		Session("UserAdmin")=""
		R "./"
	End If
	If Request.form("username")="" Then Alert "请输入用户名。",""
	If Request.form("password")="" Then Alert "请输入密码。",""
	RSR.Open "Select * From [Users] WHERE [UserName]='"&fieldReplace(Request.form("username"))&"'",conn,3,3
	If RSR.EOF Then
		Alert "用户名不存在！",""
		Session("UserName")=""
	End IF
	If RSR("UserPass")<>Request.form("password") Then
		Alert "密码错误！",""
		Session("UserName")=""
	End If
	Session("UserName")=ReplaceXYZ(RSR("UserName"))
	If RSR("IsAdmin")=True Then Session("UserAdmin")=ReplaceXYZ(RSR("UserName")) Else Session("UserAdmin")=""
	RSR.Close
	Set RSR=nothing
	Set conn=nothing
	If Request("p")<>"" Then R "./?p="&Request("p")
	R "./"
%>