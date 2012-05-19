<!-- #include file="@universal.asp" --><%
	If Session("UserAdmin")="" OR Session("UserName")="" Then Alert "µÇÂ¼³¬Ê±»òÎ´µÇÂ¼£¬ÇëµÇÂ¼ºó¼ÌÐø¡£","./"
	Set conn=Server.CreateObject("ADODB.Connection")
	conn.Provider="Microsoft.Jet.OLEDB.4.0"
	conn.Open Server.Mappath(vDatabase)
	Set RSR=Server.CreateObject("ADODB.RecordSet")
	NewCN=Trim(fieldReplace(unescape(Request.form("NewCategoryName"))))
	If len(NewCN)>0 AND Len(NewCN)<=50 Then
	RSR.Open "Select * From [Categories] WHERE [Name]='"&NewCN&"'",conn,3,3
	If RSR.EOF Then
		conn.execute("INSERT INTO [Categories] ([Name]) VALUES ('"&NewCN&"')")
		W "SUCCESSFULLYCREATED"
	ELSE
		W "CATEGORYALREADYEXISTS"
	End If
	RSR.Close
	End If
	Set RSR=nothing
	Set conn=nothing
%>