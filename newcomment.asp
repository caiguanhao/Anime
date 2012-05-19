<!-- #include file="@universal.asp" --><%
	If Session("UserName")="" Then Alert "µÇÂ¼³¬Ê±»òÎ´µÇÂ¼£¬ÇëµÇÂ¼ºó¼ÌÐø¡£","./"
	Set conn=Server.CreateObject("ADODB.Connection")
	conn.Provider="Microsoft.Jet.OLEDB.4.0"
	conn.Open Server.Mappath(vDatabase)
	Set RSR=Server.CreateObject("ADODB.RecordSet")
	NewCP=Trim(Request.form("p"))
	If NewCP<>"" and Isnumeric(NewCP)=True Then NewCP=INT(NewCP) Else NewCP=0
	NewCN=Trim(fieldReplace((Session("UserName"))))
	NewCC=Trim(fieldReplace(unescape(Request.form("comment"))))
	NewCC=Replace(NewCC,vbCrlf," ")
	
	Set RegEx = New RegExp
	RegEx.Pattern = "<[^>]*>"
	RegEx.Global = True
	NewCC = RegEx.Replace(NewCC, "")
	Set RegEx = Nothing
	
	If len(NewCN)>0 AND Len(NewCN)<=50 Then
	RSR.Open "Select * From [Comments] WHERE [UserName]='"&NewCN&"' and [Content]='"&NewCC&"'",conn,3,3
	If RSR.EOF Then
		conn.execute("INSERT INTO [Comments] ([Post],[UserName],[Content],[DateCreated]) VALUES ("&NewCP&",'"&NewCN&"','"&NewCC&"',#"&Now&"#)")
		W "SUCCESSFULLYCREATED"
	ELSE
		W "COMMENTALREADYEXISTS"
	End If
	RSR.Close
	End If
	Set RSR=nothing
	Set conn=nothing
%>