<!-- #include file="@universal.asp" --><%
	If Session("UserAdmin")="" OR Session("UserName")="" Then Alert "µÇÂ¼³¬Ê±»òÎ´µÇÂ¼£¬ÇëµÇÂ¼ºó¼ÌÐø¡£","./"
	Set conn=Server.CreateObject("ADODB.Connection")
	conn.Provider="Microsoft.Jet.OLEDB.4.0"
	conn.Open Server.Mappath(vDatabase)
	NewTitle=Trim(fieldReplace((Request.form("title"))))
	If len(NewTitle)>0 Then
		If Len(NewTitle)>50 Then NewTitle=Left(NewTitle,50)
		NewCategories=","
		Set RSR=Server.CreateObject("ADODB.RecordSet")
		
		EditData=0
		If Request("p")<>"" and Isnumeric(Request("p"))=True Then
			RSR.Open "Select * From [Data] WHERE [ID]="&Request("p"),conn,3,3
			If RSR.EOF Then
				R "./"
			Else
				EditData=RSR("ID")
			End If
			RSR.Close
		End If
		
		For Each NewCategory in Request.form("category")
			RSR.Open "Select * From [Categories] WHERE [Name]='"&Trim(fieldReplace(unescape(NewCategory)))&"'",conn,3,3
			If not RSR.EOF Then
				NewCategories=NewCategories&RSR("ID")&","
			End If
			RSR.Close
		Next
		
		NewContent=Trim(fieldReplace((Request.form("content"))))
		NewPubOrDraft="False"
		If Request.form("pubordraft")="1" Then NewPubOrDraft="True"
		If EditData=0 Then
			conn.execute("INSERT INTO [Data] ([Title],[Category],[Content],[Published],[DateCreated]) VALUES ('"&NewTitle&"','"&NewCategories&"','"&NewContent&"',"&NewPubOrDraft&",#"&Now&"#)")
			RSR.Open "Select Top 1 * From [Data] ORDER BY [ID] DESC",conn,3,3
			If not RSR.EOF Then DataID=RSR("ID")
			R "./addoredit.asp?m=csuccess&p="&DataID
		Else
			conn.execute("UPDATE [Data] SET [Title]='"&NewTitle&"',[Category]='"&NewCategories&"',[Content]='"&NewContent&"',[Published]="&NewPubOrDraft&" WHERE [ID]="&EditData)
			R "./addoredit.asp?m=esuccess&p="&EditData
		End If
	End If
	Set RSR=nothing
	Set conn=nothing
	R "./x"
%>