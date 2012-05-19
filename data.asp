<!-- #include file="@universal.asp" --><%
	oOutput=""
	Set conn=Server.CreateObject("ADODB.Connection")
	conn.Provider="Microsoft.Jet.OLEDB.4.0"
	conn.Open Server.Mappath(vDatabase)
	Set RSR=Server.CreateObject("ADODB.RecordSet")
Select Case Request("t")
Case "comment"
	CCount=1
	HowManyEveryPage=10
	CurrentPage=Request("page")
	If CurrentPage<>"" and Isnumeric(CurrentPage)=True Then CurrentPage=abs(int(CurrentPage)) Else CurrentPage=1
	
	If Request("p")<>"" and Isnumeric(Request("p"))=True Then SQLADD=" WHERE [Post]="&Request("p") Else SQLADD=" WHERE [Post]=0"
	RSR.Open "Select * From [Comments]"&SQLADD&" ORDER BY [ID] DESC",conn,3,3
	If not RSR.EOF Then
		RSR.Move (CurrentPage-1)*HowManyEveryPage
		Do While Not RSR.EOF AND CCount<=HowManyEveryPage
			oOutput=oOutput&Server.HTMLEncode(ReplaceXYZ(RSR("Content")))&"<X>"&ReplaceXYZ(RSR("UserName"))&"<X>"&RSR("DateCreated")&"<X>"&RSR("ID")&"<Y>"
		RSR.MoveNext
		CCount=CCount+1
		Loop
	
	oOutput=oOutput&"<N"&(CurrentPage+1)&">"
	If RSR.RecordCount>HowManyEveryPage and (CurrentPage)*HowManyEveryPage<RSR.RecordCount Then
	Else
		oOutput=oOutput&"<END>"
	End If
	
	W escape(oOutput)
	End If
	
	RSR.Close
Case else
	RSR.Open "Select * From [Categories]",conn,3,3
	If not RSR.EOF Then
		Do While Not RSR.EOF
			oOutput=oOutput&ReplaceXYZ(RSR("Name"))&","
		RSR.MoveNext
		Loop
		W escape(oOutput)
	End If
	RSR.Close
End Select
	Set RSR=nothing
	Set conn=nothing
%>