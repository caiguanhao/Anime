<!-- #include file="@universal.asp" --><%
reUN=trim(fieldReplace(Request.form("username")))
reEM=trim(fieldReplace(Request.form("email")))
rePW=trim(fieldReplace(Request.form("password")))
rePW2=trim(fieldReplace(Request.form("password2")))

If reUN<>"" and Request.form("agree")<>"" Then
	Set RegEx = New RegExp
	RegEx.Global = True
	If len(reUN)<6 or len(reUN)>20 then alert "�û�������Ӧ��6��20���ַ���",""
	RegEx.Pattern = "^\w+$"
	If not RegEx.Test(reUN) Then alert "�û���ֻ������Ӣ����ĸ�����ֺ��»�����ɵĳ�����6��20���ַ�����",""
	
	Set conn=Server.CreateObject("ADODB.Connection")
	conn.Provider="Microsoft.Jet.OLEDB.4.0"
	conn.Open Server.Mappath(vDatabase)
	Set RSR=Server.CreateObject("ADODB.RecordSet")
	RSR.Open "Select * From [Users] WHERE [UserName]='"&reUN&"'",conn,3,3
	If not RSR.EOF Then
		Alert "�û����Ѵ��ڣ���ѡ�������û�����",""
	End IF
	RSR.Close
	Set RSR=nothing
	
	If len(reEM)<6 or len(reEM)>50 then alert "Email����Ӧ��6��50���ַ���",""
	RegEx.Pattern = "^\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$"
	If not RegEx.Test(reEM) Then alert "�����Ϲ����Email��ַ��",""
	
	If len(rePW)<6 or len(rePW)>50 then alert "���볤��Ӧ��6��20���ַ���",""
	RegEx.Pattern = "^[A-Za-z0-9]+$"
	If not RegEx.Test(rePW) Then alert "����ֻ������Ӣ����ĸ��������ɵĳ�����6��20���ַ�����",""
	If (rePW)<>(rePW2) then alert "������������������������롣",""
	Set RegEx = Nothing
	
	conn.execute("INSERT INTO [Users] ([UserName],[UserPass],[Email]) VALUES ('"&reUN&"','"&rePW&"','"&reEM&"')")
	Set conn=nothing
	
	Alert "��ϲ��ע��ɹ���","./"
ElseIf reUN<>"" Then
	R "register.asp"
End If
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" rev="stylesheet" type="text/css" href="images/style.css">
<title>ע��</title>
<script type="text/javascript" src="js/time.js"></script>
</head>

<body onload="">
<div align="center">
<table cellpadding="0" cellspacing="0" width="1000">
	<tr>
		<td>
		<!-- #include file="menu.asp" -->
		</td>
	</tr>
	<tr><td height="10"></td></tr>
	<tr>
		<td align="center">
		<table width="300">
		<tr>
			<td>
	<div class="roundedcornr_box_123135">
			   <div class="roundedcornr_top_123135"><div></div></div>
			      <div class="roundedcornr_content_123135">
					<form action="register.asp" method="post" style="margin:0px">
				      <table style="font-size:10pt">
						<tr>
							<td>�û���</td><td><input name="username" size="20" style="font-size:15pt" maxlength="20"></td>
						</tr>
						<tr>
							<td>Email</td><td><input name="email" size="20" style="font-size:15pt" maxlength="50"></td>
						</tr>
						<tr>
							<td>����</td><td><input name="password" type="password" size="20" style="font-size:15pt" maxlength="20"></td>
						</tr>
						<tr>
							<td>����<br>����</td><td><input name="password2" type="password" size="20" style="font-size:15pt" maxlength="20"></td>
						</tr>
						<tr>
							<td>�ض�</td><td><textarea rows="3" style="width:220px;font-size:9pt" readonly>�벻Ҫ������������������ˮ�����ݡ� ��ͬ���ݲ�Ҫ�ط������ݲ�Ҫ�������������衢á������ۣ�����Ӧ���Ϸ��ɹ涨��</textarea></td>
						</tr>
						<tr><td></td><td><input id="btnRegister" name="agree" type="submit" value="ͬ�Ⲣע��"></td></tr>
						</table>
				      </form>
			   </div>
			   <div class="roundedcornr_bottom_123135"><div></div></div>
			</div>
					</td>
		</tr>
		</table>
		</td>
	</tr>
		<tr>
			<td height="10"></td>
		</tr>
		<tr>
			<td>
			<!-- #include file="bottom.asp" --></td>
		</tr>
</table>
</div>
</body>

</html>