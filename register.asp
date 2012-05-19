<!-- #include file="@universal.asp" --><%
reUN=trim(fieldReplace(Request.form("username")))
reEM=trim(fieldReplace(Request.form("email")))
rePW=trim(fieldReplace(Request.form("password")))
rePW2=trim(fieldReplace(Request.form("password2")))

If reUN<>"" and Request.form("agree")<>"" Then
	Set RegEx = New RegExp
	RegEx.Global = True
	If len(reUN)<6 or len(reUN)>20 then alert "用户名长度应在6～20个字符。",""
	RegEx.Pattern = "^\w+$"
	If not RegEx.Test(reUN) Then alert "用户名只能是由英文字母、数字和下划线组成的长度在6～20的字符串。",""
	
	Set conn=Server.CreateObject("ADODB.Connection")
	conn.Provider="Microsoft.Jet.OLEDB.4.0"
	conn.Open Server.Mappath(vDatabase)
	Set RSR=Server.CreateObject("ADODB.RecordSet")
	RSR.Open "Select * From [Users] WHERE [UserName]='"&reUN&"'",conn,3,3
	If not RSR.EOF Then
		Alert "用户名已存在，请选择其他用户名！",""
	End IF
	RSR.Close
	Set RSR=nothing
	
	If len(reEM)<6 or len(reEM)>50 then alert "Email长度应在6～50个字符。",""
	RegEx.Pattern = "^\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$"
	If not RegEx.Test(reEM) Then alert "不符合规则的Email地址。",""
	
	If len(rePW)<6 or len(rePW)>50 then alert "密码长度应在6～20个字符。",""
	RegEx.Pattern = "^[A-Za-z0-9]+$"
	If not RegEx.Test(rePW) Then alert "密码只能是由英文字母和数字组成的长度在6～20的字符串。",""
	If (rePW)<>(rePW2) then alert "两次密码输入错误，请重新输入。",""
	Set RegEx = Nothing
	
	conn.execute("INSERT INTO [Users] ([UserName],[UserPass],[Email]) VALUES ('"&reUN&"','"&rePW&"','"&reEM&"')")
	Set conn=nothing
	
	Alert "恭喜，注册成功。","./"
ElseIf reUN<>"" Then
	R "register.asp"
End If
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" rev="stylesheet" type="text/css" href="images/style.css">
<title>注册</title>
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
							<td>用户名</td><td><input name="username" size="20" style="font-size:15pt" maxlength="20"></td>
						</tr>
						<tr>
							<td>Email</td><td><input name="email" size="20" style="font-size:15pt" maxlength="50"></td>
						</tr>
						<tr>
							<td>密码</td><td><input name="password" type="password" size="20" style="font-size:15pt" maxlength="20"></td>
						</tr>
						<tr>
							<td>重输<br>密码</td><td><input name="password2" type="password" size="20" style="font-size:15pt" maxlength="20"></td>
						</tr>
						<tr>
							<td>必读</td><td><textarea rows="3" style="width:220px;font-size:9pt" readonly>请不要发表包含广告宣传、灌水的内容。 相同内容不要重发。内容不要包含攻击、侮辱、谩骂的言论，内容应符合法律规定。</textarea></td>
						</tr>
						<tr><td></td><td><input id="btnRegister" name="agree" type="submit" value="同意并注册"></td></tr>
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