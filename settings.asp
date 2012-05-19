<!-- #include file="@universal.asp" --><%
	If Session("UserAdmin")="" OR Session("UserName")="" Then Alert "登录超时或未登录，请登录后继续。","./"
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" rev="stylesheet" type="text/css" href="images/style.css">
<title>网站配置</title>
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
					<form action="settings.asp?v=<%=Request("v")%>" method="post" style="margin:0px" onsubmit="<%If Request("v")="users" Then%>if(!confirm('确认？注：删除用户将无法恢复；提升管理员后需要另行打开数据库修改降级。')){return false}else{if(/管理员/.test(document.getElementById('usrsel').options[document.getElementById('usrsel').selectedIndex].text)){alert('不能删除管理员。');return false;}}<%End If%>">
				      <table style="font-size:10pt">
						<tr><td colspan="2"><h2>网站配置<!-- #include file="control.asp" --></h2></td></tr>
<%Select Case Request("v")%>
<%Case "cate"
	Set conn=Server.CreateObject("ADODB.Connection")
	conn.Provider="Microsoft.Jet.OLEDB.4.0"
	conn.Open Server.Mappath(vDatabase)
If UCase(Request.ServerVariables("HTTP_METHOD")) = "POST" Then
CCC=1
For Each Cate In Request.form("category")
	If Request.form("id")(CCC)<>"" and Isnumeric(Request.form("id")(CCC))=True Then
	Cate=fieldReplace(Trim(Cate))
	If Len(Cate)>20 Then Cate=Left(Cate,20)
	If Cate="" Then
		conn.execute("DELETE FROM [Categories] WHERE [ID]="&Request.form("id")(CCC))
	Else
		conn.execute("UPDATE [Categories] SET [Name]='"&Cate&"' WHERE [ID]="&Request.form("id")(CCC))
	End If
	End If
	CCC=CCC+1
Next
%><tr><td colspan="2">分类已更新。<a href="settings.asp?v=cate">点击这里刷新</a>。</td></tr><%
End If
%>
						<tr>
							<td colspan="2">
							<table>
						<tr><td>ID</td><td>名称</td></tr>
<%
	Set RSR=Server.CreateObject("ADODB.RecordSet")
	RSR.Open "Select * From [Categories]",conn,3,3
	Do While not RSR.EOF
	%>
	<tr><td><%=RSR("ID")%></td><td><input type="hidden" name="id" value="<%=RSR("ID")%>"><input name="category" value="<%=ReplaceXYZ(RSR("Name"))%>"></td></tr>
	<%
	RSR.movenext
	Loop
	RSR.Close
	Set RSR=nothing
	Set conn=nothing
%>
							</table>
							</td>
						</tr>
						<tr><td></td><td>要删除分类，请清空改分类名称。</td></tr>
						<tr><td></td><td><input type="submit" name="save" value="保存"></td></tr>
<%Case "delcom"
	Set conn=Server.CreateObject("ADODB.Connection")
	conn.Provider="Microsoft.Jet.OLEDB.4.0"
	conn.Open Server.Mappath(vDatabase)
	RqP=Request("p")
If Request.form("delete")<>"" and Request.form("p")<>"" and Isnumeric(Request.form("p"))=True Then
	conn.execute("DELETE FROM [Comments] WHERE [ID]="&Request.form("p"))
	RqP=""
End If
If RqP<>"" and isnumeric(RqP)=True Then
%>
<%
	Set RSR=Server.CreateObject("ADODB.RecordSet")
	RSR.Open "Select * From [Comments] WHERE [ID]="&Request("p"),conn,3,3
	%>
						<tr>
							<td>作者</td>
							<td><%=ReplaceXYZ(RSR("UserName"))%></td>
						</tr>
						<tr>
							<td>内容</td>
							<td><%=ReplaceXYZ(RSR("Content"))%></td>
						</tr>
						<tr>
							<td>日期</td>
							<td><%=RSR("DateCreated")%></td>
						</tr>
						<tr><td></td><td><input type="hidden" name="p" value="<%=RSR("ID")%>"><input type="submit" name="delete" value="删除评论"></td></tr>
<%
	RSR.Close
	Set RSR=nothing
	Set conn=nothing
Else
%><tr><td colspan="2">成功删除。</td></tr><%
End If
Case "users"
	Set conn=Server.CreateObject("ADODB.Connection")
	conn.Provider="Microsoft.Jet.OLEDB.4.0"
	conn.Open Server.Mappath(vDatabase)
If Request.form("delete")<>"" and Request.form("user")<>"" and Isnumeric(Request.form("user"))=True Then
	conn.execute("DELETE FROM [Users] WHERE [IsAdmin]=False AND [ID]="&Request.form("user"))
End If
If Request.form("promote")<>"" Then
	conn.execute("UPDATE [Users] SET [IsAdmin]=True WHERE [ID]="&Request.form("user"))
End If
%>
						<tr>
							<td>用户</td><td>
							<select name="user" id="usrsel" size="10" style="width:220px" ondblclick="alert(this.options[this.selectedIndex].text)">
<%
	Set RSR=Server.CreateObject("ADODB.RecordSet")
	RSR.Open "Select * From [Users] ORDER BY [IsAdmin] ASC",conn,3,3
	Do While not RSR.EOF
	%>
	<option value="<%=RSR("ID")%>"><%If RSR("IsAdmin") Then W "[管理员] "%><%=ReplaceXYZ(RSR("UserName"))%> (<%=ReplaceXYZ(RSR("Email"))%>)</option>
	<%
	RSR.movenext
	Loop
	RSR.Close
	Set RSR=nothing
	Set conn=nothing
%>
							</select>
							</td>
						</tr>
						<tr><td></td><td><input type="submit" name="delete" value="删除用户"> <input type="submit" name="promote" value="提升为管理员"></td></tr>
<%Case "cordb"%>
<%
If UCase(Request.ServerVariables("HTTP_METHOD")) = "POST" Then
    oldDB = Server.Mappath(vDatabase)
    bakDB = Server.Mappath(vDatabase)&"_backup"
    newDB = Server.Mappath(vDatabase)&"_compact"
%>
						<tr>
							<td align="center" colspan="2">压缩前数据库大小：
<%
set fs=Server.CreateObject("Scripting.FileSystemObject")
set f=fs.GetFile(Server.Mappath(vDatabase))
CompactBefore=f.Size
%><%=formatnumber(CompactBefore,0)%>字节<%
set f=nothing
set fs=nothing
%>
							</td>
						</tr>
<%
 
    Set FSO = CreateObject("Scripting.FileSystemObject") 
    FSO.CopyFile oldDB, bakDB, true 
    Set Engine = CreateObject("JRO.JetEngine") 
    prov = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" 
    Engine.CompactDatabase prov & OldDB, prov & newDB 
    set Engine = nothing 
    FSO.DeleteFile oldDB 
    FSO.MoveFile newDB, oldDB 
    set FSO = nothing  
End If
%>
						<tr>
							<td align="center" colspan="2">当前数据库大小：
<%
set fs=Server.CreateObject("Scripting.FileSystemObject")
set f=fs.GetFile(Server.Mappath(vDatabase))
%><%=formatnumber(f.Size,0)%>字节<%
If CompactBefore<>"" Then W " ("&INT((f.Size/CompactBefore)*100)&"%)"
set f=nothing
set fs=nothing
%>
							</td>
						</tr>
						<tr>
							<td>必读</td><td><textarea rows="3" style="width:220px;font-size:9pt" readonly>删除数据库内容并不能即时减少数据库大小，为了减少文件大小、重置表的自动编号，请定期点击“压缩和修复数据库”。过程将产生备份文件。</textarea></td>
						</tr>
						<tr><td></td><td><input id="btnRegister" type="submit" value="压缩和修复数据库"></td></tr>
<%End Select%>
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