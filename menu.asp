<table cellpadding="0" cellspacing="0" width="100%">
	<tr>
		<td width="150" align="center"><a href="./"><img alt="动漫小组" src="images/logo.png" width="105" height="35"></a></td>
		<td>
		<div class="roundedcornr_box_123135">
		   <div style="height: 10px;"><div></div></div>
		      <div class="roundedcornr_content_123135" style="height:25px">
		       <a href="./">首页</a>&nbsp;&nbsp;<a href="./?view=latest">最新</a><%
	Set conn=Server.CreateObject("ADODB.Connection")
	conn.Provider="Microsoft.Jet.OLEDB.4.0"
	conn.Open Server.Mappath(vDatabase)
	Set RSR=Server.CreateObject("ADODB.RecordSet")
	RSR.Open "Select * From [Categories]",conn,3,3
	If not RSR.EOF Then
		Do While Not RSR.EOF
%>&nbsp;&nbsp;<a href="./?view=latest&c=<%=RSR("ID")%>"><%=ReplaceXYZ(RSR("Name"))%></a><%
		RSR.MoveNext
		Loop
	End If
	RSR.Close
	Set RSR=nothing
	Set conn=nothing
%>
		       </div>
		   <div class="roundedcornr_bottom_123135"><div></div></div>
		</div>
		</td>
		<td width="10"></td>
		<td width="250" align="center">
		<form action="./" method="get" style="margin:0px"><input type="hidden" name="view" value="latest"><input type="hidden" name="c" value="<%=Request("c")%>">
		<div class="roundedcornr_box_810239">
		   <div style="height: 10px;"><div></div></div>
		      <div class="roundedcornr_content_810239" style="height:25px">
		       <input style="width:160px" name="q" value="<%=Replace(Request("q"),"""","")%>"> <input type="submit" style="width:40px;" value="搜索">
		      </div>
		   <div class="roundedcornr_bottom_810239"><div></div></div>
		</div>
		</form>
		</td>
	</tr>
</table>