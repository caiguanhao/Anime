<!-- #include file="@universal.asp" --><%
	If Session("UserAdmin")="" OR Session("UserName")="" Then Alert "��¼��ʱ��δ��¼��Ȩ�޲��㣬���¼�������","./"
	Set conn=Server.CreateObject("ADODB.Connection")
	conn.Provider="Microsoft.Jet.OLEDB.4.0"
	conn.Open Server.Mappath(vDatabase)
	Set RSR=Server.CreateObject("ADODB.RecordSet")
	If Request("p")<>"" and Isnumeric(Request("p"))=True Then
		RSR.Open "Select * From [Data] WHERE [ID]="&Request("p"),conn,3,3
		If not RSR.EOF Then
			ssID=RSR("ID")
			ssTitle=ReplaceXYZ(RSR("Title"))
			ssCate=ReplaceXYZ(RSR("Category"))
			ssContent=ReplaceXYZ(RSR("Content"))
			ssCategory=(RSR("Category"))
			ssPublished=(RSR("Published"))
		Else
			R "./"
		End If
		RSR.Close
		ssCategoryX=""
		ssCategory=Split(ssCategory&",",",")
		For k=0 to UBound(ssCategory)
			If ssCategory(k)<>"" and Isnumeric(ssCategory(k))=True Then
				RSR.Open "Select * From [Categories] WHERE [ID]="&ssCategory(k),conn,3,3
				If not RSR.EOF Then
					RSR.Close
					RSR.Open "Select * From [Categories] ORDER BY [ID] ASC",conn,3,3
					CateNO=0
					Do while not RSR.EOF
						If RSR("ID")=INT(ssCategory(k)) Then
							ssCategoryX=ssCategoryX&CateNO&","
						End If
						CateNO=CateNO+1
						RSR.Movenext
					Loop
				End IF
				RSR.Close
			End If
		Next
		ssCategory=ssCategoryX
	ElseIf Request("p")<>"" and Isnumeric(Request("p"))=False Then
		R "./"
	End If
	Set RSR=nothing
	Set conn=nothing
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" rev="stylesheet" type="text/css" href="images/style.css">
<title>���ݹ���</title>
<script type="text/javascript" src="js/time.js"></script>
<script type="text/javascript">
var xmlhttp;
if (window.XMLHttpRequest){xmlhttp=new XMLHttpRequest();}
else{xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");}

function loadCategory(str) {
var xmlRT, CateInput, CateSeled;
CateInput="";
CateSeled="<%=ssCategory%>";
xmlhttp.onreadystatechange=function(){
if (xmlhttp.readyState==4 && xmlhttp.status==200){
  xmlRT=unescape(xmlhttp.responseText).split(',');
  for (i=0;i<=xmlRT.length-1;i++){if(xmlRT[i]!=""){CateInput+=("<span style='white-space: nowrap'><input type='checkbox' name='category' value='"+escape(xmlRT[i])+"' id='Cate"+i+"'><label for='Cate"+i+"'>"+xmlRT[i]+"</label></span> ")}}
  for (i=0;i<=xmlRT.length-1;i++){if(document.getElementById('Cate'+i)){if(document.getElementById('Cate'+i).checked){CateSeled+=(i+",")}}}
  if(CateInput!=""){document.getElementById('fldCategory').innerHTML=CateInput}
  if(CateSeled!=""){CateSeled=CateSeled.split(',');for (i=0;i<CateSeled.length-1;i++){document.getElementById('Cate'+CateSeled[i]).checked=true}}
}}
xmlhttp.open("GET","data.asp?sid="+Math.random(),true);
xmlhttp.send();
}

function createCategory() {
var NewCN;
NewCN=prompt('�������µ��������','').replace(/(^\s*)|(\s*$)/g, "");
if(NewCN!=null&&NewCN!=""){
xmlhttp.onreadystatechange=function(){
if (xmlhttp.readyState==4 && xmlhttp.status==200){
	if (xmlhttp.responseText=='SUCCESSFULLYCREATED'){loadCategory('2');}
	else if (xmlhttp.responseText=='CATEGORYALREADYEXISTS'){alert('����Ѵ���');}
	else {alert('����δ֪����');}
}else if(xmlhttp.readyState==4){alert('����δ֪����(STATUS:'+xmlhttp.status+')');}}
NewCN=("NewCategoryName="+escape(NewCN));
xmlhttp.open("POST","update.asp",true);
xmlhttp.setRequestHeader("Content-Length", NewCN.length);
xmlhttp.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
xmlhttp.send(NewCN);
}}


function fillTime() {
document.getElementById('timeE').innerHTML=elapsedTime(fixDate('<%=Now%>'));
}


window.onload=function(){loadCategory('2');<%If Request("m")<>"" Then%>fillTime();setInterval('fillTime()',1000);<%End If%>}


</script>
</head>

<body>
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
		<table width="800">
		<%If ssID<>"" Then%>
	<tr><td>
	<div class="roundedcornr_box_406176">
	   <div class="roundedcornr_top_406176"><div></div></div>
	      <div class="roundedcornr_content_406176" style="height:14px;">
	      <div style="float:left">
	      <small>
	      <%Select Case Request("m")%>
	      <%Case "csuccess"%>
	      <b>�ɹ��������ݡ�</b>
	      <%Case "esuccess"%>
	      <b>�ɹ��������ݡ�</b>
	      <%End Select%>
	      ״̬��<%If ssPublished=False Then%>��δ������<%Else%>�ѷ�����<a href="./?p=<%=ssID%>">�鿴</a><%End If%></div><div style="float:right;font-size:10pt" id="timeE"></div>
	      </small></div>
	   <div class="roundedcornr_bottom_406176"><div></div></div>
	</div>
	</td></tr>
	<%End If%>
	<tr>
		<td>
<div class="roundedcornr_box_123135">
		   <div class="roundedcornr_top_123135"><div></div></div>
		      <div class="roundedcornr_content_123135">
				<table cellpadding="0" cellspacing="0">
						<tr>
							<td>
							<form id="pub" style="margin:0px" action="save.asp<%If ssID<>"" Then W "?p="&ssID%>" method="post"><input type="hidden" id="pubordraft" name="pubordraft" value="1" disabled>
							<table cellpadding="0" cellspacing="0">
							<tr>
								<td colspan="2"><h2><%If ssID="" Then%>������Դ<%Else%>�༭��Դ<%End If%><!-- #include file="control.asp" --></h2></td>
							</tr>
							<tr valign="top">
								<td nowrap>���⣺</td>
								<td width="500"><input name="title" value="<%=ssTitle%>" style="font-size:20pt;width:500px" onkeydown="if(event.keyCode==13){return false}"></td>
							</tr>
							<tr valign="top">
								<td><div style="padding:5px 0 5px 0;">���</div></td>
								<td width="500"><div style="padding:5px 0 5px 0;" onkeydown="if(event.keyCode==13){return false}"><span id="fldCategory">
									�������������ȷ�����������JavaScript��</span><span style='white-space: nowrap'><a href="javascript:void(0);" onclick="createCategory()">+�������</a></span></div></td>
							</tr>
							<tr valign="top">
								<td>���ܣ�</td>
								<td width="500"><textarea id="fldIntro" name="content" rows="20" cols="30" style="font-size:10pt;width:500px"><%=ssContent%></textarea>
								<div style="padding:0 0 10px 0"><em><small>
									֧�ֲ�����HTML���Զ�ʶ��URL���ſᡢ������ed2k����</small></em></span></td>
							</tr>
							</table>
							</form>
							</td>
							<td valign="bottom" width="250">
								<div id="uploadfilesinput">
							    <form style="margin:0 0 10px 10px" name="frmSend" method="POST" enctype="multipart/form-data" accept-charset="utf-8" action="upload.asp" target="uploadframe">
								<div id="fileinput"><input name="attach1" id="attach1" type="file" size="20" onchange="document.getElementById('txtwait').style.display='';document.getElementById('fileinput').style.display='none';this.form.submit();"></div>
								<div id="txtwait" style="display:none;text-align:center"><img border="0" src="images/loading.gif" width="32" height="32"><br>
									��ȴ��ļ��ϴ���ɡ�<a href="javascript:void(0);" onclick="window.open('about:blank','uploadframe');document.getElementById('attach1').value='';document.getElementById('txtwait').style.display='none';document.getElementById('fileinput').style.display='';">ȡ��</a></div>
								<em><small>�����ϴ���С����20MB���ļ�</small></em>
								</form>
								<div id="uploadfrm" style="display:none"><iframe name="uploadframe" width="500" height="50"></iframe></div>
								</div>
							</td>
						</tr>
						<tr><td height="30"></td></tr>
						<tr><td align="center" valign="top" height="50" colspan="2" bgcolor="#C3DAF9">
						<div style="float:left;padding-left:50px"><input type="submit" value="Ԥ��" name="publish" onclick="document.getElementById('pub').target='_blank';document.getElementById('pub').action='index.asp';document.getElementById('pub').submit();"></div>
						<div style="float:right;padding-right:50px"><input type="submit" value="<%If ssID="" Then%>��������<%Else%>���沢��������<%End If%>" name="publish" style="font-weight:bold;" onclick="document.getElementById('pub').target='';document.getElementById('pub').action='save.asp<%If ssID<>"" Then W "?p="&ssID%>';document.getElementById('pubordraft').disabled=false;document.getElementById('pub').submit();"> <input type="submit" value="<%If ssID="" Then%>��Ϊ�ݸ�<%Else%>���浫������<%End If%>" onclick="document.getElementById('pub').target='';document.getElementById('pub').action='save.asp<%If ssID<>"" Then W "?p="&ssID%>';document.getElementById('pub').submit();"></div>
						</td></tr>
						</table>
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