<!-- #include file="@universal.asp" --><%
	NewTitle=Trim(fieldReplace((Request.form("title"))))
	NewContent=Trim(fieldReplace((Request.form("content"))))
	If NewTitle="" and NewContent="" Then
		Set conn=Server.CreateObject("ADODB.Connection")
		conn.Provider="Microsoft.Jet.OLEDB.4.0"
		conn.Open Server.Mappath(vDatabase)
		Set RSR=Server.CreateObject("ADODB.RecordSet")
		If Request("f")<>"" Then
			RSR.Open "Select * From [Data] WHERE Instr([Content],'"&fieldReplace(Request("f"))&"')<>0",conn,3,3
			If not RSR.EOF Then
				R "./?p="&RSR("ID")
			Else
				R "./"
			End If
			RSR.Close
		End If
		If Request("p")<>"" and Isnumeric(Request("p"))=True Then
			RSR.Open "Select * From [Data] WHERE [ID]="&Request("p"),conn,3,3
			If not RSR.EOF Then
				ssID=RSR("ID")
				ssPublished=(RSR("Published"))
				If ssPublished=False Then
					ssTitle="尚未发布。"
					ssContent="你所查看的内容尚未发布或已被删除。"
				Else
					conn.execute("UPDATE [Data] SET [Views]="&(RSR("Views")+1)&" WHERE [ID]="&RSR("ID"))
					ssTitle=ReplaceXYZ(RSR("Title"))
					ssCate=ReplaceXYZ(RSR("Category"))
					ssContent=ReplaceXYZ(RSR("Content"))
					ssDate=RSR("DateCreated")
					ssCategory=(RSR("Category"))
					ssCategory=Split(ssCategory&",",",")
					ssCategoryX=""
					For k=0 to UBound(ssCategory)
						If ssCategory(k)<>"" and Isnumeric(ssCategory(k))=True Then
							RSR.Close
							RSR.Open "Select * From [Categories] WHERE [ID]="&ssCategory(k),conn,3,3
							If not RSR.EOF Then
								ssCategoryX=ssCategoryX&"<a href=""./?view=latest&c="&RSR("ID")&""">"&ReplaceXYZ(RSR("Name"))&"</a>、"
							End IF
						End If
					Next
					If Right(ssCategoryX,1)="、" Then ssCategoryX=Left(ssCategoryX,Len(ssCategoryX)-1)
					ssCategory=ssCategoryX
				End If
			Else
				R "./"
			End If
			RSR.Close
		ElseIf Request("p")<>"" and Isnumeric(Request("p"))=False Then
			R "./"
		End If
		Set RSR=nothing
		Set conn=nothing
	Else
		ssID=-1
		ssTitle=NewTitle
		ssContent=NewContent
	End If
	
Function ViewsCon(numb)
ViewsCon=numb
Select Case True
Case numb<=10
Case numb<1000: ViewsCon=int(numb/10)*10&"+"
Case else: ViewsCon="1000+"
End Select
End Function
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" rev="stylesheet" type="text/css" href="images/style.css">
<title><%If ssTitle<>"" Then W ssTitle&" - "%>动漫小组</title>
<link rel="stylesheet" type="text/css" href="css/slideshow.css" media="screen" />
<script type="text/javascript" src="js/time.js"></script>
<script>
var xmlhttp;
if (window.XMLHttpRequest){xmlhttp=new XMLHttpRequest();}
else{xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");}

function loadComment(ssID,ssPAGE) {
var AllComm, SiComm, CommOut, NextPage;
CommOut='';
xmlhttp.onreadystatechange=function(){
if (xmlhttp.readyState==4 && xmlhttp.status==200){
  AllComm=(unescape(xmlhttp.responseText)).split('<Y>');
  for (i=0;i<=AllComm.length-1;i++){
    if(AllComm[i].indexOf('<X>')!=-1){SiComm=AllComm[i].split('<X>');CommOut+='<p>'+SiComm[0]+'<%If Session("UserAdmin")<>"" AND Session("UserName")<>"" Then W " [<a href=""./settings.asp?v=delcom&p='+SiComm[3]+'"">删除</a>]"%><br><span id="commdate'+i+'" title="'+SiComm[2]+'">'+elapsedTime(fixDate(SiComm[2]))+'</span>&nbsp;by @<a href="javascript:atme(\''+SiComm[1]+'\')">'+SiComm[1]+'</a></p>'}
    else if(/\<N(\d*)\>/.test(AllComm[i])) {NextPage=parseInt(AllComm[i].replace(/\<N(\d*)\>/,"$1"));CommOut+='<p>';if(NextPage>2){CommOut+='[<a href="javascript:loadComment('+ssID+','+(NextPage-2)+')">上一页</a>]'};if(!/\<END\>/.test(AllComm[i])){CommOut+='[<a href="javascript:loadComment('+ssID+','+NextPage+')">下一页</a>]'};CommOut+='</p>';}
  }
  if(CommOut){document.getElementById('allcomments').innerHTML=CommOut;
  if(document.getElementById('allcomments').scrollHeight>300){document.getElementById('allcomments').style.height='300px'}
  if(document.getElementById('mycomm')){document.getElementById('mycomm').value='';document.getElementById('wdcount').innerHTML='140';}
  };
}}
xmlhttp.open("GET","data.asp?t=comment&p="+ssID+"&page="+ssPAGE+"&sid="+Math.random(),true);
xmlhttp.send();
fillTime();
}

var fillFirstTime=0;
function fillTime() {
for (i=0;i<=5;i++){
if(document.getElementById('commdate'+i)){document.getElementById('commdate'+i).innerHTML=elapsedTime(fixDate(document.getElementById('commdate'+i).title));}else{escape}
}
if(fillFirstTime==0){setInterval('fillTime()',1000);}
fillFirstTime=1;
}

function newComment(ssID) {
var NewCN;
NewCN=document.getElementById('mycomm').value;
if(NewCN.length<5||NewCN.length>140){alert('评论长度过短或过长，最短不少于5个字符，最长140个字符。');return false;}
NewCN=("p="+ssID+"&comment="+escape(NewCN));
xmlhttp.onreadystatechange=function(){
if (xmlhttp.readyState==4 && xmlhttp.status==200){
	if (xmlhttp.responseText=='SUCCESSFULLYCREATED'){loadComment(ssID,1);}
	else if (xmlhttp.responseText=='COMMENTALREADYEXISTS'){alert('请不要重复评论。');}
	else {alert('发生未知错误。');}
}else if(xmlhttp.readyState==4){alert('发生未知错误。(STATUS:'+xmlhttp.status+')');}}
xmlhttp.open("POST","newcomment.asp",true);
xmlhttp.setRequestHeader("Content-Length", NewCN.length);
xmlhttp.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
xmlhttp.send(NewCN);
}

</script>
<script type="text/javascript" src="js/mootools.js"></script>
<script type="text/javascript" src="js/slideshow.js"></script>
<script type="text/javascript">		
//<![CDATA[
  window.addEvent('domready', function(){
    var data = {
<%
AllImageFiles=""
set fs=Server.CreateObject("Scripting.FileSystemObject")
set fo=fs.GetFolder(Server.Mappath("files/201103"))

for each x in fo.files
Select Case LCASE(Split(x.Name&".",".")(UBound(Split(x.Name&".","."))-1))
Case "jpg","jpeg","jpe","png","gif"
AllImageFiles=AllImageFiles&x.Name&vbCrlf
End Select
next
AllImageFiles=Split(vbCrlf&AllImageFiles,vbCrlf)
For i=1 to UBound(AllImageFiles)-1
%>
'201103/<%=AllImageFiles(UBound(AllImageFiles)-i)%>': { caption: '', href: './?f=<%=AllImageFiles(UBound(AllImageFiles)-i)%>' },
<%
Next

set fo=nothing
set fs=nothing
%>
'welcome.png': { }
    };
    var myShow = new Slideshow('show', data, {controller: false, height: 300, hu: 'files/', thumbnails: false, width: 710});
  });
//]]>
</script>
<script type="text/javascript">
function resizeImages(){
  maxSize = 600;
  for (x=0;x<document.images.length;x++){
     iHeight = document.images[x].height;
     iWidth = document.images[x].width;

     if (iHeight > iWidth){
        sizeGuide = iHeight;
        size2 = iWidth;
     }else{
        sizeGuide = iWidth;
        size2 = iHeight;
     }
     if (sizeGuide > maxSize){
        sizeRatio = sizeGuide / size2;
        newSize1 = maxSize;
        newSize2 = newSize1 / sizeRatio;
        if (iHeight > iWidth){
           document.images[x].height = newSize1;
           document.images[x].width = newSize2;
        }else{
           document.images[x].width = newSize1;
           document.images[x].height = newSize2;
        }
     }
  }
}

window.onload=function(){resizeImages();loadComment(<%If ssID="" Then W 0 Else W ssID%>,1);};

function textCounter(field,cntfield,maxlimit) {
if (field.value.length > maxlimit){field.value = field.value.substring(0, maxlimit);}
else {cntfield.innerHTML = maxlimit - field.value.length;}
}

function atme(mename) {
if(mename.toLowerCase()=='<%=LCASE(Session("UserName"))%>'){alert('请不要回复自己。')}else{document.getElementById('mycomm').value='回复@'+mename+': '+document.getElementById('mycomm').value}
}

</script>
</head>

<body>

<div align="center">
	<table cellpadding="0" cellspacing="0" width="1000">
		<tr>
			<td>
			<!-- #include file="menu.asp" --></td>
		</tr>
		<tr>
			<td height="10"></td>
		</tr>
		<tr valign="top">
			<td>
			<table cellpadding="0" cellspacing="0" width="100%">
				<tr valign="top">
					<td width="741">
<%Select Case LCASE(Request("view"))%>
<%Case "latest"

ssC=Request("c")

Set conn=Server.CreateObject("ADODB.Connection")
conn.Provider="Microsoft.Jet.OLEDB.4.0"
conn.Open Server.Mappath(vDatabase)
Set RSR=Server.CreateObject("ADODB.RecordSet")
If ssC<>"" and Isnumeric(ssC)=True then
RSR.Open "Select * From [Categories] WHERE [ID]="&Request("c"),conn,3,3
If Not RSR.EOF Then
ssC=RSR("ID")
ssC2=" - "&ReplaceXYZ(RSR("Name"))
ssC3=" AND Instr([Category],',"&ssC&",')<>0"
Else
ssC=""
ssC2=""
ssC3=""
End If
RSR.Close
Else
ssC=""
ssC2=""
ssC3=""
End If

If Request("q")<>"" Then
	ssQ=Left(Request("q"),50)
	ssC2=ssC2&" - "&ssQ
	ssC3=ssC3&" AND (Instr([Content],'"&ssQ&"')<>0 OR Instr([Title],'"&ssQ&"')<>0)"
End If
%>
					<table cellpadding="0" cellspacing="0" width="100%">
					<tr><td colspan="3">
						<div class="roundedcornr_box_123135"><div class="roundedcornr_top_123135"><div></div></div>
						<div class="roundedcornr_content_123135">
						<table style="font-size:10pt">
						<tr>
							<td>
							<h2 style="margin-bottom:0px">最新内容<%=ssC2%></h2>
							</td>
						</tr>
						<tr><td height="20"></td></tr>
<%
CurrentPage=Request("page")
If CurrentPage<>"" and Isnumeric(CurrentPage)=True Then CurrentPage=abs(int(CurrentPage)) Else CurrentPage=1
ItemsPerPage=10
Set RSR2=Server.CreateObject("ADODB.RecordSet")
RSR.Open "Select * From [Data] WHERE [Published]=True"&ssC3&" ORDER BY [ID] DESC",conn,3,3
If not RSR.EOF Then
RSR.Move (CurrentPage-1)*ItemsPerPage
CCount=1
Do while not RSR.EOF and CCount<=ItemsPerPage
	ssTitle=ReplaceXYZ(RSR("Title"))
	ssCate=ReplaceXYZ(RSR("Category"))
	ssContent=ReplaceXYZ(RSR("Content"))
	ssDate=RSR("DateCreated")
	ssCategory=(RSR("Category"))
	ssCategory=Split(ssCategory&",",",")
	ssCategoryX=""
	For k=0 to UBound(ssCategory)
		If ssCategory(k)<>"" and Isnumeric(ssCategory(k))=True Then
			RSR2.Open "Select * From [Categories] WHERE [ID]="&ssCategory(k),conn,3,3
			If not RSR2.EOF Then
				ssCategoryX=ssCategoryX&"<a href=""./?view=latest&c="&RSR2("ID")&""">"&ReplaceXYZ(RSR2("Name"))&"</a>、"
			End IF
			RSR2.Close
		End If
	Next
	If Right(ssCategoryX,1)="、" Then ssCategoryX=Left(ssCategoryX,Len(ssCategoryX)-1)
	ssCategory=ssCategoryX
%>
						<tr>
							<td>
							<h2 style="margin-bottom:0px"><a href="./?p=<%=RSR("ID")%>"><%=ssTitle%></a></h2>
							<small><%If ssDate<>"" Then%>于<%=ssDate%>发布。<%End If%><%If ssCategory<>"" Then%>分类：<%=ssCategory%>。<%End If%><%If ssID>0 Then%><a href="addoredit.asp?p=<%=ssID%>">编辑</a><%End If%></small>
							<p><%=Replace(Server.HTMLEncode(ssContent),vbCrlf,"<br>")%></p>
							</td>
						</tr>
						<tr><td height="20"></td></tr>
<%
	CCount=CCount+1
	RSR.Movenext
Loop%>
						<tr><td><%If CurrentPage>1 Then%>[<a href="./?view=latest&c=<%=ssC%>&q=<%=Server.URLEncode(ssQ)%>&page=<%=CurrentPage-1%>">上一页</a>]<%End If%><%If RSR.RecordCount>ItemsPerPage and (CurrentPage)*ItemsPerPage<RSR.RecordCount Then%>[<a href="./?view=latest&c=<%=ssC%>&q=<%=Server.URLEncode(ssQ)%>&page=<%=CurrentPage+1%>">下一页</a>]<%End If%></td></tr>
<%
Else
%><tr><td>没有内容。</td></tr><%
End If
Set RSR=nothing
Set RSR2=nothing
Set conn=nothing
%>
						</table>
						</div>
						<div class="roundedcornr_bottom_123135">
							<div>
							</div>
						</div>
					</div>
					</td></tr>
					</table>
<%Case Else%>
<%If ssID="" Then%>
					<table cellpadding="0" cellspacing="0">
					<tr><td colspan="3">
						<div class="roundedcornr_box_123135"><div class="roundedcornr_top_123135"><div></div></div>
						<div class="roundedcornr_content_123135">
						<table style="font-size:10pt">
						<tr>
							<td>
  <div id="show" class="slideshow">
  </div>
							</td>
						</tr>
						</table>
						</div>
						<div class="roundedcornr_bottom_123135">
							<div>
							</div>
						</div>
					</div>
					</td></tr>
					<tr><td height="10"></td></tr>
					<tr>
						<td width="365">
						<div class="roundedcornr_box_123135"><div class="roundedcornr_top_123135"><div></div></div>
						<div class="roundedcornr_content_123135">
						<table style="font-size:10pt">
<%
Set conn=Server.CreateObject("ADODB.Connection")
conn.Provider="Microsoft.Jet.OLEDB.4.0"
conn.Open Server.Mappath(vDatabase)
Set RSR=Server.CreateObject("ADODB.RecordSet")
RSR.Open "Select Top 8 * From [Data] WHERE [Published]=True ORDER BY [ID] DESC",conn,3,3
Do while not RSR.EOF
%>
						<tr>
							<td><a href="./?p=<%=(RSR("ID"))%>"><%=ReplaceXYZ(RSR("Title"))%></a>&nbsp;&nbsp;<i><small><%=DateValue(RSR("DateCreated"))%></small></i></td>
						</tr>
<%
	RSR.Movenext
Loop
Set RSR=nothing
Set conn=nothing
%>
						</table>
						</div>
						<div class="roundedcornr_bottom_123135">
							<div>
							</div>
						</div>
					</div>
						</td>
						<td width="10"></td>
						<td width="365">
						<div class="roundedcornr_box_123135"><div class="roundedcornr_top_123135"><div></div></div>
						<div class="roundedcornr_content_123135">
						<table style="font-size:10pt">
<%
Set conn=Server.CreateObject("ADODB.Connection")
conn.Provider="Microsoft.Jet.OLEDB.4.0"
conn.Open Server.Mappath(vDatabase)
Set RSR=Server.CreateObject("ADODB.RecordSet")
RSR.Open "Select Top 8 * From [Data] WHERE [Published]=True ORDER BY [Views] DESC",conn,3,3
Do while not RSR.EOF
%>
						<tr>
							<td><a href="./?p=<%=(RSR("ID"))%>"><%=ReplaceXYZ(RSR("Title"))%></a>&nbsp;&nbsp;<i><small><%=ViewsCon(RSR("Views"))%>人次浏览</small></i></td>
						</tr>
<%
	RSR.Movenext
Loop
Set RSR=nothing
Set conn=nothing
%>
						</table>
						</div>
						<div class="roundedcornr_bottom_123135">
							<div>
							</div>
						</div>
					</div>
						</td>
					</tr>
					</table>
<%Else%>
						<div class="roundedcornr_box_123135"><div class="roundedcornr_top_123135"><div></div></div>
						<div class="roundedcornr_content_123135">
					<table cellpadding="0" cellspacing="0">
					<tr>
						<td>
						<h2 style="margin-bottom:0px"><%=ssTitle%></h2>
						<small><%If ssDate<>"" Then%>于<%=ssDate%>发布。<%End If%><%If ssCategory<>"" Then%>分类：<%=ssCategory%>。<%End If%><%If ssID>0 Then%><a href="addoredit.asp?p=<%=ssID%>">编辑</a><%End If%></small>
						<p><%=ConCon(ssContent)%></p>
						</td>
					</tr>
					</table>
						</div>
						<div class="roundedcornr_bottom_123135">
							<div>
							</div>
						</div>
					</div>
<%End If%>
<%End Select%>
					</td>
					<td width="10"></td>
					<td>
					<table cellpadding="0" cellspacing="0" width="100%">
					<tr>
						<td>
					<div class="roundedcornr_box_406176">
					   <div class="roundedcornr_top_406176"><div></div></div>
					      <div class="roundedcornr_content_406176">
<%If Session("UserName")="" Then%>
					      用户登录
					      <form action="login.asp<%If ssID<>"" Then W "?p="&ssID%>" method="post" style="margin:0px">
					      <table style="font-size:10pt">
							<tr>
								<td>用户名</td><td><input name="username" size="15"></td><td><input type="submit" value="登录"></td>
							</tr>
							<tr>
								<td>密码</td><td><input name="password" type="password" size="15"></td><td align="center"><a href="register.asp">注册</a></td>
							</tr>
							</table>
					      </form>
<%Else%>
					      欢迎回来，<%=Session("UserName")%>！<a href="./login.asp?logout=1">退出</a><br><%If Session("UserAdmin")<>"" Then%>
					      <!-- #include file="control.asp" -->
					      <%End If%><%End If%>
					      </div>
					   <div class="roundedcornr_bottom_406176"><div></div></div>
					</div>
						</td>
					</tr>
					<tr>
						<td height="10"></td>
					</tr>
					<tr>
						<td>
					<div class="roundedcornr_box_406176">
					   <div class="roundedcornr_top_406176"><div></div></div>
					      <div class="roundedcornr_content_406176">
					      最新评论
					      <div id="allcomments" style="font-size:9pt;margin-top:3px;margin-bottom:3px;max-height:300px;overflow-y:auto">尚没有人发表过评论。</div>
					      <table style="font-size:10pt">
							<tr>
								<td>
								<textarea style="font-size:9pt;width:200px;height:22px" onclick="this.style.height='80px';" id="mycomm" name="comment" wrap="physical" onKeyDown="textCounter(document.getElementById('mycomm'),document.getElementById('wdcount'),140)"
onKeyUp="textCounter(document.getElementById('mycomm'),document.getElementById('wdcount'),140)" onchange="this.value=this.value.replace(/<.*?>/g, '');"></textarea>
								</td>
							</tr>
							<tr><td><input type="button" onclick="newComment(<%If ssID<>"" and isnumeric(ssID)=true Then W ssID Else W 0%>);"<%If Session("UserName")<>"" Then%> value="发表"<%Else%> value="请登录后发表评论" disabled<%End If%>> <a href="javascript:loadComment(<%If ssID="" Then W 0 Else W ssID%>,1)" title="点击重新载入评论"><span id="wdcount">140</span></a></td></tr>
							</table>
					      </div>
					   <div class="roundedcornr_bottom_406176"><div></div></div>
					</div>
						</td>
					</tr>
					</table>
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
