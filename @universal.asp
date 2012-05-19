<%
Session.CodePage = 936
Session.Timeout=1440
vDatabase="database/data.mdb"

Function W(STR)
Response.Write STR
End Function

Function R(STR)
Response.Redirect STR
End Function

Function Alert(STR,STR2)
Response.Write "<script>alert('"&STR&"');"
If STR2="" Then Response.Write "history.back();" Else Response.Write "location.href='"&STR2&"';"
Response.Write "</script>"
Response.end
End Function
                          
Function fieldReplace(item)
 item = replace(item, chr(34), "{#34}")
 item = replace(item, chr(39), "{#39}")
 item = replace(item, "%", "%")
 item = replace(item, "ゴ", "{#JP01}")
 item = replace(item, "ガ", "{#JP02}")
 item = replace(item, "ギ", "{#JP03}")
 item = replace(item, "グ", "{#JP04}")
 item = replace(item, "ゲ", "{#JP05}")
 item = replace(item, "ザ", "{#JP06}")
 item = replace(item, "ジ", "{#JP07}")
 item = replace(item, "ズ", "{#JP08}")
 item = replace(item, "ヅ", "{#JP09}")
 item = replace(item, "デ", "{#JP10}")
 item = replace(item, "ド", "{#JP11}")
 item = replace(item, "ポ", "{#JP12}")
 item = replace(item, "ベ", "{#JP13}")
 item = replace(item, "プ", "{#JP14}")
 item = replace(item, "ビ", "{#JP15}")
 item = replace(item, "パ", "{#JP16}")
 item = replace(item, "ヴ", "{#JP17}")
 item = replace(item, "ボ", "{#JP18}")
 item = replace(item, "ペ", "{#JP19}")
 item = replace(item, "ブ", "{#JP20}")
 item = replace(item, "ピ", "{#JP21}")
 item = replace(item, "バ", "{#JP22}")
 item = replace(item, "ヂ", "{#JP23}")
 item = replace(item, "ダ", "{#JP24}")
 item = replace(item, "ゾ", "{#JP25}")
 item = replace(item, "ゼ", "{#JP26}")
 fieldReplace = Trim(item)
End Function

Function ReplaceXYZ(item)
If len(item)=0 or isnull(item)=true Then
	ReplaceXYZ=""
	Exit Function
Else
	item = replace(item, "{#34}", chr(34))
	item = replace(item, "{#39}", chr(39))
	item = replace(item, "%", "%")
	item = replace(item, "{#JP01}", "ゴ")
	item = replace(item, "{#JP02}", "ガ")
	item = replace(item, "{#JP03}", "ギ")
	item = replace(item, "{#JP04}", "グ")
	item = replace(item, "{#JP05}", "ゲ")
	item = replace(item, "{#JP06}", "ザ")
	item = replace(item, "{#JP07}", "ジ")
	item = replace(item, "{#JP08}", "ズ")
	item = replace(item, "{#JP09}", "ヅ")
	item = replace(item, "{#JP10}", "デ")
	item = replace(item, "{#JP11}", "ド")
	item = replace(item, "{#JP12}", "ポ")
	item = replace(item, "{#JP13}", "ベ")
	item = replace(item, "{#JP14}", "プ")
	item = replace(item, "{#JP15}", "ビ")
	item = replace(item, "{#JP16}", "パ")
	item = replace(item, "{#JP17}", "ヴ")
	item = replace(item, "{#JP18}", "ボ")
	item = replace(item, "{#JP19}", "ペ")
	item = replace(item, "{#JP20}", "ブ")
	item = replace(item, "{#JP21}", "ピ")
	item = replace(item, "{#JP22}", "バ")
	item = replace(item, "{#JP23}", "ヂ")
	item = replace(item, "{#JP24}", "ダ")
	item = replace(item, "{#JP25}", "ゾ")
	item = replace(item, "{#JP26}", "ゼ")
 	ReplaceXYZ = Trim(item)
End If
End Function

Sub Include(a_sFileName)
Dim sASPCode,sFileName,sContent
Dim oStream,oReg
Dim iAspStart,iAspEnd

sFileName = Trim(a_sFileName)

Set oStream = server.createobject("ADODB.Stream") 
oStream.Charset = "utf-8"
oStream.Open

oStream.LoadFromFile Server.Mappath(sFileName)

sContent=oStream.ReadText
oStream.Close
Set oStream = Nothing

Set oReg=new RegExp
oReg.pattern="^\s*="
iAspEnd=1
iAspStart=inStr(iAspEnd,sContent,"<%")+2
Do while iAspStart>iAspEnd+1
iAspEnd = InStr(iAspStart,sContent,"%\>")+2
Execute(oReg.Replace(Mid(sContent,iAspStart,iAspEnd-iAspStart-2),"Response.Write "))
iAspStart=inStr(iAspEnd,sContent,"<%")+2
Loop
Set oReg=nothing
End Sub

Function URLDecode(strIn)
URLDecode = ""
Dim sl: sl = 1
Dim tl: tl = 1
Dim key: key = "%"
Dim kl: kl = Len(key)
sl = InStr(sl, strIn, key, 1)
Do While sl>0
	If (tl=1 And sl<>1) Or tl<sl Then
		URLDecode = URLDecode & Mid(strIn, tl, sl-tl)
	End If
	Dim hh, hi, hl
	Dim a
	Select Case UCase(Mid(strIn, sl+kl, 1))
		Case "U":
			a = Mid(strIn, sl+kl+1, 4)
			URLDecode = URLDecode & ChrW("&H" & a)
			sl = sl + 6
		Case "E":
			hh = Mid(strIn, sl+kl, 2)
			a = Int("&H" & hh)
			If Abs(a)<128 Then
				sl = sl + 3
				URLDecode = URLDecode & Chr(a)
			Else
				hi = Mid(strIn, sl+3+kl, 2)
				hl = Mid(strIn, sl+6+kl, 2)
				a = ("&H" & hh And &H0F) * 2 ^12 Or ("&H" & hi And &H3F) * 2 ^ 6 Or ("&H" & hl And &H3F)
				If a<0 Then a = a + 65536
				URLDecode = URLDecode & ChrW(a)
				sl = sl + 9
			End If
		Case Else:
			hh = Mid(strIn, sl+kl, 2)
			a = Int("&H" & hh)
			If Abs(a)<128 Then
				sl = sl + 3
			Else
				hi = Mid(strIn, sl+3+kl, 2)
				a = Int("&H" & hh & hi)
				sl = sl + 6
			End If
			URLDecode = URLDecode & Chr(a)
	End Select
	tl = sl
	sl = InStr(sl, strIn, key, 1)
Loop
URLDecode = URLDecode & Mid(strIn, tl)
End Function

Function ConCon(content)
ConCon=content

Set regEx = New RegExp
regEx.IgnoreCase = True
regEx.Global = True

HTMLCount=0
Dim HTMLContainer(100)
regEx.Pattern = "<\s*\w*[^>]*>(.*?)<\s*/\s*\w*>"
Set Matches = regEx.Execute(ConCon)
For Each Match in Matches
If HTMLCount<=UBound(HTMLContainer) Then HTMLContainer(HTMLCount)=Replace(Match,vbCrlf,"")
ConCon=Replace(ConCon,Match,"{#HTML#"&HTMLCount&"#}")
HTMLCount=HTMLCount+1
Next
regEx.Pattern = "<\s*\w.*?>"
Set Matches = regEx.Execute(ConCon)
For Each Match in Matches
If HTMLCount<=UBound(HTMLContainer) Then HTMLContainer(HTMLCount)=Replace(Match,vbCrlf,"")
ConCon=Replace(ConCon,Match,"{#HTML#"&HTMLCount&"#}")
HTMLCount=HTMLCount+1
Next

regEx.Pattern = "(\[File:((\d{6})(\d{8})\.(.+?))\])"
Set Matches = regEx.Execute(ConCon)
For Each Match in Matches
Select Case LCASE(regEx.replace(Match,"$5"))
Case "jpg","jpeg","jpe","gif","png"
ConCon=Replace(ConCon,Match,"<a href=""files/"&regEx.replace(Match,"$3")&"/"&regEx.replace(Match,"$2")&"""><img src=""files/"&regEx.replace(Match,"$3")&"/"&regEx.replace(Match,"$2")&"""></a>")
Case Else
set fs=Server.CreateObject("Scripting.FileSystemObject")
FileInfo="files/"&regEx.replace(Match,"$3")&"/"&regEx.replace(Match,"$2")
If fs.FileExists(Server.Mappath(FileInfo))=true then
	set f=fs.GetFile(Server.Mappath(FileInfo))
	FileInfo=formatnumber(f.Size,0)&"字节"
	set f=nothing
Else
	FileInfo="不存在"
End If
set fs=nothing
ConCon=Replace(ConCon,Match,"<span style=""white-space: nowrap"">[下载文件] <a href=""files/"&regEx.replace(Match,"$3")&"/"&regEx.replace(Match,"$2")&""">"&regEx.replace(Match,"$2")&"</a>（"&FileInfo&"）</span>")
End Select
Next

regEx.Pattern = "(ed2k://\|file\|(.+?)\|(\d+)\|[a-fA-F0-9]{32}\|(((p=[a-fA-F0-9]{32}(:[a-fA-F0-9]{32})*\|)?(h=\w{32}\|)?(s=http://[\w\.-_&%/]+\|)*)|((p=[a-fA-F0-9]{32}(:[a-fA-F0-9]{32})*\|)?(s=http://[\w\.-_&%/]+\|)*(h=\w{32}\|)?)|((h=\w{32}\|)?(p=[a-fA-F0-9]{32}(:[a-fA-F0-9]{32})*\|)?(s=http://[\w\.-_&%/]+\|)*)|((h=\w{32}\|)?(s=http://[\w\.-_&%/]+\|)*(p=[a-fA-F0-9]{32}(:[a-fA-F0-9]{32})*\|)?)|((s=http://[\w\.-_&%/]+\|)*(p=[a-fA-F0-9]{32}(:[a-fA-F0-9]{32})*\|)?(h=\w{32}\|)?)|((s=http://[\w\.-_&%/]+\|)*(h=\w{32}\|)?(p=[a-fA-F0-9]{32}(:[a-fA-F0-9]{32})*\|)?))/(\|sources,[\w\.-_]+:\d{1,5}\|/)?)"
Set Matches = regEx.Execute(ConCon)
For Each Match in Matches
ConCon=Replace(ConCon,Match,"<span style=""white-space: nowrap"">[ed2k链接] <a href="""&regEx.replace(Match,"$1")&""">"&URLDecode(regEx.replace(Match,"$2"))&"</a>（"&formatnumber(regEx.replace(Match,"$3"),0)&"字节）</span>")
Next

URLPatt="((https?|ftp|gopher|telnet|file|notes|ms-help):((//)|(\\\\))+[\w\d:#@%/;$()~_?\+-=\\\.&]*)"
regEx.Pattern = URLPatt
Set Matches = regEx.Execute(ConCon)
For Each Match in Matches

	regEx.Pattern = "(http:\/\/v\.youku\.com\/v_show\/id_(.+?)\.html)"
	If regEx.Test(Match) Then
		ConCon=Replace(ConCon,Match,"<embed src=""http://player.youku.com/player.php/Type/Folder/Fid/2765974/Ob/1/Pt/0/sid/"&regEx.replace(Match,"$2")&"/v.swf"" quality=""high"" width=""480"" height=""400"" align=""middle"" allowScriptAccess=""allways"" mode=""transparent"" type=""application/x-shockwave-flash""></embed>")
	End If

	regEx.Pattern = "(http:\/\/www\.tudou\.com\/programs\/view\/(.+?)\/)"
	If regEx.Test(Match) Then
		ConCon=Replace(ConCon,Match,"<embed src=""http://www.tudou.com/v/"&regEx.replace(Match,"$2")&"/&rpid=13826951/v.swf"" type=""application/x-shockwave-flash"" allowscriptaccess=""always"" allowfullscreen=""true"" wmode=""opaque"" width=""480"" height=""400""></embed>")
	End If
	
	regEx.Pattern = URLPatt
	If regEx.Test(Match) Then
		ConCon=Replace(ConCon,Match,"<a href="""&regEx.replace(Match,"$1")&""" target=""_blank"">"&regEx.replace(Match,"$1")&"</a>")
	End If
Next

ConCon=Replace(ConCon,vbCrlf,"<br>")

regEx.Pattern = "(\{\#HTML\#(\d*)\#\})"
Set Matches = regEx.Execute(ConCon)
For Each Match in Matches
HTMLNO=regEx.replace(Match,"$2")
If HTMLNO<>"" and Isnumeric(HTMLNO)=True Then
	HTMLNO=int(HTMLNO)
	If HTMLNO>=0 AND HTMLNO<=UBound(HTMLContainer) Then
	ConCon=Replace(ConCon,Match,HTMLContainer(int(HTMLNO)))
	End If
End If
Next

Set Matches = nothing
Set regEx = nothing

End Function
%>