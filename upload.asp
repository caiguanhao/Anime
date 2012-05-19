<%@ Language=VBScript %>
<% 
option explicit 
Response.Expires = -1
Server.ScriptTimeout = 600
' All communication must be in UTF-8, including the response back from the request
Session.CodePage  = 65001

If Session("UserAdmin")="" OR Session("UserName")="" Then Response.end
%>
<!-- #include file="freeaspupload.asp" -->
<%
  Dim uploadsDirVar
  uploadsDirVar = Server.Mappath("files")
  
function TestEnvironment()
    Dim fso, fileName, testFile, streamTest
    TestEnvironment = ""
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    if not fso.FolderExists(uploadsDirVar) then
        TestEnvironment = "<B>Folder " & uploadsDirVar & " does not exist.</B><br>The value of your uploadsDirVar is incorrect. Open uploadTester.asp in an editor and change the value of uploadsDirVar to the pathname of a directory with write permissions."
        exit function
    end if
    fileName = uploadsDirVar & "\test.txt"
    on error resume next
    Set testFile = fso.CreateTextFile(fileName, true)
    If Err.Number<>0 then
        TestEnvironment = "<B>Folder " & uploadsDirVar & " does not have write permissions.</B><br>The value of your uploadsDirVar is incorrect. Open uploadTester.asp in an editor and change the value of uploadsDirVar to the pathname of a directory with write permissions."
        exit function
    end if
    Err.Clear
    testFile.Close
    fso.DeleteFile(fileName)
    If Err.Number<>0 then
        TestEnvironment = "<B>Folder " & uploadsDirVar & " does not have delete permissions</B>, although it does have write permissions.<br>Change the permissions for IUSR_<I>computername</I> on this folder."
        exit function
    end if
    Err.Clear
    Set streamTest = Server.CreateObject("ADODB.Stream")
    If Err.Number<>0 then
        TestEnvironment = "<B>The ADODB object <I>Stream</I> is not available in your server.</B><br>Check the Requirements page for information about upgrading your ADODB libraries."
        exit function
    end if
    Set streamTest = Nothing
end function

function SaveFiles
    Dim Upload, fileName, fileSize, ks, i, fileKey

    Set Upload = New FreeASPUpload
    Upload.Save(uploadsDirVar)

	' If something fails inside the script, but the exception is handled
	If Err.Number<>0 then Exit function

    SaveFiles = ""
    ks = Upload.UploadedFiles.keys
    if (UBound(ks) <> -1) then
        for each fileKey in Upload.UploadedFiles.keys
            SaveFiles = SaveFiles & "[File:" & Upload.UploadedFiles(fileKey).FileName & "]\n"
        next
        SaveFiles = "<html><head><script>alert(unescape('"&escape("文件上传成功。")&"'));window.parent.document.getElementById('attach1').value='';window.parent.document.getElementById('txtwait').style.display='none';window.parent.document.getElementById('fileinput').style.display='';window.parent.document.getElementById('fldIntro').value+='"&SaveFiles&"';</script></head><body></body></html>"
    else
        SaveFiles = "No file selected for upload or the file name specified in the upload form does not correspond to a valid file in the system."
    end if
end function
%>
<%
Dim diagnostics
if Request.ServerVariables("REQUEST_METHOD") <> "POST" then
    diagnostics = TestEnvironment()
    if diagnostics<>"" then
        response.write "<div>"
        response.write diagnostics
        response.write "<p>After you correct this problem, reload the page."
        response.write "</div>"
    end if
else
    response.write "<div>"
    response.write SaveFiles()
    response.write "</div>"
end if

%>
