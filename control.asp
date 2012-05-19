&nbsp;<select onchange="location.href=this.value" id="ctrlsel">
<option value="./">选择任务</option>
<option value="addoredit.asp">发布资源</option>
<option value="settings.asp?v=cate">管理分类</option>
<option value="settings.asp?v=users">管理用户</option>
<option value="settings.asp?v=cordb">压缩/修复数据库</option>
</select><input type="button" onclick="location.href=document.getElementById('ctrlsel').value" value="go">