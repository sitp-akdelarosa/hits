<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
	' File System Object の生成
	Set fs=Server.CreateObject("Scripting.FileSystemobject")

	' 着離岸情報照会
	WriteLog fs, "d101","着離岸情報照会","01", ","
%>
<html>

<frameset rows="308,*,66" border="0">
	<frame src="./arvdep_input.asp" name="search" noresize scrolling="no">
	<frame src="./arvdep_list.asp" name="list" noresize scrolling="auto">
	<frame src="./arvdep_bottom.asp" name="bottom" noresize scrolling="no">
</frameset>

</html>
