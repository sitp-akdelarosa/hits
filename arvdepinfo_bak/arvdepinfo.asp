<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
	' File System Object の生成
	Set fs=Server.CreateObject("Scripting.FileSystemobject")

	' 着離岸情報照会
	WriteLog fs, "b001","着離岸情報照会","00", ","
%>
<html>

<frameset rows="308,*,66" border="0">
	<frame src="./arvdep_input2.asp" name="search" noresize scrolling="no">
	<frame src="./arvdep_list2.asp" name="list" noresize scrolling="auto">
	<frame src="./arvdep_bottom2.asp" name="bottom" noresize scrolling="no">
</frameset>

</html>
