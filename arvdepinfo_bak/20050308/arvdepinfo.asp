<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
	' File System Object ‚Ì¶¬
	Set fs=Server.CreateObject("Scripting.FileSystemobject")

	' ’…—£ŠÝî•ñÆ‰ï
	WriteLog fs, "b001","’…—£ŠÝî•ñÆ‰ï","00", ","
%>
<html>

<frameset rows="233,*,68" border="0">
	<frame src="./arvdep_input.asp" name="search" noresize scrolling="no">
	<frame src="./arvdep_list.asp" name="list" noresize>
	<frame src="./arvdep_bottom.asp" name="bottom" noresize scrolling="no">
</frameset>

</html>
