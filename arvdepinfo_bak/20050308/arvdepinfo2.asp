<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
	' File System Object ‚Ì¶¬
	Set fs=Server.CreateObject("Scripting.FileSystemobject")

	' ’…—£ŠÝî•ñÆ‰ï
	WriteLog fs, "b001","’…—£ŠÝî•ñÆ‰ï","00", ","
%>
<html>

<frameset rows="308,*,66" border="0">
	<frame src="./arvdep_input2.asp" name="search" noresize scrolling="no">
	<frame src="./arvdep_list2.asp" name="list" noresize scrolling="auto">
	<frame src="./arvdep_bottom2.asp" name="bottom" noresize scrolling="no">
</frameset>

</html>
