<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
	' File System Object �̐���
	Set fs=Server.CreateObject("Scripting.FileSystemobject")

	' �����ݏ��Ɖ�
	WriteLog fs, "d101","�����ݏ��Ɖ�","01", ","
%>
<html>

<frameset rows="308,*,66" border="0">
	<frame src="./arvdep_input.asp" name="search" noresize scrolling="no">
	<frame src="./arvdep_list.asp" name="list" noresize scrolling="auto">
	<frame src="./arvdep_bottom.asp" name="bottom" noresize scrolling="no">
</frameset>

</html>
