<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
	' File System Object �̐���
	Set fs=Server.CreateObject("Scripting.FileSystemobject")

	' �����ݏ��Ɖ�
	WriteLog fs, "b001","�����ݏ��Ɖ�","00", ","
%>
<html>

<frameset rows="233,*,68" border="0">
	<frame src="./arvdep_input.asp" name="search" noresize scrolling="no">
	<frame src="./arvdep_list.asp" name="list" noresize>
	<frame src="./arvdep_bottom.asp" name="bottom" noresize scrolling="no">
</frameset>

</html>
