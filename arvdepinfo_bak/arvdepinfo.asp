<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
	' File System Object �̐���
	Set fs=Server.CreateObject("Scripting.FileSystemobject")

	' �����ݏ��Ɖ�
	WriteLog fs, "b001","�����ݏ��Ɖ�","00", ","
%>
<html>

<frameset rows="308,*,66" border="0">
	<frame src="./arvdep_input2.asp" name="search" noresize scrolling="no">
	<frame src="./arvdep_list2.asp" name="list" noresize scrolling="auto">
	<frame src="./arvdep_bottom2.asp" name="bottom" noresize scrolling="no">
</frameset>

</html>
