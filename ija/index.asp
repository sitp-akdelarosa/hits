<%@ LANGUAGE="VBScript" %>
<%Option Explicit%>
<%
' �C����уV�X�e��(�g�ѓd�b��)�@�J�ڐ}
' 
' �@�@�@�@�@�@�@�@�@�R���e�iNo�Ɖ�
' �@�@�@�@�@�@�@�@��������(mcont01.asp) ����(mcont02.asp)
' �@�@�@�@�@�@�@�@���@�@�@�R���e�iNo���́@�@�@�Ɖ��
' �@�@�@�@�@�@�@�@���@�@�@�@�@�@�@�@�@�@�@�@�@�@��
' �@�@�@�@�@�@�@�@��BL�ԍ��Ɖ�@�@�@�@�@�@�@�@�@��
' ���C�����j���[�@��������(mblno01.asp) ����(mblno02.asp)
'  (index.asp)�@�����@�@�@�@BL�ԍ����́@�@�@�R���e�i�ꗗ
' �@�@�@�@�@�@�@�@��
' �@�@�@�@�@�@�@�@��������������
' �@�@�@�@�@�@�@�@��������(muser.asp) ����(mrung01.asp) ����(mrung02.asp) ����[mrung03.asp]
' �@�@�@�@�@�@�@�@���@�@�@���[�UID���́@�@�R���e�iNo���́@�@������ƑI���@�@�@�t�@�C���o��
' �@�@�@�@�@�@�@�@��
' �@�@�@�@�@�@�@�@���Q�[�g�����p����
' �@�@�@�@�@�@�@�@��������(mterm01.asp)
' �@�@�@�@�@�@�@�@��
' �@�@�@�@�@�@�@�@���f���\��(�����ߑ勴,�ҋ@��,�Q�[�g�O �e����)
' �@�@�@�@�@�@�@�@��������(mpict01.asp)
' 
%>
<!--#include file="common.inc"-->
<!--#include file="mcommon.inc"-->
<%
Dim sPhoneType
sPhoneType = GetPhoneType()

' Log�o��
Dim oFs
Set oFS = Server.CreateObject("Scripting.FileSystemObject")
WriteLogM oFS, "Unknown", "0200", "�g��-�s�n�o���", "00" , sPhoneType, ","
Set oFS = Nothing

Dim sTBorder
If sPhoneType = "E" Then
	' EzWeb�p�^�O��ҏW
	Response.ContentType = "text/x-hdml; charset=Shift_JIS hdml"
%>
	<hdml version="3.0" public="true" markable="true">
	
	<display title="HiTS ver.2">
		<center>
		<img src="mtitle<%=GetImageExt()%>" alt="HiTS ver.2"><br>
		���Ŕ��o���Ɖ�<br>
		<a task="gosub" dest="mcont01.asp" accesskey="1">�@���Ŕԍ��Ɖ�</a><br>
		<a task="gosub" dest="./cam/mcont01cam.asp">�@�@�@�@�����ӓ�</a><br>
		<a task="gosub" dest="mblno01.asp" accesskey="2">�@BL�ԍ��Ɖ�</a><br>
		<a task="gosub" dest="./cam/mblno01cam.asp">�@�@�@�@�����ӓ�</a><br>
		�^�s������<br>
		<a task="gosub" dest="muser.asp" accesskey="3">�@������������</a><br>
		�^�[�~�i�����<br>
		<a task="gosub" dest="mterm01.asp" accesskey="4">�@��-ē�����</a><br>
		(����)<br>
		<a task="gosub" dest="mpict01.asp?pict=1" accesskey="5">�@�����ߑ勴</a><br>
		<a task="gosub" dest="mpict01.asp?pict=2" accesskey="6">�@�ҋ@��f��</a><br>
		<a task="gosub" dest="mpict01.asp?pict=3" accesskey="7">�@��-đO�f��</a><br>
		(ICCT)<br>
		<a task="gosub" dest="mpict01.asp?pict=4" accesskey="8">�@��-đO�f��</a><br>
	</display>
	</hdml>
<%
Else
	' EzWeb�ȊO�̃^�O��ҏW
%>
	<html>
	<head>
		<meta http-equiv="Content-Language" content="ja">
		<meta http-equiv="Content-Type" content="text/html; charset=shift_jis">
		<%=GetTitleTag("HiTS ver.2")%>
	</head>
	
	<body>
	<center>
	<img src="mtitle<%=GetImageExt()%>" alt="HiTS ver.2"><br>
	<hr>
<%
	If sPhoneType = "I" Then
		'i-mode�ł̓e�[�u���������Ȃ��̂ō��l��
%>
		<div align="left">
			���Ŕ��o���Ɖ�<br>
			�@<a href="mcont01.asp" <%=GetKeyTag("1")%>><%=GetKeyLabel("1")%>���Ŕԍ��Ɖ�</a><br>
			�@<a href="mblno01.asp" <%=GetKeyTag("2")%>><%=GetKeyLabel("2")%>BL�ԍ��Ɖ�</a><br>
			�^�s������<br>
			�@<a href="muser.asp" <%=GetKeyTag("3")%>><%=GetKeyLabel("3")%>������������</a><br>
			�^�[�~�i�����<br>
			�@<a href="mterm01.asp" <%=GetKeyTag("4")%>><%=GetKeyLabel("4")%>��-ē�����</a><br>
			�@(����)<br>
			�@<a href="mpict01.asp?pict=1" <%=GetKeyTag("5")%>><%=GetKeyLabel("5")%>�����ߑ勴</a><br>
			�@<a href="mpict01.asp?pict=2" <%=GetKeyTag("6")%>><%=GetKeyLabel("6")%>�ҋ@��f��</a><br>
			�@<a href="mpict01.asp?pict=3" <%=GetKeyTag("7")%>><%=GetKeyLabel("7")%>��-đO�f��</a><br>
			�@(ICCT)<br>
			�@<a href="mpict01.asp?pict=4" <%=GetKeyTag("8")%>><%=GetKeyLabel("8")%>��-đO�f��</a><br>
		</div>
<%
	Else
		If sPhoneType = "J" Then
			sTBorder = ""
		Else
			sTBorder = " border=""0"" "
		End If
%>
		<table <%=sTBorder%>>
			<tr><td>
				���Ŕ��o���Ɖ�<br>
			</td></tr>
			<tr><td>
				�@<a href="mcont01.asp" <%=GetKeyTag("1")%>><%=GetKeyLabel("1")%>���Ŕԍ��Ɖ�</a><br>
			</td></tr>
			<tr><td>
				�@<a href="mblno01.asp" <%=GetKeyTag("2")%>><%=GetKeyLabel("2")%>BL�ԍ��Ɖ�</a><br>
			</td></tr>
			<tr><td>
				�^�s������<br>
			</td></tr>
			<tr><td>
				�@<a href="muser.asp" <%=GetKeyTag("3")%>><%=GetKeyLabel("3")%>������������</a><br>
			</td></tr>
			<tr><td>
				�^�[�~�i�����<br>
			</td></tr>
			<tr><td>
				�@<a href="mterm01.asp" <%=GetKeyTag("4")%>><%=GetKeyLabel("4")%>��-ē�����</a><br>
			</td></tr>
			<tr><td>
				�@(����)<br>
			</td></tr>
			<tr><td>
				�@<a href="mpict01.asp?pict=1" <%=GetKeyTag("5")%>><%=GetKeyLabel("5")%>�����ߑ勴</a><br>
			</td></tr>
			<tr><td>
				�@<a href="mpict01.asp?pict=2" <%=GetKeyTag("6")%>><%=GetKeyLabel("6")%>�ҋ@��f��</a><br>
			</td></tr>
			<tr><td>
				�@<a href="mpict01.asp?pict=3" <%=GetKeyTag("7")%>><%=GetKeyLabel("7")%>��-đO�f��</a><br>
			</td></tr>
			<tr><td>
				�@(ICCT)<br>
			</td></tr>
			<tr><td>
				�@<a href="mpict01.asp?pict=4" <%=GetKeyTag("8")%>><%=GetKeyLabel("8")%>��-đO�f��</a><br>
			</td></tr>
		</table>
<%
	End If
%>
	<hr>
	</body>
	</html>
<%
End If
%>




