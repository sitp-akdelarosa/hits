<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:sst000F.asp				_/
'_/	Function	:�X�e�[�^�X�z�M�ꗗ��ʃt���[��		_/
'_/	Date			:2003/12/25				_/
'_/	Code By		:aspLand HARA			_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTP�R���e���c�^�C�v�ݒ�
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->
<%
'''�Z�b�V�����̗L�������`�F�b�N
	CheckLoginH
'''�f�[�^�擾
	Dim USER,UType
	USER = Session.Contents("userid")
'''�G���[�g���b�v�J�n
	on error resume next
'''DB�ڑ�
	Dim ObjConn, ObjRS, StrSQL
	ConnDBH ObjConn, ObjRS

'''���[�U�F�؁A�����f�[�^����
	StrSQL = "select UserType,FullName,NameAbrev from mUsers where UserCode='" & USER &"'"
	ObjRS.Open StrSQL, ObjConn
	if err <> 0 then
		DisConnDBH ObjConn, ObjRS	'DB�ؒf
		jumpErrorP "0","c101","01","�X�e�[�^�X�z�M�˗����ꗗ","102",""
	end if

	UType = Trim(ObjRS("UserType"))		'���[�U�^�C�v
	Session.Contents("UType") = Utype
	Session.Contents("LinUN") = Trim(ObjRS("FullName"))		'���O�C�����[�U����
	Session.Contents("sUN") = Trim(ObjRS("NameAbrev"))		'���O�C�����[�U����

'''DB�ڑ�����
	DisConnDBH ObjConn, ObjRS
'''�G���[�g���b�v����
	on error goto 0
%>
<!-------------�X�e�[�^�X�z�M�ꗗ���Frame--------------------------->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE>�X�e�[�^�X�z�M���ꗗ</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<frameset rows="73,*,33" border="0" frameborder="0">
	<noframes>
	���̃y�[�W�̓t���[���Ή��̃u���E�U�ł������������B
	</noframes>
	<frame src="./sst000T.asp" name="Top" scrolling="no" noresize>
	<frameset cols="120,*" border="0" frameborder="0" border="0" frameborder="0">
		<frame src="./sst000M.asp" name="Menu">
		<frame src="./top.html" name="List" scrolling="no" noresize>
	</frameset>
	<frame src="./sst000B.asp" name="Bottom" scrolling="no" noresize>
</frameset>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!-------------��ʏI���--------------------------->
</BODY>
</HTML>
