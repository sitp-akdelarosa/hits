<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi000F.asp				_/
'_/	Function	:���O���ꗗ��ʃt���[��		_/
'_/	Date		:2003/05/26				_/
'_/	Code By		:SEIKO Electric.Co ��d			_/
'_/	Modify		:					_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTP�R���e���c�^�C�v�ݒ�
	Response.ContentType = "text/html; charset=Shift_JIS"
	Response.AddHeader "Pragma","No-Cache"
	Response.AddHeader "Cache-Control","No-Cache"
%>
<!--#include File="Common.inc"-->
<%
'�Z�b�V�����̗L�������`�F�b�N
  CheckLoginH
'�f�[�^�擾
  dim USER,UType
  USER       = Session.Contents("userid")
'�G���[�g���b�v�J�n
    on error resume next
'DB�ڑ�
  dim ObjConn, ObjRS, StrSQL
  ConnDBH ObjConn, ObjRS


'���[�U�F�؁A�����f�[�^����
  StrSQL = "select UserType,FullName,HeadCompanyCode,NameAbrev from mUsers " &_ 
           "where UserCode='" & USER &"'"
  ObjRS.Open StrSQL, ObjConn
  if err <> 0 then
    DisConnDBH ObjConn, ObjRS	'DB�ؒf
    jampErrerP "0","b101","01","���ʁF���[�U�f�[�^�擾","102",""
  end if

    UType                      = Trim(ObjRS("UserType"))		'���[�U�^�C�v
    Session.Contents("UType")  = Utype
    Session.Contents("LinUN")  = Trim(ObjRS("FullName"))		'���O�C�����[�U����
    Session.Contents("sUN")    = Trim(ObjRS("NameAbrev"))		'���O�C�����[�U����
    If UType=5 Then
      Session.Contents("COMPcd") = Trim(ObjRS("HeadCompanyCode"))	'�w�b�h��ЃR�[�h
    Else 
      '2010/04/16 Upd-S Tanaka &nbsp;�̕�����SQL���Ɏg����̂ŏC��
      Session.Contents("COMPcd") = "&nbsp;�@"
      'Session.Contents("COMPcd") = ""
      '2010/04/16 Upd-S Tanaka
    End If
'DB�ڑ�����
  DisConnDBH ObjConn, ObjRS
'�G���[�g���b�v����
  on error goto 0
%>
<!-------------���O���ꗗ���Frame--------------------------->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<TITLE>���O���ꗗ</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<frameset rows="40,*,33" border="0" frameborder="0">
  <noframes>
  ���̃y�[�W�̓t���[���Ή��̃u���E�U�ł������������B
  </noframes>
  <frame src="./dmi000T.asp" name="Top" scrolling="no" noresize>
  <frameset cols="100,*" border="0" frameborder="0" border="0" frameborder="0">
    <frame src="./dmi000M.asp" name="Menu">
<!--
    <frame src="./dmo010F.asp" name="List" scrolling="no" noresize>
-->
    <frame src="./top.asp" name="List" scrolling="no" noresize>
  </frameset>
  <frame src="./dmi000B.asp" name="Bottom" scrolling="no" noresize>
</frameset>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!-------------��ʏI���--------------------------->
</BODY></HTML>
