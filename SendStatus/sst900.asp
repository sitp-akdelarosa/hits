<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:sst900.asp				_/
'_/	Function	:�X�e�[�^�X���z�M���ʏ���			_/
'_/	Date			:2004/1/15				_/
'_/	Code By		:aspLand HARA		_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'''HTTP�R���e���c�^�C�v�ݒ�
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->
<%
	'''�Z�b�V�����̗L�������`�F�b�N
	CheckLoginH

	'''�f�[�^�擾
	Dim CONnum,Flag,BLnum
	Dim inPutStr,strNums
	CONnum = Request.Form("ContBLNo")
	Flag   = Request.Form("ContORBL")

	'''�G���[�g���b�v�J�n
	on error resume next
	'''DB�ڑ�
	Dim ObjConn, ObjRS, StrSQL
	ConnDBH ObjConn, ObjRS

	Select Case Flag
		Case "1"		'''�R���e�i�ԍ��w��
			inPutStr="<INPUT type=hidden name='cntnrno' value='"& CONnum &"'>"
		Case "2"		'�a�k�ԍ��w��
			inPutStr="<INPUT type=hidden name='blno' value='"& CONnum &"'>"
	End Select

	if Flag=1 Then
		Session.Contents("route") = "�A���R���e�i���Ɖ�i��ƑI���j "
	Else
		Session.Contents("route") = "Top > �A���R���e�i���Ɖ�i��ƑI���j "
	End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>�]����</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT language=JavaScript>
<!--
function opnewin(){
  window.focus();
  document.sst900.submit();
}
// -->
</SCRIPT>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY onLoad="opnewin()">
<P>�]����...���΂炭���҂����������B</P>
<FORM action="../impcntnr.asp" name="sst900">
<%= inPutStr %>
</FORM>
</BODY>
</HTML>
