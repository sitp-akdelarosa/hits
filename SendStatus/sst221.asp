<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:sst221.asp				_/
'_/	Function	:�X�e�[�^�X�z�M�Ώۍ폜			_/
'_/	Date			:2004/01/15				_/
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

	'''�f�[�^�폜���܂�����ʂɂāu�ŐV�̏��ɍX�V�v��Submit���ꂽ�ꍇ�̑΍�
	if Session.Contents("DeleteSubmitted")="False" then

	'''�f�[�^�擾
	Dim USER, KIND, NUMBER
	USER   = UCase(Session.Contents("userid"))
	KIND = Request.Form("ContORBL")
	NUMBER = Request.Form("ContBLNo")

	'''�G���[�g���b�v�J�n
	on error resume next
	'''DB�ڑ�
	Dim ObjConn, ObjRS, StrSQL
	ConnDBH ObjConn, ObjRS

	'''�f�[�^�폜�i�����敪Process��'D'�ɂ���B���ۂ̃��R�[�h�폜�͓��������ɂčs���B�j
	StrSQL = "UPDATE TargetContainers SET UpdtTime='" & Now() & "', UpdtPgCd='STATUS01',"
	StrSQL =  StrSQL & " UpdtTmnl='" & USER & "', Process='D' "
	if KIND = 1 then		'''�폜�Ώۂ��R���e�i�ԍ�
		StrSQL =  StrSQL & " WHERE ContNo='" & NUMBER & "' AND UserCode='" & USER & "'"
	elseif KIND = 2 then		'''�폜�Ώۂ��a�k�ԍ�
		StrSQL =  StrSQL & " WHERE BLNo='" & NUMBER & "' AND UserCode='" & USER & "'"
	else
		response.write("KIND error!")
		response.end
	end if

	ObjConn.Execute(StrSQL)
	if err <> 0 then
		Set ObjRS = Nothing
		jumpErrorPDB ObjConn,"1","c102","14","�X�e�[�^�X�z�M�Ώۍ폜","104","SQL:<BR>"&StrSQL
	end if

	'''���O�o��
	WriteLogH "c102", "�X�e�[�^�X�z�M�Ώۍ폜","14",""
	ObjRS.close

	'''DB�ڑ�����
	DisConnDBH ObjConn, ObjRS
	'''�G���[�g���b�v����
	on error goto 0

	Session.Contents("DeleteSubmitted") = "True"

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>�X�e�[�^�X�z�M�Ώۍ폜</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT Language="JavaScript">
<!--
function CloseWin(){
	try{
		window.opener.parent.DList.location.href="sst100L.asp"
	}catch(e){}
	window.close();
}
// -->
</SCRIPT>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin="0" topMargin="0" marginheight="0" marginwidth="0">
<!-------------�X�e�[�^�X�z�M�Ώۍ폜--------------------------->
<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%">
<FORM name="sst221">
	<TR><TD>�@</TD></TR>
	<TR>
		<TD align="center">
			�폜���܂����B
	</TR>
	<TR><TD>�@</TD></TR>
	<TR>
		<TD align="center">
			<INPUT type="button" value="����" onClick="CloseWin()">
		</TD>
	</TR>
</FORM>
</TABLE>
<!-------------��ʏI���--------------------------->
</BODY>
</HTML>

<%'''if Session.Contents("DeleteSubmitted")="False"��else���� %>
<% else %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>�X�e�[�^�X�z�M�Ώۍ폜</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT Language="JavaScript">
<!--
function CloseWin(){
	try{
		window.opener.parent.DList.location.href="sst100L.asp"
	}catch(e){}
	window.close();
}
// -->
</SCRIPT>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin="0" topMargin="0" marginheight="0" marginwidth="0">
<!-------------�X�e�[�^�X�z�M�Ώۍ폜--------------------------->
<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%">
<FORM name="sst221">
	<TR><TD>�@</TD></TR>
	<TR>
		<TD align="center">
			�폜�͊��Ɋ������Ă��܂��B<BR><BR><BR>
			<INPUT type="button" value="����" onClick="CloseWin()">
		</TD>
	</TR>
</FORM>
</TABLE>
<!-------------��ʏI���--------------------------->
</BODY>
</HTML>
<%'''if Session.Contents("DeleteSubmitted")="False"��endif���� %>
<% end if %>
