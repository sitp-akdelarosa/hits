<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits										    _/
'_/	FileName	:dmi411.asp									    _/
'_/	Function	:��Ɣ���mail�Ώۍ��ڐݒ���͊m�F			    _/
'_/	Date		:2009/03/10									    _/
'_/	Code By		:Shibuta									    _/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTP�R���e���c�^�C�v�ݒ�
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->
<%
'�Z�b�V�����̗L�������`�F�b�N
	CheckLoginH
  
'�f�[�^�擾
	Dim F_DelResults(4), F_RecEmp(4), F_RecResults(4), F_DelEmp(4)
	Dim Email1, Email2, Email3, Email4, Email5
	Dim iCnt
	Dim NoEntered, ItemsToSend
	DIm strWork
	
	F_DelResults(0) = Request.Form("F_DelResults1")
	F_DelResults(1) = Request.Form("F_DelResults2")
	F_DelResults(2) = Request.Form("F_DelResults3")
	F_DelResults(3) = Request.Form("F_DelResults4")
	F_DelResults(4) = Request.Form("F_DelResults5")
	
	F_RecEmp(0) = Request.Form("F_RecEmp1")
	F_RecEmp(1) = Request.Form("F_RecEmp2")
	F_RecEmp(2) = Request.Form("F_RecEmp3")
	F_RecEmp(3) = Request.Form("F_RecEmp4")
	F_RecEmp(4) = Request.Form("F_RecEmp5")
	
	F_RecResults(0) = Request.Form("F_RecResults1")
	F_RecResults(1) = Request.Form("F_RecResults2")
	F_RecResults(2) = Request.Form("F_RecResults3")
	F_RecResults(3) = Request.Form("F_RecResults4")
	F_RecResults(4) = Request.Form("F_RecResults5")
	
	F_DelEmp(0) = Request.Form("F_DelEmp1")
	F_DelEmp(1) = Request.Form("F_DelEmp2")
	F_DelEmp(2) = Request.Form("F_DelEmp3")
	F_DelEmp(3) = Request.Form("F_DelEmp4")
	F_DelEmp(4) = Request.Form("F_DelEmp5")
	
	Email1 = Request.Form("Email1")
	Email2 = Request.Form("Email2")
	Email3 = Request.Form("Email3")
	Email4 = Request.Form("Email4")
	Email5 = Request.Form("Email5")
 	
	Session.Contents("dmi411") = "true"
	
 	'�������͂���Ă��Ȃ��ꍇ
	For iCnt = 0 To 4
		if F_DelResults(iCnt) = "0" and F_RecEmp(iCnt) = "0" and F_RecResults(iCnt) = "0" and F_DelEmp(iCnt) = "0" _
			and Email1 = "" and Email2 = "" and Email3 = "" and Email4 = "" and Email5 = "" then
			NoEntered = "1"
		else
			NoEntered = "0"
		end if
		if NoEntered = "0" then
			Exit For
		end if
	Next
 	
 	'���[�����M�Ώۍ��ڐ�
	ItemsToSend = 0
 	
 	'���O�o��
 	WriteLogH "c402", "��Ɣ���mail�ݒ�","02",""


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>��Ɣ���mail�Ώۍ��ڐݒ���͊m�F</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>

//�o�^
function GoEntry(){
	f=document.dmi411;
	f.action="dmi412.asp";
	return true;
}

//�߂�
function GoBack(){
	f=document.dmi411;
	f.action="dmi410.asp";
	return true;
}

</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!-------------�X�e�[�^�X�z�M�Ώۍ��ڐݒ���--------------------------->
<%'�f�[�^�o�^�^�X�V���܂�����ʂɂāu�ŐV�̏��ɍX�V�v��Submit���ꂽ�ꍇ�̑΍� %>
<% if NoEntered = "0" then %>
<% Session.Contents("ItemsSubmitted")="False"  %>
<FORM name="dmi411" method="POST">
<TABLE border="0" cellPadding="5" cellSpacing="0" width="100%">

	<TR>
		<TD width="5%" colspan="20">�@<B>��Ɣ���mail�i�ݒ�j���ڊm�F</B></TD>
	</TR>
	
	<TR><TD>�@</TD></TR>
	
	<TR>
		<TD width="5%" colspan="20">�@�ȉ��̍�ƈ˗������������ꍇ��mail�ŘA�����܂��B</TD>
	</TR>

	<% For iCnt = 0 To 4 %>
		<% if F_DelResults(iCnt) = "1" then %>
			<TR>
				<TD width="40%">�@�@�i�P�j�����o���<TD>
			</TR>
			<% ItemsToSend = ItemsToSend + 1 %>
			<% Exit For %>
		<% end if %>
	<% Next %>  
		
	<% For iCnt = 0 To 4 %>
		<% if F_RecEmp(iCnt) = "1" then %>
			<TR>
				<TD width="40%">�@�@�i�Q�j��������<TD>
			</TR>
			<% ItemsToSend = ItemsToSend + 1 %>
			<% Exit For %>
		<% end if %>
	<% Next %>
	
	<% For iCnt = 0 To 4 %>
		<% if F_RecResults(iCnt) = "1" then %>
			<TR>
				<TD width="40%">�@�@�i�R�j���������<TD>
			</TR>
			<% ItemsToSend = ItemsToSend + 1 %>
			<% Exit For %>
		<% end if %>
	<% Next %>
	
	<% For iCnt = 0 To 4 %>
		<% if F_DelEmp(iCnt) = "1" then %>
			<TR>
				<TD width="40%">�@�@�i�S�j����o���<TD>
			</TR>
			<% ItemsToSend = ItemsToSend + 1 %>
			<% Exit For %>
		<% end if %>
	<% Next %>
	
	<TR><TD>�@</TD></TR>

	<TR>
		<TD width="20%">�@�����M��</TD>
	</TR>
	
<% if Email1 <> "" then %>
	<TR>
		<TD width="5%" colspan="3">�@�@<%=Email1%></TD>
		
		<% if F_DelResults(0) = "1" then %>
				<TD width="30%" colspan="1">(1)
		<% end if %>
		
		<% if F_RecEmp(0) = "1" then %>
			<% if F_DelResults(0) = "1" then %>
				(2)
			<% else %>
				<TD width="30%" colspan="1">(2)
			<% end if%>
		<% end if %>

		<% if F_RecResults(0) = "1" then %>
			<% if F_DelResults(0) = "1" Or F_RecEmp(0) = "1" then %>
				(3)
			<% else %>
				<TD width="30%" colspan="1">(3)
			<% end if%>
		<% end if %>

		<% if F_DelEmp(0) = "1" then %>
			<% if F_DelResults(0) = "1" Or F_RecEmp(0) = "1" Or F_RecResults(0) = "1" then %>
				(4)
			<% else %>
				<TD width="30%" colspan="1">(4)
			<% end if %>
		<% end if %>
		</TD>
	</TR>
<% end if %>

<% if Email2 <> "" then %>
	<TR>
		<TD width="5%" colspan="3">�@�@<%=Email2%></TD>
		
		<% if F_DelResults(1) = "1" then %>
				<TD width="30%" colspan="1">(1)
		<% end if %>
		
		<% if F_RecEmp(1) = "1" then %>
			<% if F_DelResults(1) = "1" then %>
				(2)
			<% else %>
				<TD width="30%" colspan="1">(2)
			<% end if%>
		<% end if %>

		<% if F_RecResults(1) = "1" then %>
			<% if F_DelResults(1) = "1" Or F_RecEmp(1) = "1" then %>
				(3)
			<% else %>
				<TD width="30%" colspan="1">(3)
			<% end if%>
		<% end if %>

		<% if F_DelEmp(1) = "1" then %>
			<% if F_DelResults(1) = "1" Or F_RecEmp(1) = "1" Or F_RecResults(1) = "1" then %>
				(4)
			<% else %>
				<TD width="30%" colspan="1">(4)
			<% end if %>
		<% end if %>
		</TD>
	</TR>
<% end if %>

<% if Email3 <> "" then %>
	<TR>
		<TD width="5%" colspan="3">�@�@<%=Email3%></TD>
		
		<% if F_DelResults(2) = "1" then %>
				<TD width="30%" colspan="1">(1)
		<% end if %>
		
		<% if F_RecEmp(2) = "1" then %>
			<% if F_DelResults(2) = "1" then %>
				(2)
			<% else %>
				<TD width="30%" colspan="1">(2)
			<% end if%>
		<% end if %>

		<% if F_RecResults(2) = "1" then %>
			<% if F_DelResults(2) = "1" Or F_RecEmp(2) = "1" then %>
				(3)
			<% else %>
				<TD width="30%" colspan="1">(3)
			<% end if%>
		<% end if %>

		<% if F_DelEmp(2) = "1" then %>
			<% if F_DelResults(2) = "1" Or F_RecEmp(2) = "1" Or F_RecResults(2) = "1" then %>
				(4)
			<% else %>
				<TD width="30%" colspan="1">(4)
			<% end if %>
		<% end if %>
		</TD>
	</TR>
<% end if %>

<% if Email4 <> "" then %>
	<TR>
		<TD width="5%" colspan="3">�@�@<%=Email4%></TD>
		
		<% if F_DelResults(3) = "1" then %>
				<TD width="30%" colspan="1">(1)
		<% end if %>
		
		<% if F_RecEmp(3) = "1" then %>
			<% if F_DelResults(3) = "1" then %>
				(2)
			<% else %>
				<TD width="30%" colspan="1">(2)
			<% end if%>
		<% end if %>

		<% if F_RecResults(3) = "1" then %>
			<% if F_DelResults(3) = "1" Or F_RecEmp(3) = "1" then %>
				(3)
			<% else %>
				<TD width="30%" colspan="1">(3)
			<% end if%>
		<% end if %>

		<% if F_DelEmp(3) = "1" then %>
			<% if F_DelResults(3) = "1" Or F_RecEmp(3) = "1" Or F_RecResults(3) = "1" then %>
				(4)
			<% else %>
				<TD width="30%" colspan="1">(4)
			<% end if %>
		<% end if %>
		</TD>
	</TR>
<% end if %>

<% if Email5 <> "" then %>
	<TR>
		<TD width="5%" colspan="3">�@�@<%=Email5%></TD>
		
		<% if F_DelResults(4) = "1" then %>
				<TD width="30%" colspan="1">(1)
		<% end if %>
		
		<% if F_RecEmp(4) = "1" then %>
			<% if F_DelResults(4) = "1" then %>
				(2)
			<% else %>
				<TD width="30%" colspan="1">(2)
			<% end if%>
		<% end if %>

		<% if F_RecResults(4) = "1" then %>
			<% if F_DelResults(4) = "1" Or F_RecEmp(4) = "1" then %>
				(3)
			<% else %>
				<TD width="30%" colspan="1">(3)
			<% end if%>
		<% end if %>

		<% if F_DelEmp(4) = "1" then %>
			<% if F_DelResults(4) = "1" Or F_RecEmp(4) = "1" Or F_RecResults(4) = "1" then %>
				(4)
			<% else %>
				<TD width="30%" colspan="1">(4)
			<% end if %>
		<% end if %>
		</TD>
	</TR>
<% end if %>

	<TR>
		<TD colspan="5" align="center">
			<INPUT type="hidden" name="F_DelResults1" value="<%=F_DelResults(0)%>">
			<INPUT type="hidden" name="F_DelResults2" value="<%=F_DelResults(1)%>">
			<INPUT type="hidden" name="F_DelResults3" value="<%=F_DelResults(2)%>">
			<INPUT type="hidden" name="F_DelResults4" value="<%=F_DelResults(3)%>">
			<INPUT type="hidden" name="F_DelResults5" value="<%=F_DelResults(4)%>">
			
			<INPUT type="hidden" name="F_RecEmp1" value="<%=F_RecEmp(0)%>">
			<INPUT type="hidden" name="F_RecEmp2" value="<%=F_RecEmp(1)%>">
			<INPUT type="hidden" name="F_RecEmp3" value="<%=F_RecEmp(2)%>">
			<INPUT type="hidden" name="F_RecEmp4" value="<%=F_RecEmp(3)%>">
			<INPUT type="hidden" name="F_RecEmp5" value="<%=F_RecEmp(4)%>">
			
			<INPUT type="hidden" name="F_RecResults1" value="<%=F_RecResults(0)%>">
			<INPUT type="hidden" name="F_RecResults2" value="<%=F_RecResults(1)%>">
			<INPUT type="hidden" name="F_RecResults3" value="<%=F_RecResults(2)%>">
			<INPUT type="hidden" name="F_RecResults4" value="<%=F_RecResults(3)%>">
			<INPUT type="hidden" name="F_RecResults5" value="<%=F_RecResults(4)%>">
			
			<INPUT type="hidden" name="F_DelEmp1" value="<%=F_DelEmp(0)%>">
			<INPUT type="hidden" name="F_DelEmp2" value="<%=F_DelEmp(1)%>">
			<INPUT type="hidden" name="F_DelEmp3" value="<%=F_DelEmp(2)%>">
			<INPUT type="hidden" name="F_DelEmp4" value="<%=F_DelEmp(3)%>">
			<INPUT type="hidden" name="F_DelEmp5" value="<%=F_DelEmp(4)%>">
			
			<INPUT type="hidden" name="Email1" value="<%=Email1%>">
			<INPUT type="hidden" name="Email2" value="<%=Email2%>">
			<INPUT type="hidden" name="Email3" value="<%=Email3%>">
			<INPUT type="hidden" name="Email4" value="<%=Email4%>">
			<INPUT type="hidden" name="Email5" value="<%=Email5%>">
			
			<INPUT type="submit" value="�n�j" onClick="return GoEntry()">
			<INPUT type="submit" value="�߂�" onClick="GoBack()">
		</TD>
	</TR>  
</TABLE>
</FORM>
<% else %>
<FORM name="dmi411" method="POST">
<TABLE border=0 cellPadding=3 cellSpacing=3 width="100%">
	<TR><TD>�@</TD></TR>
	
	<TR>
		<TD width="5%" colspan="20">�@<B>��Ɣ���mail�i�ݒ�j���ڊm�F</B></TD>
	</TR>
	
	<TR><TD>�@</TD></TR>
	
	<TR>
		<TD width="5%" colspan="20">�@�����w�肳��Ă��܂���B��낵����΁u�n�j�v�{�^�����N���b�N���Ă��������B</TD>
	</TR>

	<TR><TD>�@</TD></TR>
	<TR>
		<TD colspan="5" align="center">
			<INPUT type="hidden" name="F_DelResults1" value="<%=F_DelResults(0)%>">
			<INPUT type="hidden" name="F_DelResults2" value="<%=F_DelResults(1)%>">
			<INPUT type="hidden" name="F_DelResults3" value="<%=F_DelResults(2)%>">
			<INPUT type="hidden" name="F_DelResults4" value="<%=F_DelResults(3)%>">
			<INPUT type="hidden" name="F_DelResults5" value="<%=F_DelResults(4)%>">

			<INPUT type="hidden" name="F_RecEmp1" value="<%=F_RecEmp(0)%>">
			<INPUT type="hidden" name="F_RecEmp2" value="<%=F_RecEmp(1)%>">
			<INPUT type="hidden" name="F_RecEmp3" value="<%=F_RecEmp(2)%>">
			<INPUT type="hidden" name="F_RecEmp4" value="<%=F_RecEmp(3)%>">
			<INPUT type="hidden" name="F_RecEmp5" value="<%=F_RecEmp(4)%>">

			<INPUT type="hidden" name="F_RecResults1" value="<%=F_RecResults(0)%>">
			<INPUT type="hidden" name="F_RecResults2" value="<%=F_RecResults(1)%>">
			<INPUT type="hidden" name="F_RecResults3" value="<%=F_RecResults(2)%>">
			<INPUT type="hidden" name="F_RecResults4" value="<%=F_RecResults(3)%>">
			<INPUT type="hidden" name="F_RecResults5" value="<%=F_RecResults(4)%>">

			<INPUT type="hidden" name="F_DelEmp1" value="<%=F_DelEmp(0)%>">
			<INPUT type="hidden" name="F_DelEmp2" value="<%=F_DelEmp(1)%>">
			<INPUT type="hidden" name="F_DelEmp3" value="<%=F_DelEmp(2)%>">
			<INPUT type="hidden" name="F_DelEmp4" value="<%=F_DelEmp(3)%>">
			<INPUT type="hidden" name="F_DelEmp5" value="<%=F_DelEmp(4)%>">

			<INPUT type="hidden" name="Email1" value="<%=Email1%>">
			<INPUT type="hidden" name="Email2" value="<%=Email2%>">
			<INPUT type="hidden" name="Email3" value="<%=Email3%>">
			<INPUT type="hidden" name="Email4" value="<%=Email4%>">
			<INPUT type="hidden" name="Email5" value="<%=Email5%>">

			<INPUT type="submit" value="�n�j" onClick="return GoEntry()">
			<INPUT type="submit" value="�߂�" onClick="GoBack()">
		</TD>
	</TR>
</TABLE>
</FORM>
<% end if %>
<!-------------��ʏI���--------------------------->
</BODY>
</HTML>
