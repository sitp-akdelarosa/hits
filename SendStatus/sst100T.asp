<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:sst100T.asp				_/
'_/	Function	:�X�e�[�^�X�z�M�˗����ꗗ��ʃg�b�v		_/
'_/	Date			:2003/12/25				_/
'_/	Code By		:aspLand HARA			_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTP�R���e���c�^�C�v�ݒ�
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<%
'���[�U�f�[�^����
	Dim USER, COMPcd, LinUN
	USER   = Session.Contents("userid")
	COMPcd = Session.Contents("COMPcd")
	LinUN  = Session.Contents("LinUN")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>�X�e�[�^�X�z�M�˗����ꗗ</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
//<!--
//����
function SearchContBL(){
	f = document.search;
	if(!f.ContORBL[0].checked && !f.ContORBL[1].checked){
		alert("�����Ώۂ�I�����Ă�������");
		return false;
	}
	Num=LTrim(f.ContBLNo.value);
	if(Num.length==0){
		alert("��������ԍ�����͂��Ă�������");
		f.ContBLNo.focus();
		return false;
	}
	if(f.ContORBL[0].checked && !CheckEisuji(f.ContBLNo.value)){
		alert("��������R���e�i�ԍ��ɔ��p�p�����ȊO�̕������w�肵�Ȃ��ł�������");
		f.ContBLNo.focus();
		return false;
	}
	if(f.ContORBL[1].checked && !CheckEisu(f.ContBLNo.value)){
		alert("��������R���e�i�ԍ��ɔ��p�p�����Ɣ��p�X�y�[�X�A�u-�v�A�u/�v�ȊO�̕������w�肵�Ȃ��ł�������");
		f.ContBLNo.focus();
		return false;
	}
	if(f.ContORBL[0].checked){
		parent.DList.SearchC("2",f.ContBLNo.value);
	} else if(f.ContORBL[1].checked){
		parent.DList.SearchC("3",f.ContBLNo.value);
	}
}
//�\�[�g
//function sort(){
//	f = document.search;
//	f.SortFlag.value=f.Sort.options[target.Sort.selectedIndex].value;
//	f.target="DList";
//	f.action="./sst100L.asp";
//	f.submit();
//}
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!-------------�X�e�[�^�X�z�M�˗����ꗗ���Top--------------------------->
<TABLE border="1" cellPadding="3" cellSpacing="0" align=center width="100%">
	<TR>
		<TD width="15%">���O�C�����[�U</TD>
		<TD><%=LinUN%></TD>
	</TR>
</TABLE>
<TABLE border="0" cellPadding="3" cellSpacing="0" width="100%">
	<FORM name="search" action="">
	<TR>
		<TD width="50%"><BR><B class=title>�X�e�[�^�X�z�M�˗����ꗗ</B></TD>
		<TD width="50%">
			<INPUT type="hidden" name="ContBLFlag" value="">
			<INPUT type="radio" name="ContORBL">�R���e�i�ԍ�
			<INPUT type="radio" name="ContORBL">�a�k�ԍ�<BR>
			<INPUT type="text"  name="ContBLNo" maxlength="20">
			<INPUT type="button" value="����" onClick="SearchContBL()">
		</TD>
	<TR>
	</FORM>
</TABLE>
<!-------------��ʏI���--------------------------->
</BODY>
</HTML>
