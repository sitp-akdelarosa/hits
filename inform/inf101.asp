<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits                                          _/
'_/	FileName	:inf101.asp                                      _/
'_/	Function	:���m�点���[���A�h���X�o�^���                  _/
'_/	Date			:2005/03/03                                      _/
'_/	Code By		:aspLand HARA                                    _/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<!--#include File="Common.inc"-->
<%
	'''HTTP�R���e���c�^�C�v�ݒ�
	Response.ContentType = "text/html; charset=Shift_JIS"

	Dim EMAIL
	EMAIL=Request.Form("email")

	'''�G���[�g���b�v�J�n
	on error resume next
	'''DB�ڑ�
	Dim cn, rs, sql, cnt
	ConnDBH cn, rs

	sql="select * from send_information where email='" & EMAIL & "'"
	rs.open sql, cn, 3, 1
	if err <> 0 then
		DisConnDBH cn, rs	'DB�ؒf
		response.write("inf101.asp:send_information�e�[�u���A�N�Z�X�G���[!")
		response.end
	end if

	Dim GROUP_CODE, COMPANY_NAME, NAME, TEL, ADDRESS, exist_flag
	exist_flag = 0
	if rs.RecordCount > 0 then
		GROUP_CODE = rs("group_code")
		COMPANY_NAME = rs("company_name")
		NAME = rs("user_name")
		TEL = rs("tel")
		ADDRESS = rs("address")
		exist_flag = 1
	end if
	rs.close

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>���[���A�h���X�o�^</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./js/common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
window.resizeTo(500,350);
window.focus();

function GoEntry(){
	f=document.inf101;

	if(f.dantai.options[f.dantai.selectedIndex].value == ""){
		alert("�Ǝ��I�����Ă��������B");
		f.dantai.focus();
		return false;
	}
	if(f.company_name.value == ""){
		alert("��Ж�����͂��Ă��������B");
		f.company_name.focus();
		return false;
	}
	if(f.user_name.value == ""){
		alert("��������͂��Ă��������B");
		f.user_name.focus();
		return false;
	}
	if(f.tel.value == ""){
		alert("�d�b�ԍ�����͂��Ă��������B");
		f.tel.focus();
		return false;
	}else{
		if(!checkPhoneNumber(f.tel.value)){
			alert("�d�b�ԍ��͔��p�����œ��͂��Ă��������B");
			f.tel.focus();
			return false;
		}
	}
	if(f.address.value == ""){
		alert("�Z������͂��Ă��������B");
		f.address.focus();
		return false;
	}

	if(<%=exist_flag%>){
		if(confirm("�X�V���܂��B��낵���ł����H")){
			f.action="inf103.asp";
			f.submit();
		}else{
			return false;
		}
	}else{
		if(confirm("�o�^���܂��B��낵���ł����H")){
			f.action="inf102.asp";
			f.submit();
		}else{
			return false;
		}
	}
}
function checkPhoneNumber(a){
	if(a==""){
		return(true);
	}
	var b=a.replace(/[0-9\-]/g,'');
	if(b.length!=0){
		return(false);
	}
	return(true);
}
function GoDelete(){
	f=document.inf101;
	if(!<%=exist_flag%>){
		alert("�܂��o�^����Ă��܂���̂ō폜�͖����ł��I");
		return false;
	}
	if(confirm("�폜���܂��B��낵���ł����H")){
		f.action="inf104.asp";
			f.submit();
	}else{
		return false;
	}
}
// -->
</SCRIPT>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY bgcolor="DEE1FF" text="#000000" link="#3300FF" leftMargin="0" topMargin="0" marginheight="0" marginwidth="0">
<!-------------���[���A�h���X���o�^���--------------------------->
<% Session.Contents("InsertSubmitted")="False"  %>
<% Session.Contents("UpdateSubmitted")="False"  %>
<% Session.Contents("DeleteSubmitted")="False"  %>
<TABLE border="0" cellPadding="3" cellSpacing="1" width="100%">
<FORM name="inf101" method="POST">
<input type="hidden" name="email" value="<%=EMAIL%>">
	<TR>
		<TD colspan="2">
			<b>
				<font color="navy">
					���V�K�o�^�A�܂��́A�X�V�̏ꍇ<br>
					�@�@�@���L�����ׂē��͂��āu�o�^�v�{�^���������Ă��������B<br>
					���폜�̏ꍇ<br>
					�@�@�@�u�폜�v�{�^���������Ă��������B
				</font>
			</b>
		</TD>
	</TR>
<% if exist_flag=0 then %>
	<tr><td colspan="2" align="center"><font color="red">�V�K�o�^�˗��ł�</font></td></tr>
<% end if %>
	<TR>
		<TD align="right">���[���A�h���X�F</TD>
		<TD>
			<%=EMAIL%>
		</TD>
	</TR>
	<TR>
		<TD width="25%" align="right">�Ǝ�F</TD>
		<TD width="75%">

			<select name="dantai">
				<option value="">--�I�����Ă�������--</option>
<%
					sql = "select * from group_name order by group_code"
					rs.open sql,cn,3,1
					if err <> 0 then
						'''DB�ؒf
						DisConnDBH cn, rs
						response.write("inf101.asp:group_name�e�[�u���A�N�Z�X�G���[!")
						response.end
					end if
					while not rs.EOF
						if GROUP_CODE = rs("group_code") then
%>
							<option value="<%=rs("group_code")%>" selected><%=rs("group_name")%></option>
<%					else	%>
							<option value="<%=rs("group_code")%>"><%=rs("group_name")%></option>
<%					end if
						rs.movenext
					wend
					rs.close
					'''DB�ڑ�����
					DisConnDBH cn, rs
					'''�G���[�g���b�v����
					on error goto 0
%>
			</select>
		</TD>
	</TR>
	<TR>
		<TD align="right">��Ж��F</TD>
		<TD>
			<INPUT type="text" name="company_name" value="<%=COMPANY_NAME%>" size="45" maxlength="25">
		</TD>
	</TR>
	<TR>
		<TD align="right">�����F</TD>
		<TD>
			<INPUT type="text" name="user_name" value="<%=NAME%>" size="17" maxlength="10">
		</TD>
	</TR>
	<TR>
		<TD align="right">�A����(�d�b�ԍ�)�F</TD>
		<TD>
			<INPUT type="text" name="tel" value="<%=TEL%>" size="17" maxlength="13">&nbsp;(�L����F092-123-4567)
		</TD>
	</TR>
	<TR>
		<TD align="right">�Z���F</TD>
		<TD>
			<INPUT type="text" name="address" value="<%=ADDRESS%>" size="60" maxlength="50">
		</TD>
	</TR>
	<TR>
		<TD colspan="2" align="center">
			<INPUT type="button" value="�߂�" onClick="javascript:history.back();">�@�@
			<INPUT type="button" value="�o�^" onClick="GoEntry()">�@
			<INPUT type="button" value="�폜" onClick="GoDelete()">�@�@
			<INPUT type="button" value="���~" onClick="window.close()">
		</TD>
	</TR>
</FORM>
</TABLE>
<!-------------��ʏI���--------------------------->
</BODY>
</HTML>
