<%@Language="VBScript" %>

<!--#include file="../Common.inc"-->

<html>
<head>
<title>�X�e�[�^�X�z�M�˗��w���v</title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript"><!--
function LinkSelect(form, sel)
{
	adrs = sel.options[sel.selectedIndex].value;
	if (adrs != "-" ) parent.location.href = adrs;
}
function OpenCodeWin()
{
	var CodeWin;
	CodeWin = window.open("../codelist.asp?user=<%=Session.Contents("userid")%>","codelist","scrollbars=yes,resizable=yes,width=300,height=330");
	CodeWin.focus();
}
// -->
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="image/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------����������--------------------------->
<table border="0" cellspacing="0" cellpadding="0" width="100%" height=100%>
<tr>
	<td valign=top>
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td rowspan=2><img src="image/sst_help.gif" width="506" height="73"></td>
			<td height="25" bgcolor="000099" align="right"><img src="image/logo_hits_ver2.gif" width="300" height="25"></td>
		</tr>
		<tr>
			<td align="right" width="100%" height="48"> 
<%
call	DisplayCodeListButton
%>
			</td>
		</tr>
		</table>
		<center>
		<BR><BR><BR>
		<table border="0">
			<tr>
				<td align="center"> 
					<table border="0" cellspacing="2" cellpadding="3">
						<tr> 
							<td colspan="2" bgcolor="#99ccFF"><b><font color="#000000">�� �X�e�[�^�X�z�M�̐V�K�o�^</font></b></td>
						</tr>
						<tr> 
							<td width="15">�@</td>
							<td width="575">��ʍ����́u�\����ސؑցv���u�V�K�˗��v���N���b�N���A�\��������ʂɂăR���e�i�ԍ��܂��͂a�k�ԍ�����͂��A
															�u�o�^�v�{�^�����N���b�N���܂��B<br>
															�R���e�i���[�h���o�ς̃R���e�i�ԍ����o�^�ł��܂����A���o��P�P���ȏ�o�߂������͓̂o�^�ł��܂���B
															�܂��A�a�k�ԍ��w��̏ꍇ�A�֘A����R���e�i�ԍ������ׂĔ��o��P�P���ȏ�o�߂������͓̂o�^�ł��܂���B<br>
															�a�k�ԍ��w��̏ꍇ�A�g���s�r�o�^��ɂ��̂a�k�ɒǉ����ꂽ�R���e�i�ɂ��Ă̓X�e�[�^�X�𑗐M�ł��܂���B
															�i�΍�j�X�e�[�^�X�z�M���炻�̂a�k�̓o�^����U�폜���A�ēx�o�^���܂��B</td>
						</tr>
						<tr>
							<td colspan="2">�@</td>
						</tr>
						<tr>
							<td colspan="2" bgcolor="#99ccFF"><b><font color="#000000">�� �X�e�[�^�X�z�M�˗����ꗗ�\��</font></b></td>
						</tr>
						<tr>
							<td width="15">�@</td>
							<td width="575">��ʍ����́u�\����ސؑցv���u�˗����ꗗ�v���N���b�N����ƃX�e�[�^�X�z�M�˗����ꗗ���\������܂��B
															�Ȃ��A�R���e�i���[�h���o��P�P���ȏ�o�߂����R���e�i�ԍ��͈ꗗ�ɂ͕\������܂���B�܂��A�a�k�ԍ��w��̏ꍇ�A
															���ׂẴR���e�i�����o��P�P���ȏ�o�߂������͈̂ꗗ�ɂ͕\������܂���B</td>
						</tr>
						<tr>
							<td colspan="2">�@</td>
						</tr>
						<tr> 
							<td colspan="2" bgcolor="#99ccFF"><b><font color="#000000">�� �X�e�[�^�X�z�M�̓o�^�f�[�^�̍폜</font></b></td>
						</tr>
						<tr> 
							<td width="15">�@</td>
							<td width="575">�X�e�[�^�X�z�M�˗����ꗗ��ʂ��폜�������uNo.�v���N���b�N���A
															�\��������ʂɂāu�폜�v�{�^�����N���b�N���܂��B</td>
						</tr>
						<tr>
							<td colspan="2">�@</td>
						</tr>
						<tr> 
							<td colspan="2" bgcolor="#99ccFF"><b><font color="#000000">�� mail�������M</font></b></td>
						</tr>
						<tr> 
							<td width="15">�@</td>
							<td width="575">�V�K�o�^��ʂɂāA�R���e�i�ԍ��܂��͂a�k�ԍ�����͂��āumail�������M�v���N���b�N����ƁA
															�w�肵���R���e�i�܂��͂a�k�Ɋ֘A����R���e�i�̌��݂̏�Ԃ����[���ɂđ��M����܂��B<br>
															�܂��A�ꗗ���獀�Ԃ̃N���b�N�ɂ��\�������ڍ׉�ʂɂāA�umail�������M�v���N���b�N����ƁA
															���l�ɃR���e�i�܂��͂a�k�Ɋ֘A����R���e�i�̌��݂̏�Ԃ����[���ɂđ��M����܂��B<br>
															�Ȃ��A���[���ɂđ��M����ɂ͗A���X�e�[�^�X�z�M�˗��i�ݒ�j��ʂɂă��[���A�h���X��o�^���Ă����K�v������܂��B
															�܂��Amail�������M�ł́A�A���X�e�[�^�X�z�M�˗��i�ݒ�j��ʂ̐ݒ���e�ɌW���Ȃ��A���ׂĂ̍��ڂɂ��ă��[�����M����܂��B</td>
						</tr>
						<tr>
							<td colspan="2">�@</td>
						</tr>
						<tr> 
							<td colspan="2" bgcolor="#99ccFF"><b><font color="#000000">�� �X�e�[�^�X�z�M�Ώۍ��ڂ̐ݒ�</font></b></td>
						</tr>
						<tr> 
							<td width="15">�@</td>
							<td width="575">��ʍ����́u�\����ސؑցv���u�ݒ�v���N���b�N����ƗA���X�e�[�^�X�z�M�˗��i�ݒ�j��ʂ��\������܂��B
															��Ԃ��ω������ꍇ�Ƀ��[�������M����Ă���悤�ɐݒ肵�������ڂ�I��ŁA�u�o�^�v�{�^�����N���b�N���܂��B
															���Y��ʂɂēo�^���ꂽ���[���A�h���X�փ��[�������M����܂��B</td>
						</tr>
						<tr>
							<td colspan="2">�@</td>
						</tr>
						<tr> 
							<td colspan="2" bgcolor="#99ccFF"><b><font color="#000000">�� ����</font></b></td>
						</tr>
						<tr> 
							<td width="15">�@</td>
							<td width="575">�R���e�i�ԍ��܂��͂a�k�ԍ��ɂ�錟�����ł��܂��B�܂��A�����v�������ł��܂��B
															�Ⴆ�΁A�R���e�i�ԍ��Ƃ��āu555]���w�肵�A�u�����v�{�^�����N���b�N�����ꍇ�A
															�uCONT0000555�v�̃R���e�i�ԍ��͒��o�̑ΏۂƂȂ�܂��B<br>
															�R���e�i�ԍ�������A���ɖ߂��Ƃ��́A�����́u�˗����ꗗ�v���N���b�N���Ă��������B</td>
						</tr>
						<tr>
							<td colspan="2">�@</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
		<table border="0">
			<form>
			<tr><td>�@</td></tr>
			<tr><input type="button" value="����" onClick="window.close()"></td></tr>
			</form>
		</table>
		</center>
	</td>
</tr>
</table>
<!-------------��ʏI���--------------------------->
</body>
</html>
