<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
Dim strUser
strUser = Request.QueryString("user")
%>

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">

</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="../gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------��������o�^�R�[�h�ꗗ���--------------------------->

<center>

<BR>

<font size=4><b>�o�^�R�[�h�ꗗ</b></font>

<BR><BR>

<table border=0>
  <tr>
	<td align=left valign=middle nowrap>
		�{�V�X�e���ɓo�^����Ă���R�[�h��<BR>���̂̈ꗗ��\�����܂��B
	</td>
  </tr>
</table>

<BR>

<% If strUser="" Then %>

	<table border=1 cellpadding=3 cellspacing=1 bgcolor="#ffffff">
		<tr>
			<td align=center nowrap>
				<font color="#ff3300"><b>���O�C�����Ă��Ȃ����͕\���ł��܂���B</b></font>
			</td>
		</tr>
	</table>
	<BR>
<% Else %>

�ړI�̃R�[�h��I�����ĉ������B

<BR><BR>

<table border=0>
  <tr>
	<td align=left valign=middle nowrap>
		�P�D <a href="codelist-detail.asp?kind=1">�׎�R�[�h</a>
	</td>
  </tr>
  <tr>
	<td align=left valign=middle nowrap>
		�Q�D <a href="codelist-detail.asp?kind=2">�C�݃R�[�h</a>
	</td>
  </tr>
  <tr>
	<td align=left valign=middle nowrap>
		�R�D <a href="codelist-detail.asp?kind=3">���^�Ǝ҃R�[�h</a>
	</td>
  </tr>
  <tr>
	<td align=left valign=middle nowrap>
		�S�D <a href="codelist-detail.asp?kind=4">�D�ЃR�[�h</a>
	</td>
  </tr>
  <tr>
	<td align=left valign=middle nowrap>
		�T�D <a href="codelist-detail.asp?kind=5">�`�^�R�[�h</a>
	</td>
  </tr>
  <tr>
	<td align=left valign=middle nowrap>
		�U�D <a href="codelist-detail.asp?kind=6">�D�Ђ���̑D���Ɖ�</a>
	</td>
  </tr>
</table>

<BR>

<table border=0 width=85%>
  <tr>
	<td align=left valign=middle>
		��ʂɕ\�������R�[�h���}�E�X�őI�����ăR�s�[�iCtrl + C�j���邱�ƂŁA�L�[���͘g�̂Ƃ���ɒ���t���iCtrl + V�j�ł��܂��B
	</td>
  </tr>
</table>

<% End If %>

<form>
	<input type=button value=" close " onClick="JavaScript:window.close()">
</form>

</center>
</body>
</html>

<%
%>
