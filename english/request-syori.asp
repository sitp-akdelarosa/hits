<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%

    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemObject")

    ' BASP21 �R���|�[�l���g�̍쐬
    Set bsp=Server.CreateObject("basp21")

'    �p�����[�^
'      sSvrName		: SMTP�T�[�o�[��
'      sMailto		: ����i���於<���惁�[���A�h���X>�j
'      sMailfrom	: ���M�ҁi���M�Җ�<���M�҃��[���A�h���X>�j
'      sSubj		: ����
'      sBody		: �{��

	sSvrName = "mail.cont-info.com"
	sMailto = "request-syori.asp<cont-info@cont-info.com>"
'	if Request.form("mail") ="" then
		sMailfrom = "�A���P�[�g<NoName@cont-info.com>"
'	else
'		sMailfrom = Request.form("mail")
'	end if
	sSubj = "�y���p�҃A���P�[�g�z"
	sBody = "�y���p�҃A���P�[�g�z" & vbcrlf & vbcrlf & _ 
	        "��Ж�," & Request.form("name1") & vbcrlf & _ 
			"���Z��,��" & Request.form("posta1") & "-" & Request.form("posta1") & vbtab & _
			              Request.form("ken") & Request.form("address1") & vbcrlf & _  
			"���S���Җ�," & Request.form("tantouname") & vbcrlf & _
			"E-mail," & Request.form("mail") & vbcrlf & _
			"�V�X�e���̏��ɂ��ĂP," & Request.form("radiobutton") & vbcrlf & _
			"�Q," & Request.form("add") & vbcrlf & _
			"�R," & Request.form("change")
	rc = bsp.SendMail(sSvrName,sMailto,sMailfrom, sSubj,sBody,"")

	if trim(rc) = "" then
		strError = "����ɑ��M����܂����B"
	    WriteLog fs, "9001", "���p�҃A���P�[�g�EQ&A", "10", "," & "���M�̌���:0(����)"
	else	
		strError = "����ɑ��M�ł��܂���ł����B" 
	    WriteLog fs, "9001", "���p�҃A���P�[�g�EQ&A", "10", "," & "���M�̌���:1(���s)" & rc
	end if	

%>

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
<%
    DispMenuJava
%>
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------��������G���[���--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
	<tr>
		<td valign=top>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td rowspan=2><img src="gif/requestt.gif" width="506" height="73"></td>
					<td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
				</tr>
				<tr>
					<td align="right" width="100%" height="48"> 
<%
' Added and Commented by seiko-denki 2003.07.18
	DisplayCodeListButton
'    DispMenu
' End of Addition by seiko-denki 2003.07.18
%>
					</td>
				</tr>
			</table>
			<br>
			<br>�@
			<br>�@
			<br>�@
			<center>
				<table>
					<tr> 
						<td nowrap>
							<dl> 
								<dt><font color="#000066" size="+1">�y���p�҃A���P�[�g���M��ʁz</font><br>
								<dd>
<%
    ' �G���[���b�Z�[�W�̕\��
    DispErrorMessage strError 
%>
							</dl>
						</td>
					</tr>
				</table>
				<form>
					<br><br>
					<input type="button" value=" ��  �� " onclick="history.back()">
				</form>
			</center>
			<br>
		</td>
	</tr>
	<tr>
		<td valign="bottom"> 
<%
    DispMenuBar
%>
		</td>
	</tr>
</table>
<!-------------�G���[��ʏI���--------------------------->
<%
    DispMenuBarBack "request.asp"
%>
</body>
</html>

