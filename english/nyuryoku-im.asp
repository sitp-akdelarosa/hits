<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    '�Z�b�V�����̃`�F�b�N
    CheckLogin "nyuryoku-kaika.asp"

    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
<%
    DispMenuJava
%>
function ClickSend() {

	return (ChkSend(document.con.ContNo.value, 
					document.con.BLNo.value, 
					document.con.Year.value, 
					document.con.Month.value, 
					document.con.Day.value, 
					document.con.Hour.value, 
					document.con.Min.value));
}
// ���̓`�F�b�N
function ChkSend(sContNo, sBLNo, sYear, sMonth, sDay, sHour, sMin ) {
	if (sContNo == "" && sBLNo == "") {	/* �R���e�iNo.�܂���BL�ԍ������̓`�F�b�N */
		window.alert("�R���e�iNo.�܂���BL�ԍ��������͂ł��B");
		return false;
	}

	if (sContNo != "" && sBLNo != "") {	/* �R���e�iNo.�܂���BL�ԍ������̓`�F�b�N */
		window.alert("�R���e�iNo.�܂���BL�ԍ��̂ǂ��炩����͂��ĉ������B");
		return false;
	}

	if (sYear == "" ||  sMonth == "" || sDay == "" || sHour == "" || sMin == "") {
		window.alert("�͂������������͂ł��B");
		return false;
	}

	if (!(sYear > 0 || sYear <= 0)|| sYear < 1990 || sYear > 2100 ) {	/* �N�̃`�F�b�N */
		window.alert("�͂������̔N�̓��͂��s���ł��B");
		return false;
	}
	if (!(sMonth > 0 || sMonth <= 0)|| sMonth < 1 || sMonth > 12 ) {	/* ���̃`�F�b�N */
		window.alert("�͂������̌��̓��͂��s���ł��B");
		return false;
	}
	if (!(sDay > 0 || sDay <= 0)|| sDay < 1 || sDay > 31  ) {		/* ���̃`�F�b�N */
		window.alert("�͂������̓��̓��͂��s���ł��B");
		return false;
	}

	if (!(sHour > 0 || sHour <= 0)|| sHour < 0 || sHour > 24  ) {	/* ���̃`�F�b�N */
		window.alert("�͂������̎��̓��͂��s���ł��B");
		return false;
	}

	if (!(sMin > 0 || sMin <= 0)|| sMin < 0 || sMin > 59  ) {		/* ���̃`�F�b�N */
		window.alert("�͂������̕��̓��͂��s���ł��B");
		return false;
	}

	if (sDay<=0 || sDay>30+((sMonth==4||sMonth==6||sMonth==9||sMonth==11)?0:1) || 
	   (sMonth==2&&sDay>28+(((sYear%4==0&&sYear%100!=0)||sYear%400==0)?1:0)) ){
		window.alert("�͂������̓��̓��͂��s���ł��B");
		return false;
	}

	return true;
}
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------��������o�^���--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
	<tr>
		<td valign=top>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td rowspan=2><img src="gif/kaika3t.gif" width="506" height="73"></td>
					<td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
				</tr>
				<tr>
					<td align="right" width="100%" height="48">
<%
' Added and Commented by seiko-denki 2003.07.18
	DisplayCodeListButton
'    DispMenu
'	Dim strScriptName,strRoute
'	strScriptName = Request.ServerVariables("SCRIPT_NAME")
'	strRoute = SetRoute(strScriptName)
'	Session.Contents("route") = strRoute
' End of Addition by seiko-denki 2003.07.18
%>
          </td>
        </tr>
      </table>
      <center>
<!-- commented by seiko-denki 2003.07.18
		<table width=95% cellpadding="0" cellspacing="0" border="0">
		  <tr>
			<td align="right">
			  <font color="#333333" size="-1">
				<%=strRoute%>
			  </font>
			</td>
		  </tr>
		</table>
End of comment by seiko-denki 2003.07.18 -->
		<BR>
		<BR>
		<BR>
<table border=0 cellpadding=0 cellspacing=0><tr><td align=left>
				<table>
					<tr> 
						<td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
						<td nowrap><b>�R���e�i������</b></td>
						<td><img src="gif/hr.gif"></td>
					</tr>
				</table>
<center>
				<table>
					<tr>
						<td align=left>���L�̍��ڂ���͂̏�A���M�{�^�����N���b�N���ĉ������B<BR>
						�R���e�iNo.�܂���BLNo.�ƁA�q�ɓ͂������i�׎�w���j�͕K�{���͂ł��B
						</td>
					</tr>
				</table>
				<FORM NAME="con" METHOD="post" action="nyuryoku-im-syori.asp" onSubmit="return ClickSend()">
					<table border=0 cellpadding=0>
						<tr> 
							<td align="center"> 
								<table border="1" cellspacing="1" cellpadding="3" bgcolor="#ffffff" width=100%>
									<tr> 
										<td background="gif/tableback.gif" bgcolor="#000099" nowrap align="center"><b><font color="#FFFFFF">
											�R���e�iNo.</font></b></td>
										<td> 
										<table border=0 cellpadding=0 cellspacing=0>
										  <tr>
											<td width=200>
												<input type="text" name="ContNo" size="22" maxlength="12">
											</td>
											<td align=left valign=middle nowrap>
												<font size=1 color="#2288ff">[ ���p�p�� ]</font>
											</td>
										  </tr>
										</table>
											
										</td>
									</tr>
									<tr> 
										<td align="center" colspan="2">�܂��́A</td>
									</tr>
									<tr> 
										<td background="gif/tableback.gif" bgcolor="#000099" nowrap align="center"><b><font color="#FFFFFF">
											BL No.</font></b></td>
										<td> 
										<table border=0 cellpadding=0 cellspacing=0>
										  <tr>
											<td width=200>
												<input type="text" name="BLNo" size="22" maxlength="20">
											</td>
											<td align=left valign=middle nowrap>
												<font size=1 color="#2288ff">[ ���p�p�� ]</font>
											</td>
										  </tr>
										</table>
											
										</td>
									</tr>
								</table>
<BR>
								<table border="1" cellspacing="1" cellpadding="3" bgcolor="#ffffff" width=100%>
									<tr> 
										<td background="gif/tableback.gif" bgcolor="#000099" nowrap align="center"><b><font color="#FFFFFF">
  											�q�ɓ͂�����</font></b></td>
										<td> 
											<input type=text name="Year" size=4 maxlength="4">�N
											<input type=text name="Month" size=2 maxlength="2">��
											<input type=text name="Day" size=2 maxlength="2">���@
											<input type=text name="Hour" size=2 maxlength="2">��
											<input type=text name="Min" size=2 maxlength="2">��
											<table border=0 cellpadding=0 cellspacing=0>
											  <tr>
												<td width=200>
													&nbsp;&nbsp;&nbsp;<font size=-1>�i��j 2002�N2��25�� 15��30��</font>
												</td>
												<td align=left valign=middle nowrap>
													<font size=1 color="#ee2200">[ �K�{���� ]</font><BR>
													<font size=1 color="#2288ff">[ ���p���l ]</font>
												</td>
											  </tr>
											</table>
										</td>
									</tr>
								</table>
								<br>
									<input type=submit value=" ��  �M " name="���Z�b�g">
							</td>
						</tr>
					</table>
				</form>
</center>
				<table>
					<tr> 
						<td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
						<td nowrap><b>CSV�t�@�C���]��</b></td>
						<td><img src="gif/hr.gif"></td>
					</tr>
				</table>
<center>
				<table border="0" cellspacing="1" cellpadding="2">
					<tr> 
						<td> 
							<p>�����t�@�C���]������ꍇ�͂������N���b�N</p>
						</td>
						<td>�c</td>
						<td><a href="nyuryoku-im-csv.asp">CSV�t�@�C���]��</a></td>
					</tr>
					<tr> 
						<td>CSV�t�@�C���]���ɂ��Ă̐����͂������N���b�N</td>
						<td>�c</td>
						<td><a href="help10.asp">�w���v</a></td>
					</tr>
				</table>
</center>
</td></tr></table>
				<br>
          �@		<br>
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
<!-------------�o�^��ʏI���--------------------------->
<%
    DispMenuBarBack "nyuryoku-kaika.asp"
%>
</body>
</html>
<%
    ' �C�ݓ��͍��ڑI��
    WriteLog fs, "4004","�C�ݓ��͎�����q�ɓ�������","00", ","
%>