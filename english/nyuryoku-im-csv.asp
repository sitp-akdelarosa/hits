<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
''    ' �Z�b�V�����̃`�F�b�N
    CheckLogin "nyuryoku-im.asp"
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
<!-------------��������o�^���--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
	<tr>
		<td valign=top>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td rowspan=2><img src="gif/csvt.gif" width="506" height="73"></td>
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
				<table>
					<tr> 
						<td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
						<td nowrap><b>������q�ɓ��������i�w���ρj</b></td>
						<td><img src="gif/hr.gif"></td>
					</tr>
				</table>
				<table>
					<tr>
						<td nowrap>
							<center><font color="#000066" size="+1">�y������q�ɓ��������p�t�@�C���]����ʁz</font></center><br>
							���͂�����������q�ɓ͂����������܂�CSV�t�@�C����I�����A���M�{�^����
								�N���b�N<br>���ĉ������B</td></tr>
				</table>
				<form action="nyuryoku-im-csvin.asp?kind=cntnr" enctype="multipart/form-data" method="post"> 
					<table border="1" cellspacing="2" cellpadding="3" bgcolor="#ffffff">
						<tr> 
							<td bgcolor="#000099" nowrap align="center" background="gif/tableback.gif" colspan="2"><font color="#FFFFFF"><b>�R���e�iNo.�𗅗񂵂�CSV�t�@�C���̓]���̏ꍇ</b></font>
							</td>
						</tr>
						<tr> 
							<td bgcolor="#000099" nowrap>
								<font color="#FFFFFF"><b>CSV�t�@�C����</b></font>
							</td>
							<td nowrap> 
								<input type=file name=csvfile size=50 accept="text/css">
							</td>
						</tr> 
					</table>
					<BR>
	  				<input type=submit value=" ��  �M ">
					<br>
				</form>
				<br><br>
				<form action="nyuryoku-im-csvin.asp?kind=bl" enctype="multipart/form-data" method="post" id=form1 name=form1> 
					<table border="1" cellspacing="2" cellpadding="3" bgcolor="#ffffff">
						<tr> 
							<td bgcolor="#000099" nowrap align="center" background="gif/tableback.gif" colspan="2"><font color="#FFFFFF"><b>BL No.�𗅗񂵂�CSV�t�@�C���̓]���̏ꍇ</b></font>
							</td>
						</tr>
						<tr> 
							<td bgcolor="#000099" nowrap>
								<font color="#FFFFFF"><b>CSV�t�@�C����</b></font>
							</td>
							<td nowrap> 
								<input type=file name=csvfile size=50 accept="text/css">
							</td>
						</tr> 
					</table>
					<BR>
	  				<input type=submit value=" ��  �M " id=submit1 name=submit1>
					<br>
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
<!-------------�o�^��ʏI���--------------------------->
<%
    DispMenuBarBack "nyuryoku-im.asp"
%>
</body>
</html>

<%
    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemObject")

    ' ������q�ɓ͂������w���p�t�@�C���]�����
    WriteLog fs, "4007","�C�ݓ��͎�����q�ɓ�������-CSV�t�@�C���]��","00",","
%>
