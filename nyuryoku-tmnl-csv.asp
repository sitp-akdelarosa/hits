<%@Language="VBScript" %>

<!--#include file="Common.inc"-->
<%
    ' �Z�b�V�����̃`�F�b�N
    CheckLogin "nyuryoku-te.asp"

    ' �Z�b�V�����ϐ����烂�[�h���擾
    strChoice = Trim(Session.Contents("choice"))
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
          <td nowrap><b>
			CSV�t�@�C���]��
			</b></td>
          <td><img src="gif/hr.gif"></td>
        </tr>
      </table>
	     <table>
          <tr>
	        <td nowrap>
                <center><font color="#000066" size="+1">�y�^�[�~�i�����͗p�t�@�C���]����ʁz</font></center><br>
			���͂����������m�F�\�莞�����(
<% If strChoice="bl" Then %>
	BL�P��
<% Else %>
	�{�D�P��
<%  End If %>
                )���܂�CSV�t�@�C����I�����A<BR>���M�{�^�����N���b�N���ĉ������B</td></tr>
		</table>
			<form action="nyuryoku-tmnl-csvin.asp" enctype="multipart/form-data" method="post"> 
				<table border="1" cellspacing="2" cellpadding="3" bgcolor="#ffffff">
					<tr>
						<td bgcolor="#000099" nowrap>
							<font color="#FFFFFF"><b>CSV�t�@�C����</b></font>
						</td>
						<td nowrap> 
							<input type=file name=csvfile size=50 accept="text/css">
						</td>
					</tr>
				</table>
				<br>
				<input type=submit value=" ��  �M ">
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
    DispMenuBarBack "nyuryoku-te1.asp"
%>
</body>
</html>

<%
    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemObject")

	If strChoice="bl" Then
	    WriteLog fs, "5003","�^�[�~�i������-CSV�t�@�C���]��(BL�P��)","00", ","
	Else
	    WriteLog fs, "5005","�^�[�~�i������-CSV�t�@�C���]��(�{�D�P��)","00", ","
	End If
%>
