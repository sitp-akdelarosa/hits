<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' �Q�[�g�O�f��
    WriteLog fs, "8007", "�Q�[�g�O�f���E���G�󋵏Ɖ�-ICCT�Q�[�g�O�f��", "00", ","
%>

<html>
<head>
<title>ICCT�Q�[�g�O�f��</title>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
<%
    DispMenuJava
%>
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------��������o�^���--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
  <tr>
  <td valign="top">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
          <td rowspan="2"><img src="gif/terminalt.gif" width="506" height="73"></td>
	      <td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
  </tr>
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
        <table width="95%" cellpadding="0" cellspacing="0" border="0">
          <tr>
            <td align="right"><font color="#333333" size="-1"><%=strRoute%></font></td>
 </tr>
</table>
End of comment by seiko-denki 2003.07.18 -->

<%
    Dim fs, f, strUpdateTime, strFileName, strPath
    Set fs = CreateObject("Scripting.FileSystemObject")
    strFileName="./camera/gate.icct.jpg"
	strPath = Server.MapPath(strFileName)
    Set f = fs.GetFile(strPath)
	dateTimeTmp = f.DateLastModified
	strUpdateTime = Year(dateTimeTmp) & "�N" & _
		Right("0" & Month(dateTimeTmp), 2) & "��" & _
		Right("0" & Day(dateTimeTmp), 2) & "��" & _
		Right("0" & Hour(dateTimeTmp), 2) & "��" & _
		Right("0" & Minute(dateTimeTmp), 2) & "�����݂̏��"
%>

		<table width=95% cellpadding=3>
			<tr>
				<td align=right>
					<font color="#224599">
					&nbsp;&nbsp;<%=strUpdateTime%>
					</font>
				</td>
			</tr>
		</table>

        <table>
          <tr> 
            <td><img src="gif/botan.gif" width="17" height="17"></td>
            <td nowrap><b>ICCT�Q�[�g�O�f��</b></td>
            <td><img src="gif/hr.gif" width="400" height="3"></td>
          </tr>
          <tr>
            <td colspan="3"><img src="camera/gate.icct.jpg" width="600" height="409"></td>
          </tr>
          <tr>
            <td colspan="3">
			<form>
				�L���b�V���@�\�ɂ��Â��摜���\������邱�Ƃ�����܂��̂ŁA
				�ŐV�̕\���ɂ���ɂ͉��̃{�^�����N���b�N���Ă��������B<br>
				<input type=button value="�\���f�[�^�̍X�V" onClick="JavaScript:window.location.reload()">
			</form>
            </td>
          </tr>
        </table>
        <br>
    
</center></td>
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
    DispMenuBarBack "camera.icct.asp"
%>
</body>
</html>