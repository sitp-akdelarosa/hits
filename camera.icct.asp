<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' �Q�[�g���ӊȈՒn�}
    WriteLog fs, "8006", "�Q�[�g�O�f���E���G�󋵏Ɖ�-ICCT�Q�[�g���ӊȈՒn�}", "00", ","
%>

<html>
<head>
<title>�Q�[�g���ӊȈՒn�}</title>
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
            <td align="right">
              <font color="#333333" size="-1"><%=strRoute%></font>
            </td>
 </tr>
</table>
End of comment by seiko-denki 2003.07.18 -->


<%
    Dim fs, f, strUpdateTime, strFileName, strPath
    Set fs = CreateObject("Scripting.FileSystemObject")
    strFileName="./camera/sam-gate.icct.jpg"
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
            <td nowrap><b>�Q�[�g���ӊȈՒn�}</b></td>
            <td><img src="gif/hr.gif" width="400" height="3"></td>
          </tr>
        </table>
<table border="0"><tr>
            <td align="left"><br>
              <br>
<center>
                <img src="gif/terminalmap.icct.gif" width="440" height="252" usemap="#TarminalMap" border="0"> 
                <br>
                �@<br>

<table border=0 cellpadding=0 cellspacing=0>
 <tr>
  <td align="center" valign=middle>

	<table border=5 cellpadding=0 cellspacing=0 bgcolor="#6666ff" bordercolor="#ffffff">
	 <tr>
	  <td align="center" valign=middle height=20>
		<font color="#ffffff"><b>�Q�[�g�O�f��</b></font>
	  </td>
	 </tr>
	 <tr>
	  <td align=center valign=middle><a href="photogate.icct.asp"><img src="camera/sam-gate.icct.jpg" border="0" width="120" height="82"></a></td>
	 </tr>
	</table>

  </td>
 </tr>
 <tr>
   <td align=center valign=bottom>
	<b>�ʐ^���N���b�N����Ɗg��ł��܂��B</b>
   </td>
 </tr>
</table>
<P>
<table border="0">
 <tr>
   <td align="left">
<form>
				�L���b�V���@�\�ɂ��Â��摜���\������邱�Ƃ�����܂��̂ŁA<br>
				�ŐV�̕\���ɂ���ɂ͉��̃{�^�����N���b�N���Ă��������B<br>
				<input type=button value="�\���f�[�^�̍X�V" onClick="JavaScript:window.location.reload()">
			</form>
   </td>
 </tr>
</table>

</center>
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
    DispMenuBarBack "terminal.asp"
%>
<map name="TarminalMap"> 
  <area shape="rect" coords="70,183,162,209" href="photogate.icct.asp">
</map>
</body>
</html>