<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' �Z�b�V�����̃`�F�b�N
    CheckLogin "nyuryoku.asp"
%>

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT LANGUAGE="JavaScript">
<%
    DispMenuJava
%>
<!--
function gotoURL(){
    var gotoUrl=document.con.select.options[document.con.select.selectedIndex].value
    document.location.href=gotoUrl 
}
-->
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------��������e�Г��͏����--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/nyuryoku.gif" width="506" height="73"></td>
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
      <br>�@
      <br>�@
      <br>�@
      <br>�@
      <center>
      <table>
        <tr>
          <td>�Ǝ��I�����A���M�{�^�����N���b�N���ĉ������B</td>
        </tr>
      </table>
      <form name=con>
        <table border="1" cellspacing="2" cellpadding="3">
          <tr> 
            <td bgcolor="#000099" nowrap><font color="#FFFFFF"><b>�Ǝ�I��</b></font></td>
            <td nowrap><SELECT NAME="select" size="2">
              <option value="nyuryoku-sp.asp">�D��</option>
              <option value="nyuryoku-ki.asp">�C��</option>
<!--
              <option value="nyuryoku-un.htm">�^�����</option>
-->
              </select>
            </td>
          </tr>
        </table>
        <br>�@<br>
        <INPUT TYPE=BUTTON VALUE=" ��  �M " onClick="gotoURL()">
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
<!-------------�e�Г��͏���ʏI���--------------------------->
</body>
</html>

<%
    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' �e�Г��͏��
    WriteLog fs, "�e�Г��͏��", ""
%>
