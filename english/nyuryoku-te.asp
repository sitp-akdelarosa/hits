<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' �Z�b�V�����̃`�F�b�N
    CheckLogin "nyuryoku-te.asp"

    ' ���̓t���O�̃N���A
    bInput = true

    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' �w������̎擾
    Dim strChoice
    strChoice = Request.QueryString("choice")
    If strChoice<>"" Then
        bInput = false
    End If

    ' �Z�b�V�����ϐ�����`�^�R�[�h���擾
    strOperator = Trim(Session.Contents("userid"))

    If bInput Then
        ' �����m�F�\�莞������
        WriteLog fs, "5001", "�^�[�~�i������", "00", ","
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
<!-------------��������`�^�o�^���--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/terminal2t.gif" width="506" height="73"></td>
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
      <table border=0>
        <tr>
          <td align=left>
            <table>
              <tr> 
                <td><img src="gif/botan.gif" width="17" height="17"></td>
                  <td nowrap><b>�����m�F�\�莞������</b></td>
                  <td><img src="gif/hr.gif"></td>
              </tr>
            </table>
            <FORM NAME="con" action="nyuryoku-te.asp">
            <br>
            <center>
              �����m�F�\�莞������͂��܂��B<br>
              ���̂����ꂩ�̕��@��I�����āw���M�x�{�^�����N���b�N���Ă��������B<br>
              <br>
              <table border="0" cellspacing="2" cellpadding="3">
                <tr>
                  <td>
                    <input type="radio" name="choice" value="bl" checked>��(BL�P��)
                  </td>
                </tr>
                <tr>
                  <td>
                    <input type="radio" name="choice" value="vsl">�ꊇ(�{�D�P��)
                  </td>
                </tr>
              </table>
              <br>
              <br>
                <input type=submit value=" ���@�M "><br><br><br>
              </center>
            </form>
          </td>
        </tr>
      </table>
      <BR>
      <br><br>
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
<!-------------�`�^�o�^��ʏI���--------------------------->
<%
    DispMenuBarBack "http://www.hits-h.com/index.asp"
%>

</body>
</html>

<%
    Else
        Session.Contents("choice")=strChoice
        ' �����m�F�\�莞�����͉�ʂփ��_�C���N�g
        Response.Redirect "nyuryoku-te1.asp"    '�����m�F�\�莞�����͉��
    End If
%>
