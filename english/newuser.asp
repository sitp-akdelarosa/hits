<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' �Z�b�V�����̗L�������`�F�b�N
    Dim strUserID
    strUserID = Trim(Session.Contents("userid"))

    ' �Z�b�V�������L���ȂƂ�
    If strUserID<>"" Then
        ' �߂��ʏ����擾
        strLinkID = Session.Contents("linkid")

        ' �߂��ʂփ��_�C���N�g
        Response.Redirect strLinkID
    Else
        ' �G���[�t���O�̃N���A
        bOK = false
        bError = false

        ' �w������̎擾(��Ж�,���[���A�h���X)
        Dim strCompany
        Dim strMailAddress
        strCompany = Trim(Request.QueryString("campany"))
        strMailAddress = Trim(Request.QueryString("mail"))

        If strCompany<>"" Then
            ' ���[�U�[�h�c�̍ő�`�F�b�N
            ConnectSvr conn, rsd

            sql = "SELECT UserID, CompanyName, MailAddress FROM lUserTable ORDER BY UserID DESC"
            'SQL�𔭍s���ă��[�U�[�h�c������
            rsd.Open sql, conn, 3, 2, 1
            If Not rsd.EOF Then
                strInputUserID = GetNumStr(CInt(rsd("UserID"))+1, 5 )
            Else
                ' ���[�U�[�h�c
                strInputUserID = "00000"
            End If

            rsd.AddNew
            rsd("UserID") = strInputUserID
            rsd("CompanyName") = strCompany
            rsd("MailAddress") = strMailAddress
            rsd.UpDate

            rsd.Close
            conn.Close

            bOK = true
            ' ���[�U�[�h�c���Z�b�V�����ϐ��ɐݒ�
            Session.Contents("userid") = strInputUserID
        Else
            If Trim(Request.QueryString("flg"))<>"" Then
                ' ��Ж��G���[�̂Ƃ�
                bError=true
                strError = "��Ж��͏ȗ��ł��܂���B"
            End If
        End If

        If Not bOK Then
            ' File System Object �̐���
            Set fs=Server.CreateObject("Scripting.FileSystemobject")

            ' ���O�C��
            WriteLog fs, "���O�C��", "�V�K���[�U�h�c���s"
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
<!-------------�������烆�[�U�[�h�c�o�^���--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/idt.gif" width="506" height="73"></td>
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
      <form action="newuser.asp">
        <table>
          <tr> 
            <td nowrap>
              <dl> 
              <dt><font color="#000066" size="+1">�y���[�U�[�h�c�V�K�o�^�z</font><br>
              <dd>���Ȃ��̃��[�U�[�h�c�𔭍s���܂��̂ŁA�ȉ��ɕK�v��������͂��Ă��������B
              <dd><br>
              <dd><br>
                <table border=1 cellspacing=2 cellpadding=3 bgcolor="#FFFFFF">
                  <tr> 
                    <td nowrap bgcolor=#FFCC33><font color="#000000">��Ж�</font></td>
                    <td> 
                      <input type=text name=campany size=50 maxlength=200>
                     �i�K�{���́j</td>
                  </tr>
                  <tr> 
                    <td nowrap bgcolor=#FFCC33><font color="#000000">E-mail</font></td>
                    <td> 
                      <input type=text name=mail size=30 maxlength=200>
                     �i���p�j</td>
                  </tr>
                </table>
              <dd><br></dl>
              <center><input type=hidden name=flg value='1'>
              <input type=submit value=" �o  �^ ">
              </center>
            </td>
          </tr>
        </table>
      </form>
<%
            If bError Then
                ' �G���[���b�Z�[�W�̕\��
                DispErrorMessage strError
            End If
%>
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
<!-------------���[�U�[�h�c�o�^��ʏI���--------------------------->
<%
    DispMenuBarBack "userchk.asp"
%>
</body>
</html>

<%
        Else
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
<!-------------�������烆�[�U�[�h�c�o�^���--------------------------->
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
      <form action="
<%
                ' �߂��ʏ����擾
                strLinkID = Session.Contents("linkid")

                Response.Write strLinkID
%>
      ">
        <table>
          <tr> 
            <td nowrap>
              <dl> 
              <dt><font color="#000066" size="+1">�y���[�U�[�h�c�V�K�o�^�����z</font><br>
              <dd>���Ȃ��̃��[�U�[�h�c��[<font color="red" size="+2">
<%
                ' ���[�U�[�h�c�̕\��
                Response.Write strInputUserID
%>
                  </font>]�ł��B
              <dd>�Y��Ȃ��悤�Ƀ������Ă����Ă��������B
              </dl>
              <br>
              <br>
              <center>
              <input type=submit value=" ��  �s ">
              </center>
            </td>
          </tr>
        </table>
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
<!-------------���[�U�[�h�c�o�^��ʏI���--------------------------->
<%
    DispMenuBarBack "http://www.hits-h.com/index.asp"
%>
</body>
</html>

<%
        End If
    End If
%>
