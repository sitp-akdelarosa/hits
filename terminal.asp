<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
''    ' �Z�b�V�����̃`�F�b�N
''    CheckLogin "terminal.asp"

    ' �^�[�~�i����񃌃R�[�h�̎擾
    ConnectSvr conn, rsd

    sql = "SELECT RecWaitTime, DelWaitTime, RDWaitTime FROM Terminal WHERE Terminal='KA'"
    'SQL�𔭍s���ă^�[�~�i����񃌃R�[�h������
    rsd.Open sql, conn, 0, 1, 1
    If Not rsd.EOF Then
        iRecWaitTime = rsd("RecWaitTime")
        iDelWaitTime = rsd("DelWaitTime")
        iRDWaitTime = rsd("RDWaitTime")
    End If
    rsd.Close
'ADD START HiTS Ver2 By SEIKO N.Ooshige
    dim IcInTime,IcOutTime
    sql = "SELECT RecWaitTime, DelInWaitTime,DelOutWaitTime FROM Terminal2 WHERE Terminal='IC'"
    'SQL�𔭍s���ăA�C�����h�V�e�B�^�[�~�i����񃌃R�[�h������
    rsd.Open sql, conn, 0, 1, 1
    If Not rsd.EOF Then
        IcInTime  = rsd("RecWaitTime")					'����(IN��t����Ɗ���)
        IcOutTime = rsd("DelInWaitTime") + rsd("DelOutWaitTime")	'���o(IN��t��OUT��������)
    End If
    rsd.Close
'ADD END HiTS Ver2 By SEIKO N.Ooshige
    conn.Close
%>

<html>
<head>
<title></title>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
<%
    DispMenuJava
%>
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------��������^�[�~�i�����v���ԉ��--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/terminalt.gif" width="506" height="73"></td>
          <td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
        </tr>
        <tr>
          <td align="right" width="100%" height="48">
<%
' Added and Commented by seiko-denki 2003.07.07
	DisplayCodeListButton
'    DispMenu
'	Dim strScriptName,strRoute
'	strScriptName = Request.ServerVariables("SCRIPT_NAME")
'	strRoute = SetRoute(strScriptName)
'	Session.Contents("route") = strRoute
' End of Addition by seiko-denki 2003.07.07
%>
          </td>
        </tr>
      </table>
      <center>
		<table width=95% cellpadding="0" cellspacing="0" border="0">
		  <tr>
			<td align="right">
			  <font color="#333333" size="-1">
				<%=strRoute%>
			  </font>
			</td>
		  </tr>
		</table>

		<table width=95% cellpadding=3>
			<tr>
				<td align=right>
					<font color="#224599">
<%
	strNowTime = Year(Now) & "�N" & _
		Right("0" & Month(Now), 2) & "��" & _
		Right("0" & Day(Now), 2) & "��" & _
		Right("0" & Hour(Now), 2) & "��" & _
		Right("0" & Minute(Now), 2) & "�����݂̏��"

%>
					&nbsp;&nbsp;<%=strNowTime%>
					</font>
				</td>
			</tr>
		</table>

      <table border=0>
        <tr>
          <td align=left colspan="2">
            <br>
            <table>
              <tr> 
                <td><img src="gif/botan.gif" width="17" height="17"></td>
                <td nowrap><b>���Ńp�[�N�|�[�g�R���e�i�^�[�~�i��</b></td>
                <td><img src="gif/hr.gif" width="300"></td>
              </tr>
            </table>
          </td></tr>
          <tr><td width="80"></td><td>
<!--
            <center>
			<BR>

			  <table border="0" cellspacing="0" cellpadding="0" width="400">
 				<tr>
				  <td lign=left>
					�ߋ��P���Ԃ̃f�[�^�ŎZ�o���Ă��܂��B<BR>
					�Q�[�g�I�����Ńf�[�^�������Ȃ��ꍇ�A�l�͕\������܂���B
				  </td>
				</tr>
			  </table>
				<BR>
-->
                <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF" width="430">
                  <tr align="center"> 
                    <td nowrap bgcolor="#FFCC33" colspan="3">�^�[�~�i�������v����</td>
                    <td nowrap bgcolor="#FFCC33" rowspan="2">�Q�[�g�O�J�����f��</td>
                  </tr>
                  <tr align="center"> 
                    <td nowrap bgcolor="#FFFFCC"> �����̂� </td>
                    <td nowrap align="center"  bgcolor="#FFFFCC"> ���o�̂� </td>
                    <td nowrap align="center"  bgcolor="#FFFFCC"> ���o�� </td>
                  </tr>
                  <tr align="center"> 
                <td nowrap bgcolor="#FFFFFF" width=90>
<% ' �����҂�����
    If iRecWaitTime>120 Then
        Response.Write "<hr width=80%" & ">"
    Else
        Response.Write iRecWaitTime & "��"
    End If
%>
                </td>
                <td nowrap bgcolor="#FFFFFF" width=90>
<% ' ���o�҂�����
    If iDelWaitTime>120 Then
        Response.Write "<hr width=80%" & ">"
    Else
        Response.Write iDelWaitTime & "��"
    End If
%>
                </td>
                <td nowrap bgcolor="#FFFFFF" width=90>
<% ' ���o���҂�����
    If iRDWaitTime>120 Then
        Response.Write "<hr width=80%" & ">"
    Else
        Response.Write iRDWaitTime & "��"
    End If
%>
                </td>
                <td nowrap align="center"  bgcolor="#FFFFFF">
					<a href="camera.asp"><img src="gif/camera.gif" width="38" height="35" border="0"></a>
				</td>
              </tr>
            </table>
<%'ADD START HiTS Ver.2 By SEIKO N.Ooshige %>
         </td></tr>
         <tr><td align=left colspan="2">
            <br>
            <table>
              <tr> 
                <td><img src="gif/botan.gif" width="17" height="17"></td>
                <td nowrap><b>�A�C�����h�V�e�B�R���e�i�^�[�~�i��</b></td>
                <td><img src="gif/hr.gif" width="300"></td>
              </tr>
            </table>
         </td></tr>
         <tr><td width="80"></td><td>
                  <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF" width="500">
                    <tr align="center"> 
                      <td nowrap bgcolor="#FFCC33" colspan="2">�^�[�~�i�������v����</td>
                      <td nowrap bgcolor="#FFCC33" rowspan="2">�Q�[�g�O�J�����f��</td>
                    </tr>
                    <tr align="center"> 
                      <td nowrap bgcolor="#FFFFCC"> ����(IN��t����Ɗ���) </td>
                      <td nowrap align="center"  bgcolor="#FFFFCC"> ���o(IN��t��OUT��������) </td>
                    </tr>
                    <tr align="center"> 
                  <td nowrap bgcolor="#FFFFFF">
<% ' IC�����҂�����
    If IcInTime<2 or IcInTime>240 Then
        Response.Write "<hr width=80%" & ">"
    Else
        Response.Write IcInTime & "��"
    End If
%>
                  </td>
                  <td nowrap bgcolor="#FFFFFF">
<% ' IC���o�҂�����
    If IcOutTime<2 or IcOutTime>240 Then
        Response.Write "<hr width=80%" & ">"
    Else
        Response.Write IcOutTime & "��"
    End If
%>
                  </td>
                  <td nowrap align="center"  bgcolor="#FFFFFF">
                    <a href="camera.icct.asp"><img src="gif/camera.gif" width="38" height="35" border="0"></a>
                  </td>
                  </tr>
            </table>
         </td></tr>
		 <tr><td>�@</td><td></td></tr>
         <tr><td colspan="2">
<%
'			<form>
'			  <table border="0" cellspacing="0" cellpadding="0" width="430">
' 				<tr>
'				  <td lign=left>
'					<input type=button value='�\���f�[�^�̍X�V' OnClick="JavaScript:location.reload()">
'				  </td>
'				</tr>
'			  </table>
'			</form>
%>
		<form>
			<P>
				�ߋ��P���Ԃ̃f�[�^�ŎZ�o���Ă��܂��B<BR>
				�Q�[�g�I�����Ńf�[�^�������Ȃ��ꍇ�A�l�͕\������܂���B<BR>
				<input type=button value='�\���f�[�^�̍X�V' OnClick="JavaScript:location.reload()">
			</P>
		</form>
<%'ADD END HiTS Ver.2 By SEIKO N.Ooshige %>
            <table>
              <tr> 
                <td><img src="gif/botan.gif" width="17" height="17"></td>
                <td nowrap><b>���ӓ��H�󋵃����N</b></td>
                <td><img src="gif/hr.gif" width="300"></td>
              </tr>
            </table><br>
            <center>
            <table>
              <tr>
                <td><a href="linklog.asp?link=http://www.fk-tosikou.or.jp" target="_blank">�����k��B�������H����</a></td>
              </tr>
              <tr>
                <td><a href="linklog.asp?link=http://www.jartic.or.jp" target="_blank">�i���j���{���H��ʏ��Z���^�[</a></td>
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
<!-------------�^�[�~�i�����v���ԉ�ʏI���--------------------------->
<%
    DispMenuBarBack "http://www.hits-h.com/index.asp"
%>
</body>
</html>

<%
    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' �^�[�~�i�����v���ԏƉ�
    WriteLog fs, "8001", "�Q�[�g�O�f���E���G�󋵏Ɖ�-�Q�[�g�����v����", "00", ","
%>
