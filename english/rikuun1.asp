<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' �Z�b�V�����̃`�F�b�N
    CheckLogin "rikunn1.asp"

    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' DB�̐ڑ�
    ConnectSvr conn, rsd

    ' ���[�U��ނ��擾����
    strUserKind=Session.Contents("userkind")
    If strUserKind="" Then
        ' �Z�b�V�������؂�Ă���Ƃ�
        Response.Redirect "http://www.hits-h.com/index.asp"             '�g�b�v
        Response.End
    End If

    ' ���^����
    WriteLog fs, "6001", "���^����-�R���e�i����", "00", ","
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
          <td rowspan=2><img src="gif/rikuunt.gif" width="506" height="73"></td>
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
<center>
<table>
          <tr> 
            <td><img src="gif/botan.gif" width="17" height="17"></td>
            <td nowrap><b>�R���e�iNo.����</b></td>
            <td><img src="gif/hr.gif" width="400" height="3"></td>
          </tr>
        </table>
        <br>
		<table border=0 cellpadding=0 nowrap><tr><td>
		�i�A�o�j��q�ɒ��A�i�A�o�j�o���j���O�����A�i�A���j�����q�ɒ��A�i�A���j�f�o�������ɂ��āA<BR>
        ��Ɗ�����������͂���R���e�iNo.�����āA���M�{�^���������ĉ������B <br>
		</td></tr></table>
        <br>
          <form name=select action="rikuun2.asp" method="get">
                <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
                  <tr> 
                    <td bgcolor="#000099" nowrap colspan=2><font color="#FFFFFF"><b>�R���e�iNo.</b></font></td>
                  </tr>
                  <tr> 
                    <td bgcolor="#000099" nowrap><font color="#FFFFFF"><b>�p���S��</b></font></td>
                    <td nowrap> 
                      <input type=text name=cntnrnoe size=6 maxlength="4">
                    </td>
                  </tr>
                  <tr> 
                    <td bgcolor="#000099" nowrap><font color="#FFFFFF"><b>����</b></font></td>
                    <td nowrap> 
                      <input type=text name=cntnrnos size=10 maxlength="8">
                    </td>
                  </tr>
                </table>
          <br>
          <input type=submit value="   ���M   ">
</form>
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
    DispMenuBarBack "http://www.hits-h.com/index.asp"
%>
</body>
</html>