<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' DB�̐ڑ�
    ConnectSvr conn, rsd

    ' ���[�U��ނ��擾����
    strUserKind=Session.Contents("userkind")
    If strUserKind="" Then
        ' �Z�b�V�������؂�Ă���Ƃ�
        Response.Redirect "http://www.hits-h.com/index.asp"             '�A�o�R���e�i�Ɖ�g�b�v
        Response.End
    End If

	Dim iNum
	If strUserKind="�C��" Then
		iNum = "a101"
	ElseIf strUserKind="���^" Then
		iNum = "a102"
	ElseIf strUserKind="�׎�" Then
		iNum = "a103"
	Else
		iNum = "a104"
	End If
    ' �A�o���Ɩ��x��-�A�o�R���e�i�Ɖ�
    WriteLog fs, iNum,"��R���s�b�N�A�b�v�V�X�e��-" & strUserKind & "�p�Ɖ�","00", ","
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
	sVslCode=document.con.vessel.value;
	sVoyCode=document.con.voyage.value;
	if ((sVslCode!="" && sVoyCode=="")||(sVslCode=="" && sVoyCode!="")) {	/* �D�̃`�F�b�N */
			window.alert("�D��(�R�[���T�C��)��Voyage No.�̓y�A�œ��͂��Ă��������B");
			return false;
	}
	return true;
}
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------��������Ɖ���--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
<td rowspan=2><%

    If strUserKind="�C��" Then
        Response.Write "<img src='gif/pickkat.gif' width='506' height='73'>"
    ElseIf strUserKind="���^" Then
        Response.Write "<img src='gif/pickrit.gif' width='506' height='73'>"
    ElseIf strUserKind="�׎�" Then
        Response.Write "<img src='gif/picknit.gif' width='506' height='73'>"
    Else
        Response.Write "<img src='gif/pickkot.gif' width='506' height='73'>"
    End If

%></td>
          <td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
        </tr>
        <tr>
          <td align="right" width="100%" height="48"> 
<%
' Added and Commented by seiko-denki 2003.07.17
	DisplayCodeListButton
'    DispMenu
'	Dim strScriptName,strRoute
'	strScriptName = Request.ServerVariables("SCRIPT_NAME")
'	strRoute = SetRoute(strScriptName)
'	Session.Contents("route") = strRoute
' End of Addition by seiko-denki 2003.07.17
%>
          </td>
        </tr>
      </table>
      <center>
<!-- commented by seiko-denki 2003.07.17
		<table width=95% cellpadding="0" cellspacing="0" border="0">
		  <tr>
			<td align="right">
			  <font color="#333333" size="-1">
				<%=strRoute%>
			  </font>
			</td>
		  </tr>
		</table>
End of comment by seiko-denki 2003.07.17 -->
		<BR>
		<BR>
		<BR>

<% If strUserKind="�C��" Or strUserKind="�׎�" Then %>
		<table border=0 cellpadding=0 cellspacing=0><tr><td>
<% End If %>

      <table>
        <tr>
          <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
          <td nowrap><b>��R���s�b�N�A�b�v���m�F
<%
    If strUserKind="�C��" Then
        Response.Write "(�C�ݗp)"
    ElseIf strUserKind="���^" Then
        Response.Write "(���^�p)"
    ElseIf strUserKind="�׎�" Then
        Response.Write "(�׎�p)"
    Else
        Response.Write "(�`�^�p)"
    End If
%>
            </b></td>
          <td><img src="gif/hr.gif" width=350 height=3></td>
        </tr>
        <tr>
        </tr>

      </table>

<% If strUserKind="�C��" Or strUserKind="�׎�" Then %>
		</td></tr><tr><td align=center>
<% End If %>

      <table width="480">
        <tr>
          <td colspan="4">
			��������͂��Ȃ��ŏ��m�F�{�^���������ƁA�S�Ẵf�[�^���\������܂��B<BR><BR>
            �f�[�^�������ꍇ�͕\���ł��Ȃ���������܂��̂ŁA
			���̏ꍇ�͉��L�t�H�[���ɓK���Ȋm�F��������͂��A
			���m�F�{�^���������ĉ������B
          </td>
        </tr>
      </table>
<%
    If strUserKind<>"���^" Then
%>
      <form name="con" method="get" action="pickcheck.asp" onSubmit="return ClickSend()">
<%
    Else
%>
      <form name="con" method="get" action="pickcheck.asp">
<%
    End If
%>
              <table border="1" cellspacing="1" cellpadding="3" bgcolor="#ffffff">
<%
    If strUserKind<>"���^" Then
%>
                <tr>
                  <td bgcolor="#000099" nowrap>
                    <table border=0 cellpaddig=0 cellspacing=0>
                      <tr><td><font color="#FFFFFF"><b>�D��(�R�[���T�C��)</b></font></td></tr>
                      <tr><td><font color="#FFFFFF"><b>Voyage No.</b></font></td></tr>
                    </table>
                    </td>
                  <td nowrap>
                    <table border=0 cellpaddig=0 cellspacing=0>
                    <tr>
						<td width=150><input type=text name=vessel size=10 maxlength="7"></td>
						<td><font size="1" color="#2288ff">[���p�p��]</font></td>
					</tr>
                    <tr>
						<td width=150><input type=text name=voyage size=18 maxlength="12"></td>
						<td><font size=1 color="#2288ff">[���p�p��]</font></td>
					</tr>
                    </table>
                  </td>
                </tr>
<%
    End If
    If strUserKind="�C��" Then
%>
                <tr>
                  <td bgcolor="#000099" nowrap><font color="#FFFFFF"><b>�׎�R�[�h</b></font></td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=150>
		                    <input type=text name=ninushi size=8 maxlength="5"> 
						</td>
						<td>
							<font size=1 color="#2288ff">[���p�p��]</font>
						</td>
					  </tr>
					</table>
                  </td>
                </tr>
                <tr>
                  <td bgcolor="#000099" nowrap><font color="#FFFFFF"><b>�w�藤�^�Ǝ҃R�[�h</b></font></td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=150>
		                    <input type=text name=rikuun size=5 maxlength="3">
						</td>
						<td>
							<font size=1 color="#2288ff">[���p�p��]</font>
						</td>
					  </tr>
					</table>
                  </td>
                </tr>
<%
    End If
%>
<%
    If strUserKind<>"�C��" Then
%>
                <tr>
                  <td bgcolor="#000099" nowrap><font color="#FFFFFF"><b>�C�݃R�[�h</b></font></td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=150>
							<input type=text name=kaika size=8 maxlength="5">
						</td>
						<td align=right valign=middle nowrap>
							<font size=1 color="#2288ff">[���p�p��]</font>
						</td>
					  </tr>
					</table>
                  </td>
                </tr>
<%
    End If
%>
<%
    If strUserKind="�`�^" Then
%>
                <tr>
                  <td bgcolor="#000099" nowrap><font color="#FFFFFF"><b>���o�w���</b></font></td>
                  <td nowrap>
                    <input type=text name=decyear size=5 maxlength=4>�N
                    <input type=text name=decmon size=3 maxlength=2>��
                    <input type=text name=decday size=3 maxlength=2>��<BR>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=150>
							&nbsp;&nbsp;&nbsp;<font size=-1>�i��j 2002�N2��25��</font>
						</td>
						<td align=left valign=middle nowrap>
							<font size=1 color="#2288ff">[ ���p���l ]</font>
						</td>
					  </tr>
					</table>
                  </td>
                </tr>
<%
    End If
%>
              </table>
              <br>
              <input type=submit value="   ���m�F   ">
      </form>

<% If strUserKind="�C��" Or strUserKind="�׎�" Then %>

		</td></tr><tr><td>

      <table>
        <tr>
          <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
          <td nowrap><b>��R���s�b�N�A�b�v�˗�����
<%
    If strUserKind="�C��" Then
        Response.Write "(�C�ݗp)"
    ElseIf strUserKind="�׎�" Then
        Response.Write "(�׎�p)"
    End If
%>
            </b></td>
          <td><img src="gif/hr.gif" width=350 height=3></td>
        </tr>
        <tr>
        </tr>

      </table>

		</td></tr><tr><td align=center>

	<form action="pickexpinfo.asp">
		��R���s�b�N�A�b�v������͂���ꍇ�́A�����̓{�^���������ĉ������B<BR><BR>
		<input type=submit value="   ������   ">
	</form>

		</td></tr></table>
<% End If %>

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
<!-------------�Ɖ��ʏI���--------------------------->
<%
    DispMenuBarBack "http://www.hits-h.com/index.asp"
%>
</body>
</html>
