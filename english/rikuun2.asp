<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' �Z�b�V�����̃`�F�b�N
    CheckLogin "rikunn1.asp"

    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' DB�̐ڑ�
    ConnectSvr conn, rs

    ' ���[�U��ނ��擾����
    strUserKind=Session.Contents("userkind")
    If strUserKind="" Then
        ' �Z�b�V�������؂�Ă���Ƃ�
        Response.Redirect "http://www.hits-h.com/index.asp"             '�g�b�v
        Response.End
    End If

Dim vCtnoE, vCtnoS
Dim sCntNo
Dim sUserID
Dim sSQL
Dim sErrMsg
Dim sErrOpt

sErrMSg = ""
sErrOpt = ""

vCtnoE = Trim(Request.QueryString("cntnrnoe"))
vCtnoS = Trim(Request.QueryString("cntnrnos"))

If (IsEmpty(vCtnoE) Or vCtnoE = "") And (IsEmpty(vCtnoS) Or vCtnoS = "") Then
	sErrMsg = "�R���e�i������"
End If

If sErrMsg = "" Then

	'�Y������R���e�i��T��
	If IsEmpty(vCtnoE) Or vCtnoE = "" Then
		'�R���e�i�ԍ��̐��l�����̂ݓ��͂���Ă���ꍇ
		sSQL = "SELECT RTrim([ContNo]) AS CT FROM Container GROUP BY RTrim([ContNo]), ContNo "
		sSQL = sSQL & "HAVING (((RTrim([ContNo])) Like '%" & vCtnoS & "'))"
	Else
		'�R���e�i�ԍ��̉p�������A���l�����Ƃ��ɓ��͂���Ă���ꍇ
		sSQL = "SELECT RTrim([ContNo]) AS CT FROM Container "
		sSQL = sSQL & "WHERE RTrim([ContNo]) = '" & UCase(vCtnoE) & vCtnoS & "'"
	End If
	rs.Open sSQL, conn, 0, 1, 1
	If rs.Eof Then
		sErrMsg = "�Y���R���e�i�Ȃ�"
		sErrOpt = vCtnoS
	Else
		sCntNo = rs("CT")		'�R���e�i�ԍ��Đݒ�
		rs.MoveNext
		Do While Not rs.EOF
			sCntNo2 = rs("CT")
			rs.MoveNext
			If sCntNo<>sCntNo2 Then
				sErrMsg = "���ŕ�������"
				sErrOpt = vCtnoS
				Exit Do
			End If
		Loop
	End If
	rs.Close

    ' ���^����
	If sErrMsg = "" Then
        WriteLog fs, "6001", "���^����-�R���e�i����", "10", vCtnoE & "/" & vCtnoS & "," & "���͓��e�̐���:0(������)"
		WriteLog fs, "6002", "���^����-������������(Web)", "00", sCntNo & ","
    Else
        WriteLog fs, "6001", "���^����-�R���e�i����", "10", vCtnoE & "/" & vCtnoS & "," & "���͓��e�̐���:1(���)" & sErrMsg
    End If

'	If sErrMsg = "" Then
'		' ���񌟍������R���e�i�ԍ������[�U�e�[�u���ɕۑ�(����Ƀf�t�H���g�ŕ\�������)
'		sSQL = "SELECT lUserTable.BeforeCntnrNo FROM lUserTable WHERE lUserTable.UserID='" & sUserID & "'"
'		rs.Open sSQL, conn, 2, 2
'		If Not rs.Eof Then
'			rs("BeforeCntnrNo") = sCntNo
'			rs.Update
'		End If
'		rs.Close
'	End If

'	conn.Close
End If
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
<%
    If sErrMsg<>"" Then
%>
<table>
          <tr> 
            <td><img src="gif/botan.gif" width="17" height="17"></td>
            <td nowrap><b>�R���e�iNo.����</b></td>
            <td><img src="gif/hr.gif" width="400" height="3"></td>
          </tr>
        </table>
		<br><br>
<%
    DispErrorMessage sErrMsg
%>

<%
    Else
%>
<table>
          <tr> 
            <td><img src="gif/botan.gif" width="17" height="17"></td>
            <td nowrap><b>������Ƒ��M���</b></td>
            <td><img src="gif/hr.gif" width="400" height="3"></td>
          </tr>
        </table>
		<br><br>
        <table border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td nowrap>
              ����������Ƃ�I�����ĉ������B <br>
			  �w����x���N���b�N����ƌ��݂̎��Ԃ����͂���܂��B</td>
          </tr>
        </table>
        <br>
		          <form name=select action="rikuun3.asp">
			  <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
                  <tr> 
                    <td bgcolor="#000099" nowrap><font color="#FFFFFF"><b>�R���e�iNo.</b></font></td>
                    <td nowrap>
<%
    Response.Write sCntNo
    Session.Contents("cntnrno")=sCntNo
%>
                    </td>
                  </tr>
                  <tr> 
                    <td bgcolor="#000099" nowrap><font color="#FFFFFF"><b>�����������</b></font></td>
                    <td nowrap> 
                      <input type="radio" name="operation" value="C" checked>
					  �i�A�o�j��q�ɒ�<br>
                      <input type="radio" name="operation" value="D">
                      �i�A�o�j�o���j���O����<br>
                      <input type="radio" name="operation" value="A">
                      �i�A���j�����q�ɒ�<br>
                      <input type="radio" name="operation" value="B">
                      �i�A���j�f�o������<br>
                    </td>
                  </tr>
                </table>
          <br>
          <input type=submit value="   ����   ">
</form>
<%
    End If
%>
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
 </td>
 </tr>
 </table>
<!-------------�o�^��ʏI���--------------------------->
<%
    DispMenuBarBack "rikuun1.asp"
%>
</body>
</html>