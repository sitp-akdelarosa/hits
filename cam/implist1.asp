<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' Temp�t�@�C�������̃`�F�b�N
    CheckTempFile "IMPORT", "impentry.asp"

    ' �\�����[�h�̎擾
    Dim bDispMode          ' true=�R���e�i���� / false=BL����
    If Session.Contents("findkind")="Cntnr" Then
        bDispMode = true
    Else
        bDispMode = false
    End If

    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' �\���t�@�C���̎擾
    Dim strFileName
    strFileName = Session.Contents("tempfile")
    If strFileName="" Then
        ' �Z�b�V�������؂�Ă���Ƃ�
        Response.Redirect "impentry.asp"             '�A���R���e�i�Ɖ�g�b�v
        Response.End
    End If
    strFileName="./temp/" & strFileName

    ' �A���R���e�i�Ɖ�X�g�\��
    WriteLog fs, "2304","�A���R���e�i�Ɖ�-�����܂ł̈ʒu���","00", ","

    ' �\���t�@�C����Open
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

    '�߂��ʎ�ʂ��L��
    Session.Contents("dispreturn")=1
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
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="../gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------��������ꗗ���--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="../gif/implistt.gif" width="506" height="73"></td>
          <td height="25" bgcolor="000099" align="right"><img src="../gif/logo_hits_ver2.gif" width="300" height="25"></td>
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

		<table width=95% cellpadding=3>
			<tr>
				<td align=right>
					<font color="#224599">
					&nbsp;&nbsp;<%=GetUpdateTime(fs)%>
					</font>
				</td>
			</tr>
		</table>

      <table>
        <tr>
          <td> 
            <table>
              <tr>
                <td><img src="../gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>�^�[�~�i�������܂ł̈ʒu���&nbsp;</b></td>
                <td><img src="../gif/hr.gif"></td>
              </tr>
            </table>
            <br>

        <table border="0" cellspacing="2" cellpadding="1">
          <tr> 
            <td width="15"><BR></td>
            <td><font color="#000000" size="-1">�i��1) �N���b�N�ŒP�ƃR���e�i����\��</font></td>
            <td width="15"><BR></td>
            <td><font color="#000000" size="-1">�i��2�j�d���`�̎����́A���n���Ԃł��B</font></td>
          </tr>
        </table>

            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center" bgcolor="#FFCC33"> 
<%
    If Not bDispMode Then
        Response.Write "<td nowrap rowspan='3'>BL "
        Response.Write "No.</td>"
    End If
%>
                <td nowrap rowspan="3">�R���e�iNo.<font size="-1"><sup>(��1)</sup></font></td>
                <td nowrap colspan="2">�{�D</td>
                <td nowrap bgcolor="#FFCC33">�d�o�`</td>
                <td nowrap colspan="7">�^�[�~�i��</td>
              </tr>
              <tr bgcolor="#FFCC33" align="center"> 
                <td nowrap rowspan="2" bgcolor="#FFFF99">�D��</td>
                <td nowrap rowspan="2" bgcolor="#FFFF99">�d�o�`��</td>
                <td nowrap rowspan="2" bgcolor="#FFFF99">���݊���<br>
                  ����<font size="-1"><sup>(��2)</sup></font></td>
                <td nowrap colspan="3" bgcolor="#FFFF99">���`����</td>
                <td nowrap colspan="2" bgcolor="#FFFF99">�����m�F���� </td>
                <td nowrap rowspan="2" bgcolor="#FFFF99">���o��</td>
                <td nowrap rowspan="2" bgcolor="#FFFF99">���[�h���o<br>��������</td>
              </tr>
              <tr bgcolor="#FFCC33" align="center"> 
                <td nowrap bgcolor="#FFFF99">�v��</td>
                <td nowrap bgcolor="#FFFF99">�\��</td>
                <td nowrap bgcolor="#FFFF99">����</td>
                <td nowrap bgcolor="#FFFF99">�\��</td>
                <td nowrap bgcolor="#FFFF99">����</td>
              </tr>
<!-- ��������f�[�^�J��Ԃ� -->
<% ' �\���t�@�C���̃��R�[�h������ԌJ��Ԃ�
    LineNo=0
    Do While Not ti.AtEndOfStream
        anyTmp=Split(ti.ReadLine,",")
        LineNo=LineNo+1
%>
              <tr bgcolor="#FFFFFF">
<% ' BL No
    If Not bDispMode Then
        Response.Write "<td nowrap align=center valign=middle>"
        If strBooking<>anyTmp(0) Then
            Response.Write anyTmp(0)
            strBooking=anyTmp(0)
        Else
            Response.Write "<br>"
        End If
        Response.Write "</td>"
    End If
%>
                <td nowrap align=center valign=middle>
<% ' �R���e�iNo.
    Response.Write "<a href='impdetail.asp?line=" & LineNo & "&return=1'>" & anyTmp(1) & "</a>"
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �D��
    If anyTmp(7)<>"" Then
        Response.Write anyTmp(7)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �d�o�`
    If anyTmp(9)<>"" Then
        Response.Write anyTmp(9)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �d�o�` - ���݊���
    Response.Write DispDateTimeCell(anyTmp(11),10)
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �^�[�~�i�� �| ���݃X�P�W���[��
    If anyTmp(31)<>"" Then
        Response.Write "<font color='#0000FF'>"
    End If
    Response.Write DispDateTimeCell(anyTmp(31),10)
    If anyTmp(31)<>"" Then
        Response.Write "</font>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �^�[�~�i�� - ���ݗ\��
    If anyTmp(2)<>"" Then
        bLate = false
        If anyTmp(3)<>"" Then
            If anyTmp(2)<anyTmp(3) Then
                bLate = true
            End If
        End If
        If anyTmp(31)<>"" Then
            If anyTmp(31)<anyTmp(2) Then
                bLate = true
            End If
        End If
        If bLate Then
            Response.Write "<font color='#FF0000'>"
        Else
            Response.Write "<font color='#0000FF'>"
        End If
        Response.Write DispDateTimeCell(anyTmp(2),10)
        Response.Write "</font>"
    Else
        Response.Write DispDateTimeCell(anyTmp(2),10)
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �^�[�~�i�� - ���݊���
    Response.Write DispDateTimeCell(anyTmp(3),10)
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �^�[�~�i�� - �����m�F�\��
    If anyTmp(32)<>"" Then
        If anyTmp(18)<>"" Then
            If anyTmp(32)<anyTmp(18) Then
                Response.Write "<font color='#FF0000'>"
            Else
                Response.Write "<font color='#0000FF'>"
            End If
        Else
            Response.Write "<font color='#0000FF'>"
        End If
        Response.Write DispDateTimeCell(anyTmp(32),10)
        Response.Write "</font>"
    Else
        Response.Write DispDateTimeCell(anyTmp(32),10)
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �^�[�~�i�� - ���[�h����(�m�F)����
    Response.Write DispDateTimeCell(anyTmp(18),5)
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �^�[�~�i�����o��
    If anyTmp(4)="Y" Then
        Response.Write "��"
    ElseIf anyTmp(4)="S" Then
        Response.Write "��"
    Else
        Response.Write "�~"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �^�[�~�i�� - ���[�h���o����
    Response.Write DispDateTimeCell(anyTmp(13),10)
%>
                </td>
              </tr>
<%
    Loop
%>
<!-- �����܂� -->
            </table>
<form>
      <input type=button value='�\���f�[�^�̍X�V' OnClick="JavaScript:window.location.href='impreload.asp?request=implist1.asp'">
</form>
          </td>
        </tr>
      </table>
      </center>
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
<!-------------�ꗗ��ʏI���--------------------------->
<%
    DispMenuBarBack "implist.asp"
%>
</body>
</html>