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
    WriteLog fs, "2302","�A���R���e�i�Ɖ�-�����R���e�i","00", ","

    ' �\���t�@�C����Open
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

    '�߂��ʎ�ʂ��L��
    Session.Contents("dispreturn")=0
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
                <td nowrap><b>�葱���y�у^�[�~�i�����o�ۏ��&nbsp;</b></td>
                <td><img src="../gif/hr.gif"></td>
              </tr>
            </table>
            <br>
            <table border="0">
              <tr>
                <td>�@<a href="implist1.asp">�� �^�[�~�i�������܂ł̈ʒu���</a></td>
              </tr>
              <tr>
                <td>�@<a href="implist2.asp">�� �^�[�~�i�����o��̈ʒu��񁕊�{���</a></td>
              </tr>
            </table>
            <table>
              <tr>
                <td>  
                  <br>

        <table border="0" cellspacing="2" cellpadding="1">
          <tr> 
            <td width="15"><BR></td>
            <td><font color="#000000" size="-1">�i��1) �N���b�N�ŒP�ƃR���e�i����\��</font></td>
          </tr>
        </table>

                  <table border="1" cellspacing="1" cellpadding="3">
                    <tr align="center" bgcolor="#FFCC33"> 
<%
    If Not bDispMode Then
        Response.Write "<td nowrap align=center valign=middle rowspan='3' width='78'>BL No.</td>"
    End If
%>
                      <td rowspan="3" nowrap bgcolor="#FFCC33" align="center">�R���e�iNo.<font size="-1"><sup>(��1)</sup></font></td>
                      <td colspan="5" nowrap bgcolor="#FFCC33" align="center">�s���葱��</td>
                      <td rowspan="3" valign="middle" nowrap align="center" bgcolor="#FFCC33">�����<br>
                        DO���s</td>
                      <td rowspan="3" valign="middle" nowrap align="center" bgcolor="#FFCC33">�t���[<br>
                        �^�C��</td>
                      <td rowspan="3" valign="middle" nowrap align="center" bgcolor="#FFCC33">�^�[�~�i��<br>
                        ���o��</td>
<%'HiTS ver2 ADD by SEIKO n.Ooshige 2003/06/26%>
<!--                      <td rowspan="3" valign="middle" nowrap align="center" bgcolor="#FFCC33">���O����<br>��Ɣԍ�</td>-->
                    </tr>
                    <tr align="center"> 
                      <td nowrap bgcolor="#FFFFCC" colspan="2" align="center">�����m�F����</td>
                      <td align="center" nowrap bgcolor="#FFFFCC" rowspan="2">���A�����u</td>
                      <td align="center" nowrap bgcolor="#FFFFCC" rowspan="2">�ʔ���</td>
                      <td align="center" nowrap bgcolor="#FFFFCC" rowspan="2">�ʊ�/<br>�ېŗA��</td>
                    </tr>
                    <tr align="center"> 
                      <td nowrap bgcolor="#FFFFCC">�\��</td>
                      <td nowrap bgcolor="#FFFFCC">����</td>
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
    Response.Write "<a href='impdetail.asp?line=" & LineNo & "'>" & anyTmp(1) & "</a>"
%>
                      </td>
                      <td nowrap align=center valign=middle>
<% ' �����m�F�\�莞��
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
<% ' �����m�F��������
    Response.Write DispDateTimeCell(anyTmp(18),5)
%>
                      </td>
                      <td nowrap align=center valign=middle>
<% ' ���A��
    If anyTmp(17)="S" Then
        Response.Write "�~"
    ElseIf anyTmp(17)="C" Then
        Response.Write "��"
    Else
        Response.Write "�|"
    End If
%>
                      </td>
                      <td nowrap align=center valign=middle>
<% ' �ʔ���
    If anyTmp(33)<>"" Then
        Response.Write "��"
    Else
        Response.Write "�|"
    End If
%>
                      </td>
                      <td nowrap align=center valign=middle>
<% ' �ʊց^�ېŗA��
    If anyTmp(19)<>"" Then
        Response.Write "��"
    Else
        Response.Write "�~"
    End If
%>
                      </td>
                      <td nowrap align=center valign=middle>
<% ' ������c�n���s
    If anyTmp(21)<>"Y" Then
        Response.Write "�~"
    Else
        Response.Write "��"
    End If
%>
                      </td>
                      <td nowrap align=center valign=middle>
<% ' �t���[�^�C��
    If anyTmp(22)<>"" Then
        If anyTmp(22)<DispDateTime(Now,10) Then
            Response.Write "<font color='#FF0000'>"
        Else
            Response.Write "<font color='#000000'>"
        End If
        Response.Write DispDateTimeCell(anyTmp(22),5)
        Response.Write "</font>"
    Else
        Response.Write DispDateTimeCell(anyTmp(22),5)
    End If
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
<%'HiTS ver2 ADD by SEIKO n.Ooshige 2003/06/26
 ' ���O���͍�Ɣԍ�
'   Response.Write "                      <td nowrap align=center valign=middle>"
'   Response.Write anyTmp(40)
'   Response.Write "                    �@</td>"
%>
                    </tr>
<%
    Loop
%>
<!-- �����܂� -->
                  </table>
<form>
      <input type=button value='�\���f�[�^�̍X�V' OnClick="JavaScript:window.location.href='impreload.asp?request=implist.asp'">
</form>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      <form action="impcsvout.asp"><input type="submit" value="CSV�t�@�C���o��">
    �@<a href="help06.asp">CSV�t�@�C���o�͂Ƃ́H</a> 
      </form>
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
    DispMenuBarBack "impentry.asp"
%>
</body>
</html>