<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' Temp�t�@�C�������̃`�F�b�N
    CheckTempFile "EXPORT", "expentry.asp"

    ' �\�����[�h�̎擾
    Dim bDispMode          ' true=�R���e�i���� / false=�u�b�L���O����
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
        Response.Redirect "expentry.asp"             '�A�o�R���e�i�Ɖ�g�b�v
        Response.End
    End If
    strFileName="../temp/" & strFileName

    ' �A�o�R���e�i�Ɖ�X�g�\��
    WriteLog fs, "1004","�A�o�R���e�i�Ɖ�-�ʒu���","00", ","

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
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------��������ꗗ���--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/explistt.gif" width="506" height="73"></td>
          <td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
        </tr>
        <tr>
          <td align="right" width="100%" height="48"> 
<%
' Added and Commented by seiko-denki 2003.07.18
'	DisplayCodeListButton
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
end of comment by seiko-denki 2003.07.18 -->
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
                <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>Position information</b></td>
                <td><img src="gif/hr.gif" hspace="3"></td>
              </tr>
            </table>
            <br>

            <table border="0" cellspacing="2" cellpadding="1">
              <tr> 
                <td width="15"><BR></td>
                <td><font color="#000000" size="-1">(*1)Display datails when clicking a container No. </font></td>
                <td width="15"><BR></td>
                <td><font color="#000000" size="-1">(*2)The time of the discharge port is local time.</font></td>
              </tr>
            </table>

            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center" bgcolor="#FFCC33"> 
<%
    If Not bDispMode Then
        Response.Write "<td nowrap rowspan='3'>Booking "
        Response.Write "No.</td>"
    End If
%>
                <td nowrap rowspan="3">Container
No.<font size="-1"><sup>(*1)</sup></font></td>
                <td colspan="4" nowrap>Conveyance by land</td>
                <td nowrap>Stock Yard </td>
                <td colspan="6" nowrap bgcolor="#FFCC33">Terminal</td>
                <td bgcolor="#FFCC33" nowrap colspan="2">Discharge Port</td>
              </tr>
              <tr align="center" bgcolor="#FFFF99"> 
                <td nowrap rowspan="2" bgcolor="#FFFF99">Empty container<br> pickup
                  time</td>
                <td nowrap colspan="2">Warehouse arrival time</td>
                <td nowrap rowspan="2">Vanning time</td>
                <td nowrap rowspan="2">SY in<br>
                  Actual<br>
                  time</td>
                <td nowrap colspan="2">CY in time</td>
                <td nowrap rowspan="2">Loading<br>
                  time</td>
                <td nowrap colspan="3">Departure time</td>
                <td nowrap colspan="2">arrival time<font size="-1"><sup>(*2)</sup></font></td>
              </tr>
              <tr align="center" bgcolor="#FFFF99">
                <td nowrap>Instruction</td>
                <td nowrap>Actual</td>
                <td nowrap>Instruction</td>
                <td nowrap>Actual</td>
                <td nowrap>Estimate</td>
                <td nowrap>Intended</td>
                <td nowrap>Actual</td>
                <td nowrap>Intended</td>
                <td nowrap>Actual</td>
              </tr>
<!-- ��������f�[�^�J��Ԃ� -->
<% ' �\���t�@�C���̃��R�[�h������ԌJ��Ԃ�
    LineNo=0
    Do While Not ti.AtEndOfStream
        anyTmp=Split(ti.ReadLine,",")
        LineNo=LineNo+1
%>
              <tr bgcolor="#FFFFFF"> 
<% ' Booking No
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
    Response.Write "<a href='expdetail.asp?line=" & LineNo & "'>" & anyTmp(1) & "</a>"
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ����^�� - ��R�����
    Response.Write DispDateTimeCell(anyTmp(16),10)
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ����^�� - �q�ɓ����X�P�W���[��
    If anyTmp(26)<>"" Then
        If anyTmp(26)<anyTmp(17) Then
            Response.Write "<font color='#FF0000'>"
        Else
            Response.Write "<font color='#0000FF'>"
        End If
    End If
    Response.Write DispDateTimeCell(anyTmp(26),10)
    If anyTmp(26)<>"" Then
        Response.Write "</font>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ����^�� - �q�ɓ���
    Response.Write DispDateTimeCell(anyTmp(17),10)
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ����^�� - �o���j���O
    Response.Write DispDateTimeCell(anyTmp(18),10)
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �X�g�b�N���[�h - ����
    Response.Write DispDateTimeCell(anyTmp(24),10)
%>
                </td>
                <td align="center" nowrap>
<% ' �^�[�~�i�� - CY�����w�� $�ǉ�
    If anyTmp(30)<>"" Then
        If Left(anyTmp(30),10)<Left(anyTmp(19),10) Then
            Response.Write "<font color='#FF0000'>"
        Else
            Response.Write "<font color='#0000FF'>"
        End If
    End If
    Response.Write DispDateTimeCell(anyTmp(30),5)
    If anyTmp(30)<>"" Then
        Response.Write "</font>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �^�[�~�i�� - CY��������
    Response.Write DispDateTimeCell(anyTmp(19),10)
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �^�[�~�i�� - �D�ϊ���
    Response.Write DispDateTimeCell(anyTmp(20),10)
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �^�[�~�i�� - ���݃X�P�W���[��
    If anyTmp(25)<>"" Then
        Response.Write "<font color='#0000FF'>"
    End If
    Response.Write DispDateTimeCell(anyTmp(25),5)
    If anyTmp(25)<>"" Then
        Response.Write "</font>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �^�[�~�i�� - ���ݗ\��
    If anyTmp(15)<>"" Then
        bLate = false
        If anyTmp(21)<>"" Then
            If anyTmp(15)<anyTmp(21) Then
                bLate = true
            End If
        End If
        If anyTmp(25)<>"" Then
            If Left(anyTmp(25),10)<Left(anyTmp(15),10) Then
                bLate = true
            End If
        End If
        If bLate Then
            Response.Write "<font color='#FF0000'>"
        Else
            Response.Write "<font color='#0000FF'>"
        End If
        Response.Write DispDateTimeCell(anyTmp(15),10)
        Response.Write "</font>"
    Else
        Response.Write DispDateTimeCell(anyTmp(15),10)
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �^�[�~�i�� - ���݊���
    Response.Write DispDateTimeCell(anyTmp(21),10)
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �d���` - ���ݗ\��
    If anyTmp(23)<>"" Then
        If anyTmp(22)<>"" Then
            If anyTmp(23)<anyTmp(22) Then
                Response.Write "<font color='#FF0000'>"
            Else
                Response.Write "<font color='#0000FF'>"
            End If
        Else
            Response.Write "<font color='#0000FF'>"
        End If
        Response.Write DispDateTimeCell(anyTmp(23),10)
        Response.Write "</font>"
    Else
        Response.Write DispDateTimeCell(anyTmp(23),10)
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �d���` - ���݊���
    Response.Write DispDateTimeCell(anyTmp(22),10)
%>
                </td>
              </tr>
<%
    Loop
%>
<!-- �����܂� -->
            </table>
<form>
      <input type=button value='Display Update' OnClick="JavaScript:window.location.href='expreload.asp?request=explist1.asp'">
</form>
          </td>
        </tr>
      </table>
      <br>
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
    DispMenuBarBack "explist.asp"
%>
</body>
</html>
