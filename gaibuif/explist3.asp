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
    strFileName="./temp/" & strFileName

    ' �A�o�R���e�i�Ɖ�X�g�\��
    WriteLog fs, "3006","�d�o�n���Ɖ�-�^�[�~�i���A�{�D�ɌW����","00", ","

    ' �\���t�@�C����Open
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

    '�߂��ʎ�ʂ��L��
    Session.Contents("dispreturn")=3
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
          <td rowspan=2><img src="../gif/explistt.gif" width="506" height="73"></td>
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
                <td><img src="../gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>�^�[�~�i���A�{�D�ɌW����@</b></td>
                <td><img src="../gif/hr.gif"></td>
              </tr>
            </table>
            <br>

        <table border="0" cellspacing="2" cellpadding="1">
          <tr> 
            <td width="15"><BR></td>
            <td><font color="#000000" size="-1">�i��1) �N���b�N�ŒP�ƃR���e�i����\��</font></td>
<!-- commented by nics 2009.02.24
            <td width="15"><BR></td>
            <td><font color="#000000" size="-1">�i��2�j�d���`�̎����́A���n���Ԃł��B</font></td>
end of comment by nics 2009.02.24 -->
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
                <td rowspan="3" bgcolor="#FFCC33" nowrap align="center">�R���e�iNo.<font size="-1"><sup>(��1)</sup></font></td>
                <td colspan="2" bgcolor="#FFCC33" nowrap align="center">������t����</td>
<!-- mod by nics 2009.02.24 -->
<!--                <td colspan="6" bgcolor="#FFCC33" nowrap align="center">�^�[�~�i��</td>-->
                <td colspan="7" bgcolor="#FFCC33" nowrap align="center">�^�[�~�i��</td>
<!-- end of mod by nics 2009.02.24 -->
                <td colspan="2" bgcolor="#FFCC33" nowrap align="center">�{�D</td>
<!-- commented by nics 2009.02.24
                <td colspan="2" bgcolor="#FFCC33" nowrap align="center">�d���`</td>
end of comment by nics 2009.02.24 -->
              </tr>
              <tr> 
                <td bgcolor="#FFFF99" align="center" nowrap rowspan="2">�I�[�v����</td>
                <td bgcolor="#FFFF99" align="center" nowrap rowspan="2">�N���[�Y��</td>
                <td bgcolor="#FFFF99" align="center" colspan="2" nowrap>��������</td>
                <td bgcolor="#FFFF99" align="center" rowspan="2" nowrap>�D��<BR>
                  ��������</td>
<!-- mod by nics 2009.02.24 -->
<!--                <td bgcolor="#FFFF99" align="center" colspan="3" nowrap>���ݎ���</td>-->
                <td bgcolor="#FFFF99" align="center" colspan="2" nowrap>���ݎ���</td>
<!-- end of mod by nics 2009.02.24 -->
<!-- add by nics 2009.02.24 -->
                <td bgcolor="#FFFF99" align="center"  nowrap rowspan="2"><font color="#000000">�����^�[�~�i��<br>(���u�ꏊ�R�[�h)</font></td>
                <td bgcolor="#FFFF99" align="center"  nowrap rowspan="2"><font color="#000000">�{�D�S��<br>�I�y���[�^</font></td>
<!-- end of add by nics 2009.02.24 -->
                <td bgcolor="#FFFF99" align="center" nowrap rowspan="2">�D��</td>
                <td bgcolor="#FFFF99" align="center" rowspan="2" nowrap>�d�o�`��</td>
<!-- commented by nics 2009.02.24
                <td bgcolor="#FFFF99" align="center" nowrap colspan="2">���ݎ���<font size="-1"><sup>(��2)</sup></font></td>
end of comment by nics 2009.02.24 -->
              </tr>
              <tr> 
                <td bgcolor="#FFFF99" align="center" nowrap>�w��</td>
                <td bgcolor="#FFFF99" align="center" nowrap>����</td>
<!-- commented by nics 2009.02.24
                <td bgcolor="#FFFF99" align="center" nowrap>�v��</td>
end of comment by nics 2009.02.24 -->
                <td bgcolor="#FFFF99" align="center" nowrap>�\��</td>
                <td bgcolor="#FFFF99" align="center" nowrap>����</td>
<!-- commented by nics 2009.02.24
                <td bgcolor="#FFFF99" align="center" nowrap>�\��</td>
                <td bgcolor="#FFFF99" align="center" nowrap>����</td>
end of comment by nics 2009.02.24 -->
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
<% ' CY�I�[�v��
    Response.Write DispDateTimeCell(anyTmp(9),5)
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' CY�N���[�Y
    Response.Write DispDateTimeCell(anyTmp(10),5)
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
<!-- commented by nics 2009.02.24
                <td nowrap align=center valign=middle>
<% ' �^�[�~�i�� - ���݃X�P�W���[��
    If anyTmp(25)<>"" Then
        Response.Write "<font color='#0000FF'>"
    End If
    Response.Write DispDateTimeCell(anyTmp(25),10)
    If anyTmp(25)<>"" Then
        Response.Write "</font>"
    End If
%>
                </td>
end of comment by nics 2009.02.24 -->
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
            If anyTmp(25)<anyTmp(15) Then
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
<!-- add by nics 2009.02.24 -->
                     <td nowrap align=center valign=middle>
<% ' �����^�[�~�i��(���u�ꏊ�R�[�h)
    strDisp = "<br>"
    If anyTmp(6) <> "" Then
        strDisp = anyTmp(6)
        If anyTmp(36) <> "" Then
            strDisp = strDisp & "<br>(" & anyTmp(36) & ")"
        End If
    End If
    Response.Write strDisp

%>
                     </td>
                     <td nowrap align=center valign=middle>
<% ' �S���I�y���[�^
    If anyTmp(37)<>"" Then
        Response.Write anyTmp(37)
    Else
        Response.Write "<br>"
    End If
%>
                     </td>
<!-- end of add by nics 2009.02.24 -->
                <td nowrap align=center valign=middle>
<% ' �D��
    If anyTmp(12)<>"" Then
        Response.Write anyTmp(12)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �d���`
    If anyTmp(14)<>"" Then
        Response.Write anyTmp(14)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- commented by nics 2009.02.24
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
end of comment by nics 2009.02.24 -->
              </tr>
<%
    Loop
%>
<!-- �����܂� -->
            </table>
<form>
      <input type=button value='�\���f�[�^�̍X�V' OnClick="JavaScript:window.location.href='expreload.asp?request=explist3.asp'">
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
