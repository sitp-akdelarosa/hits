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
    WriteLog fs, "1005","�A�o�R���e�i�Ɖ�-�R���e�i�쐬�ɌW����","00", ","

    ' �\���t�@�C����Open
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

    '�߂��ʎ�ʂ��L��
    Session.Contents("dispreturn")=2
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
                <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>�R���e�i�쐬�ɌW����@</b></td>
                <td><img src="gif/hr.gif"></td>
              </tr>
            </table>
            <br>

        <table border="0" cellspacing="2" cellpadding="1">
          <tr> 
            <td width="15"><BR></td>
            <td><font color="#000000" size="-1">�i��1) �N���b�N�ŒP�ƃR���e�i����\��</font></td>
            <td width="15"><BR></td>
            <td><font color="#000000" size="-1">�i��2�j96=HC</font></td>
          </tr>
        </table>

            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center" bgcolor="#FFCC33"> 
<%
    If Not bDispMode Then
        Response.Write "<td nowrap rowspan='2'>Booking "
        Response.Write "No.</td>"
    End If
%>
                <td rowspan="2" bgcolor="#FFCC33" nowrap align="center">�R���e�iNo.<font size="-1"><sup>(��1)</sup></font></td>
<!-- mod by mes(2005/3/28) �e�A�E�F�C�g�ǉ� -->
<!--                <td colspan="5" bgcolor="#FFCC33" nowrap align="center">��R���e�i</td>-->
<!--                <td colspan="6" bgcolor="#FFCC33" nowrap align="center">��R���e�i</td> -->
		<td colspan="7" bgcolor="#FFCC33" nowrap align="center">��R���e�i</td>
<!-- end mes -->
                <td colspan="6" bgcolor="#FFCC33" nowrap align="center">�o���j���O��R���e�i</td>
              </tr>
              <tr> 
                <td bgcolor="#FFFF99" align="center" nowrap>��掞��</td>
                <td bgcolor="#FFFF99" align="center" nowrap>���ꏊ</td>
                <td bgcolor="#FFFF99" align="center" nowrap>�T�C�Y</td>
<!-- Add-S MES Aoyagi 2010.11.23 -->
		<td bgcolor="#FFFF99" align="center" nowrap>�^�C�v</td>
<!-- Add-E MES Aoyagi 2010.11.23 -->
                <td bgcolor="#FFFF99" align="center" nowrap>����<BR><font size="-1"><sup>(��2)</sup></font></td>
<!-- add by mes(2005/3/28) �e�A�E�F�C�g�ǉ� -->
                <td bgcolor="#FFFF99" align="center" nowrap>TW</td>
<!-- end mes -->
                <td bgcolor="#FFFF99" align="center" nowrap>���[�t�@</td>
<!-- commented by nics 2009.02.24
                <td bgcolor="#FFFF99" align="center" nowrap>�q��<br>
                  ��������</td>
end of comment by nics 2009.02.24 -->
                <td bgcolor="#FFFF99" align="center" nowrap>�V�[��No.</td>
                <td bgcolor="#FFFF99" align="center" nowrap>�ݕ�<br>�d��(t)</td>
                <td bgcolor="#FFFF99" align="center" nowrap>��<br>�d��(t)</td>
                <td bgcolor="#FFFF99" align="center" nowrap>�o���j���O<br>
                  ��������</td>
<!-- commented by nics 2009.02.24
                <td bgcolor="#FFFF99" align="center" nowrap>����<br>
                  �^�[�~�i����</td>
end of comment by nics 2009.02.24 -->
<!-- add by nics 2009.02.24 -->
                <td bgcolor="#FFFF99" align="center" nowrap><font color="#000000">�����^�[�~�i��<br>(���u�ꏊ�R�[�h)</font></td>
                <td bgcolor="#FFFF99" align="center" nowrap><font color="#000000">�{�D�S��<br>�I�y���[�^</font></td>
<!-- end of add by nics 2009.02.24 -->
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
<% ' ��R�����ꏊ
    If anyTmp(2)<>"" Then
        Response.Write anyTmp(2)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �T�C�Y
    If anyTmp(3)<>"" Then
        Response.Write anyTmp(3)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>

<!-- Add-S MES Aoyagi 2010.11.23 -->
<% ' �^�C�v
    If anyTmp(39)<>"" Then
        Response.Write anyTmp(39)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<!-- Add-E MES Aoyagi 2010.11.23 -->

<% ' ����
    If anyTmp(4)<>"" Then
        Response.Write anyTmp(4)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- add by mes(2005/3/28) �e�A�E�F�C�g�ǉ� -->
                <td nowrap align=center valign=middle>
<% ' �e�A�E�F�C�g(TW)
    If anyTmp(32)<>"" And anyTmp(32)>0 Then
    	If anyTmp(32)<100 then
	        dWeight=anyTmp(32) * 100
	    Else
	        dWeight=anyTmp(32)
	    End If
        Response.Write dWeight
    Else
        Response.Write "�|"
    End If
%>
                </td>
<!-- end mes -->
                <td nowrap align=center valign=middle>
<% ' ���[�t�@�[
    If anyTmp(5)="R" Then
        Response.Write "��"
    ElseIf anyTmp(5)<>"" Then
        Response.Write "�|"
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- commented by nics 2009.02.24
                <td nowrap align=center valign=middle>
<% ' ����^�� - �q�ɓ���
    Response.Write DispDateTimeCell(anyTmp(17),10)
%>
                </td>
end of comment by nics 2009.02.24 -->
                <td nowrap align=center valign=middle>
<% ' �V�[��No.
    If anyTmp(7)<>"" Then
        Response.Write anyTmp(7)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �ݕ��d�� $�ǉ�
    If anyTmp(27)<>"" And anyTmp(27)<>"0" Then
        dWeight=anyTmp(27) / 10
        Response.Write dWeight
    Else
        Response.Write "�|"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ���d��
    If anyTmp(8)<>"" And anyTmp(8)<>"0" Then
        dWeight=anyTmp(8) / 10
        Response.Write dWeight
    Else
        Response.Write "�|"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ����^�� - �o���j���O
    Response.Write DispDateTimeCell(anyTmp(18),10)
%>
                </td>
<!-- commented by nics 2009.02.24
                <td nowrap align=center valign=middle>
<% ' �����^�[�~�i����
    If anyTmp(6)<>"" Then
        Response.Write anyTmp(6)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
end of comment by nics 2009.02.24 -->
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
              </tr>
<%
    Loop
%>
<!-- �����܂� -->
            </table>
<form>
      <input type=button value='�\���f�[�^�̍X�V' OnClick="JavaScript:window.location.href='expreload.asp?request=explist2.asp'">
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
    DispMenuBarBack "explist.asp"
%>
</body>
</html>
