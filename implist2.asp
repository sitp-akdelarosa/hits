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
'������ Add by nics 2010.02.02
    Dim USER
	USER    = Session.Contents("userid")
'������ end of Add by nics 2010.02.02

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
    WriteLog fs, "2005","�A���R���e�i�Ɖ�-���o��̈ʒu��񁕊�{���","00", ","

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
          <td rowspan=2><img src="gif/implistt.gif" width="506" height="73"></td>
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
                <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>�^�[�~�i�����o��̈ʒu��񁕊�{���&nbsp;</b></td>
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
            <td width="15"><BR></td>
            <td><font color="#000000" size="-1">�i��3�j���h�@�Ɋւ��댯���̗L��</font></td>
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
<!-- Mod-S MES Aoyagi 2010.11.23 -->
<!--                <td nowrap colspan="5">��{���</td> -->
		<td nowrap colspan="6">��{���</td>
<!-- Mod-E MES Aoyagi 2010.11.23 -->
<!-- mod by nics 2009.02.24 -->
<!--                <td nowrap colspan="2">�^�[�~�i��</td>-->
                <td nowrap colspan="3">�^�[�~�i��</td>
<!-- end of mod by nics 2009.02.24 -->
                <td nowrap colspan="2">�X�g�b�N���[�h</td>
<!-- mod by nics 2009.02.24 -->
<!--                <td nowrap colspan="5">����A��</td>-->
<!-- mod by nics 2010.02.16 -->
<!--                <td nowrap colspan="4">����A��</td>-->
<%
    If Not bDispMode  or USER <> "" Then
        Response.Write "<td nowrap colspan='4'>"
        Response.Write "����A��</td>"
	Else
        Response.Write "<td nowrap colspan='3'>"
        Response.Write "����A��</td>"
    End If
%>
<!-- mod by nics 2010.02.16 -->
<!-- end of mod by nics 2009.02.24 -->
              </tr>
              <tr bgcolor="#FFCC33" align="center"> 
                <td nowrap bgcolor="#FFFFCC" align="center" rowspan="2">�T�C�Y</td>
<!-- Add-S MES Aoyagi 2010.11.23 -->
		<td nowrap bgcolor="#FFFFCC" align="center" rowspan="2">�^�C�v</td>
<!-- Add-E MES Aoyagi 2010.11.23 -->
                <td nowrap bgcolor="#FFFFCC" align="center" rowspan="2">����<BR><font size="-1"><sup>(��2)</sup></font></td>
                <td nowrap bgcolor="#FFFFCC" align="center" rowspan="2">���[�t�@</td>
                <td nowrap bgcolor="#FFFFCC" align="center" rowspan="2">�d��(t)</td>
                <td nowrap align="center" bgcolor="#FFFFCC" rowspan="2">�댯��<BR><font size="-1"><sup>(��3)</sup></font></td>
                <td nowrap bgcolor="#FFFFCC" rowspan="2">���o<br>��</td>
<!-- add by nics 2009.02.24 -->
                <td nowrap rowspan="2" bgcolor="#FFFFCC"><font color="#000000">���o�^�[�~�i��<br>(���u�ꏊ�R�[�h)</font></td>
                <td nowrap rowspan="2" bgcolor="#FFFFCC"><font color="#000000">�{�D�S��<br>�I�y���[�^</font></td>
<!-- end of add by nics 2009.02.24 -->
<!-- commented by nics 2009.02.24
                <td nowrap bgcolor="#FFFFCC" rowspan="2">���o<br>�ꏊ</td>
end of comment by nics 2009.02.24 -->
                <td nowrap bgcolor="#FFFFCC" colspan="2">���o����</td>
<!-- commented by nics 2009.02.24
                <td nowrap bgcolor="#FFFFCC" colspan="2">�q�ɓ�������<br></td>
                <td nowrap bgcolor="#FFFFCC" rowspan="2">�f�o���j���O<br>
                  ��������</td>
end of comment by nics 2009.02.24 -->
                <td nowrap bgcolor="#FFFFCC" rowspan="2">��R��<br>
                  �ԋp����</td>
                <td nowrap bgcolor="#FFFFCC" rowspan="2">��R��<br>
                  �ԋp�ꏊ</td>
<!-- add by nics 2009.02.24 -->
<!-- mod by nics 2010.02.16 -->
<!--                <td nowrap bgcolor="#FFFFCC" rowspan="2">�f�B�e���V����<br>�t���[�^�C��</td>	-->
<%
    If Not bDispMode  or USER <> "" Then
        Response.Write "<td nowrap bgcolor='#FFFFCC' rowspan='2'>"
        Response.Write "�f�B�e���V����<br>�t���[�^�C��</td>"
    End If
%>
<!-- end of mod by nics 2010.02.16 -->
<!-- end of add by nics 2009.02.24 -->
              </tr>
              <tr bgcolor="#FFFFCC" align="center"> 
                <td nowrap>�\��</td>
                <td nowrap>����</td>
<!-- commented by nics 2009.02.24
                <td nowrap>�w��</td>
                <td nowrap>����</td>
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
    Response.Write "<a href='impdetail.asp?line=" & LineNo & "&return=2'>" & anyTmp(1) & "</a>"
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �T�C�Y
    If anyTmp(23)<>"" Then
        Response.Write anyTmp(23)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>

<!-- Add-S MES Aoyagi 2010.11.23 -->
<% ' �^�C�v
    If anyTmp(46)<>"" Then
        Response.Write anyTmp(46)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<!-- Add-E MES Aoyagi 2010.11.23 -->

<% ' ����
    If anyTmp(24)<>"" Then
        Response.Write anyTmp(24)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ���[�t�@�[
    If anyTmp(25)="R" Then
        Response.Write "��"
    ElseIf anyTmp(25)<>"" Then
        Response.Write "�|"
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �d��
    If anyTmp(26)<>"" And anyTmp(26)<>"0" Then
        dWeight=anyTmp(26) / 10
        Response.Write dWeight
    Else
        Response.Write "�|"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �댯��
    If anyTmp(27)="H" Then
        Response.Write "��"
    ElseIf anyTmp(27)<>"" Then
        Response.Write "�|"
    Else
        Response.Write "<br>"
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
<!-- commented by nics 2009.02.24
                <td nowrap align=center valign=middle>
<% ' �^�[�~�i�� - ���o�ꏊ
    If anyTmp(5)<>"" Then
        Response.Write anyTmp(5)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
end of comment by nics 2009.02.24 -->
<!-- add by nics 2009.02.24 -->
                     <td nowrap align=center valign=middle>
<% ' ���o�^�[�~�i��(���u�ꏊ�R�[�h)
    strDisp = "<br>"
    If anyTmp(5) <> "" Then
        strDisp = anyTmp(5)
        If anyTmp(43) <> "" Then
            strDisp = strDisp & "<br>(" & anyTmp(43) & ")"
        End If
    End If
    Response.Write strDisp
%>
                     </td>
                     <td nowrap align=center valign=middle>
<% ' �S���I�y���[�^
    If anyTmp(45)<>"" Then
        Response.Write anyTmp(45)
    Else
        Response.Write "<br>"
    End If
%>
                     </td>
<!-- end of add by nics 2009.02.24 -->
                <td align="center" nowrap>
<% ' �X�g�b�N���[�h - ���o�\�� $�ǉ�
    sTemp=DispReserveCell(anyTmp(35),anyTmp(36),sColor)
    If Left(sTemp,1)="<" Then
        Response.Write sTemp
    Else
        Response.Write sColor
        Response.Write Left(sTemp,5) & "<br>" & Mid(sTemp,7)
        If sColor<>"" Then
            Response.Write "</font>"
        End If
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �X�g�b�N���[�h - ���o����
    Response.Write DispDateTimeCell(anyTmp(30),10)
%>
                </td>
<!-- commented by nics 2009.02.24
                <td align="center" nowrap>
<% ' ����^�� - �q�ɓ����X�P�W���[��
    If anyTmp(34)<>"" Then
        If anyTmp(34)<anyTmp(14) Then
            Response.Write "<font color='#FF0000'>"
        Else
            Response.Write "<font color='#0000FF'>"
        End If
    End If
    Response.Write DispDateTimeCell(anyTmp(34),10)
    If anyTmp(34)<>"" Then
        Response.Write "</font>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ����A�� - �q�ɓ�������
    Response.Write DispDateTimeCell(anyTmp(14),10)
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ����A�� - �f�o������
    Response.Write DispDateTimeCell(anyTmp(15),10)
%>
                </td>
end of comment by nics 2009.02.24 -->
                <td nowrap align=center valign=middle>
<% ' ����A�� - ��R���ԋp����
    Response.Write DispDateTimeCell(anyTmp(16),10)
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �ԋp�ꏊ
    If anyTmp(10)<>"" Then
        Response.Write anyTmp(10)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- add by nics 2009.02.24 -->
<!-- mod by nics 2010.02.16 -->
<!--                <td align="center" nowrap>	-->
<% ' �f�B�e���V�����t���[�^�C��
    ' anyTmp(39) �� �f�B�e���V�����t���[�^�C��
    ' anyTmp(16) �� ��o���ԋp����[yyyy/mm/dd hh:nn]
    ' anyTmp(44) �� ��o���ԋp�\���[yyyy/mm/dd]
	'mod by nics 2010.02.16
	If Not bDispMode  or USER <> "" Then
	    Response.Write "<td align='center' nowrap>"
	    strDisp = anyTmp(39)
	    strColor = "#000000"    ' ��
	    ' ��o���ԋp�������ݒ肳��Ă���ꍇ
	    If anyTmp(16) <> "" Then
	        ' ��o���ԋp�������V�X�e�����t�̏ꍇ
	        If Left(anyTmp(16),10) < DispDateTime(Now,10) Then
	            strDisp = "�|"
	        End If
	    ' ��o���ԋp�������ݒ肳��Ă��Ȃ��ꍇ
	    Else
	        ' ��o���ԋp�\��������ݒ肳��Ă���ꍇ
	        If IsDate(anyTmp(44)) Then
	            ' ��o���ԋp�\������V�X�e�����t�̏ꍇ
	            If anyTmp(44) <= DispDateTime(Now,10) Then
	                strColor = "#FF0000"    ' ��
	            ' (��o���ԋp�\����|2��)���V�X�e�����t�̏ꍇ
	            ElseIf DispDateTime(DateAdd("d", -2, cDate(anyTmp(44))),10) <= DispDateTime(Now,10) Then
	                strColor = "#FFA500"    ' ��
	            End If
	        End If
	    End If
	    Response.Write "<font color='" & strColor & "'>"
	    Response.Write strDisp
	    Response.Write "</font>"
	'add by nics 2010.02.16
	    Response.Write "�@</td>"
	End If
	'end of mod by nics 2010.02.16
%>
<!-- del by nics 2010.02.16 -->
<!--              �@</td>	-->
<!-- del by nics 2010.02.16 -->
<!-- end of add by nics 2009.02.24 -->
              </tr>
<%
    Loop
%>
<!-- �����܂� -->
            </table>
<form>
      <input type=button value='�\���f�[�^�̍X�V' OnClick="JavaScript:window.location.href='impreload.asp?request=implist2.asp'">
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