<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' Temp�t�@�C�������̃`�F�b�N
    CheckTempFile "MSIMPORT", "impentry.asp"

    ' �\�[�g���[�h�̎擾
    Dim strSortKey
    strSortKey=Session.Contents("sortkey")

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
    WriteLog fs, "2105","�A���R���e�i�Ɖ�-���^�p���ꗗ","00", ","

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
          <td rowspan=2><img src="gif/imprikuun.gif" width="506" height="73"></td>
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
                <td nowrap><b>�A���R���e�i���ꗗ(���^�p)&nbsp;</b></td>
                <td><img src="gif/hr.gif"></td>
              </tr>
            </table>
            <br>
			&nbsp;&nbsp;&nbsp;&nbsp;���ړ��̃{�^���������ƁA���̍��ڂ̒l�Ń\�[�g����܂��B<BR><BR>

        <table border="0" cellspacing="2" cellpadding="1">
          <tr> 
            <td width="15"><BR></td>
            <td><font color="#000000" size="-1">�i��1) �N���b�N�ŒP�ƃR���e�i����\��</font></td>
            <td width="15"><BR></td>
            <td><font color="#000000" size="-1">�i��2�j96=HC</font></td>
            <td width="15"><BR></td>
            <td><font color="#000000" size="-1">�i��3�j���h�@�Ɋւ��댯���̗L��</font></td>
            <td width="15"><BR></td>
            <td><font color="#000000" size="-1">�i��4�j�N���b�N�Ŋ����������͉�ʂ�</font></td>
          </tr>
        </table>

            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center" bgcolor="#FFCC33"> 
                <td nowrap rowspan="3" valign=bottom>
<%
    If strSortKey="�C��" Then
        Response.Write "�C��<BR><img src='gif/1.gif' height=13><BR><img src='gif/sort-r.gif' vspace=2></td>"
    Else
        Response.Write "�C��<BR><img src='gif/1.gif' height=13><BR>"
        Response.Write "<a href='ms-impreload.asp?request=ms-implist2.asp&sort=�C��'><img src='gif/sort-b.gif' border=0 vspace=2></a>"
        Response.Write "</td>"
    End If
%>
                <td nowrap rowspan="3" valign=bottom>BL No.<BR><img src="gif/1.gif" height=25></td>
                <td nowrap rowspan="3" valign=bottom>�R���e�iNo.<font size="-1"><sup>(��1)</sup></font><BR><img src="gif/1.gif" height=25></td>
                <td colspan="5" nowrap>��{���</td>
                <td colspan="2" nowrap>�^�[�~�i��</td>
                <td nowrap>�X�g�b�N���[�h</td>
                <td colspan="6" nowrap>����A��</td>
              </tr>
              <tr align="center" bgcolor="#FFFF99"> 
                <td nowrap rowspan="2"><font color="#000000">�T�C�Y</font></td>
                <td nowrap rowspan="2"><font color="#000000">����<BR><font size="-1"><sup>(��2)</sup></font></font></td>
                <td nowrap rowspan="2"><font color="#000000">���[�t�@�[</font></td>
                <td nowrap rowspan="2"><font color="#000000">�d��(t)</font></td>
                <td nowrap rowspan="2"><font color="#000000">�댯��<BR><font size="-1"><sup>(��3)</sup></font></font></td>
                <td nowrap rowspan="2"><font color="#000000">���o<br>��</font></td>
                <td nowrap rowspan="2"><font color="#000000">���o<br>�ꏊ</font></td>
                <td nowrap rowspan="2"><font color="#000000">���o<br>��������</font></td>
                <td colspan="2" nowrap><font color="#000000">�q�ɓ�������</font></td>
                <td nowrap rowspan="2"><font color="#000000">�q�ɗ���</font></td>
                <td nowrap rowspan="2"><font color="#000000">�f�o���j���O<br>��������<font size="-1"><sup>(��4)</sup></font></font></td>
                <td nowrap rowspan="2"><font color="#000000">��R��<br>�ԋp����</font></td>
                <td nowrap rowspan="2"><font color="#000000">��R��<br>�ԋp�ꏊ</font></td>
              </tr>
              <tr align="center" bgcolor="#FFFF99"> 
                <td nowrap valign=bottom><font color="#000000">
<%
    If strSortKey="�q�ɓ���" Then
        Response.Write "�w��<BR><img src='gif/sort-r.gif' vspace=2></font></td>"
    Else
        Response.Write "�w��<BR>"
        Response.Write "<a href='ms-impreload.asp?request=ms-implist2.asp&sort=�q�ɓ���'><img src='gif/sort-b.gif' border=0 vspace=2></a>"
        Response.Write "</font></td>"
    End If
%>
                <td nowrap valign=top><font color="#000000">����<BR><font size="-1"><sup>(��4)</sup></font></font></td>
              </tr>
<!-- ��������f�[�^�J��Ԃ� -->
<% ' �\���t�@�C���̃��R�[�h������ԌJ��Ԃ�
    LineNo=0
    Do While Not ti.AtEndOfStream
        anyTmp=Split(ti.ReadLine,",")
        LineNo=LineNo+1
%>
              <tr bgcolor="#FFFFFF"> 
                <td nowrap align=center valign=middle>
<% ' �C�ݖ�
    If anyTmp(8)<>"" Then
        Response.Write anyTmp(8)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' BL No
    If strBooking<>anyTmp(0) Then
        Response.Write anyTmp(0)
        strBL=anyTmp(0)
    Else
        Response.Write "<br>"
    End If
%>
				</td>
                <td nowrap align=center valign=middle>
<% ' �R���e�iNo.
    If anyTmp(1)<>"" Then
        Response.Write "<a href='ms-impdetail.asp?line=" & LineNo & "'>" & anyTmp(1) & "</a>"
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ��{��� - �T�C�Y
    If anyTmp(53)<>"" Then
        Response.Write anyTmp(53)
    Else
        If anyTmp(10)<>"" Then
            Response.Write "<font color='#0000FF'>" & anyTmp(10) & "</font>"
        Else
            Response.Write "<br>"
        End If
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ��{��� - ����
    If anyTmp(54)<>"" Then
        Response.Write anyTmp(54)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ��{��� - ���[�t�@�[
    If anyTmp(55)<>"" Then
        If anyTmp(55)="R" Then
            Response.Write "��"
        Else
            Response.Write "�|"
        End If
    Else
        Response.Write "<font color='#0000FF'>"
        If anyTmp(11)<>"" Then
            If anyTmp(11)<>"RF" Then
                Response.Write "�|"
            Else
                Response.Write "��"
            End If
        Else
            Response.Write "<br>"
        End If
        Response.Write "</font>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ��{��� - ���d��
    If anyTmp(56)<>"" And anyTmp(56)<>"0" Then
        dWeight=anyTmp(56) / 10
        Response.Write dWeight
    Else
        Response.Write "�|"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ��{��� - �댯��
    If anyTmp(57)<>"" Then
        If anyTmp(57)<>"H" Then
            Response.Write "�|"
        Else
            Response.Write "��"
        End If
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �^�[�~�i�� - ���o��
    If anyTmp(34)="Y" Then
        Response.Write "��"
    ElseIf anyTmp(34)="S" Then
        Response.Write "��"
    Else
        Response.Write "�~"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �^�[�~�i�� - ���o�ꏊ
    If anyTmp(35)<>"" Then
        Response.Write anyTmp(35)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �X�g�b�N���[�h - ���o��������
    Response.Write DispDateTimeCell(anyTmp(60),10)
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ����^�� - �q�ɓ����X�P�W���[��
    If anyTmp(64)<>"" Then
        strTemp=anyTmp(44)
    Else
        strTemp=anyTmp(13)
    End If
    If strTemp<>"" Then
        If strTemp<anyTmp(44) Then
            Response.Write "<font color='#FF0000'>"
        Else
            Response.Write "<font color='#0000FF'>"
        End If
    End If
    Response.Write DispDateTimeCell(strTemp,10)
    If strTemp<>"" Then
        Response.Write "</font>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ����^�� - �q�ɓ���
    If anyTmp(5)<>"" Then
        Response.Write "<a href='ms-impinput.asp?kind=1&line=" & LineNo & "&request=ms-implist2.asp'>"
    End If
    strTemp = DispDateTimeCell(anyTmp(44),10)
    If Left(strTemp,1)="<" And anyTmp(5)<>"" Then
        Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"
    Else
        Response.Write strTemp
    End If
    If anyTmp(5)<>"" Then
        Response.Write "</a>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ����^�� - �q�ɗ���
    If anyTmp(12)<>"" Then
        Response.Write anyTmp(12)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �f�o���j���O - ��������
    If anyTmp(5)<>"" Then
        Response.Write "<a href='ms-impinput.asp?kind=2&line=" & LineNo & "&request=ms-implist2.asp'>"
    End If
    strTemp = DispDateTimeCell(anyTmp(45),10)
    If Left(strTemp,1)="<" And anyTmp(5)<>"" Then
        Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"
    Else
        Response.Write strTemp
    End If
    If anyTmp(5)<>"" Then
        Response.Write "</a>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ����^�� - ��R���ԋp����
    Response.Write DispDateTimeCell(anyTmp(46),10)
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ����^�� - ��R���ԋp�ꏊ
    If anyTmp(40)<>"" Then
        Response.Write anyTmp(40)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
              </tr>
<%
    Loop
%>
<!-- �����܂� -->
            </table>
<form>
      <input type=button value='�\���f�[�^�̍X�V' OnClick="JavaScript:window.location.href='ms-impreload.asp?request=ms-implist2.asp'">
</form>
          </td>
        </tr>
      </table>
      <form action="ms-impcsvout.asp"><input type="submit" value="CSV�t�@�C���o��">
    �@<a href="help17.asp">CSV�t�@�C���o�͂Ƃ́H</a> 
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
    DispMenuBarBack "ms-impentry.asp"
%>
</body>
</html>
