<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' Temp�t�@�C�������̃`�F�b�N
    CheckTempFile "MSIMPORT", "impentry.asp"

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

    ' ���[�U��ނ��`�F�b�N����
    strUserKind=Session.Contents("userkind")
    ' Sort������ނ��`�F�b�N����
    strSortKey=Session.Contents("sortkey")

	Dim iNum
	If strUserKind="�C��" Then
		iNum = "2104"
	ElseIf strUserKind="���^" Then
		iNum = "2105"
	Else
		iNum = "2106"
	End If

    ' �A���R���e�i�Ɖ�X�g�\��
    WriteLog fs, iNum,"�A���R���e�i�Ɖ�-" & strUserKind & "�p���ꗗ","00", ","

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
<td rowspan=2><%

    If strUserKind="�C��" Then
        Response.Write "<img src='gif/impkaika.gif' width='506' height='73'>"
    ElseIf strUserKind="���^" Then
        Response.Write "<img src='gif/imprikuun.gif' width='506' height='73'>"
    Else
        Response.Write "<img src='gif/impninushi.gif' width='506' height='73'>"
    End If

%></td>
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
                <td nowrap><b>
<%
        Response.Write "�A���R���e�i���ꗗ(" & strUserKind & "�p)"
%>
                &nbsp;</b></td>
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
            <td><font color="#000000" size="-1">�i��2�j�d���`�̎����́A���n���Ԃł��B</font></td>
            <td width="15"><BR></td>
            <td><font color="#000000" size="-1">�i��3�j�F���� &nbsp; ���F�Ɖ��</font></td>
          </tr>
        </table>

            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center" bgcolor="#FFCC33"> 
                <td nowrap colspan="2" valign=bottom>
<%
    If strSortKey="�D��" Then
        Response.Write "�{�D<BR><img src='gif/sort-r.gif' vspace=2></td>"
    Else
        Response.Write "�{�D<BR>"
        Response.Write "<a href='ms-impreload.asp?request=ms-implist1.asp&sort=�D��'><img src='gif/sort-b.gif' border=0 vspace=2></a>"
        Response.Write "</td>"
    End If
%>
<%
    If strUserKind="�C��" Then
        Response.Write "<td nowrap rowspan='3' valign=bottom>"
        If strSortKey="�׎�" Then
            Response.Write "�׎�<BR><img src='gif/1.gif' height=18><BR><img src='gif/sort-r.gif' vspace=2></td>"
        Else
            Response.Write "�׎�<BR><img src='gif/1.gif' height=18><BR>"
            Response.Write "<a href='ms-impreload.asp?request=ms-implist1.asp&sort=�׎�'><img src='gif/sort-b.gif' border=0 vspace=2></a>"
            Response.Write "</td>"
        End If
    Else
        Response.Write "<td nowrap rowspan='3' valign=bottom>"
        If strSortKey="�C��" Then
            Response.Write "�C��<BR><img src='gif/1.gif' height=18><BR><img src='gif/sort-r.gif' vspace=2></td>"
        Else
            Response.Write "�C��<BR><img src='gif/1.gif' height=18><BR>"
            Response.Write "<a href='ms-impreload.asp?request=ms-implist1.asp&sort=�C��'><img src='gif/sort-b.gif' border=0 vspace=2></a>"
            Response.Write "</td>"
        End If
    End If
%>
                <td nowrap rowspan="3" valign=bottom>�D��<BR><img src="gif/1.gif" height=30></td>
                <td nowrap rowspan="3" valign=bottom>BL No.<BR><img src="gif/1.gif" height=30></td>
                <td nowrap rowspan="3" valign=bottom>�R���e�iNo.<font size="-1"><sup>(��1)</sup></font><BR><img src="gif/1.gif" height=30></td>
                <td nowrap bgcolor="#FFCC33">�d�o�`</td>
                <td nowrap colspan="7">�^�[�~�i��</td>
                <td nowrap colspan="3">����A��</td>
              </tr>
              <tr bgcolor="#FFCC33" align="center"> 
                <td nowrap rowspan="2" bgcolor="#FFFF99">�D��</td>
                <td nowrap rowspan="2" bgcolor="#FFFF99">Voyage No.</td>
                <td nowrap rowspan="2" bgcolor="#FFFF99">���݊���<br>
                  ����<font size="-1"><sup>(��2)</sup></font></td>
                <td nowrap colspan="3" bgcolor="#FFFF99">���ݎ���</td>
                <td nowrap colspan="2" bgcolor="#FFFF99">�����m�F����</td>
                <td nowrap rowspan="2" bgcolor="#FFFF99">���o<BR>��</td>
                <td nowrap rowspan="2" bgcolor="#FFFF99">���o<br>����</td>
                <td nowrap rowspan="2" bgcolor="#FFFF99" valign=bottom>
<%
    If strSortKey="���^�Ǝ�" Then
        Response.Write "�w�藤�^<br>�Ǝ�<font size=-1><sup>(��3)</sup></font><BR><img src='gif/sort-r.gif' vspace=2></td>"
    Else
        Response.Write "�w�藤�^<br>�Ǝ�<font size=-1><sup>(��3)</sup></font><BR>"
        Response.Write "<a href='ms-impreload.asp?request=ms-implist1.asp&sort=���^�Ǝ�'><img src='gif/sort-b.gif' border=0 vspace=2></a>"
        Response.Write "</td>"
    End If
%>
                <td nowrap colspan="2" bgcolor="#FFFF99">�q�ɓ���</td>
              </tr>
              <tr bgcolor="#FFCC33" align="center"> 
                <td nowrap bgcolor="#FFFF99">�v��</td>
                <td nowrap bgcolor="#FFFF99">�\��</td>
                <td nowrap bgcolor="#FFFF99">����</td>
                <td nowrap bgcolor="#FFFF99">�\��</td>
                <td nowrap bgcolor="#FFFF99">����</td>
                <td nowrap bgcolor="#FFFF99" valign=bottom>
<%
    If strSortKey="�q�ɓ���" Then
        Response.Write "�w��<BR><img src='gif/sort-r.gif' vspace=2></td>"
    Else
        Response.Write "�w��<BR>"
        Response.Write "<a href='ms-impreload.asp?request=ms-implist1.asp&sort=�q�ɓ���'><img src='gif/sort-b.gif' border=0 vspace=2></a>"
        Response.Write "</td>"
    End If
%>
                <td nowrap bgcolor="#FFFF99" valign=top>����</td>
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
<% ' �D��
    If anyTmp(6)<>"" Then
        Response.Write anyTmp(6)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' Voyage No.
    If anyTmp(3)<>"" Then
        Response.Write anyTmp(3)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<%
    If strUserKind="�C��" Then
        strTemp=anyTmp(7)
    Else
        strTemp=anyTmp(8)
    End If
    If strTemp<>"" Then
        Response.Write strTemp
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �D��
    If anyTmp(20)<>"" Then
        Response.Write anyTmp(20)
    ElseIf anyTmp(15)<>"" Then
        Response.Write anyTmp(15)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' BL No
    If anyTmp(0)<>"" Then
        Response.Write anyTmp(0)
    Else
        Response.Write "<br>"
    End If
%>
				</td>
                <td nowrap align=center valign=middle>
<% ' �R���e�iNo.
    Response.Write "<a href='ms-impdetail.asp?line=" & LineNo & "&return=1'>" & anyTmp(1) & "</a>"
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �d�o�` - ���݊���
    Response.Write DispDateTimeCell(anyTmp(41),10)
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �^�[�~�i�� �| ���݃X�P�W���[��
    If anyTmp(61)<>"" Then
        Response.Write "<font color='#0000FF'>"
    End If
    Response.Write DispDateTimeCell(anyTmp(61),10)
    If anyTmp(61)<>"" Then
        Response.Write "</font>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �^�[�~�i�� - ���ݗ\��
    If anyTmp(32)<>"" Then
        bLate = false
        If anyTmp(33)<>"" Then
            If anyTmp(32)<anyTmp(33) Then
                bLate = true
            End If
        End If
        If anyTmp(61)<>"" Then
            If anyTmp(61)<anyTmp(32) Then
                bLate = true
            End If
        End If
        If bLate Then
            Response.Write "<font color='#FF0000'>"
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
<% ' �^�[�~�i�� - ���݊���
    Response.Write DispDateTimeCell(anyTmp(33),10)
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �^�[�~�i�� - �����m�F�\��
    If anyTmp(62)<>"" Then
        If anyTmp(48)<>"" Then
            If anyTmp(62)<anyTmp(48) Then
                Response.Write "<font color='#FF0000'>"
            Else
                Response.Write "<font color='#0000FF'>"
            End If
        Else
            Response.Write "<font color='#0000FF'>"
        End If
        Response.Write DispDateTimeCell(anyTmp(62),10)
        Response.Write "</font>"
    Else
        Response.Write DispDateTimeCell(anyTmp(62),10)
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �^�[�~�i�� - ���[�h����(�m�F)����
    Response.Write DispDateTimeCell(anyTmp(48),10)
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �^�[�~�i�����o��
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
<% ' �^�[�~�i�� - ���[�h���o����
    Response.Write DispDateTimeCell(anyTmp(43),10)
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ����A�� - �w�藤�^�Ǝ�
    If anyTmp(9)<>"" Then
        If anyTmp(14)<>"" Then
            Response.Write anyTmp(9)
        Else
            Response.Write "<font color='#0000FF'>" & anyTmp(9) & "</font>"
        End If
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ����A�� - �w��
    If anyTmp(64)<>"" Then
        strTemp=anyTmp(64)
    Else
        strTemp=anyTmp(13)
    End If
    If strTemp<>"" Then
        If strTemp<anyTmp(45) Then
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
<% ' ����A�� - ����
    Response.Write DispDateTimeCell(anyTmp(45),10)
%>
                </td>
              </tr>
<%
    Loop
%>
<!-- �����܂� -->
            </table>
<form>
      <input type=button value='�\���f�[�^�̍X�V' OnClick="JavaScript:window.location.href='ms-impreload.asp?request=ms-implist1.asp'">
</form>
          </td>
        </tr>
      </table>
      <form action="ms-impcsvout.asp"><input type="submit" value="CSV�t�@�C���o��">
<%
    If strUserKind="�C��" Then
        Response.Write "<a href='help16.asp'>CSV�t�@�C���o�͂Ƃ́H</a>"
    Else
        Response.Write "<a href='help18.asp'>CSV�t�@�C���o�͂Ƃ́H</a>"
    End If
%>
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
