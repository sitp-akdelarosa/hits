<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' Temp�t�@�C�������̃`�F�b�N
    CheckTempFile "MSEXPORT", "expentry.asp"

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
        Response.Redirect "http://www.hits-h.com/index.asp"             '���j���[��ʂ�
        Response.End
    End If
    strFileName="./temp/" & strFileName

    ' �A�o�R���e�i�Ɖ�X�g�\��
    WriteLog fs, "1104","�A�o�R���e�i�Ɖ�-�C�ݗp���ꗗ","00", ","

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
          <td rowspan=2><img src="gif/expkaika.gif" width="506" height="73"></td>
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
                <td nowrap><b>�A�o�R���e�i���ꗗ(�C�ݗp)&nbsp;</b></td>
                <td><img src="gif/hr.gif"></td>
              </tr>
            </table>
			<BR>
			&nbsp;&nbsp;&nbsp;&nbsp;���ړ��̃{�^���������ƁA���̍��ڂ̒l�Ń\�[�g����܂��B<BR><BR>
        <table border="0" cellspacing="2" cellpadding="1">
          <tr> 
            <td width="15"><BR></td>
            <td><font color="#000000" size="-1">�i��1) �N���b�N�ŒP�ƃR���e�i����\��</font></td>
            <td width="15"><BR></td>
            <td><font color="#000000" size="-1">�i��2�j�F���� &nbsp;&nbsp; ���F�Ɖ��</font></td>
            <td width="15"><BR></td>
            <td><font color="#000000" size="-1">�i��3�j96=HC</font></td>
          </tr>
        </table>

            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center" bgcolor="#FFCC33"> 
                <td nowrap valign=bottom rowspan=2><font color="#000000">
<%
    If strSortKey="�׎喼" Then
        Response.Write "�׎�Ǘ��ԍ�<BR><img src='gif/1.gif' height=8><BR><img src='gif/sort-r.gif' vspace=2></font></td>"
    Else
        Response.Write "�׎�Ǘ��ԍ�<BR><img src='gif/1.gif' height=8><BR>"
        Response.Write "<a href='ms-expreload.asp?request=ms-explist1.asp&sort=�׎喼'><img src='gif/sort-b.gif' border=0 vspace=2></a>"
        Response.Write "</font></td>"
    End If
%>
                <td nowrap rowspan="2" valign=bottom>Booking No.<BR><img src='gif/1.gif' height=20></td>
                <td nowrap rowspan="2" valign=bottom>�R���e�iNo.<font size="-1"><sup>(��1)</sup></font><BR><img src='gif/1.gif' height=20></td>
                <td nowrap rowspan="2" valign=bottom>
<%
    If strSortKey="���^�Ǝ�" Then
        Response.Write "�w�藤�^<br>�Ǝ�<font size='-1'><sup>(��2)</sup></font><BR><img src='gif/sort-r.gif' vspace=2></td>"
    Else
        Response.Write "�w�藤�^<br>�Ǝ�<font size='-1'><sup>(��2)</sup></font><BR>"
        Response.Write "<a href='ms-expreload.asp?request=ms-explist1.asp&sort=���^�Ǝ�'><img src='gif/sort-b.gif' border=0 vspace=2></a>"
        Response.Write "</td>"
    End If
%>
                <td colspan="4" nowrap>��R���e�i</td>
                <td colspan="2" nowrap>�q�ɓ���</td>
                <td colspan="2" nowrap>CY����</td>
                <td colspan="2" nowrap>�{�D</td>
                <td colspan="3" nowrap>�o���j���O��R���e�i</td>
              </tr>
              <tr align="center" bgcolor="#FFFF99"> 
                <td nowrap><font color="#000000">��R��<br>���ꏊ</font></td>
                <td nowrap><font color="#000000">�T�C�Y</font></td>
                <td nowrap><font color="#000000">����<BR><font size="-1"><sup>(��3)</sup></font></font></td>
                <td nowrap><font color="#000000">���[�t�@�[</font></td>
                <td nowrap valign=bottom><font color="#000000">
<%
    If strSortKey="�q�ɓ���" Then
        Response.Write "�w��<BR><img src='gif/1.gif' height=3><BR><img src='gif/sort-r.gif' vspace=2></font></td>"
    Else
        Response.Write "�w��<BR><img src='gif/1.gif' height=3><BR>"
        Response.Write "<a href='ms-expreload.asp?request=ms-explist1.asp&sort=�q�ɓ���'><img src='gif/sort-b.gif' border=0 vspace=2></a>"
        Response.Write "</font></td>"
    End If
%>
                <td nowrap valign=bottom><font color="#000000">����</font><BR><img src="gif/1.gif" height=15></td>
                <td nowrap valign=bottom><font color="#000000">
<%
    If strSortKey="CY����" Then
        Response.Write "�w��<BR><img src='gif/1.gif' height=3><BR><img src='gif/sort-r.gif' vspace=2></font></td>"
    Else
        Response.Write "�w��<BR><img src='gif/1.gif' height=3><BR>"
        Response.Write "<a href='ms-expreload.asp?request=ms-explist1.asp&sort=CY����'><img src='gif/sort-b.gif' border=0 vspace=2></a>"
        Response.Write "</font></td>"
    End If
%>
                <td nowrap valign=bottom><font color="#000000">����</font><BR><img src="gif/1.gif" height=15></td>
                <td nowrap><font color="#000000">�D��</font></td>
                <td nowrap><font color="#000000">�d���`��</font></td>
                <td nowrap><font color="#000000">�V�[��<br>No.</font></td>
                <td nowrap><font color="#000000">�ݕ�<br>�d��(t)</font></td>
                <td nowrap><font color="#000000">��<br>�d��(t)</font></td>
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
<% ' �׎��� - ���́A�Ǘ��ԍ�
    Response.Write anyTmp(7) & "<br>" & anyTmp(14)
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' Booking No
    Response.Write anyTmp(0)
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �R���e�iNo.
    If anyTmp(1)<>"" Then
        Response.Write "<a href='ms-expdetail.asp?line=" & LineNo & "'>" & anyTmp(1) & "</a>"
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ���^�Ǝ҃R�[�h
    If anyTmp(17)="" Then
        Response.Write "<font color='#0000FF'>"
    End If
    If anyTmp(9)<>"" Then
        Response.Write anyTmp(9)
    Else
        Response.Write "<br>"
    End If
    If anyTmp(17)="" Then
        Response.Write "</font>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ��R���e�i - ��R�����ꏊ
    If anyTmp(32)<>"" Then
        Response.Write anyTmp(32)
    Else
        If anyTmp(20)<>"" Then
            Response.Write "<font color='#0000FF'>" & anyTmp(20) & "</font>"
        Else
            Response.Write "<br>"
        End If
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ��R���e�i - �T�C�Y
    If anyTmp(33)<>"" Then
        Response.Write anyTmp(33)
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
<% ' ��R���e�i - ����
    If anyTmp(34)<>"" Then
        Response.Write anyTmp(34)
    Else
        If anyTmp(12)<>"" Then
            Response.Write "<font color='#0000FF'>" & anyTmp(12) & "</font>"
        Else
            Response.Write "<br>"
        End If
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ��R���e�i - ���[�t�@�[
    If anyTmp(35)<>"" Then
        If anyTmp(35)="R" Then
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
<% ' ����^�� - �q�ɓ����X�P�W���[��
    If anyTmp(56)<>"" Then
        strTemp=anyTmp(56)
    Else
        strTemp=anyTmp(15)
    End If
    If strTemp<>"" Then
        If strTemp<anyTmp(47) Then
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
    Response.Write DispDateTimeCell(anyTmp(47),10)
%>
				</td>
                <td nowrap align=center valign=middle>
<% ' ����^�� - CY�����X�P�W���[��
    If anyTmp(60)<>"" Then
        strTemp=anyTmp(60)
    Else
        strTemp=anyTmp(16)
    End If
    If strTemp<>"" Then
        If strTemp<anyTmp(49) Then
            Response.Write "<font color='#FF0000'>"
        Else
            Response.Write "<font color='#0000FF'>"
        End If
    End If
    Response.Write DispDateTimeCell(strTemp,5)
    If strTemp<>"" Then
        Response.Write "</font>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ����^�� - CY����
    Response.Write DispDateTimeCell(anyTmp(49),5)
%>
                </td>
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
<% ' �d���`��
    If anyTmp(44)<>"" Then
        Response.Write anyTmp(44)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �V�[��No.
    If anyTmp(37)<>"" Then
        Response.Write anyTmp(37)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �ݕ��d��
    If anyTmp(57)<>"" And anyTmp(57)<>"0" Then
        dWeight=anyTmp(57) / 10
        Response.Write dWeight
    Else
        Response.Write "�|"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ���d��
    If anyTmp(38)<>"" And anyTmp(38)<>"0" Then
        dWeight=anyTmp(38) / 10
        Response.Write dWeight
    Else
        Response.Write "�|"
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
      <input type=button value='�\���f�[�^�̍X�V' OnClick="JavaScript:window.location.href='ms-expreload.asp?request=ms-explist1.asp'">
      </form>
          </td>
        </tr>
      </table>
      <form action="ms-expcsvout.asp"><input type="submit" value="CSV�t�@�C���o��">
    �@<a href="help13.asp">CSV�t�@�C���o�͂Ƃ́H</a> 
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
    DispMenuBarBack "ms-expentry.asp"
%>
</body>
</html>
