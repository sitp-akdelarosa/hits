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
    WriteLog fs, "1106","�A�o�R���e�i�Ɖ�-�׎�p���ꗗ","00", ","

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
          <td rowspan=2><img src="gif/expninushi.gif" width="506" height="73"></td>
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
                <td nowrap><b>�A�o�R���e�i���ꗗ(�׎�p)&nbsp;</b></td>
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
            <td><font color="#000000" size="-1">�i��2�j�F���� &nbsp;&nbsp; ���F�Ɖ��</font></td>
          </tr>
        </table>

            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center" bgcolor="#FFCC33">
                <td nowrap rowspan="2" valign=bottom>
<%
    If strSortKey="�׎�Ǘ��ԍ�" Then
        Response.Write "�׎�Ǘ��ԍ�<BR><img src='gif/1.gif' height=6><BR><img src='gif/sort-r.gif' vspace=2></td>"
    Else
        Response.Write "�׎�Ǘ��ԍ�<BR><img src='gif/1.gif' height=6><BR>"
        Response.Write "<a href='ms-expreload.asp?request=ms-explist3.asp&sort=�׎�Ǘ��ԍ�'><img src='gif/sort-b.gif' border=0 vspace=2></a>"
        Response.Write "</td>"
    End If
%>
                <td nowrap rowspan="2" valign=bottom>Booking No.<BR><img src="gif/1.gif" height=18></td>
                <td nowrap rowspan="2" valign=bottom>�R���e�iNo.<font size="-1"><sup>(��1)</sup></font><BR><img src="gif/1.gif" height=18></td>
                <td nowrap rowspan="2" valign=bottom>
<%
    If strSortKey="�C��" Then
        Response.Write "�C��<font size=-1><sup>(��2)</sup></font><BR><img src='gif/1.gif' height=6><BR><img src='gif/sort-r.gif' vspace=2></td>"
    Else
        Response.Write "�C��<font size=-1><sup>(��2)</sup></font><BR><img src='gif/1.gif' height=6><BR>"
        Response.Write "<a href='ms-expreload.asp?request=ms-explist3.asp&sort=�C��'><img src='gif/sort-b.gif' border=0 vspace=2></a>"
        Response.Write "</td>"
    End If
%>
                <td colspan="2" nowrap>�q�ɓ���</td>
                <td nowrap rowspan="2">�o���j���O<br>����</td>
                <td colspan="2" nowrap>CY����</td>
                <td nowrap rowspan="2">�D��<br>����</td>
                <td nowrap rowspan="2">����<br>����</td>
                <td colspan="2" nowrap>�d���`����</td>
              </tr>
              <tr align="center" bgcolor="#FFFF99"> 
                <td nowrap valign=bottom><font color="#000000">
<%
    If strSortKey="�q�ɓ���" Then
        Response.Write "�w��<BR><img src='gif/sort-r.gif' vspace=2></a></font></td>"
    Else
        Response.Write "�w��<BR>"
        Response.Write "<a href='ms-expreload.asp?request=ms-explist3.asp&sort=�q�ɓ���'><img src='gif/sort-b.gif' border=0 vspace=2></a>"
        Response.Write "</font></td>"
    End If
%>
                <td nowrap valign=top><font color="#000000">����</font></td>
                <td nowrap><font color="#000000">�w��</font></td>
                <td nowrap><font color="#000000">����</font></td>
                <td nowrap><font color="#000000">�\��</font></td>
                <td nowrap><font color="#000000">����</font></td>
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
<% ' �׎��� - �Ǘ��ԍ�
    Response.Write anyTmp(14)
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
<% ' �C�ݖ�
    If anyTmp(17)="" Then
        Response.Write "<font color='#0000FF'>"
    End If
    If anyTmp(8)<>"" Then
        Response.Write anyTmp(8)
    Else
        Response.Write "<br>"
    End If
    If anyTmp(17)="" Then
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
<% ' �o���j���O - ��������
    Response.Write DispDateTimeCell(anyTmp(48),10)
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
    Response.Write DispDateTimeCell(anyTmp(49),10)
%>
                </td>
                <td align="center" nowrap> 
<% ' �^�[�~�i�� - �D�ϊ���
    Response.Write DispDateTimeCell(anyTmp(50),10)
%>
                </td>
                <td align="center" nowrap>
<% ' �^�[�~�i�� - ���݊���
    Response.Write DispDateTimeCell(anyTmp(51),10)
%>
                </td>
                <td align="center" nowrap>
<% ' �d���` - ���ݗ\��
    If anyTmp(53)<>"" Then
        If anyTmp(52)<>"" Then
            If anyTmp(53)<anyTmp(52) Then
                Response.Write "<font color='#FF0000'>"
            Else
                Response.Write "<font color='#0000FF'>"
            End If
        Else
            Response.Write "<font color='#0000FF'>"
        End If
        Response.Write DispDateTimeCell(anyTmp(53),10)
        Response.Write "</font>"
    Else
        Response.Write DispDateTimeCell(anyTmp(53),10)
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' �d���` - ���݊���
    Response.Write DispDateTimeCell(anyTmp(52),10)
%>
                </td>
              </tr>
<%
    Loop
%>
<!-- �����܂� -->
            </table>
<form>
      <input type=button value='�\���f�[�^�̍X�V' OnClick="JavaScript:window.location.href='ms-expreload.asp?request=ms-explist3.asp'">
</form>
          </td>
        </tr>
      </table>
      <form action="ms-expcsvout.asp"><input type="submit" value="CSV�t�@�C���o��">
    �@<a href="help15.asp">CSV�t�@�C���o�͂Ƃ́H</a> 
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
