<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' Temp�t�@�C�������̃`�F�b�N
    CheckTempFile "MSEXPORT", "index.asp"

    strSortKey=Session.Contents("sortkey")

	'���
	Dim iLoginKind,sLoginKind
	iLoginKind = Request.QueryString("kind")
	Select Case iLoginKind
		Case "1"	sLoginKind = "�C��"
					iNum = "a105"
		Case "2"	sLoginKind = "���^"
					iNum = "a106"
		Case "3"	sLoginKind = "�׎�"
					iNum = "a107"
		Case "4"	sLoginKind = "�`�^"
					iNum = "a108"
		Case Else
	End Select

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
    WriteLog fs, iNum,"��R���s�b�N�A�b�v�V�X�e��-" & sLoginKind & "�p���ꗗ","00", ","

    ' �\���t�@�C����Open
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

    '�߂��ʎ�ʂ��L��
    Session.Contents("dispreturn")=4
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
    If sLoginKind="�C��" Then
        Response.Write "<img src='gif/pickkat.gif' width='506' height='73'>"
    ElseIf sLoginKind="���^" Then
        Response.Write "<img src='gif/pickrit.gif' width='506' height='73'>"
    ElseIf sLoginKind="�׎�" Then
        Response.Write "<img src='gif/picknit.gif' width='506' height='73'>"
    Else
        Response.Write "<img src='gif/pickkot.gif' width='506' height='73'>"
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
                <td nowrap><b>��R���s�b�N�A�b�v���ꗗ(<%=sLoginKind%>�p)&nbsp;</b></td>
                <td><img src="gif/hr.gif"></td>
              </tr>
            </table>
			<BR>
      <form method=post action="picklist-syori.asp">

<table border=0 cellpadding=0 cellspacing=0 width=500><tr><td>

        <table border="0" cellspacing="2" cellpadding="1">
          <tr> 
            <td width="15"><BR></td>
            <td valign=top nowrap><font color="#000000" size="-1">�i��1�j96=HC</font></td>
            <td width="15"><BR></td>
            <td valign=top nowrap><font color="#000000" size="-1">�i��2�j���F�m�F�� &nbsp;&nbsp; �ԁF�ύX &nbsp;&nbsp; �F���m�F</font></td>

<% If sLoginKind="�`�^" Then %>
            <td width="15"><BR></td>
            <td valign=top nowrap><font color="#000000" size="-1">�i��3�j</font></td>
			<td valign=top><font color="#000000" size="-1">���e���m�F��������́����N���b�N���ă`�F�b�N�}�[�N�����A���ꏊ�A�܂��́A�w�����ύX�������ꍇ�͕ύX�{�^�����A���Ȃ���Ίm�F�{�^���������Ă��������B</font></td>
<% End If %>
<% If sLoginKind="���^" Then %>
            <td width="15"><BR></td>
            <td valign=top nowrap><font color="#000000" size="-1">�i��3�j</font></td>
			<td valign=top><font color="#000000" size="-1">�w�����ύX�������ꍇ�́A�����N���b�N���ă`�F�b�N�}�[�N�����A���w����ύX�{�^���������Ă��������B</font></td>
<% End If %>

          </tr>
        </table>

            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center" bgcolor="#FFCC33"> 

<% If sLoginKind<>"�C��" Then %>
                <td nowrap>�֌W��</td>
<% Else %>
                <td nowrap colspan=3>�֌W��</td>
<% End If %>
                <td nowrap colspan=3>�D�Ϗ��</td>
                <td nowrap colspan=3>�K�v��R��</td>
                <td nowrap colspan=2>��R�����</td>
                <td nowrap colspan=2>�q�ɓ���</td>
                <td nowrap colspan=2>CY����</td>
<% If sLoginKind="�`�^" Then %>
                <td nowrap rowspan=3 colspan=2>�m�F�^�ύX<font size=-1><sup>�i��3�j</sup></font></td>
<% End If %>
<% If sLoginKind="���^" Then %>
                <td nowrap rowspan=3>���w���<BR>�ύX<font size=-1><sup>�i��3�j</sup></font></td>
<% End If %>
              </tr>
              <tr align="center" bgcolor="#FFFF99"> 
<% If sLoginKind<>"�C��" Then %>
                <td nowrap rowspan=2 valign=bottom>
					�C��<BR>
					<img src="gif/1.gif" height=4><BR>
	<% If strSortKey="�C��" Then %>
					<img src="gif/sort-r.gif" vspace=2></td>
	<% Else %>
					<a href="pickreload.asp?kind=<%=iLoginKind%>&sort=�C��"><img src="gif/sort-b.gif" border=0 vspace=2></a></td>
	<% End If %>
				
<% Else %>
                <td nowrap rowspan=2 valign=bottom>
					�׎�<BR>
					<img src="gif/1.gif" height=4><BR>
	<% If strSortKey="�׎�" Then %>
					<img src="gif/sort-r.gif" vspace=2></td>
	<% Else %>
					<a href="pickreload.asp?kind=<%=iLoginKind%>&sort=�׎�"><img src="gif/sort-b.gif" border=0 vspace=2></a></td>
	<% End If %>
				
                <td nowrap rowspan=2 valign=bottom>
					�`�^<BR>
					<img src="gif/1.gif" height=4><BR>
	<% If strSortKey="�`�^" Then %>
					<img src="gif/sort-r.gif" vspace=2></td>
	<% Else %>
					<a href="pickreload.asp?kind=<%=iLoginKind%>&sort=�`�^"><img src="gif/sort-b.gif" border=0 vspace=2></a></td>
	<% End If %>
				
                <td nowrap rowspan=2 valign=bottom>
					���^<BR>
					<img src="gif/1.gif" height=4><BR>
	<% If strSortKey="���^" Then %>
					<img src="gif/sort-r.gif" vspace=2></td>
	<% Else %>
					<a href="pickreload.asp?kind=<%=iLoginKind%>&sort=���^"><img src="gif/sort-b.gif" border=0 vspace=2></a></td>
	<% End If %>
				
<% End If %>
                <td nowrap rowspan=2>�D���^<BR>VoyageNo.</td>
                <td nowrap rowspan=2>BookingNo.</td>
                <td nowrap rowspan=2>�׎�Ǘ��ԍ�</td>
                <td nowrap rowspan=2>�T�C�Y</td>
                <td nowrap rowspan=2>����<BR><font size=-1><sup>�i��1�j</sup></font></td>
                <td nowrap rowspan=2>�^�C�v</td>
                <td nowrap rowspan=2 valign=bottom>���ꏊ<font size=-1><sup>�i��2�j</sup></font><BR><img src="gif/1.gif" height=14></td>
                <td nowrap rowspan=2 valign=bottom>
					�w���<font size=-1><sup>�i��2�j</sup></font><BR>
					<img src="gif/1.gif" height=2><BR>
	<% If strSortKey="�w���" Then %>
					<img src="gif/sort-r.gif" vspace=2></td>
	<% Else %>
					<a href="pickreload.asp?kind=<%=sLoginKind%>&sort=�w���"><img src="gif/sort-b.gif" border=0 vspace=2></a></td>
	<% End If %>
                <td nowrap rowspan=2>�ꏊ</td>
                <td nowrap rowspan=2>�w�����</td>
                <td nowrap rowspan=2>�ꏊ</td>
                <td nowrap>���݌v���</td>
              </tr>
              <tr align="center" bgcolor="#FFFF99"> 
                <td nowrap>�J�b�g��</td>
              </tr>

<!-- ��������f�[�^�J��Ԃ� -->
<% ' �\���t�@�C���̃��R�[�h������ԌJ��Ԃ�
    LineNo=0
    Do While Not ti.AtEndOfStream
        anyTmp=Split(ti.ReadLine,",")
        LineNo=LineNo+1
%>
              <tr bgcolor="#FFFFFF"> 

<% If sLoginKind<>"�C��" Then %>
                <td nowrap align=center valign=middle rowspan=2>
<% ' �C��
    If anyTmp(8)<>"" Then
        Response.Write anyTmp(8)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<% Else %>
                <td nowrap align=center valign=middle rowspan=2>
<% ' �׎�
    If anyTmp(7)<>"" Then
        Response.Write anyTmp(7)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle rowspan=2>
<% ' �`�^
    If anyTmp(16)<>"" Then
        Response.Write anyTmp(16)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle rowspan=2>
<% ' ���^
    If anyTmp(9)<>"" Then
        Response.Write anyTmp(9)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<% End If %>

                <td nowrap align=center valign=middle rowspan=2>
<% ' �D���^Voyage
    If anyTmp(2)<>"" Then
        Response.Write anyTmp(2) & "<BR>" & anyTmp(43)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle rowspan=2>
<% ' Booking
    If anyTmp(0)<>"" Then
        Response.Write anyTmp(0)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle rowspan=2>
<% ' �׎�Ǘ��ԍ�
    If anyTmp(14)<>"" Then
        Response.Write anyTmp(14)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle rowspan=2>
<% ' �T�C�Y
    If anyTmp(10)<>"" Then
        Response.Write anyTmp(10)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle rowspan=2>
<% ' ����
    If anyTmp(12)<>"" Then
        Response.Write anyTmp(12)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle rowspan=2>
<% ' �^�C�v
    If anyTmp(11)<>"" Then
        Response.Write anyTmp(11)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle rowspan=2>
<% ' ��R���|���ꏊ
    If anyTmp(20)<>"" Then
		If anyTmp(26)="1" Then
	        Response.Write anyTmp(20)
		ElseIf anyTmp(27)="1" Then
	        Response.Write "<font color=""#ff0000"">" & anyTmp(20) & "</font>"
		Else
	        Response.Write "<font color=""#0000ff"">" & anyTmp(20) & "</font>"
		End If
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle rowspan=2>
<% ' ��R���|�w���
    If anyTmp(24)<>"" Then
		If anyTmp(26)="1" Then
			Response.Write DispDateTimeCell(anyTmp(24),5)
		ElseIf anyTmp(28)="1" Then
	        Response.Write "<font color=""#ff0000"">" & DispDateTimeCell(anyTmp(24),5) & "</font>"
		Else
	        Response.Write "<font color=""#0000ff"">" & DispDateTimeCell(anyTmp(24),5) & "</font>"
		End If
    Else
        Response.Write "<hr width=80% >"
    End If
%>
                </td>
                <td nowrap align=center valign=middle rowspan=2>
<% ' �q�Ɂ|�ꏊ
    If anyTmp(13)<>"" Then
        Response.Write anyTmp(13)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle rowspan=2>
<% ' �q�Ɂ|�w�����
    If anyTmp(15)<>"" Then
        Response.Write DispDateTimeCell(anyTmp(15),10)
    Else
        Response.Write "<hr width=80% >"
    End If
%>
                </td>
                <td nowrap align=center valign=middle rowspan=2>
<% ' �b�x�|�ꏊ
    If anyTmp(22)<>"" Then
        Response.Write anyTmp(22)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �b�x�|���݌v���
    If anyTmp(45)<>"" Then
        Response.Write DispDateTimeCell(anyTmp(45),5)
    Else
        Response.Write "<hr width=80% >"
    End If
%>
                </td>

<% If sLoginKind="�`�^" Then %>
                <td nowrap align=center valign=middle rowspan=2>
<%		If anyTmp(26)="1" Then %>
					��</td>
<%		ElseIf anyTmp(26)="2" Then  %>
					�ύX</td>
<%		Else  %>
					<font color="#0000ff">��</font></td>
<%		End If  %>
<% End If %>

<% If sLoginKind="�`�^" Or sLoginKind="���^" Then %>
                <td nowrap align=center valign=middle rowspan=2>
					<input type=checkbox name="check<%=LineNo%>"></td>
<% End If %>

			</tr>
			<tr bgcolor="#FFFFFF">
                <td nowrap align=center valign=middle>
<% ' �b�x�|�J�b�g��
    If anyTmp(40)<>"" Then
        Response.Write DispDateTimeCell(anyTmp(40),5)
    Else
        Response.Write "<hr width=80% >"
    End If
%>
                </td>
			</tr>





<%
    Loop
%>
<!-- �����܂� -->
            </table>

<% If sLoginKind="�`�^" Then %>
</td></tr><tr><td align=right>
			<input type=hidden name="allline" value="<%=LineNo%>">
			<input type=submit name="ok" value=" �m �F "><input type=submit name="pickinput" value=" �� �X ">
</td></tr><tr><td>
<% ElseIf sLoginKind="���^" Then %>
</td></tr><tr><td align=right>
			<input type=hidden name="allline" value="<%=LineNo%>">
			<input type=submit name="pickinput" value=" ���w����ύX ">
</td></tr><tr><td>
<% Else %>
		<BR>
<% End If %>

		<input type=button value="�\���f�[�^�̍X�V" onClick="JavaScript:window.location.href='pickreload.asp?kind=<%=iLoginKind%>&sort=<%=sLoginKind%>'">

</td></tr></table>

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
    DispMenuBarBack "pickselect.asp"
%>
</body>
</html>
