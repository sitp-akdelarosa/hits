<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' Temp�t�@�C�������̃`�F�b�N
    CheckTempFile "MSEXPORT", "expentry.asp"

    ' �w������̎擾
    Dim iLineNo
    iLineNo = CInt(Request.QueryString("line"))
    Dim iReturn
    iReturn = Session.Contents("dispreturn")

    ' ���[�U��ނ��`�F�b�N����
    strUserKind=Session.Contents("userkind")

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

    ' �\���t�@�C����Open
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

    ' �ڍו\���s�̃f�[�^�̎擾
    LineNo=0
    Do While Not ti.AtEndOfStream
        anyTmp=Split(ti.ReadLine,",")
        LineNo=LineNo+1
        If iLineNo=LineNo Then
           Exit Do
        End If
    Loop
    ti.Close

    ' �A�o�R���e�i�Ɖ�ڍ�
    WriteLog fs, "1108","�A�o�R���e�i�Ɖ�-�R���e�i�ڍ�", "00", anyTmp(1) & ","

    Session.Contents("dispexpctrl")=anyTmp(14)     ' �\���׎�Ǘ��ԍ����L��
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
<body bgcolor="DEE1FF" text="#000000" link="#0000FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" vlink="#0000FF">
<!-------------��������ڍ׉��--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/expdetailt.gif" width="506" height="73"></td>
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
          <td>�@<br>
            <table border=1 cellpadding="3" cellspacing="1">
              <tr> 
<% ' Booking No
    Response.Write "<td bgcolor='#003399' background='gif/tableback.gif' nowrap><font color='#FFFFFF'><b>Booking No</b></font></td>"
    Response.Write "<td bgcolor='#FFFFFF' nowrap>" & anyTmp(0) & "</td>"
%>
                <td bgcolor="#003399" background="gif/tableback.gif" nowrap><font color="#FFFFFF"><b>�R���e�iNo.</b></font></td>
                <td bgcolor="#FFFFFF" nowrap>
<% ' �R���e�iNo.
    Response.Write anyTmp(1)
%>
                </td>
              </tr>
            </table>
<BR>
            <table>
              <tr>
                <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>��{���@</b></td>
                <td><img src="gif/hr.gif"></td>
              </tr>
            </table>
            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center" bgcolor="#FFCC33"> 
                <td colspan="4" nowrap>��R���e�i</td>
                <td colspan="5" nowrap bgcolor="#FFCC33">�o���j���O��R���e�i</td>
                <td bgcolor="#FFCC33" nowrap colspan="2">������t����</td>
              </tr>
              <tr align="center" bgcolor="#FFFF99"> 
                <td nowrap><font color="#000000">��R�����ꏊ</font></td>
                <td nowrap><font color="#000000">�T�C�Y</font></td>
                <td nowrap>����<font size="-1"><sup>(��1)</sup></font></td>
                <td nowrap><font color="#000000">���[�t�@�[</font></td>
                <td nowrap><font color="#000000">�V�[��No.</font></td>
                <td nowrap><font color="#000000">�ݕ��d��(t)</font></td>
                <td nowrap><font color="#000000">���d��(t)</font></td>
                <td nowrap><font color="#000000">�댯�i</font></td>
                <td nowrap><font color="#000000">�����^�[�~�i����</font></td>
                <td nowrap><font color="#000000">�I�[�v����</font></td>
                <td nowrap>�N���[�Y��</td>
              </tr>
              <tr> 
                <td nowrap align="center">
<% ' ��R�����ꏊ
    If anyTmp(32)<>"" Then
        Response.Write anyTmp(32)
    ElseIf anyTmp(20)<>"" Then
        Response.Write "<font color='#0000FF'>" & anyTmp(20) & "</font>"
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' �T�C�Y
    If anyTmp(33)<>"" Then
        Response.Write anyTmp(33)
    ElseIf anyTmp(10)<>"" Then
        Response.Write "<font color='#0000FF'>" & anyTmp(10) & "</font>"
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' ����
    If anyTmp(34)<>"" Then
        Response.Write anyTmp(34)
    ElseIf anyTmp(12)<>"" Then
        Response.Write "<font color='#0000FF'>" & anyTmp(12) & "</font>"
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' ���[�t�@�[
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
                <td align="center" nowrap>
<% ' �V�[��No.
    If anyTmp(37)<>"" Then
        Response.Write anyTmp(37)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' �ݕ��d�� $�ǉ�
    If anyTmp(57)<>"" And anyTmp(57)<>"0" Then
        dWeight=anyTmp(57) / 10
        Response.Write dWeight
    Else
        Response.Write "�|"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' ���d��
    If anyTmp(38)<>"" And anyTmp(38)<>"0" Then
        dWeight=anyTmp(38) / 10
        Response.Write dWeight
    Else
        Response.Write "�|"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' �댯�i
    If anyTmp(61)="H" Then
        Response.Write "��"
    ElseIf anyTmp(61)<>"" Then
        Response.Write "�|"
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' �����^�[�~�i����
    If anyTmp(36)<>"" Then
        Response.Write anyTmp(36)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' CY�I�[�v��
    Response.Write DispDateTimeCell(anyTmp(39),5)
%>
                </td>
                <td align="center" nowrap>
<% ' CY�N���[�Y
    Response.Write DispDateTimeCell(anyTmp(40),5)
%>
                </td>
              </tr>
            </table>
            <table border="0" cellspacing="2" cellpadding="1">
              <tr> 
                <td width="15">&nbsp;</td>
                <td><font color="#000000" size="-1">(��1)96=HC</font></td>
              </tr>
            </table>
<BR>
            <table>
              <tr>
                <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>�ʒu���@</b></td>
                <td><img src="gif/hr.gif"></td>
              </tr>
            </table> 
            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center" bgcolor="#FFCC33"> 
                <td nowrap>�ꏊ</td>
                <td colspan="3" nowrap>����A��</td>
                <td nowrap bgcolor="#FFCC33">�X�g�b�N���[�h</td>
                <td colspan="4" nowrap bgcolor="#FFCC33">�^�[�~�i��</td>
                <td bgcolor="#FFCC33" nowrap>�d���`</font></td>
              </tr>
              <tr align="center" bgcolor="#FFFF99"> 
                <td nowrap rowspan="2"><font color="#000000">�H��</font></td>
                <td nowrap rowspan="2"><font color="#000000">��R�����</font></td>
                <td nowrap><font color="#000000">�q�ɓ���</font></td>
<% 
	Dim iSupNum
	If anyTmp(34)<>"" And strUserKind="���^" Then
		iSupNum = 3
%>
                <td nowrap rowspan="2"><font color="#000000">�o���j���O</font><font size="-1"><sup>(��2)</sup></font></td>
<%
	Else
		iSupNum = 2
%>
                <td nowrap rowspan="2"><font color="#000000">�o���j���O</font></td>
<% End If %>
                <td nowrap><font color="#000000">����</font></td>
                <td nowrap><font color="#000000">CY����</font></td>
                <td nowrap rowspan="2"><font color="#000000">�D�ϊ���</font></td>
                <td nowrap colspan="2"><font color="#000000">����</font></td>
                <td nowrap><font color="#000000">���ݎ���</font><font size="-1"><sup>(��<%=iSupNum%>)</sup></font></td>
              </tr>
              <tr align="center" bgcolor="#FFFF99"> 
<% If anyTmp(34)<>"" And strUserKind="���^" Then %>
                <td nowrap><font color="#000000">�w���^����</font><font size="-1"><sup>(��2)</sup></font></td>
<% Else %>
                <td nowrap><font color="#000000">�w���^����</font></td>
<% End If %>
                <td nowrap><font color="#000000">�\��^����</font></td>
                <td nowrap><font color="#000000">�w���^����</font></td>
                <td nowrap><font color="#000000">�v��</font></td>
                <td nowrap><font color="#000000">�\��^����</font></td>
                <td nowrap><font color="#000000">�\��^����</font></td>
              </tr>
              <tr> 
                <td nowrap rowspan="2" bgcolor="#FFFFCC" align="center"><font color="#000000">����</font></td>
                <td rowspan="2" align="center" nowrap>
<% ' ����^�� - ��R�����
    Response.Write DispDateTimeCell(anyTmp(46),11)
%>
                </td>
                <td align="center" nowrap>
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
    Response.Write DispDateTimeCell(strTemp,11)
    If strTemp<>"" Then
        Response.Write "</font>"
    End If
%>
                </td>
                <td rowspan="2" align="center" nowrap> 
<% ' ����^�� - �o���j���O
    If anyTmp(34)<>"" And strUserKind="���^" Then
        Response.Write "<a href='ms-expinput.asp?kind=2&line=" & LineNo & "&request=ms-expdetail.asp'>"
    End If
    strTemp = DispDateTimeCell(anyTmp(48),11)
    If Left(strTemp,1)="<" And anyTmp(34)<>"" Then
        Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"
    Else
        Response.Write strTemp
    End If
    If anyTmp(34)<>"" And strUserKind="���^" Then
        Response.Write "</a>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' �X�g�b�N���[�h - �����\�� $�ǉ�
    sTemp=DispReserveCell(anyTmp(58),anyTmp(59),sColor)
    Response.Write sColor
    Response.Write sTemp
    If sColor<>"" Then
        Response.Write "</font>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' �^�[�~�i�� - CY�����w�� $�ǉ�
    If anyTmp(60)<>"" Then
        strTemp=anyTmp(60)
    Else
        strTemp=anyTmp(16)
    End If
    If strTemp<>"" Then
        If Left(strTemp,10)<Left(anyTmp(49),10) Then
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
                <td rowspan="2" align="center" nowrap> 
<% ' �^�[�~�i�� - �D�ϊ���
    Response.Write DispDateTimeCell(anyTmp(50),11)
%>
                </td>
                <td rowspan="2" align="center" nowrap>
<% ' �^�[�~�i�� - ���݃X�P�W���[��
    If anyTmp(55)<>"" Then
        Response.Write "<font color='#0000FF'>"
    End If
    Response.Write DispDateTimeCell(anyTmp(55),5)
    If anyTmp(55)<>"" Then
        Response.Write "</font>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' �^�[�~�i�� - ���ݗ\��
    If anyTmp(45)<>"" Then
        bLate = false
        If anyTmp(51)<>"" Then
            If anyTmp(45)<anyTmp(51) Then
                bLate = true
            End If
        End If
        If anyTmp(55)<>"" Then
            If Left(anyTmp(55),10)<Left(anyTmp(45),10) Then
                bLate = true
            End If
        End If
        If bLate Then
            Response.Write "<font color='#FF0000'>"
        Else
            Response.Write "<font color='#0000FF'>"
        End If
        Response.Write DispDateTimeCell(anyTmp(45),11)
        Response.Write "</font>"
    End If
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
        Response.Write DispDateTimeCell(anyTmp(53),11)
        Response.Write "</font>"
    Else
        Response.Write DispDateTimeCell(anyTmp(53),11)
    End If
%>
                </td>
              </tr>
              <tr>
                <td align="center" nowrap>
<% ' ����^�� - �q�ɓ���
    If anyTmp(34)<>"" And strUserKind="���^" Then
        Response.Write "<a href='ms-expinput.asp?kind=1&line=" & LineNo & "&request=ms-expdetail.asp'>"
    End If
    strTemp = DispDateTimeCell(anyTmp(47),11)
    If Left(strTemp,1)="<" And anyTmp(34)<>"" Then
        Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"
    Else
        Response.Write strTemp
    End If
    If anyTmp(34)<>"" And strUserKind="���^" Then
        Response.Write "</a>"
    End If
%>
                </td>
                <td align="center" nowrap> 
<% ' �X�g�b�N���[�h - ��������
    Response.Write DispDateTimeCell(anyTmp(54),11)
%>
                </td>
                <td align="center" nowrap> 
<% ' �^�[�~�i�� - CY��������
    Response.Write DispDateTimeCell(anyTmp(49),11)
%>
                </td>
                <td align="center" nowrap>
<% ' �^�[�~�i�� - ���݊���
    Response.Write DispDateTimeCell(anyTmp(51),11)
%>
                </td>
                <td align="center" nowrap>
<% ' �d���` - ���݊���
    Response.Write DispDateTimeCell(anyTmp(52),11)
%>
                </td>
              </tr>
            </table>
            <table border="0" cellspacing="2" cellpadding="1">
              <tr> 
<% If anyTmp(34)<>"" And strUserKind="���^" Then %>
                <td width="15">&nbsp;</td>
                <td><font color="#000000" size="-1">�i��2�j�N���b�N�Ŋ����������͉�ʂ�</font></td>
<% End If %>
                <td width="15">&nbsp;</td>
                <td><font color="#000000" size="-1">�i��<%=iSupNum%>�j�d���`�̎����́A���n���Ԃł��B</font></td>
              </tr>
            </table>
<BR>
            <table>
              <tr>
                <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>�{�D���@</b></td>
                <td><img src="gif/hr.gif"></td>
              </tr>
            </table>
            <table border=1 cellpadding="3" cellspacing="1">
              <tr> 
                <td bgcolor="#FFCC33" nowrap>�D��</td>
                <td bgcolor="#FFFFFF">
<% ' �D��
    If anyTmp(41)<>"" Then
        Response.Write anyTmp(41)
    ElseIf anyTmp(24)<>"" Then
        Response.Write anyTmp(24)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td bgcolor="#FFCC33" nowrap><font color="#000000">�D��</font></td>
                <td bgcolor="#FFFFFF" nowrap>
<% ' �D��
    If anyTmp(42)<>"" Then
        Response.Write anyTmp(42)
    ElseIf anyTmp(2)<>"" Then
        Response.Write anyTmp(2)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td bgcolor="#FFCC33" nowrap>Voyage No.<font color="#FFFFFF"><b> 
                </b></font></td>
                <td bgcolor="#FFFFFF" nowrap>
<% ' ���q
    If anyTmp(43)<>"" Then
        Response.Write anyTmp(43)
    ElseIf anyTmp(3)<>"" Then
        Response.Write anyTmp(3)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td bgcolor="#FFCC33" nowrap>�d���`</td>
                <td bgcolor="#FFFFFF" nowrap>
<% ' �d���`
    If anyTmp(44)<>"" Then
        Response.Write anyTmp(44)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
              </tr>
            </table>
            <br>
<form>
      <input type=button value='�\���f�[�^�̍X�V' OnClick="JavaScript:window.location.href='ms-expreload.asp?request=ms-expdetail.asp'">
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
<!-------------�ڍ׉�ʏI���--------------------------->
<%
    If iReturn=1 Then
        DispMenuBarBack "ms-explist1.asp"
    ElseIf iReturn=2 Then
        DispMenuBarBack "ms-explist2.asp"
    ElseIf iReturn=3 Then
        DispMenuBarBack "ms-explist3.asp"
    End If
%>
</body>
</html>
