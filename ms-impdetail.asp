<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' Temp�t�@�C�������̃`�F�b�N
    CheckTempFile "MSIMPORT", "impentry.asp"

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
        Response.Redirect "impentry.asp"             '�A���R���e�i�Ɖ�g�b�v
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

    ' �A���R���e�i�Ɖ�ڍ�
    WriteLog fs, "2108","�A���R���e�i�Ɖ�-�R���e�i�ڍ�", "00", anyTmp(1) & ","

    Session.Contents("dispcntnr")=anyTmp(1)     ' �\���R���e�iNo.���L��
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
<SCRIPT LANGUAGE="JavaScript">
function winOpen(winName,url,W,H){
  var WinD11=window.open(url,winName,'scrollbars=auto,resizable=yes,width='+W+',height='+H+'');
  WinD11.focus();
  WinD11.document.close();
}
</Script>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#0000ff" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" vlink="#0000ff">
<!-------------��������ڍ׉��--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/impdetailt.gif" width="506" height="73"></td>
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
<% ' BL No
    Response.Write "<td bgcolor='#003399' background='gif/tableback.gif' nowrap><font color='#FFFFFF'><b>BL No</b></font></td>"
    Response.Write "<td bgcolor='#FFFFFF' nowrap>" & anyTmp(0) & "</td>"
%>
                <td bgcolor="#003399" background="gif/tableback.gif" nowrap><font color="#FFFFFF"><b>�R���e�iNo</b></font></td>
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
                <td nowrap><b>�ʒu���@</b></td>
                <td><img src="gif/hr.gif"></td>
              </tr>
            </table>

            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center"> 
                <td nowrap align="center" bgcolor="#FFCC33">�ꏊ</td>
                <td nowrap bgcolor="#FFCC33">�d�o�`<font size="-1"><sup>(��1)</sup></font></td>
                <td nowrap bgcolor="#FFCC33">�O�`<font size="-1"><sup>(��1)</sup></font></td>
                <td colspan="4" nowrap bgcolor="#FFCC33">�^�[�~�i��</td>
                <td nowrap bgcolor="#FFCC33">�X�g�b�N���[�h</td>
                <td colspan="3" nowrap bgcolor="#FFCC33">����A��</td>
              </tr>
              <tr align="center"> 
                <td nowrap rowspan="2" bgcolor="#FFFFCC">�H��</td>
                <td nowrap rowspan="2" bgcolor="#FFFFCC">���݊���</td>
                <td nowrap rowspan="2" bgcolor="#FFFFCC">���݊���</td>
                <td nowrap colspan="2" bgcolor="#FFFFCC">����</td>
                <td nowrap colspan="2" bgcolor="#FFFFCC">���[�h</td>
                <td nowrap bgcolor="#FFFFCC">���o����</td>
                <td nowrap bgcolor="#FFFFCC">�q�ɓ���</td>
<% If anyTmp(54)<>"" And strUserKind="���^" Then %>
                <td nowrap rowspan="2" bgcolor="#FFFFCC">�f�o������<font size="-1"><sup>(��2)</sup></font></td>
<% Else %>
                <td nowrap rowspan="2" bgcolor="#FFFFCC">�f�o������</td>
<% End If %>
                <td nowrap rowspan="2" bgcolor="#FFFFCC">��R��<BR>�ԋp����</td>
              </tr>
              <tr align="center" bgcolor="#FFFFCC">
                <td nowrap>�v��</td>
                <td nowrap>�\��^����</td>
                <td nowrap>��������</td>
                <td nowrap>���o����</td>
                <td nowrap>�\��^����</td>
<% If anyTmp(54)<>"" And strUserKind="���^" Then %>
                <td nowrap>�w���^����<font size="-1"><sup>(��2)</sup></font></td>
<% Else %>
                <td nowrap>�w���^����</td>
<% End If %>
              </tr>
              <tr align="center"> 
                <td bgcolor="#FFFFCC" rowspan="2" nowrap>����</td>
                <td align="center" rowspan="2" nowrap>
<% ' �d�o�` - ���݊���
    Response.Write DispDateTimeCell(anyTmp(41),11)
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' �O�` - ���݊��� $�ǉ�
    Response.Write DispDateTimeCell(anyTmp(67),11)
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' �^�[�~�i�� �| ���݃X�P�W���[��
    If anyTmp(61)<>"" Then
        Response.Write "<font color='#0000FF'>"
    End If
    Response.Write DispDateTimeCell(anyTmp(61),5)
    If anyTmp(61)<>"" Then
        Response.Write "</font>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' �^�[�~�i�� - ���ݗ\��
    If anyTmp(32)<>"" Then
        bLate = false
        If anyTmp(33)<>"" Then
            If anyTmp(32)<anyTmp(33) Then
                bLate = true
            End If
        End If
        If anyTmp(61)<>"" Then
            If Left(anyTmp(61),10)<Left(anyTmp(32),10) Then
                bLate = true
            End If
        End If
        If bLate Then
            Response.Write "<font color='#FF0000'>"
        Else
            Response.Write "<font color='#0000FF'>"
        End If
        Response.Write DispDateTimeCell(anyTmp(32),11)
        Response.Write "</font>"
    Else
        Response.Write DispDateTimeCell(anyTmp(32),11)
    End If
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' �^�[�~�i�� - ���[�h����(�m�F)����
    Response.Write DispDateTimeCell(anyTmp(42),11)
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' �^�[�~�i�� - ���[�h���o����
    Response.Write DispDateTimeCell(anyTmp(43),11)
%>
                </td>
                <td align="center" nowrap>
<% ' �X�g�b�N���[�h - ���o�\�� $�ǉ�
    sTemp=DispReserveCell(anyTmp(65),anyTmp(66),sColor)
    Response.Write sColor
    Response.Write sTemp
    If sColor<>"" Then
        Response.Write "</font>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' ����^�� - �q�ɓ����X�P�W���[��
    If anyTmp(64)<>"" Then
        strTemp=anyTmp(64)
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
    Response.Write DispDateTimeCell(strTemp,11)
    If strTemp<>"" Then
        Response.Write "</font>"
    End If
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' ����A�� - �f�o������
    If anyTmp(54)<>"" And strUserKind="���^" Then
        Response.Write "<a href='ms-impinput.asp?kind=2&line=" & LineNo & "&request=ms-impdetail.asp'>"
    End If
    strTemp = DispDateTimeCell(anyTmp(45),11)
    If Left(strTemp,1)="<" And anyTmp(54)<>"" Then
        Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"
    Else
        Response.Write strTemp
    End If
    If anyTmp(54)<>"" And strUserKind="���^" Then
        Response.Write "</a>"
    End If
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' ����A�� - ��R���ԋp����
    Response.Write DispDateTimeCell(anyTmp(46),11)
%>
                </td>
              </tr>
              <tr>
                <td align="center" nowrap>
<% ' �^�[�~�i�� - ���݊���
    Response.Write DispDateTimeCell(anyTmp(33),11)
%>
                </td>
                <td align="center" nowrap>
<% ' �X�g�b�N���[�h - ���o����
    Response.Write DispDateTimeCell(anyTmp(60),11)
%>
                </td>
                <td align="center" nowrap>
<% ' ����A�� - �q�ɓ�������
    If anyTmp(54)<>"" And strUserKind="���^" Then
        Response.Write "<a href='ms-impinput.asp?kind=1&line=" & LineNo & "&request=ms-impdetail.asp'>"
    End If
    strTemp = DispDateTimeCell(anyTmp(44),11)
    If Left(strTemp,1)="<" And anyTmp(54)<>"" Then
        Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"
    Else
        Response.Write strTemp
    End If
    If anyTmp(54)<>"" And strUserKind="���^" Then
        Response.Write "</a>"
    End If
%>
                </td>
              </tr>
            </table>
        <table border="0" cellspacing="2" cellpadding="1">
          <tr> 
            <td width="15"><BR></td>
            <td><font color="#000000" size="-1">�i��1�j�d���`�A�O�`�̎����́A���n���Ԃł��B</font></td>
<% If anyTmp(54)<>"" And strUserKind="���^" Then %>
            <td width="15"><BR></td>
            <td><font color="#000000" size="-1">�i��2�j�N���b�N�Ŋ����������͉�ʂ�</font></td>
<% End If %>
          </tr>
        </table>
            <br>
<!-----�葱���---------------->
            <table>
              <tr>
                <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>�葱���y�у^�[�~�i�����o�ۏ��</b></td>
                <td><img src="gif/hr.gif"></td>
              </tr>
            </table>
            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center"> 
                <td rowspan="3" nowrap bgcolor="#FFCC33">����</td>
                <td colspan="4" nowrap bgcolor="#FFCC33">�s���葱��</td>
                <td rowspan="3" nowrap bgcolor="#FFCC33">�����<br>
                  DO���s</td>
                <td rowspan="3" nowrap bgcolor="#FFCC33">�t���[<br>
                  �^�C��</td>
                <td rowspan="3" nowrap bgcolor="#FFCC33">�^�[�~�i��<br>
                  ���o��</td>
              </tr>
              <tr> 
                <td align="center" nowrap bgcolor="#FFFFCC">�����m�F����</td>
                <td align="center" nowrap bgcolor="#FFFFCC" rowspan="2">���A��</td>
                <td align="center" nowrap bgcolor="#FFFFCC" rowspan="2">�ʔ���</td>
                <td align="center" nowrap bgcolor="#FFFFCC" rowspan="2">�ʊ�/<br>�ېŗA��</td>
              </tr>
              <tr> 
                <td align="center" nowrap bgcolor="#FFFFCC">�\��^����</td>
              </tr>
              <tr align="center"> 
                <td bgcolor="#FFFFCC" rowspan="2" nowrap>���</td>
                <td align="center" nowrap>
<% ' �����m�F�\�莞��
    If anyTmp(62)<>"" Then
        If anyTmp(48)<>"" Then
            If Left(anyTmp(62),10)<Left(anyTmp(48),10) Then
                Response.Write "<font color='#FF0000'>"
            Else
                Response.Write "<font color='#0000FF'>"
            End If
        Else
            Response.Write "<font color='#0000FF'>"
        End If
        Response.Write DispDateTimeCell(anyTmp(62),11)
        Response.Write "</font>"
    Else
        Response.Write DispDateTimeCell(anyTmp(62),11)
    End If
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' ���A��
    If anyTmp(47)="S" Then
        Response.Write "�~"
    ElseIf anyTmp(47)="C" Then
        Response.Write "��"
    Else
        Response.Write "�|"
    End If
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' �ʔ���
    If anyTmp(63)<>"" Then
        Response.Write "��"
    Else
        Response.Write "�|"
    End If
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' �ʊց^�ېŗA��
    If anyTmp(49)<>"" Then
        If anyTmp(49)="O" Or anyTmp(49)="T" Then
            Response.Write "<a href='#"
            Response.Write iLineNo
            Response.Write "' onClick=""winOpen('win1','ms-impdetail-h.asp?line="
            Response.Write iLineNo
            Response.Write "',150,150)"">��</a>"
        Else
            Response.Write "��"
        End If
    Else
        Response.Write "�~"
    End If
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' ������c�n���s
    If anyTmp(51)<>"Y" Then
        Response.Write "�~"
    Else
        Response.Write "��"
    End If
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' �t���[�^�C��
    If anyTmp(52)<>"" Then
        If anyTmp(52)<DispDateTime(Now,10) Then
            Response.Write "<font color='#FF0000'>"
        Else
            Response.Write "<font color='#000000'>"
        End If
        Response.Write DispDateTimeCell(anyTmp(52),5)
        Response.Write "</font>"
    Else
        Response.Write DispDateTimeCell(anyTmp(52),5)
    End If
%>
                </td>
                <td align="center" rowspan="2" nowrap>
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
              </tr>
              <tr>
                <td align="center" nowrap>
<% ' �����m�F��������
    Response.Write DispDateTimeCell(anyTmp(48),5)
%>
                </td>
              </tr>
            </table>
            <br>
<!---------------��{���--------------------------------------------->
            <table>
              <tr>
                <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>��{���</b></td>
                <td><img src="gif/hr.gif"></td>
              </tr>
            </table>
            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center" bgcolor="#FFCC33"> 
                <td valign="top" nowrap>����</td>
                <td nowrap bgcolor="#FFCC33">�T�C�Y</td>
<%
	Dim iSupNum
	If anyTmp(54)<>"" And strUserKind="���^" Then
		iSupNum = 3
	Else
		iSupNum = 2
	End If
%>
                <td nowrap bgcolor="#FFCC33">����<font size="-1"><sup>(��<%=iSupNum%>)</sup></font></td>
                <td nowrap bgcolor="#FFCC33">���[�t�@�[</td>
                <td nowrap bgcolor="#FFCC33">���d��(t)</td>
                <td valign="top" nowrap>�댯��<font size="-1"><sup>(��<%=iSupNum+1%>)</sup></font></td>
                <td nowrap bgcolor="#FFCC33">���o�^�[�~�i��</td>
                <td nowrap bgcolor="#FFCC33">�X�g�b�N���[�h���p</td>
                <td nowrap bgcolor="#FFCC33">�ԋp�ꏊ</td>
              </tr>
              <tr align="center"> 
                <td bgcolor="#FFFFCC" nowrap>���</td>
                <td align="center" nowrap>
<% ' �T�C�Y
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
                <td align="center" nowrap>
<% ' ����
    If anyTmp(54)<>"" Then
        Response.Write anyTmp(54)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' ���[�t�@�[
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
                <td align="center" nowrap>
<% ' ���d��
    If anyTmp(56)<>"" And anyTmp(56)<>"0" Then
        dWeight=anyTmp(56) / 10
        Response.Write dWeight
    Else
        Response.Write "�|"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' �댯��
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
                <td align="center" nowrap>
<% ' ���o�^�[�~�i��
    If anyTmp(35)<>"" Then
        Response.Write anyTmp(35)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' �X�g�b�N���[�h���p $�ǉ�
    If anyTmp(65)>="1" And anyTmp(65)<="4" Then
        Response.Write "��"
    Else
        Response.Write "�~"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' �ԋp�ꏊ
    If anyTmp(40)<>"" Then
        Response.Write anyTmp(40)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
              </tr>
            </table>
        <table border="0" cellspacing="2" cellpadding="1">
          <tr> 
            <td width="15"><BR></td>
            <td><font color="#000000" size="-1">�i��<%=iSupNum%>) 96=HC</font></td>
            <td width="15"><BR></td>
            <td><font color="#000000" size="-1">�i��<%=iSupNum+1%>�j���h�@�Ɋւ��댯���̗L��</font></td>
          </tr>
        </table>
            <br>
<!---------------�{�D���--------------------------------------------->
            <table>
              <tr> 
                <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>�{�D���&nbsp;&nbsp;</b></td>
                <td><img src="gif/hr.gif"></td>
              </tr>
            </table>
            <table border=1 cellpadding="3" cellspacing="1">
              <tr> 
                <td bgcolor="#FFCC33" nowrap><font color="#000000">�D��</font></td>
                <td bgcolor="#FFFFFF" nowrap>
<% ' �D��
    If anyTmp(36)<>"" Then
        Response.Write anyTmp(36)
    ElseIf anyTmp(21)<>"" Then
        Response.Write anyTmp(21)
    ElseIf anyTmp(15)<>"" Then
        Response.Write anyTmp(15)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td bgcolor="#FFCC33" nowrap><font color="#000000">�D��</font></td>
                <td bgcolor="#FFFFFF" nowrap>
<% ' �D��
    If anyTmp(37)<>"" Then
        Response.Write anyTmp(37)
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
    If anyTmp(38)<>"" Then
        Response.Write anyTmp(38)
    ElseIf anyTmp(3)<>"" Then
        Response.Write anyTmp(3)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td bgcolor="#FFCC33" nowrap>�d�o�`</td>
                <td bgcolor="#FFFFFF" nowrap>
<% ' �d�o�`
    If anyTmp(39)<>"" Then
        Response.Write anyTmp(39)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td bgcolor="#FFCC33" nowrap>�O�`</td>
                <td bgcolor="#FFFFFF" nowrap>
<% ' �O�`
    If anyTmp(68)<>"" Then
        Response.Write anyTmp(68)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
              </tr>
            </table>
<form>
      <input type=button value='�\���f�[�^�̍X�V' OnClick="JavaScript:window.location.href='ms-impreload.asp?request=ms-impdetail.asp'">
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
        DispMenuBarBack "ms-implist1.asp"
    ElseIf iReturn=2 Then
        DispMenuBarBack "ms-implist2.asp"
    End If
%>
</body>
</html>
