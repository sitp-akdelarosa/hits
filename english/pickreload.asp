<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<!--#include file="pickcom.inc"-->

<%
    ' ���O�C����ʂ̎擾�Ƃ��̏���
    strUserKind=Session.Contents("userkind")
    If strUserKind="" Then
        ' �Z�b�V�������؂�Ă���Ƃ�
        Response.Redirect "http://www.hits-h.com/index.asp"         '���j���[
        Response.End
    End If

	sSortKey = Request.QueryString("sort")

    ' �Z�b�V�����̃`�F�b�N
	Dim sLoginKind
    If strUserKind="�C��" Then
        CheckLogin "picklist.asp?kind=1"
		sLoginKind = "1"
    ElseIf strUserKind="���^" Then
        CheckLogin "picklist.asp?kind=2"
		sLoginKind = "2"
    ElseIf strUserKind="�׎�" Then
        CheckLogin "picklist.asp?kind=3"
		sLoginKind = "3"
    Else
        CheckLogin "picklist.asp?kind=4"
		sLoginKind = "4"
    End If

    ' Temp�t�@�C�������̃`�F�b�N
    CheckTempFile "MSEXPORT", "expentry.asp"

    ' �L�����Ă��錟�����������[�h
    strShipper=Session.Contents("findkey1")       '�׎�R�[�h
    strForwader=Session.Contents("findkey2")      '�C�݃R�[�h
    strTrucker=Session.Contents("findkey3")       '���^�R�[�h
    strVslCode=Session.Contents("findkey4")       '�D���R�[�h
    strVoyCtrl=Session.Contents("findkey5")       'Voyage No.
    strPickDate=Session.Contents("findkey6")      '��R�����o��
    strOpeCode=Session.Contents("findkey7")       '�`�^�R�[�h

    ' �G���[�t���O�̃N���A
    bError = false

    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemObject")

    ' �\���t�@�C���̎擾
    Dim strFileName
    strFileName = Session.Contents("tempfile")
    If strFileName="" Then
        ' �Z�b�V�������؂�Ă���Ƃ�
        Response.Redirect "http://www.hits-h.com/index.asp"         '���j���[
        Response.End
    End If

    ' �f�[�^�x�[�X�̐ڑ�
    ConnectSvr conn, rsd

    ' ���������̍쐬
    sWhere = ""

    '�׎�R�[�h
    If strShipper<>"" Then
        If sWhere<>"" Then
            sWhere = sWhere & " And ExportCargoInfo.Shipper='" & strShipper & "'"
            strOption = strOption & ",�׎�R�[�h," & strShipper
        Else
            sWhere = "ExportCargoInfo.Shipper='" & strShipper & "'"
            strOption = "�׎�R�[�h," & strShipper
        End If
    End If
    '�C�݃R�[�h
    If strForwader<>"" Then
        If sWhere<>"" Then
            sWhere = sWhere & " And ExportCargoInfo.Forwarder='" & strForwader & "'"
            strOption = strOption & ",�C�݃R�[�h," & strForwader
        Else
            sWhere = "ExportCargoInfo.Forwarder='" & strForwader & "'"
            strOption = "�C�݃R�[�h," & strForwader
        End If
    End If
    '���^�R�[�h
    If strTrucker<>"" Then
        If sWhere<>"" Then
            sWhere = sWhere & " And ExportCargoInfo.Trucker='" & strTrucker & "'"
            strOption = strOption & ",���^�R�[�h," & strTrucker
        Else
            sWhere = "ExportCargoInfo.Trucker='" & strTrucker & "'"
            strOption = "���^�R�[�h," & strTrucker
        End If
    End If
    '�D���R�[�h
    If strVslCode<>"" Then
        If sWhere<>"" Then
            sWhere = sWhere & " And ExportCargoInfo.VslCode='" & strVslCode & "'"
            strOption = strOption & ",�D���R�[�h," & strVslCode
        Else
            sWhere = sWhere & "ExportCargoInfo.VslCode='" & strVslCode & "'"
            strOption = "�D���R�[�h," & strVslCode
        End If
    End If
    '�`�^�R�[�h
    If strOpeCode<>"" Then
        If sWhere<>"" Then
            sWhere = sWhere & " And ExportCargoInfo.OpeCode='" & strOpeCode & "'"
        Else
            sWhere = sWhere & "ExportCargoInfo.OpeCode='" & strOpeCode & "'"
        End If
    End If
    'Voyage No.
    If strVoyCtrl<>"" Then
        If sWhere<>"" Then
            sWhere = sWhere & " And ExportCargoInfo.DsVoyage='" & strVoyCtrl & "'"
            strOption = strOption & ",Voyage No.," & strVoyCtrl
        Else
            sWhere = sWhere & "ExportCargoInfo.DsVoyage='" & strVoyCtrl & "'"
            strOption = "Voyage No.," & strVoyCtrl
        End If
    End If
   '��R�����o�w���
    If strPickDate<>"//" Then
        If sWhere<>"" Then
            sWhere = sWhere & " And ExportCargoInfo.PickDate<='" & strPickDate & " 23:59:59'" &_
							  " And ExportCargoInfo.PickDate>='" & strPickDate & " 00:00:00'"
        Else
            sWhere = sWhere & "ExportCargoInfo.PickDate<='" & strPickDate & " 23:59:59'" &_
						 " And ExportCargoInfo.PickDate>='" & strPickDate & " 00:00:00'"
        End If
    End If

    ' Sort�����̍쐬
	If sSortKey="�C��" Then
		sSort="ExportCargoInfo.Forwarder,ExportCargoInfo.PickDate"
		Session.Contents("sortkey")="�C��"
	ElseIf sSortKey="�׎�" Then
		sSort="ExportCargoInfo.Shipper,ExportCargoInfo.PickDate"
		Session.Contents("sortkey")="�׎�"
	ElseIf sSortKey="���^" Then
		sSort="ExportCargoInfo.Trucker,ExportCargoInfo.PickDate"
		Session.Contents("sortkey")="���^"
	ElseIf sSortKey="�`�^" Then
		sSort="ExportCargoInfo.OpeCode,ExportCargoInfo.PickDate"
		Session.Contents("sortkey")="�`�^"
	Else
		sSort="ExportCargoInfo.PickDate"
		Session.Contents("sortkey")="�w���"
	End If

    ' �擾�����R���e�i��񃌃R�[�h���e���|�����t�@�C���ɏ����o��
    strFileName="./temp/" & strFileName
    ' �]���t�@�C����Open
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),2,True)

    bWriteFile = SerchMSExpCntnr(conn, rsd, ti, sWhere, sSort)

    ' �t�@�C����DB�̃N���[�Y
    ti.Close
    conn.Close

    If bWriteFile = 0 Then
        ' �Y�����R�[�h�Ȃ��Ƃ�
        bError = true
        strError = "�w������ɊY������R���e�i�͂Ȃ��Ȃ�܂����B"
    End If


    If bError Then
%>

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
<%
    DispMenuJava
%>
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------��������G���[���--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2>
<%
    If strUserKind="�C��" Then
        Response.Write "<img src='gif/pickkat.gif' width='506' height='73'>"
    ElseIf strUserKind="���^" Then
        Response.Write "<img src='gif/pickrit.gif' width='506' height='73'>"
    ElseIf strUserKind="�׎�" Then
        Response.Write "<img src='gif/picknit.gif' width='506' height='73'>"
    Else
        Response.Write "<img src='gif/pickkot.gif' width='506' height='73'>"
    End If
%>
          </td>
          <td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
        </tr>
        <tr>
          <td align="right" width="100%" height="48"> 
<%
' Added and Commented by seiko-denki 2003.07.18
	DisplayCodeListButton
'    DispMenu
'	Dim strRoute
'	strRoute = Session.Contents("route")
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
		<BR>
		<BR>
		<BR>
      <table>
        <tr>
          <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
          <td nowrap><b>��R���s�b�N�A�b�v���ꗗ
<%
    If strUserKind="�C��" Then
        Response.Write "(�C�ݗp)"
    ElseIf strUserKind="���^" Then
        Response.Write "(���^�p)"
    ElseIf strUserKind="�׎�" Then
        Response.Write "(�׎�p)"
    Else
        Response.Write "(�`�^�p)"
    End If
%>
            </b></td>
          <td><img src="gif/hr.gif"></td>
        </tr>
      </table>

      <table>
        <tr> 
          <td nowrap>
<%
    ' �G���[���b�Z�[�W�̕\��
    DispErrorMessage strError
%>
          </td>
        </tr>
      </table>
      </center>
      <br>
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
<!-------------�G���[��ʏI���--------------------------->
<%
    DispMenuBarBack "pickselect.asp"
%>
</body>
</html>

<%
    Else
         ' �ꗗ��ʂփ��_�C���N�g
        Response.Redirect "picklist.asp?kind=" & sLoginKind
    End If
%>
