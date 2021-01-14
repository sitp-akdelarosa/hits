<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<!--#include file="pickcom.inc"-->

<!--#include file="vessel.inc"-->

<%

    ' �G���[�t���O�̃N���A
    bError = false

    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    '���͉�ʂ��L��
    Session.Contents("findcsv")="no"    ' ���ړ��͂ł��邱�Ƃ��L��

    Session.Contents("sortkey")="�w���"

    ' �w������̎擾
    Dim strShipper       '�׎�R�[�h
    Dim strTrucker       '���^�R�[�h
    Dim strForwader      '�C�݃R�[�h
    Dim strVslCode       '�D���R�[�h
    Dim strOpeCode       '�`�^�R�[�h
    Dim strVoyCtrl       'Voyage No.
    Dim strPickDate      '��R�����o�w���
    strShipper = UCase(Trim(Request.QueryString("ninushi")))
    strTrucker = UCase(Trim(Request.QueryString("rikuun")))
    strForwader = UCase(Trim(Request.QueryString("kaika")))
    strVslCode = UCase(Trim(Request.QueryString("vessel")))
    strVoyCtrl = UCase(Trim(Request.QueryString("voyage")))
    strPickDate = Trim(Request.QueryString("decyear")) & "/" & Trim(Request.QueryString("decmon")) & "/" &_
				  Trim(Request.QueryString("decday"))

	Dim iNum,strOption
    strOption = ""
   ' ���O�C����ʂ̎擾�Ƃ��̏���
    strUserKind=Session.Contents("userkind")
    If strUserKind="�C��" Then
		iNum = "a101"
        strForwader=Session.Contents("userid")
        strOption = strVslCode & "/" & strVoyCtrl & "/" & strShipper & "/" & strTrucker
    ElseIf strUserKind="���^" Then
		iNum = "a102"
        strTrucker=Session.Contents("userid")
        strOption = strForwader
    ElseIf strUserKind="�׎�" Then
		iNum = "a103"
        strShipper=Session.Contents("userid")
        strOption = strVslCode & "/" & strVoyCtrl & "/" & strForwader
    ElseIf strUserKind="�`�^" Then
		iNum = "a104"
        strOpeCode=Session.Contents("userid")
        strOption = strVslCode & "/" & strVoyCtrl & "/" & strForwader & "/" & strPickDate
    End If

    ' �Q��Key���L��
    Session.Contents("findkey1")=strShipper       '�׎�R�[�h
    Session.Contents("findkey2")=strForwader      '�C�݃R�[�h
    Session.Contents("findkey3")=strTrucker       '���^�R�[�h
    Session.Contents("findkey4")=strVslCode       '�D���R�[�h
    Session.Contents("findkey5")=strVoyCtrl       'Voyage No.
    Session.Contents("findkey6")=strPickDate      '��R�����o�w���
    Session.Contents("findkey7")=strOpeCode       '�`�^�R�[�h

    ' �e���|�����t�@�C�������쐬���āA�Z�b�V�����ϐ��ɐݒ�
    Dim strFileName
    strFileName = GetNumStr(Session.SessionID, 8) & ".csv"
    Session.Contents("tempfile")=strFileName

    ' �R���e�i��񃌃R�[�h�̎擾
    ConnectSvr conn, rsd

    ' ���������̍쐬
    sWhere = ""

    '�׎�R�[�h
    If strShipper<>"" Then
        If sWhere<>"" Then
            sWhere = sWhere & " And ExportCargoInfo.Shipper='" & strShipper & "'"
        Else
            sWhere = "ExportCargoInfo.Shipper='" & strShipper & "'"
        End If
    End If
    '�C�݃R�[�h
    If strForwader<>"" Then
        If sWhere<>"" Then
            sWhere = sWhere & " And ExportCargoInfo.Forwarder='" & strForwader & "'"
        Else
            sWhere = "ExportCargoInfo.Forwarder='" & strForwader & "'"
        End If
    End If
    '���^�R�[�h
    If strTrucker<>"" Then
        If sWhere<>"" Then
            sWhere = sWhere & " And ExportCargoInfo.Trucker='" & strTrucker & "'"
        Else
            sWhere = "ExportCargoInfo.Trucker='" & strTrucker & "'"
        End If
    End If
    '�D���R�[�h
    If strVslCode<>"" Then
        If sWhere<>"" Then
            sWhere = sWhere & " And ExportCargoInfo.VslCode='" & strVslCode & "'"
        Else
            sWhere = sWhere & "ExportCargoInfo.VslCode='" & strVslCode & "'"
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
            sWhere = sWhere & " And ExportCargoInfo.LdVoyage='" & strVoyCtrl & "'"
        Else
            sWhere = sWhere & "ExportCargoInfo.LdVoyage='" & strVoyCtrl & "'"
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

    sSort="ExportCargoInfo.PickDate"

    ' �擾�����R���e�i��񃌃R�[�h���e���|�����t�@�C���ɏ����o��
    strFileName="./temp/" & strFileName
    ' �e���|�����t�@�C����Open
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),2,True)

    bWriteFile = SerchMSExpCntnr(conn, rsd, ti, sWhere, sSort)

    ti.Close
    conn.Close

    ' Temp�t�@�C�������ݒ�
    SetTempFile "MSEXPORT"

    If bWriteFile = 0 Then
        ' �Y�����R�[�h�Ȃ��Ƃ�
        bError = true
        strError = "�w������ɊY������R���e�i�͂���܂���ł����B"
        strOption = strOption & "," & "���͓��e�̐���:1(���)"
    Else
        strOption = strOption & "," & "���͓��e�̐���:0(������)"

        ' DT02�g�����U�N�V�����𔭍s����
        If strUserKind="���^" Then
            ' �g�����U�N�V�����t�@�C���̊g���q 
            Const SEND_EXTENT = "snd"
            sTranID = "DT02"
            ' �����敪
            Const sSyori = "R"
            ' ���M�ꏊ
            Const sPlace = ""

            ' �e���|�����t�@�C�������쐬���āA�Z�b�V�����ϐ��ɐݒ�
            Dim sDT02, iSeqNo, strFileName_01, sTusin
            '�V�[�P���X�ԍ�
            iSeqNo = GetDailyTransNo
            '�ʐM�����擾
            sTusin  = SetTusinDate

            sDT02 = iSeqNo & "," & sTranID & "," & sSyori & ","  & sTusin & ",Web - " & _
                    strTrucker & "," & sPlace & ",X," & strTrucker & "," & strForwader
            sFileName = ArrangeNumV(Month(Now), 2) & ArrangeNumV(Day(Now), 2) & iSeqNo
            strFileName_01 = "./send/" & sFileName & "." & SEND_EXTENT
            Set ti=fs.OpenTextFile(Server.MapPath(strFileName_01),2,True)
            ti.WriteLine sDT02
            ti.Close
        End If
    End If

    ' �A�o�R���e�i�Ɖ�
    WriteLog fs, iNum,"��R���s�b�N�A�b�v�V�X�e��-" & strUserKind & "�p�Ɖ�","10", strOption

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
<!-------------��������Ɖ�G���[���--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
<td rowspan=2><%
    If strUserKind="�C��" Then
        Response.Write "<img src='gif/pickkat.gif' width='506' height='73'>"
    ElseIf strUserKind="���^" Then
        Response.Write "<img src='gif/pickrit.gif' width='506' height='73'>"
    ElseIf strUserKind="�׎�" Then
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
		<BR><BR><BR>

      <table>
        <tr>
          <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
          <td nowrap><b>��R���s�b�N�A�b�v���Ɖ�
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
          <td>
<%
    ' �G���[���b�Z�[�W�̕\��
    DispErrorMessage strError
%>
          </td></tr>
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
<!-------------�Ɖ�G���[��ʏI���--------------------------->
<%
    DispMenuBarBack "JavaScript:window.history.back()"
%>
</body>
</html>

<%
    Else
        ' �ꗗ��ʂփ��_�C���N�g
        If strUserKind="�C��" Then
            Response.Redirect "picklist.asp?kind=1"
        ElseIf strUserKind="���^" Then
            Response.Redirect "picklist.asp?kind=2"
        ElseIf strUserKind="�׎�" Then
            Response.Redirect "picklist.asp?kind=3"
        ElseIf strUserKind="�`�^" Then
            Response.Redirect "picklist.asp?kind=4"
        End If
    End If
%>
