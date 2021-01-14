<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<!--#include file="ms-ImpCom.inc"-->

<!--#include file="vessel.inc"-->

<%
    ' �Z�b�V�����̃`�F�b�N
    CheckLogin "impentry.asp"

    ' �G���[�t���O�̃N���A
    bError = false

    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    '���͉�ʂ��L��
    Session.Contents("findcsv")="no"    ' ���ړ��͂ł��邱�Ƃ��L��

    ' �w������̎擾
    Dim strShipper       '�׎�R�[�h
    Dim strTrucker       '���^�R�[�h
    Dim strForwader      '�C�݃R�[�h
    Dim strVslCode       '�D���R�[�h
    Dim strVoyCtrl       'Voyage No.
    strShipper = UCase(Trim(Request.QueryString("ninushi")))
    strTrucker = UCase(Trim(Request.QueryString("rikuun")))
    strForwader = UCase(Trim(Request.QueryString("kaika")))
    strVslCode = UCase(Trim(Request.QueryString("vessel")))
    strVoyCtrl = UCase(Trim(Request.QueryString("voyage")))

	Dim iNum,strOption
    strOption = ""
    ' ���O�C����ʂ̎擾�Ƃ��̏���
    strUserKind=Session.Contents("userkind")
    If strUserKind="�C��" Then
		iNum = "2101"
        strForwader=Session.Contents("userid")
        Session.Contents("sortkey")="�D��"           ' �\�[�g�L�[���w��
        strOption = strVslCode & "/" & strVoyCtrl & "/" & strShipper & "/" & strTrucker
    ElseIf strUserKind="���^" Then
		iNum = "2102"
        strTrucker=Session.Contents("userid")
        Session.Contents("sortkey")="�C��"           ' �\�[�g�L�[���w��
        strOption = strForwader
    ElseIf strUserKind="�׎�" Then
		iNum = "2103"
        strShipper=Session.Contents("userid")
        Session.Contents("sortkey")="�D��"           ' �\�[�g�L�[���w��
        strOption = strVslCode & "/" & strVoyCtrl & "/" & strForwader
    End If

    ' �Q��Key���L��
    Session.Contents("findkey1")=strShipper       '�׎�R�[�h
    Session.Contents("findkey2")=strForwader      '�C�݃R�[�h
    Session.Contents("findkey3")=strTrucker       '���^�R�[�h
    Session.Contents("findkey4")=strVslCode       '�D���R�[�h
    Session.Contents("findkey5")=strVoyCtrl       'Voyage No.

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
            sWhere = sWhere & " And ImportCargoInfo.Shipper='" & strShipper & "'"
        Else
            sWhere = "ImportCargoInfo.Shipper='" & strShipper & "'"
        End If
    End If
    '�C�݃R�[�h
    If strForwader<>"" Then
        If sWhere<>"" Then
            sWhere = sWhere & " And ImportCargoInfo.Forwarder='" & strForwader & "'"
        Else
            sWhere = "ImportCargoInfo.Forwarder='" & strForwader & "'"
        End If
    End If
    '���^�R�[�h
    If strTrucker<>"" Then
        If sWhere<>"" Then
            sWhere = sWhere & " And ImportCargoInfo.Trucker='" & strTrucker & "'"
        Else
            sWhere = "ImportCargoInfo.Trucker='" & strTrucker & "'"
        End If
    End If
    '�D���R�[�h
    If strVslCode<>"" Then
        If sWhere<>"" Then
            sWhere = sWhere & " And ImportCargoInfo.VslCode='" & strVslCode & "'"
        Else
            sWhere = sWhere & "ImportCargoInfo.VslCode='" & strVslCode & "'"
        End If
    End If
    'Voyage No.
    If strVoyCtrl<>"" Then
        If sWhere<>"" Then
            sWhere = sWhere & " And ImportCargoInfo.DsVoyage='" & strVoyCtrl & "'"
        Else
            sWhere = sWhere & "ImportCargoInfo.DsVoyage='" & strVoyCtrl & "'"
        End If
    End If

    ' Sort�����̍쐬
    strSortKey=Session.Contents("sortkey")
    If strSortKey="�׎�" Then
        sSort="ImportCargoInfo.Shipper"
    ElseIf strSortKey="�C��" Then
        sSort="ImportCargoInfo.Forwarder"
    ElseIf strSortKey="�D��" Then
        sSort="ImportCargoInfo.VslCode, ImportCargoInfo.DsVoyage"
    ElseIf strSortKey="�q�ɓ���" Then
        sSort="ImportCargoInfo.WHArTime"
    ElseIf strSortKey="���^�Ǝ�" Then
        sSort="ImportCargoInfo.Trucker"
    End If

    ' �擾�����R���e�i��񃌃R�[�h���e���|�����t�@�C���ɏ����o��
    strFileName="./temp/" & strFileName
    ' �e���|�����t�@�C����Open
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),2,True)

    bWriteFile = SerchMSImpCntnr(conn, rsd, ti, sWhere, sSort)

    ti.Close
    conn.Close

    ' Temp�t�@�C�������ݒ�
    SetTempFile "MSIMPORT"

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
                    strTrucker & "," & sPlace & ",I," & strTrucker & "," & strForwader
            sFileName = ArrangeNumV(Month(Now), 2) & ArrangeNumV(Day(Now), 2) & iSeqNo
            strFileName_01 = "./send/" & sFileName & "." & SEND_EXTENT
            Set ti=fs.OpenTextFile(Server.MapPath(strFileName_01),2,True)
            ti.WriteLine sDT02
            ti.Close
        End If
    End If

    ' �A���R���e�i�Ɖ�
    WriteLog fs, iNum,"�A���R���e�i�Ɖ�-" & strUserKind & "�p�Ɖ�","10", strOption

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
		<BR>
		<BR>
		<BR>
      <table>
        <tr>
          <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
          <td nowrap><b>�A���R���e�i���Ɖ�
<%
    If strUserKind="�C��" Then
        Response.Write "(�C�ݗp)"
    ElseIf strUserKind="���^" Then
        Response.Write "(���^�p)"
    Else
        Response.Write "(�׎�p)"
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
            Response.Redirect "ms-implist1.asp"          '�A���R���e�i�ꗗ
        ElseIf strUserKind="���^" Then
            Response.Redirect "ms-implist2.asp"          '�A���R���e�i�ꗗ
        ElseIf strUserKind="�׎�" Then
            Response.Redirect "ms-implist1.asp"          '�A���R���e�i�ꗗ
        End If
    End If
%>
