<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<!--#include file="ImpCom.inc"-->

<%
''    ' �Z�b�V�����̃`�F�b�N
''    CheckLogin "impentry.asp"

    ' �G���[�t���O�̃N���A
    bError = false

    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    '���͉�ʂ��L��
    Session.Contents("findcsv")="no"    ' ���ړ��͂ł��邱�Ƃ��L��

    ' �w������̎擾
    Dim strCntnrNo,strCntnrNoLog
    Dim strBLNo,strBLNoLog
    strCntnrNo = UCase(Trim(Request.QueryString("cntnrno")))
    strBLNo = UCase(Trim(Request.QueryString("blno")))
	strCntnrNoLog = strCntnrNo
	strBLNoLog = strBLNo
    If strCntnrNo="" And strBLNo="" Then
        ' �����w��̂Ȃ��Ƃ� �G���[���b�Z�[�W��\��
        bError = true
        strError = "�Q�Ƃ������R���e�iNo.���́AB/L No�̂����A<br>�ꍀ�ڂ͓��͂��Ă��������B"
        strOption = "," & "," & "���͓��e�̐���:1(���)"
        ' �����w��̂Ȃ��Ƃ� �T���v����ʂ�\��
        Response.Redirect "implist.html"
        Responce.End
    Else
        ' �e���|�����t�@�C�������쐬���āA�Z�b�V�����ϐ��ɐݒ�
        Dim strFileName
        strFileName = GetNumStr(Session.SessionID, 8) & ".csv"
        Session.Contents("tempfile")=strFileName

        ' �R���e�i��񃌃R�[�h�̎擾
        ConnectSvr conn, rsd
        sWhere = ""
        If strBLNo<>"" Then        ' BL�ԍ��̓��͂��D��
            strInput = "," & "���͓��e," & strBookingNo
            strOption = "���͕��@����,3(BL�ԍ�1��)" & strInput

            Session.Contents("findkey")=strBLNo         ' �Q��Key���L��
            iCanma = InStr(strBLNo,",")
            Do While iCanma>0
                strOption = "���͕��@����,4(BL�ԍ�����)" & strInput
                sTemp = Trim(Left(strBLNo,iCanma-1))
                strBLNo = Right(strBLNo,Len(strBLNo)-iCanma)
                If sWhere<>"" Then
                    sWhere = sWhere & " Or ImportCont.BLNo='" & sTemp & "'"
                Else
                    sWhere = "ImportCont.BLNo='" & sTemp & "'"
                End If
                iCanma = InStr(strBLNo,",")
            Loop
            If sWhere<>"" Then
                sWhere = sWhere & " Or ImportCont.BLNo='" & Trim(strBLNo) & "'"
            Else
                sWhere = "ImportCont.BLNo='" & Trim(strBLNo) & "'"
            End If
            Session.Contents("findkind")="Blno"       ' �Q�ƃ��[�h
        Else
            strInput = "," & "���͓��e," & strCntnrNo
            strOption = "���͕��@����,0(�R���e�iNo.1��)" & strInput

            Session.Contents("findkey")=strCntnrNo       ' �Q��Key���L��
            iCanma = InStr(strCntnrNo,",")
            Do While iCanma>0
                strOption = "���͕��@����,1(�R���e�iNo.����)" & strInput
                sTemp = Trim(Left(strCntnrNo,iCanma-1))
                strCntnrNo = Right(strCntnrNo,Len(strCntnrNo)-iCanma)
                If sWhere<>"" Then
                    sWhere = sWhere & " Or ImportCont.ContNo='" & sTemp & "'"
                Else
                    sWhere = "ImportCont.ContNo='" & sTemp & "'"
                End If
                iCanma = InStr(strCntnrNo,",")
            Loop
            If sWhere<>"" Then
                sWhere = sWhere & " Or ImportCont.ContNo='" & Trim(strCntnrNo) & "'"
            Else
                sWhere = "ImportCont.ContNo='" & Trim(strCntnrNo) & "'"
            End If
            Session.Contents("findkind")="Cntnr"         ' �Q�ƃ��[�h
        End If

        ' �擾�����R���e�i��񃌃R�[�h���e���|�����t�@�C���ɏ����o��
        strFileName="./temp/" & strFileName
        ' �e���|�����t�@�C����Open
        Set ti=fs.OpenTextFile(Server.MapPath(strFileName),2,True)

        bWriteFile = SerchImpCntnr(conn, rsd, ti, sWhere)

        ti.Close
        conn.Close

        ' Temp�t�@�C�������ݒ�
        SetTempFile "IMPORT"

        If bWriteFile = 0 Then
            ' �Y�����R�[�h�Ȃ��Ƃ�
            bError = true
            strError = "No container that corresponded to a specified condition."
            strOption = "���͓��e�̐���:1(���)"
        Else
            strOption = "���͓��e�̐���:0(������)"
        End If

    End If

	Dim iWrkNum
	If strBLNoLog="" Then
		iWrkNum = 11
		Do While InStr(strCntnrNoLog,",")>0
			strCntnrNoLog = Left(strCntnrNoLog,InStr(strCntnrNoLog,",")-1) & _
							"/" & Right(strCntnrNoLog,Len(strCntnrNoLog)-InStr(strCntnrNoLog,","))
		Loop
		strOption = strCntnrNoLog & "," & strOption
	Else
		iWrkNum = 12
		Do While InStr(strBLNoLog,",")>0
			strBLNoLog = Left(strBLNoLog,InStr(strBLNoLog,",")-1) & _
							"/" & Right(strBLNoLog,Len(strBLNoLog)-InStr(strBLNoLog,","))
		Loop
		strOption = strBLNoLog & "," & strOption
	End If

    ' �A���R���e�i�Ɖ�
    WriteLog fs, "2301","�A���R���e�i�Ɖ�",iWrkNum, strOption

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
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="../gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------��������Ɖ�G���[���--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="../gif/shokait.gif" width="506" height="73"></td>
          <td height="25" bgcolor="000099" align="right"><img src="../gif/logo_hits_ver2.gif" width="300" height="25"></td>
        </tr>
        <tr>
          <td align="right" width="100%" height="48"> 
<%
' Added and Commented by seiko-denki 2003.07.18
'	DisplayCodeListButton
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
          <td><img src="../gif/botan.gif" width="17" height="17" vspace="4"></td>
          <td nowrap><b>Container Information (Imp.)</b></td>
          <td><img src="../gif/hr.gif"></td>
        </tr>
      </table>
		<BR>
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
    DispMenuBarBack "impentry.asp"
%>
</body>
</html>

<%
    Else
        If bWriteFile = 1 Then
            '�߂��ʎ�ʂ��L��
            Session.Contents("dispreturn")=0
            ' �ڍ׉�ʂփ��_�C���N�g
            Response.Redirect "impdetail.asp?line=1"    '�A���R���e�i�ڍ�
        Else
            '�߂��ʎ�ʂ��L��
            Session.Contents("dispreturn")=0
            ' �ꗗ��ʂփ��_�C���N�g
            Response.Redirect "implist.asp"             '�A���R���e�i�ꗗ
        End If
    End If
%>
