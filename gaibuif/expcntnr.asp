<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<!--#include file="ExpCom.inc"-->

<%
''    ' �Z�b�V�����̃`�F�b�N
''    CheckLogin "expentry.asp"
    ' �G���[�t���O�̃N���A
    bError = false

    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    '���͉�ʂ��L��
    Session.Contents("findcsv")="no"    ' ���ړ��͂ł��邱�Ƃ��L��

    ' �w������̎擾
    Dim strCntnrNo,strCntnrNoLog
    Dim strBookingNo,strBookingNoLog
' 2009/05/09 add-s �`�����ǉ�
    Dim strUserCodeEx,strUserCodeExNoLog
' 2009/05/09 add-s �`�����ǉ�
    strCntnrNo = UCase(Trim(Request.QueryString("cntnrno")))
    strBookingNo = UCase(Trim(Request.QueryString("booking")))
' 2009/05/09 add-s �`�����ǉ�
    strUserCodeEx = UCase(Trim(Request.QueryString("portcode")))
	strUserCodeExNoLog = strUserCodeEx
' 2009/05/09 add-s �`�����ǉ�
	strCntnrNoLog = strCntnrNo
	strBookingNoLog = strBookingNo
    If strCntnrNo="" And strBookingNo="" Then
        ' �����w��̂Ȃ��Ƃ� �G���[���b�Z�[�W��\��
        bError = true
        strError = "�Q�Ƃ������R���e�iNo.���́ABooking No�̂����A<br>�ꍀ�ڂ͓��͂��Ă��������B"
        strOption = "," & "," & ",���͓��e�̐���:1(���)"
        ' �����w��̂Ȃ��Ƃ� �T���v����ʂ�\��
        Response.Redirect "explist.html"
        Responce.End
    Else
        ' �e���|�����t�@�C�������쐬���āA�Z�b�V�����ϐ��ɐݒ�
        Dim strFileName
        strFileName = GetNumStr(Session.SessionID, 8) & ".csv"
        Session.Contents("tempfile")=strFileName

' 2009/05/09 add-s �`�����ǉ�
        Session.Contents("usercodeex")=strUserCodeEx     ' �Q�ƃ��[�U�[�R�[�h���L��
' 2009/05/09 add-e �`�����ǉ�

        ' �R���e�i��񃌃R�[�h�̎擾
        ConnectSvr conn, rsd
        sWhere = ""
'������ Add_S  by nics 200902����
        dim bWriteFile
        bWriteFile = 0
        ' �擾�����R���e�i��񃌃R�[�h���e���|�����t�@�C���ɏ����o��
        strFileName="./temp/" & strFileName
        ' �e���|�����t�@�C����Open
        Set ti=fs.OpenTextFile(Server.MapPath(strFileName),2,True)
'������ Add_E  by nics 200902����
        If strBookingNo<>"" Then        ' Booking�ԍ��̓��͂��D��
            strInput = "," & "���͓��e," & strBookingNo
            strOption = "���͕��@����,3(Booking�ԍ�1��)" & strInput

            Session.Contents("findkey")=strBookingNo     ' �Q��Key���L��
'������ Mod_S  by nics 200902����
'            iCanma = InStr(strBookingNo,",")
'            Do While iCanma>0
'                strOption = "���͕��@����,4(Booking�ԍ�����)" & strInput
'                sTemp = Trim(Left(strBookingNo,iCanma-1))
'                strBookingNo = Right(strBookingNo,Len(strBookingNo)-iCanma)
'                If sWhere<>"" Then
'                    sWhere = sWhere & " Or ExportCont.BookNo='" & sTemp & "'"
'                Else
'                    sWhere = "ExportCont.BookNo='" & sTemp & "'"
'                End If
'                iCanma = InStr(strBookingNo,",")
'            Loop
'            If sWhere<>"" Then
'                sWhere = sWhere & " Or ExportCont.BookNo='" & Trim(strBookingNo) & "'"
'            Else
'                sWhere = "ExportCont.BookNo='" & Trim(strBookingNo) & "'"
'            End If
'������
            Do While strBookingNo <> ""
                iCanma = InStr(strBookingNo, ",")
                If iCanma > 0 Then
                    strOption = "���͕��@����,4(Booking�ԍ�����)" & strInput
                    sTemp = Left(strBookingNo, iCanma-1)
                    strBookingNo = Mid(strBookingNo, iCanma+1)
                Else
                    sTemp = strBookingNo
                    strBookingNo = ""
                End If
' 2009/05/09 mod-s �`�����ǉ�/SQL�C���W�F�N�V�����Ή�
'                sWhere = "ExportCont.BookNo='" & Trim(sTemp) & "'"
'                bWriteFile = bWriteFile + SerchExpCntnr(conn, rsd, ti, sWhere)
                bRtn = ChkSQLInjectionBookNo(sTemp)
                If bRtn Then
                    sWhere = "ExportCont.BookNo='" & Trim(sTemp) & "' and Container.UserCode='" & Trim(strUserCodeEx) & "'"
                    bWriteFile = bWriteFile + SerchExpCntnr(conn, rsd, ti, sWhere)
                End If
' 2009/05/09 mod-e �`�����ǉ�/SQL�C���W�F�N�V�����Ή�
            Loop
'������ Mod_E  by nics 200902����
            Session.Contents("findkind")="Booking"       ' �Q�ƃ��[�h
        Else
            strInput = "," & "���͓��e," & strCntnrNo
            strOption = "���͕��@����,0(�R���e�iNo.1��)" & strInput

            Session.Contents("findkey")=strCntnrNo       ' �Q��Key���L��
'������ Mod_S  by nics 200902����
'            iCanma = InStr(strCntnrNo,",")
'            Do While iCanma>0
'                strOption = "���͕��@����,1(�R���e�iNo.����)" & strInput
'                sTemp = Trim(Left(strCntnrNo,iCanma-1))
'                strCntnrNo = Right(strCntnrNo,Len(strCntnrNo)-iCanma)
'                If sWhere<>"" Then
'                    sWhere = sWhere & " Or ExportCont.ContNo='" & sTemp & "'"
'                Else
'                    sWhere = "ExportCont.ContNo='" & sTemp & "'"
'                End If
'                iCanma = InStr(strCntnrNo,",")
'            Loop
'            If sWhere<>"" Then
'                sWhere = sWhere & " Or ExportCont.ContNo='" & Trim(strCntnrNo) & "'"
'            Else
'                sWhere = "ExportCont.ContNo='" & Trim(strCntnrNo) & "'"
'            End If
'������
            Do While strCntnrNo <> ""
                iCanma = InStr(strCntnrNo, ",")
                If iCanma > 0 Then
                    strOption = "���͕��@����,1(�R���e�iNo.����)" & strInput
                    sTemp = Left(strCntnrNo, iCanma-1)
                    strCntnrNo = Mid(strCntnrNo, iCanma+1)
                Else
                    sTemp = strCntnrNo
                    strCntnrNo = ""
                End If
' 2009/05/09 mod-s �`�����ǉ�/SQL�C���W�F�N�V�����Ή�
'                sWhere = "ExportCont.ContNo='" & Trim(sTemp) & "'"
'                bWriteFile = bWriteFile + SerchExpCntnr(conn, rsd, ti, sWhere)
                bRtn = ChkSQLInjectionCntnrNo(sTemp)
                If bRtn Then
                    sWhere = "ExportCont.ContNo='" & Trim(sTemp) & "' and Container.UserCode='" & Trim(strUserCodeEx) & "'"
                    bWriteFile = bWriteFile + SerchExpCntnr(conn, rsd, ti, sWhere)
                End If
' 2009/05/09 mod-e �`�����ǉ�/SQL�C���W�F�N�V�����Ή�
            Loop
'������ Mod_E  by nics 200902����
            Session.Contents("findkind")="Cntnr"         ' �Q�ƃ��[�h
        End If

'������ Del_S  by nics 200902����
'        ' �擾�����R���e�i��񃌃R�[�h���e���|�����t�@�C���ɏ����o��
'        strFileName="./temp/" & strFileName
'        ' �e���|�����t�@�C����Open
'        Set ti=fs.OpenTextFile(Server.MapPath(strFileName),2,True)
'
'        bWriteFile = SerchExpCntnr(conn, rsd, ti, sWhere)
'������ Del_E  by nics 200902����

        ti.Close
        conn.Close

        ' Temp�t�@�C�������ݒ�
        SetTempFile "EXPORT"

        If bWriteFile = 0 Then
            ' �Y�����R�[�h�Ȃ��Ƃ�
            bError = true
            strError = "�w������ɊY������R���e�i�͂���܂���ł����B"
            strOption = "���͓��e�̐���:1(���)"
        Else
            strOption = "���͓��e�̐���:0(������)"
        End If

    End If

	Dim iWrkNum
	If strBookingNoLog="" Then
		iWrkNum = 11
		Do While InStr(strCntnrNoLog,",")>0
			strCntnrNoLog = Left(strCntnrNoLog,InStr(strCntnrNoLog,",")-1) & _
							"/" & Right(strCntnrNoLog,Len(strCntnrNoLog)-InStr(strCntnrNoLog,",")) & _
							"/" & strUserCodeExNoLog
		Loop
		strOption = strCntnrNoLog & "," & strOption
	Else
		iWrkNum = 12
		Do While InStr(strBookingNoLog,",")>0
			strBookingNoLog = Left(strBookingNoLog,InStr(strBookingNoLog,",")-1) & _
							"/" & Right(strBookingNoLog,Len(strBookingNoLog)-InStr(strBookingNoLog,",")) & _
							"/" & strUserCodeExNoLog
		Loop
		strOption = strBookingNoLog & "," & strOption
	End If

    ' �A�o�R���e�i�Ɖ�
    WriteLog fs, "1001","�A�o�R���e�i�Ɖ�(�O��)",iWrkNum, strOption

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
<!--
function FancBack()
{
        window.history.back();
}
// -->
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
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
          <td><img src="../gif/botan.gif" width="17" height="17" vspace="4"></td>
          <td nowrap><b>�d�o�n���</b></td>
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
    DispMenuBarBack "JavaScript:FancBack()"
%>
</body>
</html>

<%
    Else
        If bWriteFile = 1 Then
            '�߂��ʎ�ʂ��L��
            Session.Contents("dispreturn")=0

            ' �ڍ׉�ʂփ��_�C���N�g
            Response.Redirect "expdetail.asp?line=1"     '�A�o�R���e�i�ڍ�
        Else
            '�߂��ʎ�ʂ��L��
            Session.Contents("dispreturn")=0
            ' �ꗗ��ʂփ��_�C���N�g
            Response.Redirect "explist.asp"             '�A�o�R���e�i�ꗗ
        End If
    End If
%>
