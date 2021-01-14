<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<!--#include file="vessel.inc"-->

<%
    ' �Z�b�V�����̃`�F�b�N
    CheckLogin "nyuryoku-kaika.asp"
	sSosin = Trim(Session.Contents("userid"))

	' �g�����U�N�V�����t�@�C���̊g���q 
	Const SEND_EXTENT = "snd"
	' �g�����U�N�V�����h�c
	Const sTranID = "IM16"
	' �����敪
	Const sSyori = "R"

	' ���M�ꏊ
	Const sPlace = ""
    ' �G���[�t���O�̃N���A
    bError = false

    ' �w������̎擾
    Dim strKind
    strKind = Request.QueryString("kind")

    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemObject")

    ' �e���|�����t�@�C�������쐬���āA�Z�b�V�����ϐ��ɐݒ�
    Dim strFileName
    strFileName = GetNumStr(Session.SessionID, 8) & ".csv"
    Session.Contents("tempfile") = strFileName

    ' �]���t�@�C���̎擾
    tb=Request.TotalBytes      :' �u���E�U����̃g�[�^���T�C�Y
    br=Request.BinaryRead(tb)  :' �u���E�U����̐��f�[�^

    ' BASP21 �R���|�[�l���g�̍쐬
    Set bsp=Server.CreateObject("basp21")

    filesize=bsp.FormFileSize(br,"csvfile")
    filename=bsp.FormFileName(br,"csvfile")

    fpath=strFileName
    fpath=fs.BuildPath(Server.MapPath("./temp"),fpath)

    lng=bsp.FormSaveAs(br,"csvfile",fpath)

    ' �t�@�C���]���Ɏ��s�����Ƃ�

    Dim sText	'�]���t�@�C��

    If lng<=0 Then
        bError=true
        strError = "'" & filename & "'�t�@�C���̓]���Ɏ��s���܂����B"
    Else
        ' �]���t�@�C����Open
        Set ti=fs.OpenTextFile(fpath,1,True)
		Dim anyTmp, iRecCount, iWriteCnt, iErrLine
		iRecCount = 0
		iWriteCnt = 0
		iErrLine = 0

        ConnectSvr conn, rsd
        ' �]���t�@�C���̃��R�[�h������ԌJ��Ԃ�
        Do While Not ti.AtEndOfStream
            strError=""
			sText = ti.ReadLine
			anyTmp = Split(sText, ",")
			If Ubound(anytmp) <> 1 Then
                ' �t�@�C���`���G���[
                strError="���ڐ����ُ�ł��B"
			Else
				'�t�@�C���`���I�ɂ͐���
				if strKind = "cntnr" then
					'�A���R���e�i�e�[�u��������
                    sql = "SELECT ImportCont.VslCode, ImportCont.VoyCtrl, ImportCont.BLNo, VslSchedule.DsVoyage" & _
                          " FROM ImportCont, VslSchedule" & _
                          " WHERE ImportCont.ContNo = '" & Trim(anyTmp(0)) & "'" & _
                          " And VslSchedule.VslCode = ImportCont.VslCode" & _
                          " And VslSchedule.VoyCtrl = ImportCont.VoyCtrl"

                    'SQL�𔭍s���ėA���R���e�i������
                    rsd.Open sql, conn, 0, 1, 1
                    If Not rsd.EOF Then
                        sVslCode = Trim(rsd("VslCode"))		'�D��
                        sVoyCtrl = Trim(rsd("DsVoyage"))	'���q
                        sBLNo = Trim(rsd("BLNo"))			'BL�ԍ�
                        sText=sVslCode & "," & sVoyCtrl & "," & Trim(anyTmp(0)) & "," & sBLNo
                    Else
                        ' �R���e�i �G���[
                        strError=strError & "�Y������A���R���e�i�����݂��܂���B(" & anyTmp(0) & ") "
                    End If
                    rsd.Close
                Else
                    sql = "SELECT BL.VslCode, BL.VoyCtrl, VslSchedule.DsVoyage" & _
                          " FROM BL, VslSchedule" & _
                          " WHERE BL.BLNo = '" & Trim(anyTmp(0)) & "'" & _
                          " And VslSchedule.VslCode = BL.VslCode" & _
                          " And VslSchedule.VoyCtrl = BL.VoyCtrl"

                    'SQL�𔭍s����BL������
                    rsd.Open sql, conn, 0, 1, 1
                    If Not rsd.EOF Then
                        sVslCode = Trim(rsd("VslCode"))		'�D��
                        sVoyCtrl = Trim(rsd("DsVoyage"))	'���q
                        sText=sVslCode & "," & sVoyCtrl & ",," & Trim(anyTmp(0))
                    Else
                        ' BL �G���[
                        strError=strError & "�Y������BL No.�����݂��܂���B(" & anyTmp(0) & ") "
                    End If
                    rsd.Close
                End If
                ' �͂������̕K�{�`�F�b�N
                sTemp=ChangeDate(Trim(anyTmp(1)),12)
                If sTemp="" Then
                    ' ���͂Ȃ��Ƃ� �G���[
                    strError=strError & "�͂��������w�肳��Ă��܂���B"
                ElseIf InStr(sTemp,"(")<>0 Then
                    ' ���̓f�[�^ �G���[
                    strError=strError & "�͂�������" & sTemp
                Else
                    sText=sText & "," & sTemp
                End If

                If strError="" Then
                    ReDim Preserve Tmp(iWriteCnt)
                    Tmp(iWriteCnt) = sText
                    iWriteCnt = iWriteCnt + 1
                End If
            End If
            iRecCount = iRecCount + 1
            If strError<>"" Then
                ReDim Preserve sErrLine(iErrLine)
                sErrLine(iErrLine) = iRecCount & "����:" & strError
                iErrLine = iErrLine + 1
            End If
        Loop
        ti.Close

        If iErrLine > 0 Then
            bError = true
            strError = "'" & filename & "'�t�@�C���̌`�����Ⴂ�܂��B" & "<br>"
            For i = 0 to iErrLine - 1
                strError = strError & sErrLine(i) & "<br>"
            Next
        Else
            iOutCount=0
            ' �o�̓t�@�C���ݒ�
			iSeqNo_IM16 = GetDailyTransNo

			sFileName = ArrangeNumV(Month(Now), 2) & ArrangeNumV(Day(Now), 2) & iSeqNo_IM16
			strFileName_01 = "./send/" & sFileName & "." & SEND_EXTENT
		    Set tout=fs.OpenTextFile(Server.MapPath(strFileName_01),2,True)

            For iCount=0 To iWriteCnt - 1
                '�V�[�P���X�ԍ�
                anyTmp1 = Split(Tmp(iCount), ",")
				If iCount <> 0  Then
					iSeqNo_IM16 = GetDailyTransNo
				End If
				'�ʐM�����擾
				sTusin  = SetTusinDate

				sIM16 = iSeqNo_IM16 & "," & sTranID & "," & sSyori & ","  & sTusin & ",Web - " & _
						sSosin & "," & sPlace & "," & anyTmp1(0) & "," &  anyTmp1(1) & "," & _
						anyTmp1(2) & "," & anyTmp1(3) & "," & anyTmp1(4) & ",," & sSosin
				tout.WriteLine sIM16
                iOutCount=iOutCount+1
			Next 

		    tout.Close

		    ' �G���[���b�Z�[�W�̕\��
			strError = "����ɍX�V����܂����B"
		End IF
    End If

    If bError Then
        strOption = filename & "," & "���͓��e�̐���:1(���)"
    Else
        strOption = filename & "," & "���͓��e�̐���:0(������) " & iOutCount & "���o��"
    End If

	Dim iWrkNum
	If strKind="bl" Then
		iWrkNum = "22"
	Else
		iWrkNum = "21"
	End If

    ' ������q�ɓ͂������w���p�t�@�C���]�����
    WriteLog fs, "4007","�C�ݓ��͎�����q�ɓ�������-CSV�t�@�C���]��",iWrkNum, strOption

'''    If bError Then
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
					<td rowspan=2><img src="gif/csvt.gif" width="506" height="73"></td>
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
						<td nowrap><b>������q�ɓ��������i�w���ρj</b></td>
						<td><img src="gif/hr.gif"></td>
					</tr>
				</table>
				<table>
					<tr> 
						<td nowrap>
								<font color="#000066" size="+1">�y������q�ɓ��������p�t�@�C���]����ʁz</font><br><BR>
<%
    ' �G���[���b�Z�[�W�̕\��
    If strError="����ɍX�V����܂����B" Then
        DispInformationMessage strError
    Else
        DispErrorMessage strError
    End If
%>
							</dl>
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
    DispMenuBarBack "JavaScript:window.history.back()"
%>
</body>
</html>

