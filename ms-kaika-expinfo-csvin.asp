<%@Language="VBScript" %>

<!--#include file="Common.inc"-->
<!--#include file="vessel.inc"-->
<!--#include file="csvcheck.inc"-->

<%
''    ' �Z�b�V�����̃`�F�b�N
    CheckLogin "nyuryoku-kaika.asp"

	sSosin = Trim(Session.Contents("userid"))

	' �g�����U�N�V�����t�@�C���̊g���q 
	Const SEND_EXTENT = "snd"
	' �g�����U�N�V�����h�c
	Const sTranID = "EX16"
	' �����敪
	Const sSyori = "R"

	' ���M�ꏊ
	Const sPlace = ""
    ' �G���[�t���O�̃N���A
    bError = false

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
			If Ubound(anytmp) <> 12 Then
                ' �t�@�C���`���G���[
                strError="���ڐ����ُ�ł��B"
			Else
				'�e���ڌ������������`�F�b�N
				'�D�� anyTmp(0)
				strError = strError & CheckParam(anyTmp(0),"�D���i�R�[���T�C���j",7,0,true,false) 
 				'Voyage No. anyTmp(1)
				strError = strError & CheckParam(anyTmp(1),"Voyage No.",12,0,true,false) 
 				'�׎�R�[�h anyTmp(2)
				strError = strError & CheckParam(anyTmp(2),"�׎�R�[�h",5,0,true,false) 
 				'�׎�Ǘ��ԍ� anyTmp(3)
				strError = strError & CheckParam(anyTmp(3),"�׎�Ǘ��ԍ�",10,0,true,false) 
 				'Booking No. anyTmp(4)
				strError = strError & CheckParam(anyTmp(4),"Booking No.",20,0,true,false) 
 				'�w�藤�^�Ǝ҃R�[�h anyTmp(5)
				strError = strError & CheckParam(anyTmp(5),"�w�藤�^�Ǝ҃R�[�h",3,0,false,false) 
 				'��o���q�ɓ����\����� anyTmp(6)
					'���t�̃X���b�V��������Č��������킹��
				sTemp=ChangeDate(Trim(anyTmp(6)),12)
           	    If InStr(sTemp,"(")<>0 Then
                    ' ���̓f�[�^ �G���[
                    strError=strError & "��R���q�ɓ����\�������" & sTemp
                End If
 				'�b�x�����w��� anyTmp(7)
					'���t�̃X���b�V��������Č��������킹��
				sTemp=ChangeDate(Trim(anyTmp(7)),8)
           	    If InStr(sTemp,"(")<>0 Then
                    ' ���̓f�[�^ �G���[
                    strError=strError & "�b�x�����w�����" & sTemp
                End If
 				'�T�C�Y anyTmp(8)
				strError = strError & CheckParam(anyTmp(8),"�T�C�Y",2,0,false,true) 
 				'�^�C�v anyTmp(9)
				strError = strError & CheckParam(anyTmp(9),"�^�C�v",2,0,false,false) 
 				'�n�C�g anyTmp(10)
				strError = strError & CheckParam(anyTmp(10),"����",2,0,false,true) 
 				'��o���s�b�N�ꏊ anyTmp(11)
				strError = strError & CheckParam(anyTmp(11),"��R���s�b�N�ꏊ",20,0,false,false) 
 				'�q�ɗ��� anyTmp(12)
				strError = strError & CheckParam(anyTmp(12),"�q�ɗ���",12,0,false,false) 

				'�G���[����SQL���Ȃ�ׂ����s���Ȃ��悤��If���Ŋ���
				If strError="" Then
					' �D�������݂��邩
					sql = "SELECT count(*) FROM mVessel WHERE VslCode='" & UCase(Trim(anyTmp(0))) & "'"
					rsd.Open sql, conn, 0, 1, 1
					If Not rsd.EOF Then
						iRecCount = rsd(0)
						If iRecCount=0 Then
						    bError = true
							strError = "�w�肳�ꂽ�D�������݂��܂���B(" & anyTmp(0) & ") "
						End If
					End If
					rsd.Close
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
			Dim sEX16, iSeqNo_EX16, sFileName, strFileName_01, sTran, sTusin
			iSeqNo_EX16 = GetDailyTransNo

			sFileName = ArrangeNumV(Month(Now), 2) & ArrangeNumV(Day(Now), 2) & iSeqNo_EX16
			strFileName_01 = "./send/" & sFileName & "." & SEND_EXTENT
		    Set tout=fs.OpenTextFile(Server.MapPath(strFileName_01),2,True)

            For iCount=0 To iWriteCnt - 1
                '�V�[�P���X�ԍ�
                anyTmp1 = Split(Tmp(iCount), ",")
				If iCount <> 0  Then
					iSeqNo_EX16 = GetDailyTransNo
				End If

'�g�����U�N�V�����쐬��CSV�t�@�C�������ڂ�Trim��UCase��������  2002/02/04
				For j=0 To 12
					anyTmp1(j) = UCase(Trim(anyTmp1(j)))
				Next
'�����܂�

				'���t�̃X���b�V��������Č��������킹��
				anyTmp1(6)=ChangeDate(Trim(anyTmp1(6)),12)

				anyTmp1(7)=ChangeDate(Trim(anyTmp1(7)),8)

				'�ʐM�����擾
				sTusin  = SetTusinDate

				sEX16 = iSeqNo_EX16 & "," & sTranID & "," & sSyori & ","  & sTusin & ",Web - " & _
						sSosin & "," & sPlace & "," & anyTmp1(0) & "," &  anyTmp1(1) & "," & _
						anyTmp1(4) & "," & anyTmp1(2) & "," & anyTmp1(3) & "," & sSosin & "," & _
						sCont & "," & anyTmp1(8) & "," & anyTmp1(9) & "," & anyTmp1(10) & "," & _
						anyTmp1(12) & "," & anyTmp1(5) & "," & anyTmp1(6) & "," & anyTmp1(7) & "," & _
						anyTmp1(11)
				tout.WriteLine sEX16
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

    ' �C�ݗp�t�@�C���]����ʏƉ�
    WriteLog fs, "4103","�C�ݓ��͗A�o�ݕ����-CSV�t�@�C���]��","20", strOption

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
          <td rowspan=2><img src="gif/kaika4t.gif" width="506" height="73"></td>
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
          <td><img src="gif/botan.gif" width="17" height="17"></td>
          <td nowrap><b>CSV�t�@�C���]��</b></td>
          <td><img src="gif/hr.gif"></td>
        </tr>
      </table>
      <table>
        <tr> 
          <td nowrap align=center>
            <font color="#000066" size="+1">�y�A�o�ݕ����p�t�@�C���]����ʁz</font><br><BR>
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
    DispMenuBarBack "ms-kaika-expinfo-csv.asp"
%>
</body>
</html>

