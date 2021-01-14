<%@Language="VBScript" %>

<!--#include file="Common.inc"-->
<!--#include file="vessel.inc"-->

<%
    ' �Z�b�V�����̃`�F�b�N
    CheckLogin "nyuryoku-in1.asp"

    ' �g�����U�N�V�����t�@�C���̊g���q 
    Const SEND_EXTENT = "snd"

    ' �����敪
    Const sSyori = "R"

    ' �g�����O�P
    Const sTran1 = "VS01"

    ' �g�����O�P
    Const sTran2 = "VS02"

    ' ���M��
    sSosin = Trim(Session.Contents("userid"))

    ' ���M�ꏊ
    Const sPlace = ""

    ' �G���[�t���O�̃N���A
    bError = false

    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemObject")
    ' File System Object �̐���
    Set fs2=Server.CreateObject("Scripting.FileSystemObject")

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

    fpath=fs.GetFileName(strFileName)
    fpath=fs.BuildPath(Server.MapPath("./temp"),fpath)

    lng=bsp.FormSaveAs(br,"csvfile",fpath)

    ' �t�@�C���]���Ɏ��s�����Ƃ�
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
            If Ubound(anytmp) <> 7 Then
                ' �t�@�C���`���G���[
                strError="���ڐ����ُ�ł��B"
            Else
                '�t�@�C���`���I�ɂ͐���
                ' �R�[���T�C���̃`�F�b�N
                sql = "SELECT FullName, ShipLine FROM mVessel WHERE VslCode='" & Trim(anyTmp(0)) & "'"
                'SQL�𔭍s���đD���}�X�^�[������
                rsd.Open sql, conn, 0, 1, 1
                If rsd.EOF Then
                    ' �Y�����R�[�h�̂Ȃ��Ƃ� �G���[
                    strError="�Y������R�[���T�C�����L��܂���B(" & anyTmp(0) & ") "
                Else
                    sText=Trim(anyTmp(0))
                    strShipLine=Trim(rsd("ShipLine"))
                End If
                rsd.Close
                ' VoyageNo�̃`�F�b�N
                If Len(Trim(anyTmp(1)))>12 Or Len(Trim(anyTmp(1)))<=0 Then
                    ' VoyageNo�̒��� �G���[
                    strError=strError & "Voyage No.�̒������ُ�ł��B(" & anyTmp(1) & ") "
                Else
                    ' VoyageNo�̃`�F�b�N
                    sql = "SELECT VoyCtrl, DsVoyage, LdVoyage FROM VslSchedule WHERE VslCode='" & Trim(anyTmp(0)) & "' And (" & _
                          "DsVoyage='" & Trim(anyTmp(1)) & "' Or LdVoyage='" & Trim(anyTmp(1)) & "')"
                    'SQL�𔭍s���č`�R�[�h�}�X�^�[������
                    rsd.Open sql, conn, 0, 1, 1
                    If rsd.EOF Then
                        ' �Y�����R�[�h�̂Ȃ��Ƃ�
                        sText=sText & "," & Trim(anyTmp(1)) & "," & Trim(anyTmp(1))
						iVoyCtrl=""
                    Else
                        sText=sText & "," & Trim(rsd("DsVoyage")) & "," & Trim(rsd("LdVoyage"))
						iVoyCtrl=rsd("VoyCtrl")
                    End If
                    rsd.Close
                End If
                ' �^�s�D�Ђ̐ݒ�
                sText=sText & "," & strShipLine
                ' �`�R�[�h�̃`�F�b�N
                sql = "SELECT FullName FROM mPort WHERE PortCode='" & Trim(anyTmp(2)) & "'"
                'SQL�𔭍s���č`�R�[�h�}�X�^�[������
                rsd.Open sql, conn, 0, 1, 1
                If rsd.EOF Then
                    ' �Y�����R�[�h�̂Ȃ��Ƃ� �G���[
                    strError=strError & "�Y������`�R�[�h���L��܂���B(" & anyTmp(2) & ") "
                Else
                    sText=sText & "," & Trim(anyTmp(2))
                End If
                rsd.Close
                ' ���ݗ\�莞���̕K�{�`�F�b�N
                sTemp=ChangeDate(Trim(anyTmp(3)),12)
                If sTemp="" Then
                    ' ���͂Ȃ��Ƃ� �G���[
                    strError=strError & "���ݗ\�莞�����w�肳��Ă��܂���B"
                ElseIf InStr(sTemp,"(")<>0 Then
                    ' ���̓f�[�^ �G���[
                    strError=strError & "���ݗ\�莞����" & sTemp
                Else
                    sText=sText & "," & sTemp
                End If
                ' ���݊��������̃`�F�b�N
                sTemp=ChangeDate(Trim(anyTmp(4)),12)
                If InStr(sTemp,"(")<>0 Then
                    ' ���̓f�[�^ �G���[
                    strError=strError & "���݊���������" & sTemp
                Else
                    sText=sText & "," & sTemp
                End If
                ' ���݊��������̃`�F�b�N
                sTemp=ChangeDate(Trim(anyTmp(5)),12)
                If InStr(sTemp,"(")<>0 Then
                    ' ���̓f�[�^ �G���[
                    strError=strError & "���݊���������" & sTemp
                Else
                    sText=sText & "," & sTemp
                End If
                ' ����Long Schedule�̃`�F�b�N
                sTemp=ChangeDate(Trim(anyTmp(6)),8)
                If InStr(sTemp,"(")<>0 Then
                    ' ���̓f�[�^ �G���[
                    strError=strError & "����Long Schedule��" & sTemp
                Else
                    sText=sText & "," & sTemp
                    If sTemp<>"" Then
                         sText=sText & "2359"
                    End If
                End If
                ' ����Long Schedule�̃`�F�b�N
                sTemp=ChangeDate(Trim(anyTmp(7)),8)
                If InStr(sTemp,"(")<>0 Then
                    ' ���̓f�[�^ �G���[
                    strError=strError & "����Long Schedule��" & sTemp
                Else
                    sText=sText & "," & sTemp
                    If sTemp<>"" Then
                         sText=sText & "2359"
                    End If
                End If

				sText=sText & "," & iVoyCtrl
				If strError="" Then
					iVesselFlg=0
					For i=0 To iWriteCnt - 1
						anyChk1 = Split(Tmp(i), ",")
						anyChk2 = Split(sText, ",")
						' �����D�̓����|�[�g�̃f�[�^������Ƃ��̃`�F�b�N
						If anyChk1(0)=anyChk2(0) And anyChk1(1)=anyChk2(1) And anyChk1(2)=anyChk2(2) And anyChk1(4)=anyChk2(4) Then
							strError=strError & "�����R�[���T�C���ɑ΂��ē����`����������w�肳��Ă��܂��B(" & anyChk1(4) & ")"
							Exit For
						End If
						' �����D�̃f�[�^�����ꂽ�Ƃ���ɂ���Ƃ��̃`�F�b�N
						If anyChk1(0)=anyChk2(0) And anyChk1(1)=anyChk2(1) And anyChk1(2)=anyChk2(2) Then
							If iVesselFlg=0 Then
								iVesselFlg=1
							ElseIf iVesselFlg=2 Then
								Exit For
							End If
						Else
							If iVesselFlg=1 Then
								iVesselFlg=2
							End If
						End If
					Next
					If iVesselFlg=2 Then
						strError=strError & "�����R�[���T�C���̃f�[�^�����ꂽ�ꏊ�Ŏw�肳��Ă��܂��B(" & anyChk1(0) & ")"
					End If
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
            iSeqNo_VS01 = GetDailyTransNo
            sFileName = ArrangeNumV(Month(Now), 2) & ArrangeNumV(Day(Now), 2) & iSeqNo_VS01
            strFileName_01 = "./send/" & sFileName & "." & SEND_EXTENT
            Set tout=fs.OpenTextFile(Server.MapPath(strFileName_01),2,True)
            Dim sVs01, sVs02, sVs02_Body								'�������݃f�[�^
            Dim sVslCode_SV, sDsVoyage_SV, sLdVoyage_SV, sShipLine_SV	'�ꌏ�O�f�[�^(��ٻ��/���q)
            sVslCode_SV = ""
            sVoyage_SV = ""
			Dim strPort()
			Dim strPortData()
			Dim inpPortCount
			Dim sBefDate
			Dim sAftDate
			Dim sWkText
			Dim bSwap
			inpPortCount=0
            For iCount=0 To iWriteCnt - 1
                '�V�[�P���X�ԍ�
                anyTmp1 = Split(Tmp(iCount), ",")
                If anyTmp1(0) = sVslCode_SV And anyTmp1(1) = sDsVoyage_SV Then
                Else
                    If sVslCode_SV <> "" Then

						iPortOutCount=0
						If Trim(iVoyCtrl_SV)<>"" Then
							' SQL�𔭍s���Ė{�D��`�n������
							sql = "SELECT VslPort.PortCode, VslPort.ETA, VslPort.TA, VslPort.ETD, VslPort.TD, VslPort.ETALong, VslPort.ETDLong " & _
									"FROM VslPort WHERE VslPort.VslCode='" & sVslCode_SV & "' And VslPort.VoyCtrl=" & iVoyCtrl_SV & _
									" ORDER BY VslPort.CallSeq "
							rsd.Open sql, conn, 0, 1, 1
							Do While Not rsd.EOF
								' ��`�n��񃌃R�[�h�̍쐻
								strRec = Trim(rsd("PortCode")) & "," & _
										 SetTusinDate2(DispDateTime(rsd("ETA"),0)) & "," & SetTusinDate2(DispDateTime(rsd("TA"),0)) & ","  & _
										 SetTusinDate2(DispDateTime(rsd("ETD"),0)) & "," & SetTusinDate2(DispDateTime(rsd("TD"),0)) & ","  & _
										 SetTusinDate2(DispDateTime(rsd("ETALong"),0)) & "," & SetTusinDate2(DispDateTime(rsd("ETDLong"),0))
								ReDim Preserve strPortData(iPortOutCount)
								strPortData(iPortOutCount) = strRec
								iPortOutCount=iPortOutCount + 1
								rsd.MoveNext
							Loop
							rsd.Close
						End If

						' �����`�̃f�[�^�𓝍�
						For i=0 To inpPortCount-1
							anyPort1 = Split(strPort(i), ",")
							iFlg=false
							For j=0 To iPortOutCount-1
								anyPort2 = Split(strPortData(j), ",")
								If anyPort1(0)=anyPort2(0) Then
									iFlg=true
									If anyPort1(1)<>"" Then
										anyPort2(5)=anyPort1(1)
									End If
									If anyPort1(2)<>"" Then
										anyPort2(6)=anyPort1(2)
									End If
									If anyPort1(3)<>"" Then
										anyPort2(1)=anyPort1(3)
									End If
									If anyPort1(4)<>"" Then
										anyPort2(2)=anyPort1(4)
									End If
									If anyPort1(5)<>"" Then
										anyPort2(3)=anyPort1(5)
									End If
									If anyPort1(6)<>"" Then
										anyPort2(4)=anyPort1(6)
									End If
									strTmp=""
									For k=0 To UBound(anyPort2)
										strTmp=strTmp & anyPort2(k) & ","
									Next
									strPortData(j)=Left(strTmp,Len(strTmp)-1)
									Exit For
								End If
							Next
							If Not iFlg Then
								ReDim Preserve strPortData(iPortOutCount)
								strPortData(iPortOutCount) = anyPort1(0) & "," & anyPort1(3) & "," & anyPort1(4) & "," & anyPort1(5) & "," & _
														anyPort1(6) & "," & anyPort1(1) & "," & anyPort1(2)
								iPortOutCount=iPortOutCount + 1
							End If
						Next
						' �f�[�^�𒅊ݗ\�莞���Ń\�[�g����(��������̗v�]�ŁA�R�����g�� 2002/02/26)
'						For i = 0 to iPortOutCount - 2
'							anyTmp=Split(strPortData(i),",")
'							sBefDate = anyTmp(1)
'							For j = (i + 1) To iPortOutCount - 1
'								anyTmp=Split(strPortData(j),",")
'								sAftDate = anyTmp(1)
'								bSwap = FALSE
'								If sAftDate <> "" Then
'									If sBefDate = "" Then
'										bSwap = TRUE
'									ElseIf sBefDate > sAftDate Then
'										bSwap = TRUE
'									End If
'								End IF
'								If bSwap = TRUE Then
'									sWkText = strPortData(i)
'									strPortData(i) = strPortData(j)
'									strPortData(j) = sWkText
'								End IF
'							Next
'						Next

                        '�������ݏ���
                        If iOutCount<>0 Then
                            iSeqNo_VS01 = GetDailyTransNo
                        End If
						'�ʐM�����擾
						sTusin  = SetTusinDate
						sVs01 = iSeqNo_VS01 & "," & sTran1 & "," & sSyori & ","  & sTusin & ",Web - " & _
								sSosin & "," & sPlace & "," & sVslCode_SV & "," &  sDsVoyage_SV & "," & _
								sLdVoyage_SV & "," &  sShipLine_SV
						tout.WriteLine sVs01

						sVs02_Body = ""
						For i=0 To iPortOutCount-1
							anyPort2 = Split(strPortData(i), ",")
							sVs02_Body = sVs02_Body & "," & anyPort2(0) & "," &  anyPort2(5) & "," & anyPort2(6) & "," & anyPort2(1) & _
										"," & anyPort2(2) & "," & anyPort2(3) & "," & anyPort2(4)
						Next

						'�擪�ɕ�����𖄂ߍ��ޖ��ו�������������
						sVs02 = iSeqNo_VS01 & "," & sTran2 & "," & sSyori & ","  & sTusin & ",Web - " & _
								sSosin & "," & sPlace & "," & sVslCode_SV & "," &  sDsVoyage_SV & "," & _
								sLdVoyage_SV & sVs02_Body
						tout.WriteLine sVs02
                        iOutCount=iOutCount+1
						inpPortCount=0
					End If
				End If

				sVs02_Body = anyTmp1(4) & "," &  anyTmp1(8) & "," & anyTmp1(9) & "," & anyTmp1(5) & _
						"," & anyTmp1(6) & ",," & anyTmp1(7) & ","
				iVoyCtrl=anyTmp1(10)
				ReDim Preserve strPort(inpPortCount)
				strPort(inpPortCount) = sVs02_Body
				inpPortCount = inpPortCount + 1

			 	'�ꌏ�O�f�[�^�Z�b�g(��ٻ��/���q��)
				iVoyCtrl_SV = anyTmp1(10)
				sVslCode_SV = anyTmp1(0)
				sDsVoyage_SV = anyTmp1(1)
				sLdVoyage_SV = anyTmp1(2)
				sShipLine_SV = anyTmp1(3)
			Next 
            '�Ō�̃f�[�^
            If sVslCode_SV<>"" Then

				iPortOutCount=0
				If Trim(iVoyCtrl)<>"" Then
					' SQL�𔭍s���Ė{�D��`�n������
					sql = "SELECT VslPort.PortCode, VslPort.ETA, VslPort.TA, VslPort.ETD, VslPort.TD, VslPort.ETALong, VslPort.ETDLong " & _
							"FROM VslPort WHERE VslPort.VslCode='" & anyTmp1(0) & "' And VslPort.VoyCtrl=" & iVoyCtrl & _
							" ORDER BY VslPort.CallSeq "
					rsd.Open sql, conn, 0, 1, 1
					Do While Not rsd.EOF
						' ��`�n��񃌃R�[�h�̍쐻
						strRec = Trim(rsd("PortCode")) & "," & _
								 SetTusinDate2(DispDateTime(rsd("ETA"),0)) & "," & SetTusinDate2(DispDateTime(rsd("TA"),0)) & ","  & _
								 SetTusinDate2(DispDateTime(rsd("ETD"),0)) & "," & SetTusinDate2(DispDateTime(rsd("TD"),0)) & ","  & _
								 SetTusinDate2(DispDateTime(rsd("ETALong"),0)) & "," & SetTusinDate2(DispDateTime(rsd("ETDLong"),0))
						ReDim Preserve strPortData(iPortOutCount)
						strPortData(iPortOutCount) = strRec
						iPortOutCount=iPortOutCount + 1
						rsd.MoveNext
					Loop
					rsd.Close
				End If

				' �����`�̃f�[�^�𓝍�
				For i=0 To inpPortCount-1
					anyPort1 = Split(strPort(i), ",")
					iFlg=false
					For j=0 To iPortOutCount-1
						anyPort2 = Split(strPortData(j), ",")
						If anyPort1(0)=anyPort2(0) Then
							iFlg=true
							If anyPort1(1)<>"" Then
								anyPort2(5)=anyPort1(1)
							End If
							If anyPort1(2)<>"" Then
								anyPort2(6)=anyPort1(2)
							End If
							If anyPort1(3)<>"" Then
								anyPort2(1)=anyPort1(3)
							End If
							If anyPort1(4)<>"" Then
								anyPort2(2)=anyPort1(4)
							End If
							If anyPort1(5)<>"" Then
								anyPort2(3)=anyPort1(5)
							End If
							If anyPort1(6)<>"" Then
								anyPort2(4)=anyPort1(6)
							End If
							strTmp=""
							For k=0 To UBound(anyPort2)
								strTmp=strTmp & anyPort2(k) & ","
							Next
							strPortData(j)=Left(strTmp,Len(strTmp)-1)
							Exit For
						End If
					Next
					If Not iFlg Then
						ReDim Preserve strPortData(iPortOutCount)
						strPortData(iPortOutCount) = anyPort1(0) & "," & anyPort1(3) & "," & anyPort1(4) & "," & anyPort1(5) & "," & _
												anyPort1(6) & "," & anyPort1(1) & "," & anyPort1(2)
						iPortOutCount=iPortOutCount + 1
					End If
				Next

				' �f�[�^�𒅊ݗ\�莞���Ń\�[�g����(��������̗v�]�ŁA�R�����g�� 2002/02/26)
'				For i = 0 to iPortOutCount - 2
'					anyTmp=Split(strPortData(i),",")
'					sBefDate = anyTmp(1)
'					For j = (i + 1) To iPortOutCount - 1
'						anyTmp=Split(strPortData(j),",")
'						sAftDate = anyTmp(1)
'						bSwap = FALSE
'						If sAftDate <> "" Then
'							If sBefDate = "" Then
'								bSwap = TRUE
'							ElseIf sBefDate > sAftDate Then
'								bSwap = TRUE
'							End If
'						End IF
'						If bSwap = TRUE Then
'							sWkText = strPortData(i)
'							strPortData(i) = strPortData(j)
'							strPortData(j) = sWkText
'						End IF
'					Next
'				Next

    			'�������ݏ���
    			iSeqNo_VS01 = GetDailyTransNo
    			'�ʐM�����擾
    			sTusin = SetTusinDate
    			sVs01 = ""
    			sVs01 = iSeqNo_VS01 & "," & sTran1 & "," & sSyori & ","  & sTusin & ",Web - " & _
    					sSosin & "," & sPlace & "," & sVslCode_SV & "," &  sDsVoyage_SV & "," & _
    					sLdVoyage_SV & "," &  sShipLine_SV
    			tout.WriteLine sVs01

				sVs02_Body = ""
				For i=0 To iPortOutCount-1
					anyPort2 = Split(strPortData(i), ",")
					sVs02_Body = sVs02_Body & "," & anyPort2(0) & "," &  anyPort2(5) & "," & anyPort2(6) & "," & anyPort2(1) & _
								"," & anyPort2(2) & "," & anyPort2(3) & "," & anyPort2(4)
				Next

    			'�擪�ɕ�����𖄂ߍ��ޖ��ו�������������
    			sVs02 = iSeqNo_VS01 & "," & sTran2 & "," & sSyori & ","  & sTusin & ",Web - " & _
    					sSosin & "," & sPlace & "," & sVslCode_SV & "," &  sDsVoyage_SV & "," & _
    					sLdVoyage_SV & sVs02_Body
    			tout.WriteLine sVs02
                iOutCount=iOutCount+1
            End If
		    tout.Close
		    ' �G���[���b�Z�[�W�̕\��
			strError = "����ɍX�V����܂����B"
        End If
    End If

    If bError Then
        strOption = filename & "," & "���͓��e�̐���:1(���)"
    Else
        strOption = filename & "," & "���͓��e�̐���:0(������) " & iOutCount & "���o��"
    End If

    ' �D�Зp/�^�[�~�i���t�@�C���]����ʏƉ�
    WriteLog fs, "3002","�D�Ё^�^�[�~�i������-CSV�t�@�C���]��","20", strOption

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
                
                  <td nowrap><b>�D�Ё^�^�[�~�i������</b></td>
                   <td><img src="gif/hr.gif"></td>
 </tr>
</table>
      <table>
        <tr> 
          <td nowrap align=center>
            <font color="#000066" size="+1">�y�D�Зp�t�@�C���]����ʁz</font>
			<BR><br>
<%
    ' �G���[���b�Z�[�W�̕\��
    If strError="����ɍX�V����܂����B" Then
        DispInformationMessage strError
    Else
        DispErrorMessage strError
    End If
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
<!-------------�o�^��ʏI���--------------------------->
<%
    DispMenuBarBack "nyuryoku-csv.asp"
%>
</body>
</html>
