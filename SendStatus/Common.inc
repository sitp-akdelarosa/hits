<%
'''SQL�T�[�o�ւ̐ڑ�
Function ConnDBH(ObjConn,ObjRS)
	Set ObjConn = Server.CreateObject("ADODB.Connection")
	ObjConn.Open "Provider=SQLOLEDB; Data Source = SI-KURT-LAPTOP\SQLEXPRESS; Initial Catalog = CIDB; User Id = sa; Password=P@55w0rd"
	'ObjConn.Open "wavedb","wave","wave"
ObjConn.CursorLocation = 3	'''3:adUseClient
	Set ObjRS = Server.CreateObject("ADODB.Recordset")
	ObjConn.BeginTrans
End Function


'''SQL�T�[�o�ւ̐ؒf
Function DisConnDBH(ObjConn,ObjRS)
	Set ObjRS = Nothing
	ObjConn.CommitTrans
	ObjConn.Close
	Set ObjConn = Nothing
End Function


'''���[�U�[�h�c���L�����ǂ����̃`�F�b�N
Function CheckLoginH()
	Dim strUserID
	strUserID = Trim(Session.Contents("userid"))

	'''�Z�b�V�����������ȂƂ�
	If strUserID="" Then
		Dim strtmp
		strtmp = "���O�C�����Ă���"&Session.Timeout&"���o�߂��܂����B"
		jumpErrorP "0","c000","00","�X�e�[�^�X�z�M","000",strtmp
	End If
End Function


'''���[�U�[�h�c���L�����ǂ����̃`�F�b�N
Function CheckLoginH2()
	Dim strUserID
	strUserID = Trim(Session.Contents("userid"))

	'''�Z�b�V�����������ȂƂ�
	If strUserID="" Then
		Dim strtmp
		strtmp = "���O�C�����Ă���"&Session.Timeout&"���o�߂��܂����B"
		jumpErrorP "1","c103","01","�X�e�[�^�X�z�M","000",strtmp
	End If
End Function


'''�A�N�Z�X���O�o��
'''Input :���ID�A��ʖ��A���ID�A�m�F�p�I�v�V�����f�[�^
Function WriteLogH(iNum, strScrName, iWrkNum, strOption)
	'�T�[�o�����̎擾
	Dim DayTime,day
	getDayTime DayTime
	'''���[�U�^�C�v�R�[�h
	Dim UtypeA
	UtypeA = Array("","�׎�","�C��","�D��","�`�^","���^")
	Dim strDate,logPath,ObjFSO,ObjTS

	strDate = DayTime(0) & DayTime(1) & DayTime(2)
'20091101 �{�ԃ����[�X��
	logPath = Server.MapPath("/log/" & strDate & ".log")
'logPath = Server.MapPath("/20090316/log/" & strDate & ".log")
'20091101 �{�ԃ����[�X��
	Set ObjFSO = Server.CreateObject("Scripting.FileSystemObject")
	Set ObjTS = ObjFSO.OpenTextFile(logPath,8,True)
	ObjTS.Write DayTime(0)&"/"&DayTime(1)&"/"&DayTime(2)&" "&DayTime(3)&":"&DayTime(4)&":"&Right("0" & Second(Now), 2)
	ObjTS.Write "," & iNum
	ObjTS.Write "," & strScrName
	ObjTS.Write "," & iWrkNum
	ObjTS.Write "," & UtypeA(Session.Contents("UType")) & "," & Session.Contents("userid")
	ObjTS.Write "," & Request.ServerVariables("SCRIPT_NAME")
	ObjTS.Write "," & Request.ServerVariables("REMOTE_ADDR")
	If strOption<>"" Then
		ObjTS.WriteLine "," & strOption
	Else
		ObjTS.WriteLine 
	End If
	ObjTS.Close
	Set ObjTS = Nothing
	Set ObjFSO = Nothing

    '2009/03/05 Add-S Tanaka
     xLogReg iNum,iWrkNum,Session.Contents("userid")
    '2009/03/05 Add-E Tanaka

End Function

'2009/03/05 Add-S Tanaka
'���O��DB�ɓo�^
Function xLogReg(iNum,iWrkNum,strUserId)
	'iNum		[ I ]���ID
	'iWrkNum	[ I ]���ID

	Dim ObjConn, ObjRS, StrSQL,fs,ti,strTemp
	Dim iAddFlag
	Dim PageNum(),WkNum(),PageTitle(),SubTitle(),anyTmpTitle
	Dim strTitleFileName
	dim DayTime,day,strDate
	getDayTime DayTime

	strDate = DayTime(0) & DayTime(1) & DayTime(2)

	Set fs=Server.CreateObject("Scripting.FileSystemObject")
	
	'DB�ڑ�
	ConnDBH ObjConn,ObjRS

	' ���O�^�C�g���擾
'20091101 �{�ԃ����[�X��
	strTitleFileName="../logweb.txt"
'strTitleFileName="/20090316/logweb.txt"
'20091101 �{�ԃ����[�X��
	Set ti=fs.OpenTextFile(Server.MapPath(strTitleFileName),1,True)
	iAddFlag=0
	
	'���O�^�C�g�������[�v
	Do While Not ti.AtEndOfStream
		strTemp=ti.ReadLine
		anyTmpTitle=Split(strTemp,",")
		'���O�^�C�g���ɊY��������̏ꍇ
		IF TRIM(anyTmpTitle(0))=TRIM(CSTR(iNum)) and TRIM(anyTmpTitle(1))=TRIM(CSTR(iWrkNum)) then
			'���O�C���R�[�h���Y���̂���
			IF strUserId<>"90000" and strUserId<>"10000" and strUserId<>"MES" and strUserId<>"MEO" Then
				iAddFlag=1
				Exit Do
			ELSE
				Exit Function
			End If
		End If
	Loop
	ti.Close

	'�o�^�\�ȃf�[�^�����݂���ꍇ
	If iAddFlag=1 Then

		StrSQL = ""
		StrSQL = StrSQL & " SELECT * FROM xLog "
		StrSQL = StrSQL & " WHERE LogDate='" & strDate &"'"
		StrSQL = StrSQL & " AND PageNum='" & TRIM(CSTR(iNum)) &"'"
		StrSQL = StrSQL & " AND WkNum='" & TRIM(CSTR(iWrkNum)) &"'"
		
		ObjRS.Open StrSQL, ObjConn

		'���R�[�h�����݂���ꍇ
		If ObjRS.EOF=False then
			StrSQL = ""
			StrSQL = StrSQL & " UPDATE xLog SET "
			StrSQL = StrSQL & " DataCount=DataCount+1 "
			StrSQL = StrSQL & " ,UpdtTime=GetDate() "
			StrSQL = StrSQL & " ,UpdtPgCd='COMMNINC' "
			StrSQL = StrSQL & " ,UpdtTmnl='" & Session.Contents("userid") & "'"
			StrSQL = StrSQL & " WHERE LogDate='" &strDate &"'"
			StrSQL = StrSQL & " AND PageNum='" & TRIM(CSTR(iNum)) &"'"
			StrSQL = StrSQL & " AND WkNum='" & TRIM(CSTR(iWrkNum)) &"'"
		ELSE
			StrSQL = ""
			StrSQL = StrSQL & " INSERT INTO xLog( "
			StrSQL = StrSQL & " LogDate, "
			StrSQL = StrSQL & " PageNum, "
			StrSQL = StrSQL & " WkNum, "
			StrSQL = StrSQL & " UpdtPgCd, "
			StrSQL = StrSQL & " UpdtTime, "
			StrSQL = StrSQL & " UpdtTmnl, "
			StrSQL = StrSQL & " DataCount "
			StrSQL = StrSQL & " )VALUES( "
			StrSQL = StrSQL & "'" & strDate & "',"
			StrSQL = StrSQL & "'" & TRIM(CSTR(iNum)) & "',"
			StrSQL = StrSQL & "'" & TRIM(CSTR(iWrkNum)) & "',"
			StrSQL = StrSQL & "'COMMNINC',"
			StrSQL = StrSQL & "GetDate(),"
			StrSQL = StrSQL & "'" & Session.Contents("userid") & "',"
			StrSQL = StrSQL & "1)"
		End If
		ObjConn.Execute(StrSQL)
		DisConnDBH ObjConn,ObjRS
	End If

End Function
'2009/03/05 Add-E Tanaka



'''�G���[��ʂւ̑J��
'''Input :�G���[��ʃt���O�A���ID�A���ID�A��Ɩ��A�G���[�R�[�h�A�I�v�V����
Function jumpErrorP(WinFlag,dispId,wkID,wkName,errorCd,etc)
	on error goto 0		'''�G���[�g���b�v����
	Session.Contents("WinFlag")=WinFlag
	Session.Contents("dispId")=dispId
	Session.Contents("wkID")=wkID
	Session.Contents("wkName")=wkName
	Session.Contents("errorCd")=errorCd
	Session.Contents("etc")=etc
	Response.Redirect "./error.asp"
End Function


'''�G���[��ʂւ̑J��
'''Input :DB�ڑ�Object�A���R�[�h�Z�b�gObject�A�G���[��ʃt���O�A���ID�A���ID�A��Ɩ��A�G���[�R�[�h�A�I�v�V����
Function jumpErrorPDB(ObjConn,WinFlag,dispId,wkID,wkName,errorCd,etc)
'''���[���o�b�N��DB�ؒf
	ObjConn.RollbackTrans
	ObjConn.Close
	Set ObjConn = Nothing
	on error goto 0		'�G���[�g���b�v����
	Session.Contents("WinFlag")=WinFlag
	Session.Contents("dispId")=dispId
	Session.Contents("wkID")=wkID
	Session.Contents("wkName")=wkName
	Session.Contents("errorCd")=errorCd
	Session.Contents("etc")=etc
	Response.Redirect "./error.asp"
End Function


'''�T�[�o�����̎擾
'''Input :Array(5)=(YYYY,MM,DD,hh,mm)
Function getDayTime(DayTime)
	Dim MyDate
	MyDate = Now
	DayTime = Array(Year(MyDate),Right("0"&Month(MyDate),2),Right(Date(),2),Right("0"&Hour(MyDate),2),Right("0"&Minute(MyDate),2))
End Function



'''�A���R���e�i�e�[�u���̃R���e�i���݃`�F�b�N
'''Input :DB�ڑ�Object�A���R�[�h�Z�b�gObject�A�R���e�i�^BL�ԍ��A��ƃt���O�A�߂�l
'''Output:ret(���݂���=>true�@���݂��Ȃ�=>false) 
Function checkImportCont(ObjConn, ObjRS,strNum,Flag,ret)
	Dim StrSQL
	If Flag = 1 Then
		StrSQL = "SELECT count(ContNo) AS num FROM ImportCont WHERE ContNo=" & strNum
	Else
		StrSQL = "SELECT count(ContNo) AS num FROM ImportCont WHERE BLNo=" & strNum
	End If
	ObjRS.Open StrSQL, ObjConn
	if err <> 0 then
		DisConnDBH ObjConn, ObjRS	'DB�ؒf
		jumpErrorP "1","b000","00","�A���R���e�i�e�[�u��","101","SQL�F<BR>"&StrSQL
	end if
	If ObjRS("num") <> 0 Then
		ret = true
	Else
		ret = false
	End If
	ObjRS.close
End Function


'''�A���R���e�i�e�[�u���̃R���e�i�����`�F�b�N
'''Input :DB�ڑ�Object�A���R�[�h�Z�b�gObject�A�R���e�i�^BL�ԍ��A��ƃt���O�A�߂�l
'''Output:ret(������=>true�@����=>false) 
Function checkImportContComp(ObjConn,ObjRS,strNum,Flag,ret)
	dim StrSQL
	If Flag <> 4 Then
		StrSQL = "SELECT count(ContNo) AS num FROM ImportCont WHERE ContNo=" & strNum 
	Else
		StrSQL = "SELECT count(ContNo) AS num FROM ImportCont WHERE BLNo=" & strNum 
	End If
	StrSQL = StrSQL & " AND CYDelTime Is Null "
	ObjRS.Open StrSQL, ObjConn
	if err <> 0 then
		DisConnDBH ObjConn, ObjRS	'DB�ؒf
		jumpErrorP "1","b000","00","�A���R���e�i�e�[�u�������`�F�b�N","101","SQL�F<BR>"&StrSQL
	end if
	If ObjRS("num") <> 0 Then
		ret = true
	Else
		ret = false
	End If
	ObjRS.close
End Function


'''INI�t�@�C������p�����[�^��Ǎ���
'''Input :Array(1)
'''OutPut:Array(0)=ALARM�AArray(1)=COMPLETED_DATA
Function getIni(param)
	Dim ObjFSO,ObjTS,tmpStr
	Set ObjFSO = Server.CreateObject("Scripting.FileSystemObject")
	'--- �t�@�C�����J���i�ǂݎ���p�j ---
	Set ObjTS = ObjFSO.OpenTextFile(Server.Mappath("./INI/PREDEF.INI"),1,false)
	'--- �t�@�C���f�[�^�̓Ǎ��� ---
	Do Until ObjTS.AtEndofStream
		tmpStr = Split(ObjTS.ReadLine, "=", 3, 1)
		Select Case tmpStr(0)
			Case "ALARM"
				param(0) = tmpStr(1)
			Case "COMPLETED_DATA"
				param(1) = tmpStr(1)
			Case "INPUT_DATA"
				param(2) = tmpStr(1)
		End Select
	Loop
	ObjTS.Close
	Set ObjTS = Nothing
	Set ObjFSO = Nothing
End Function


'''�f�[�^��n�����\�b�h
'''Input :�f�[�^���A�f�[�^�e�[�u��
Function Set_Data(Num,DtTbl)
	Dim i,j
	Response.Write "  <INPUT type=hidden name='DataNum' value='" & Num & "'>" & vbCrLf
	For i = 0 to Num
		Response.Write "  <INPUT type=hidden name='Datatbl" & i & "' value='"
		For j=0 To UBound(DtTbl(i))-1
			Response.Write DtTbl(i)(j) & ","
		Next
		Response.Write DtTbl(i)(j) & "'>" & vbCrLf
	Next
End Function

'''�f�[�^���胁�\�b�h
'''Input :�f�[�^���A�f�[�^�e�[�u��
Function Get_Data(Num,DtTbl)
	Dim param
	Num=Request("DataNum")
	ReDim DtTbl(Num)
	For Each param In Request.Form
		If Left(param, 7) = "Datatbl" Then
			DtTbl(Mid(param,8))=Split(Request.Form(param), ",", -1, 1)
		End If
	Next
End Function

''' Added 20040329
Function DispDateTime(vDate,length)
	Dim strTmp
	If VarType(vDate)<>vbDate Then
		strTmp = ""
	Else
		strTmp = _
		Year(vDate) & "/" & _
		Right("0" & Month(vDate), 2) & "/" & _
		Right("0" & Day(vDate), 2) & " " & _
		Right("0" & Hour(vDate), 2) & ":" & _
		Right("0" & Minute(vDate), 2)
		If length=11 Then
			strTmp = Right(strTmp, 11)
		ElseIf length=10 Then
			strTmp = Left(strTmp, 10)
		ElseIf length=8 Then
			strTmp = Mid(strTmp, 3, 8)
		ElseIf length=5 Then
			strTmp = Mid(strTmp, 6, 5)
		End If
		DispDateTime = strTmp
	End If
End Function
%>