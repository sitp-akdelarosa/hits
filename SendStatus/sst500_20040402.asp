<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:sst500.asp				_/
'_/	Function	:�X�e�[�^�X�z�Mmail�������M			_/
'_/	Date			:2004/01/07				_/
'_/	Code By		:aspLand HARA			_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'''HTTP�R���e���c�^�C�v�ݒ�
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->
<%
	'''Microsoft ADO�p��adovbs.inc�ɂĒ񋟂���Ă���
	Const adBoolean = 11
	Const adDBTimeStamp = 135
	Const adInteger = 3
	Const adChar = 129
	Const adParamInput = &H0001
	Const adParamReturnValue = &H0004

	'''�Z�b�V�����̗L�������`�F�b�N
	CheckLoginH

	'''���M���܂�����ʂɂāu�ŐV�̏��ɍX�V�v��Submit���ꂽ�ꍇ�̑΍�
	if Session.Contents("SendMailSubmitted") = "False" then

		'''�f�[�^�擾
		Dim USER, KIND, NUMBER, ErrCode, NewDelMode
		Dim Email1, Email2, Email3, Email4, Email5
		Dim UserName

		USER = Session.Contents("userid")
		KIND = Request.Form("ContORBL")
		NUMBER = Request.Form("ContBLNo")
		NewDelMode = Request.Form("Mode")
		ErrCode = 0

		'''�T�[�o���t�̎擾
		Dim DayTime
		getDayTime DayTime

		'''DB�ڑ�
		Dim ObjConn, ObjRS, StrSQL
		ConnDBH ObjConn, ObjRS

		'''�w��R���e�i�ԍ�,�a�k�ԍ��̑��݃`�F�b�N
		Dim Num
		if KIND = 1 then '''�R���e�i�ԍ��w��
			StrSQL = "SELECT count(ContNo) AS CNUM FROM ImportCont WHERE ContNo='" & NUMBER & "'"
		elseif KIND = 2 then	'''�a�k�ԍ��w��̏ꍇ�B
			StrSQL = "SELECT count(BLNo) AS CNUM FROM BL WHERE BLNo='"& NUMBER & "'"
		else
			response.write("KIND error!")
			response.end
		end if
		ObjRS.Open StrSQL, ObjConn
		if err <> 0 then
			'''DB�ؒf
			DisConnDBH ObjConn, ObjRS
			jumpErrorP "2","c104","01","�X�e�[�^�X�z�Mmail�������M","101","SQL:<BR>"&strSQL
		end if
		Num = ObjRS("CNUM")
		ObjRS.close

		'''�w�肳�ꂽ�R���e�i�ԍ��܂��͂a�k�ԍ������݂���ꍇ
		if Num > 0 then
			'''�X�e�[�^�X�z�M�惁�[���A�h���X�ƃ��O�C�����[�U���̒��o
			StrSQL = "SELECT TI.Email1, TI.Email2, TI.Email3, TI.Email4, TI.Email5, MU.FullName "
			StrSQL = StrSQL & " FROM TargetItems TI, mUsers MU "
			StrSQL = StrSQL & " WHERE TI.UserCode='" & USER & "' AND TI.UserCode=MU.UserCode "
			ObjRS.Open StrSQL, ObjConn
			if err <> 0 then
				'''DB�ؒf
				DisConnDBH ObjConn, ObjRS
				jumpErrorP "2","c104","01","�X�e�[�^�X�z�Mmail�������M","101","SQL:<BR>"&strSQL
			end if
			if ObjRS.EOF then		'''���O�C�����[�U�p�̃X�e�[�^�X�z�M���ڒ�`���R�[�h�����݂��Ȃ��ꍇ
				ObjRS.close
				ErrCode = 1
			else	'''���O�C�����[�U�p�̃X�e�[�^�X�z�M���ڒ�`���R�[�h�����݂���ꍇ
				Email1=Trim(ObjRS("Email1"))
				Email2=Trim(ObjRS("Email2"))
				Email3=Trim(ObjRS("Email3"))
				Email4=Trim(ObjRS("Email4"))
				Email5=Trim(ObjRS("Email5"))
				UserName=Trim(ObjRS("FullName"))
				ObjRS.close

				if IsNull(Email1) and IsNull(Email2) and IsNull(Email3) and IsNull(Email4) and IsNull(Email5) then
				'''���O�C�����[�U�p�̃X�e�[�^�X�z�M���ڒ�`���R�[�h�����݂��邪�A���[���A�h���X���P���o�^����Ă��Ȃ��ꍇ
					ErrCode = 2

				'''�P�ł����[���A�h���X�̒�`�����݂���ꍇ�A���[�����M�ΏۂƂȂ�R���e�i�����ׂĒ��o����B
				else
					Dim ETA, TA, InTime, ListDate, DOStatus, DelPermitDate, FreeTime, FreeTimeExt
					Dim CYDelTime, DetentionFreeTime, ReturnTime
					Dim OLTICFlag, OLTICNo, OLTDateFrom, OLTDateTo
					Dim ContainerNumber, RcdNum, i
					Dim VslCode, VoyCtrl
					Dim sp, p0, p1, p2, p3, p4

					if KIND = 1 then		'''�R���e�i�ԍ��w��̏ꍇ
						StrSQL = "SELECT VslCode, VoyCtrl FROM ImportCont "
						StrSQL = StrSQL & " WHERE ContNo='"& NUMBER &"'"
						StrSQL = StrSQL & " AND UpdtTime = (SELECT max(UpdtTime) FROM ImportCont WHERE ContNo='"& NUMBER &"') "

						ObjRS.Open StrSQL, ObjConn, 3, 1
						if err <> 0 then
							'''DB�ؒf
							DisConnDBH ObjConn, ObjRS
							jumpErrorP "2","c104","01","�X�e�[�^�X�z�Mmail�������M","101","SQL:<BR>"&strSQL
						end if
						ReDim ContainerNumber(1), VslCode(1), VoyCtrl(1)
						ContainerNumber(0) = NUMBER
						VslCode(0) = ObjRS("VslCode")
						VoyCtrl(0) = ObjRS("VoyCtrl")
						RcdNum = 1
						ObjRS.close

					elseif KIND = 2 then		'''�a�k�ԍ��w��̏ꍇ�A�ΏۃR���e�i�ԍ������ׂĎ��o��
						StrSQL = "SELECT IC.VslCode, IC.VoyCtrl, IC.ContNo FROM BL, ImportCont IC "
						StrSQL = StrSQL & " WHERE BL.BLNo='"& NUMBER &"'"
						StrSQL = StrSQL & " AND BL.VslCode = IC.VslCode "
						StrSQL = StrSQL & " AND BL.VoyCtrl = IC.VoyCtrl "
						StrSQL = StrSQL & " AND BL.BLNo = IC.BLNo "
						StrSQL = StrSQL & " AND BL.UpdtTime = (SELECT max(BL.UpdtTime) FROM BL WHERE BL.BLNo='"& NUMBER &"') "

						ObjRS.Open StrSQL, ObjConn, 3, 1
						if err <> 0 then
							'''DB�ؒf
							DisConnDBH ObjConn, ObjRS
							jumpErrorP "2","c104","01","�X�e�[�^�X�z�Mmail�������M","101","SQL:<BR>"&strSQL
						end if
						RcdNum = ObjRS.RecordCount
						ReDim ContainerNumber(RcdNum), VslCode(RcdNum), VoyCtrl(RcdNum)
						for i=0 to RcdNum-1
							ContainerNumber(i) = ObjRS("ContNo")
							VslCode(i) = ObjRS("VslCode")
							VoyCtrl(i) = ObjRS("VoyCtrl")
							ObjRS.MoveNext
						next
						ObjRS.close
					end if

					Dim svName, mailTo, mailFrom, attachedFiles, ObjMail
					Dim mailFlag1, mailFlag2, mailFlag3, mailFlag4

					'''SMTP�T�[�o���̐ݒ�
					svName   = "slitdns2.hits-h.com"
					attachedFiles = ""
					mailFlag1 = 0
					mailFlag2 = 0
					mailFlag3 = 0
					mailFlag4 = 0
					'''���[�����M���A�h���X�̐ݒ�
					mailFrom = "mrhits@hits-h.com"
					mailTo = ""

					if IsNull(Email1) = false then
						mailTo = mailTo & Email1
						mailFlag1 = 1
					else
						mailFlag1 = 0
					end if

					if IsNull(Email2) = false then
						if mailFlag1 = 1 then
							mailTo = mailTo & vbtab & Email2
						else
							mailTo = mailTo & Email2
						end if
						mailFlag2 = 1
					else
						mailFlag2 = 0
					end if

					if IsNull(Email3) = false then
						if mailFlag1 = 1 or mailFlag2 = 1 then
							mailTo = mailTo & vbtab & Email3
						else
							mailTo = mailTo & Email3
						end if
						mailFlag3 = 1
					else
						mailFlag3 = 0
					end if

					if IsNull(Email4) = false then
						if mailFlag1 = 1 or mailFlag2 = 1 or mailFlag3 = 1 then
							mailTo = mailTo & vbtab & Email4
						else
							mailTo = mailTo & Email4
						end if
						mailFlag4 = 1
					else
						mailFlag4 = 0
					end if

					if IsNull(Email5) = false then
						if mailFlag1 = 1  or mailFlag2 = 1 or mailFlag3 = 1 or mailFlag4 = 1 then
							mailTo = mailTo & vbtab & Email5
						else
							mailTo = mailTo & Email5
						end if
					end if

					Dim rc, fp, fobj, tfile, sendTime
					Set ObjMail = Server.CreateObject("BASP21")

					Dim S_Flag

					'''�e�p�����[�^�̊i�[�p�z��̐錾
					ReDim ETA(RcdNum), TA(RcdNum), InTime(RcdNum), ListDate(RcdNum), DOStatus(RcdNum)
					ReDim DelPermitDate(RcdNum), FreeTime(RcdNum), FreeTimeExt(RcdNum)
					ReDim CYDelTime(RcdNum), DetentionFreeTime(RcdNum), ReturnTime(RcdNum)
					ReDim OLTICFlag(RcdNum), OLTICNo(RcdNum), OLTDateFrom(RcdNum), OLTDateTo(RcdNum)
					ReDim rc(RcdNum), sendTime(RcdNum)

					'''���o�۔���p�X�g�A�[�h�v���V�W���̌Ăяo���̂��߂̐ݒ�
					set sp = Server.CreateObject("ADODB.Command")
					set sp.ActiveConnection = ObjConn
					sp.CommandText = "{?=call DelPermitCheck(?,?,?)}"
					Set p0 = sp.CreateParameter("ret", adBoolean, adParamReturnValue)
					sp.Parameters.Append p0
					Set p1 = sp.CreateParameter("VslCode", adChar, adParamInput, 7)
					sp.Parameters.Append p1
					Set p2 = sp.CreateParameter("VoyCtrl", adInteger, adParamInput)
					sp.Parameters.Append p2
					Set p3 = sp.CreateParameter("ContNo", adChar, adParamInput, 12)
					sp.Parameters.Append p3

					'''���o�����R���e�i�̐��������[�v�����āA�R���e�i���ɏ�Ԃ����[�����M����B
					for i=0 to RcdNum-1
						StrSQL = "SELECT VP.ETA, VP.TA, IC.InTime, CT.ListDate, IC.DOStatus, IC.DelPermitDate, IC.FreeTime, "
						StrSQL = StrSQL & " IC.FreeTimeExt, IC.CYDelTime, IC.DetentionFreeTime, IC.ReturnTime, "
						StrSQL = StrSQL & " IC.OLTICFlag, IC.OLTICNo, IC.OLTDateFrom, IC.OLTDateTo "
						StrSQL = StrSQL & " FROM VslPort VP, ImportCont IC, Container CT "
						StrSQL = StrSQL & " WHERE IC.ContNo='" & ContainerNumber(i) & "'"
						StrSQL = StrSQL & " AND VP.PortCode='JPHKT' "
						StrSQL = StrSQL & " AND IC.VslCode=VP.VslCode "
						StrSQL = StrSQL & " AND IC.VoyCtrl=VP.VoyCtrl "
						StrSQL = StrSQL & " AND CT.ContNo=IC.ContNo "
						StrSQL = StrSQL & " AND IC.VslCode=CT.VslCode "
						StrSQL = StrSQL & " AND IC.VoyCtrl=CT.VoyCtrl "

						ObjRS.Open StrSQL, ObjConn
						if err <> 0 then
							'''DB�ؒf
							DisConnDBH ObjConn, ObjRS
							jumpErrorP "2","c104","01","�X�e�[�^�X�z�Mmail�������M","101","SQL:<BR>"&strSQL  & i
						end if

						ETA(i)=ObjRS("ETA")
						TA(i)=ObjRS("TA")
						InTime(i)=ObjRS("InTime")
						ListDate(i)=ObjRS("ListDate")
						DOStatus(i)=ObjRS("DOStatus")
						DelPermitDate(i)=ObjRS("DelPermitDate")
						FreeTime(i)=ObjRS("FreeTime")
						FreeTimeExt(i)=ObjRS("FreeTimeExt")
						CYDelTime(i)=ObjRS("CYDelTime")
						DetentionFreeTime(i)=ObjRS("DetentionFreeTime")
						ReturnTime(i)=ObjRS("ReturnTime")
						OLTICFlag(i)=ObjRS("OLTICFlag")
						OLTICNo(i)=ObjRS("OLTICNo")
						OLTDateFrom(i)=ObjRS("OLTDateFrom")
						OLTDateTo(i)=ObjRS("OLTDateTo")
						ObjRS.close

						Dim mailSubject, mailBody
						'''���[���^�C�g���̐ݒ�
						if KIND = 1 then
							mailSubject = "�A���X�e�[�^�X�̂��m�点(�R���e�i�ԍ��F" & ContainerNumber(i) & ")"
						elseif KIND = 2 then
							mailSubject = "�A���X�e�[�^�X�̂��m�点(�a�k�ԍ��F" & NUMBER & ")"
						end if

						'''���[���{���̍쐬
						mailBody = ""
						mailBody = UserName & " �a" & vbCrLf & vbCrLf
						mailBody = mailBody & "�A���X�e�[�^�X�̂��m�点�@�@�@" & DayTime(0) & "�N" & DayTime(1) & "��" & DayTime(2) & "��" & DayTime(3) & "������"  & vbCrLf & vbCrLf
						mailBody = mailBody & "���ΏۃR���e�i" & vbCrLf
						mailBody = mailBody & "�@" & ContainerNumber(i) & vbCrLf & vbCrLf
						mailBody = mailBody & "���X�e�[�^�X" & vbCrLf

						mailBody = mailBody & "�@(1)���`����" & vbCrLf
						if IsNull(ETA(i)) = false then
							if Hour(ETA(i)) = 0 and Minute(ETA(i)) = 0 and Second(ETA(i)) = 0 then
								mailBody = mailBody & "�@�@�\��E�E�E" & Year(ETA(i)) & "�N" & Right("0"&Month(ETA(i)),2) & "��" & Right("0"&Day(ETA(i)),2) & "��" & vbCrLf
							else
								mailBody = mailBody & "�@�@�\��E�E�E" & Year(ETA(i)) & "�N" & Right("0"&Month(ETA(i)),2) & "��" & Right("0"&Day(ETA(i)),2) & "�� " & Right("0"&Hour(ETA(i)),2) & ":" & Right("0"&Minute(ETA(i)),2) & vbCrLf
							end if
						elseif IsNull(TA(i)) = false then
							if Hour(TA(i)) = 0 and Minute(TA(i)) = 0 and Second(TA(i)) = 0 then
								mailBody = mailBody & "�@�@�����E�E�E" & Year(TA(i)) & "�N" & Right("0"&Month(TA(i)),2) & "��" & Right("0"&Day(TA(i)),2) & "��" & vbCrLf
							else
								mailBody = mailBody & "�@�@�����E�E�E" & Year(TA(i)) & "�N" & Right("0"&Month(TA(i)),2) & "��" & Right("0"&Day(TA(i)),2) & "�� " & Right("0"&Hour(TA(i)),2) & ":" & Right("0"&Minute(TA(i)),2) & vbCrLf
							end if
						else
							mailBody = mailBody & vbCrLf
						end if
						mailBody = mailBody & vbCrLf

						mailBody = mailBody & "�@(2)�b�x��������" & vbCrLf
						if IsNull(InTime(i)) = false then
							mailBody = mailBody & "�@�@" & Year(InTime(i)) & "�N" & Right("0"&Month(InTime(i)),2) & "��" & Right("0"&Day(InTime(i)),2) & "�� " & Right("0"&Hour(InTime(i)),2) & ":" & Right("0"&Minute(InTime(i)),2) & vbCrLf
						else
							mailBody = mailBody & vbCrLf
						end if
						mailBody = mailBody & vbCrLf

						mailBody = mailBody & "�@(3)�ʊ֋���" & vbCrLf
						if IsNull(ListDate(i)) = false then
							mailBody = mailBody & "�@�@���@�ʊ֋���=" & Year(ListDate(i)) & "�N" & Right("0"&Month(ListDate(i)),2) & "��" & Right("0"&Day(ListDate(i)),2) & "��" & vbCrLf
						else
							mailBody = mailBody & "�@�@�~" & vbCrLf
						end if
						mailBody = mailBody & vbCrLf

						mailBody = mailBody & "�@(4)�c�n�N���A��" & vbCrLf
						if DOStatus(i) = "Y" then
							mailBody = mailBody & "�@�@��" & vbCrLf
						else
							mailBody = mailBody & "�@�@�~" & vbCrLf
						end if
						mailBody = mailBody & vbCrLf

						'''���o�۔���
						mailBody = mailBody & "�@(5)���o��" & vbCrLf
						'''�b�x���o����Ă���ꍇ�́u�ρv�𑗐M����  Modified 20040312
						if IsNull(CYDelTime(i)) = false then
							mailBody = mailBody & "�@�@��" & vbCrLf
						else
						'''ImportCont�e�[�u����VslCode, VoyCtrl, ContNo��������BLNo�������قȂ郌�R�[�h�����݂���ꍇ�A
						'''���Y���R�[�h�ɂ��Ă��������N���A�ł��Ă��邩�`�F�b�N����B
							sp("VslCode") = VslCode(i)
							sp("VoyCtrl") = VoyCtrl(i)
							sp("ContNo") = ContainerNumber(i)
							'''�X�g�A�[�h�v���V�W���̌Ăяo��
							sp.Execute
							'''�X�g�A�[�h�v���V�W���̌Ăяo�����ʂ̔���
							if sp("ret") = True then 
								mailBody = mailBody & "�@�@���@���o�\��=" & Year(DelPermitDate(i)) & "�N" & Right("0"&Month(DelPermitDate(i)),2) & "��" & Right("0"&Day(DelPermitDate(i)),2) & "��" & vbCrLf
							else
								mailBody = mailBody & "�@�@�~" & vbCrLf
							end if
						end if
						mailBody = mailBody & vbCrLf

						''''''���Ɖ����̕\��������̂�FreeTimeExt�܂���FreeTime��mail�������M���s����菫���̏ꍇ�Ƃ��Ă���
						mailBody = mailBody & "�@(6)�f�}���[�W�t���[�^�C��" & vbCrLf
						if IsNull(FreeTimeExt(i)) = false then
							if FreeTimeExt(i) > Date then
								mailBody = mailBody & "�@�@" & Year(FreeTimeExt(i)) & "�N" & Right("0"&Month(FreeTimeExt(i)),2) & "��" & Right("0"&Day(FreeTimeExt(i)),2) & "���@����" & DateDiff("d",Date,FreeTimeExt(i)) & "��" & vbCrLf
							else
								mailBody = mailBody & "�@�@" & Year(FreeTimeExt(i)) & "�N" & Right("0"&Month(FreeTimeExt(i)),2) & "��" & Right("0"&Day(FreeTimeExt(i)),2) & "��" & vbCrLf
							end if
						elseif IsNull(FreeTime(i)) = false then
							if FreeTime(i) > Date then
								mailBody = mailBody & "�@�@" & Year(FreeTime(i)) & "�N" & Right("0"&Month(FreeTime(i)),2) & "��" & Right("0"&Day(FreeTime(i)),2) & "���@����" & DateDiff("d",Date,FreeTime(i)) & "��" & vbCrLf
							else
								mailBody = mailBody & "�@�@" & Year(FreeTime(i)) & "�N" & Right("0"&Month(FreeTime(i)),2) & "��" & Right("0"&Day(FreeTime(i)),2) & "��" & vbCrLf
							end if
						else
							mailBody = mailBody & vbCrLf
						end if
						mailBody = mailBody & vbCrLf

						mailBody = mailBody & "�@(7)�b�x���o����" & vbCrLf
						if IsNull(CYDelTime(i)) = false then
							mailBody = mailBody & "�@�@" & Year(CYDelTime(i)) & "�N" & Right("0"&Month(CYDelTime(i)),2) & "��" & Right("0"&Day(CYDelTime(i)),2) & "�� " & Right("0"&Hour(CYDelTime(i)),2) & ":" & Right("0"&Minute(CYDelTime(i)),2) & vbCrLf
						else
							mailBody = mailBody & vbCrLf
						end if
						mailBody = mailBody & vbCrLf

						'''���Ɖ����̕\��������̂̓f�B�e���V�����t���[�^�C���������ƂȂ�ꍇ�Ƃ��Ă���B
						'''�܂��ADetentionFreeTime�Ɂu0�v���ݒ肳��Ă���ꍇ�A���Ȃ킿�ԋp�\������Ƃ���
						'''�u�����́v�u�T���ȏ�v�܂��́u���X�g�I�t�v���w�肳��Ă���ꍇ�A���Ɖ����̕\���͂��Ȃ��B
						mailBody = mailBody & "�@(8)�f�B�e���V�����t���[�^�C��" & vbCrLf
						if not IsNull(DetentionFreeTime(i)) and not IsNull(CYDelTime(i)) then
							if DateAdd("d",DetentionFreeTime(i),DateValue(CYDelTime(i)))>Date then
								mailBody = mailBody & "�@�@���o������" & Trim(DetentionFreeTime(i)) & "���ȓ��@����" & DateDiff("d",Date,DateAdd("d",DetentionFreeTime(i),DateValue(CYDelTime(i)))) & "��" & vbCrLf
							else
								mailBody = mailBody & "�@�@���o������" & Trim(DetentionFreeTime(i)) & "���ȓ�" & vbCrLf
							end if
						else
							mailBody = mailBody & vbCrLf
						end if
						mailBody = mailBody & vbCrLf

						mailBody = mailBody & "�@(9)��R���ԋp��" & vbCrLf
						if IsNull(ReturnTime(i)) = false then
							mailBody = mailBody & "�@�@���@��R���ԋp����=" & Year(ReturnTime(i)) & "�N" & Right("0"&Month(ReturnTime(i)),2) & "��" & Right("0"&Day(ReturnTime(i)),2) & "�� " & Right("0"&Hour(ReturnTime(i)),2) & ":" & Right("0"&Minute(ReturnTime(i)),2) & vbCrLf
						else
							mailBody = mailBody & "�@�@�~" & vbCrLf
						end if

						'''���[�����M����

						rc(i)=ObjMail.Sendmail(svName, mailTo, mailFrom, mailSubject, mailBody, attachedFiles)
						sendTime(i)=Now
					Next

					for i=0 to RcdNum-1
						if rc(i)="" then
							S_Flag = 0
						else
							S_Flag = 1
							exit for
						end if
					next

					if S_Flag = 0 then		'''���[�����M����
						'''�폜��ʂ���mail�������M��������ꍇTargetContainers�e�[�u���̍ŏI���M�������X�V����B
						'''�V�K�o�^��ʂ���mail�������M��������ꍇ�͑Ώۃ��R�[�h���܂�insert����Ă��Ȃ��̂ōŏI���M�����̍X�V�͕s�v�B
						if NewDelMode = 2 then
							StrSQL = "UPDATE TargetContainers SET UpdtTime='" & Now() & "', UpdtPgCd='STATUS01',"
							StrSQL =  StrSQL & " UpdtTmnl='" & USER & "', LatestSentTime='" & Now() & "'"
							if KIND = 1 then		'''�Ώۂ��R���e�i�ԍ�
								StrSQL =  StrSQL & " WHERE ContNo='" & NUMBER & "' AND UserCode='" & USER & "'"
							elseif KIND = 2 then		'''�Ώۂ��a�k�ԍ�
								StrSQL =  StrSQL & " WHERE BLNo='" & NUMBER & "' AND UserCode='" & USER & "'"
							end if
							StrSQL =  StrSQL & " AND Process='R' OR Process='N'"
						end if
						ObjConn.Execute(StrSQL)
						if err <> 0 then
							Set ObjRS = Nothing
							jumpErrorPDB ObjConn,"1","c104","14","�X�e�[�^�X�z�Mmail�������M","104","SQL:<BR>"&StrSQL
						end if

						'''���O�o��
						WriteLogH "c104", "�X�e�[�^�X�z�Mmail�������M","01",""
						ErrCode = 0

					else		'''���[�����M���s
						fp = Server.MapPath("./mailerror") & "\error.txt"
						set fobj = Server.CreateObject("Scripting.FileSystemObject")

						for i=0 to RcdNum-1
							if rc(i)<>"" then
								if fobj.FileExists(fp) = True then
									set tfile = fobj.OpenTextFile(fp,8)
								else
									set tfile = fobj.CreateTextFile(fp,True,False)
								end if
								tfile.WriteLine sendTime(i) & " " & rc(i)
								tfile.Close
								ErrCode = 8
							end if
						next

					end if		'''���[�����M�����A���s�����̏I���
				end if		'''���[���A�h���X���P�ł���`����Ă���ꍇ�̏����̏I���
			end if		'''�X�e�[�^�X�z�M���ڂ���`����Ă���ꍇ�̏����̏I���
		else		'''�w�肳�ꂽ�R���e�i�ԍ��A�a�k�ԍ������݂��Ȃ�
			ErrCode = 9
		end if

		'''DB�ڑ�����
		DisConnDBH ObjConn, ObjRS
		'''�G���[�g���b�v����
		on error goto 0

		Session.Contents("SendMailSubmitted") = "True"


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>�X�e�[�^�X�z�Mmail�������M</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
function CloseWin(){
	try{
		window.opener.parent.DList.location.href="sst100L.asp"
	}catch(e){}
	window.close();
}
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin="0" topMargin="0" marginheight="0" marginwidth="0">
<!-------------�X�e�[�^�X�z�Mmail�������M���ʉ��--------------------------->
<TABLE border="0" cellPadding="3" cellSpacing="1" width="100%">
<FORM name="sst500" method="POST">
	<TR><TD>�@</TD></TR>
<% if ErrCode=0 then %>
	<TR>
		<TD align="center">
			���[�����M���܂����B<BR>
		</TD>
	</TR>
	<TR><TD>�@</TD></TR>
	<TR>
		<TD align="center">
			<INPUT type="button" value="����" onClick="CloseWin()">
		</TD>
	</TR>
<% elseif ErrCode=1 or ErrCode=2 then %>
	<TR>
		<TD align="center">
			���[�����M�悪�ݒ肳��Ă��܂���B<BR>�u�ݒ�v���j���[�ɂă��[���A�h���X��o�^���Ă��������B<BR>
		</TD>
	</TR>
	<TR><TD>�@</TD></TR>
	<TR>
		<TD align="center">
			<INPUT type="button" value="����" onClick="window.close()">
		</TD>
	</TR>
<% elseif ErrCode=8 then %>
	<TR>
		<TD align="center">
			���[�����M�Ɏ��s���܂����B<BR>
		</TD>
	</TR>
	<TR><TD>�@</TD></TR>
	<TR>
		<TD align="center">
			<INPUT type="button" value="����" onClick="window.close()">
		</TD>
	</TR>
<% elseif ErrCode=9 then %>
	<TR>
		<TD align="center">
			�w�肳�ꂽ�R���e�i�ԍ��܂��͂a�k�ԍ��͑��݂��܂���B<BR>
		</TD>
	</TR>
	<TR><TD>�@</TD></TR>
	<TR>
		<TD align="center">
			<INPUT type="button" value="����" onClick="window.close()">
		</TD>
	</TR>
<% end if %>
</FORM>
</TABLE>
<!-------------��ʏI���--------------------------->
</BODY>
</HTML>

<%''' if Session.Contents("SendMailSubmitted") = "False"��else���� %>
<% else %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>�X�e�[�^�X�z�Mmail�������M</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
function CloseWin(){
	try{
		window.opener.parent.DList.location.href="sst100L.asp"
	}catch(e){}
	window.close();
}
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin="0" topMargin="0" marginheight="0" marginwidth="0">
<!-------------�X�e�[�^�X�z�Mmail�������M���ʉ��--------------------------->
<TABLE border="0" cellPadding="3" cellSpacing="1" width="100%">
<FORM name="sst500" method="POST">
	<TR><TD>�@</TD></TR>
	<TR>
		<TD align="center">
			�����͊��Ɋ������Ă��܂��B<BR><BR><BR>
		</TD>
	</TR>
	<TR>
		<TD align="center">
			<INPUT type="button" value="����" onClick="CloseWin()">
		</TD>
	</TR>
</FORM>
</TABLE>
<!-------------��ʏI���--------------------------->
</BODY>
</HTML>
<%'''if Session.Contents("SendMailSubmitted") = "False"��endif���� %>
<% end if %>
