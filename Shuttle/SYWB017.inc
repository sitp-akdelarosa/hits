<% 
'�ߋ��R�����̓��t�̎擾
Function GetBefore3Month(sDate,sMonthStart,sDate1,sDate2,sDate3,sDate4)
	'sDate		[ I ]���ݓ��t������iYYYYMMDD�`���j
	'sMonthStart	[ I ]���x�J�n���t������iDD�`���j
	'sDate1		[ O ]����������iYYYYMM�`���j
	'sDate2		[ O ]�P�����O������iYYYYMM�`���j
	'sDate3		[ O ]�Q�����O������iYYYYMM�`���j
	'sDate4		[ O ]�R�����O������iYYYYMM�`���j
	'�߂�l		[ O ]�Ȃ�

	dim dWkdate1, dWkYear1, dWkMonth1
	dim dWkdate2, dWkYear2, dWkMonth2
	dim dWkdate3, dWkYear3, dWkMonth3
	dim dWkdate4, dWkYear4, dWkMonth4

	if int(sMonthStart) <> 1 then		'�J�n�����P�łȂ���
        if int(day(sdate)) >= int(sMonthStart) then
		 	dWkdate1 = DateAdd("m",1,sDate)	'��+1�𓖌��Ƃ���
        else
		 	dWkdate1 = sDate		
        end if
    else
		dWkdate1 = sDate			
    end if

'����
	dWkYear1 = Year(dWkdate1)
	dWkMonth1 = Month(dWkdate1)
	sDate1 = dWkYear1 & dWkMonth1
'�P�����O
	dWkdate2 = DateAdd("m",-1,dWkdate1)	'
	dWkYear2 = Year(dWkdate2)
	dWkMonth2 = Month(dWkdate2)
	sDate2 = dWkYear2 & dWkMonth2
'�Q�����O
	dWkdate3 = DateAdd("m",-2,dWkdate1)	'
	dWkYear3 = Year(dWkdate3)
	dWkMonth3 = Month(dWkdate3)
	sDate3 = dWkYear3 & dWkMonth3
'�R�����O
	dWkdate4 = DateAdd("m",-3,dWkdate1)	'
	dWkYear4 = Year(dWkdate4)
	dWkMonth4 = Month(dWkdate4)
	sDate4 = dWkYear4 & dWkMonth4

End Function

'�V���[�V���p�񐔂̎擾
Function GetStartEnd(conn, rsd, sGrpID, sWorkDate, iMonthStar,  sStartDate, sEndDate)
	'conn		[ I ]DB�ڑ�Object
	'rsd		[I/O]���R�[�h�Z�b�gObject
	'sGrpID		[ I ]�O���[�v�h�c
	'sWorkDate	[ I ]�I��N��(YYYYMM)
	'iMonthStar	[ I ]���x�J�n��(dd)
	'�߂�l		[ O ]�Ȃ�

	Dim sql
	Dim sMonth, sEndDay
	Dim sWkDate

	sStartDate = ""
	sEndDate = ""
	select case int(iMonthStar)
		case	1
			sWkDate		=	mid(sWorkDate,1,4) & "/"  & mid(sWorkDate,5,2) & "/" &  "01" 
'D20040204 S
'			sStartDate	=	"20" & DateAdd("d",0,cDate(sWkDate))
'			sEndDate	=	"20" & DateAdd("d",-1,DateAdd("m",1,cDate(sStartDate)))
'D20040204 E
'I20040204 S
			sStartDate	=	DateAdd("d",0,cDate(sWkDate))
			sEndDate	=	DateAdd("d",-1,DateAdd("m",1,cDate(sStartDate)))
'I20040204 E
		'�ĕϊ�
			sStartDate	=	mid(sStartDate,1,4) & mid(sStartDate,6,2) & mid(sStartDate,9,2)
			sEndDate	=	mid(sEndDate,1,4) & mid(sEndDate,6,2) & mid(sEndDate,9,2)
		case else
			sMonth		=	mid(sWorkDate,1,4) & "/"  & mid(sWorkDate,5,2)	'�����擾
			sMonth		=	DateAdd("m",-1,cDate(sMonth))					'�O���擾

			if len(iMonthStar)	=	1 then
				iMonthStar = "0" & 	iMonthStar
			end if
			sStartDate	=   "20" & mid(sMonth,1,2) & mid(sMonth,4,2) & iMonthStar	'D20040212
			sStartDate	=   mid(sMonth,1,2) & mid(sMonth,4,2) & iMonthStar			'I20040212
			sEndDay		=	iMonthStar - 1
			if len(sEndDay) = 1 then
				sEndDay	=	"0" & sEndDay
			end if

			if len(Trim(sWorkDate)) = 5 then
				sWorkDate = left(sWorkDate,4) & "0" & Right(sWorkDate,1)
			end if
			sEndDate	=	sWorkDate & sEndDay
	end select

	sql = "SELECT * FROM sSYUseCount" & _
		  " WHERE RTRIM(GroupID) = '" & sGrpID & "'" & _
		  "   AND RTRIM(RecDelDate) >= '" & sStartDate & "'" & _
		  "   AND RTRIM(RecDelDate) <= '" & sEndDate & "'" & _
		  " Order By RecDelDate, AmPm"
	rsd.Open sql, conn, 0, 1, 1

End Function

%>