<% 
'過去３ヶ月の日付の取得
Function GetBefore3Month(sDate,sMonthStart,sDate1,sDate2,sDate3,sDate4)
	'sDate		[ I ]現在日付文字列（YYYYMMDD形式）
	'sMonthStart	[ I ]月度開始日付文字列（DD形式）
	'sDate1		[ O ]当月文字列（YYYYMM形式）
	'sDate2		[ O ]１ヶ月前文字列（YYYYMM形式）
	'sDate3		[ O ]２ヶ月前文字列（YYYYMM形式）
	'sDate4		[ O ]３ヶ月前文字列（YYYYMM形式）
	'戻り値		[ O ]なし

	dim dWkdate1, dWkYear1, dWkMonth1
	dim dWkdate2, dWkYear2, dWkMonth2
	dim dWkdate3, dWkYear3, dWkMonth3
	dim dWkdate4, dWkYear4, dWkMonth4

	if int(sMonthStart) <> 1 then		'開始日が１でない時
        if int(day(sdate)) >= int(sMonthStart) then
		 	dWkdate1 = DateAdd("m",1,sDate)	'月+1を当月とする
        else
		 	dWkdate1 = sDate		
        end if
    else
		dWkdate1 = sDate			
    end if

'当月
	dWkYear1 = Year(dWkdate1)
	dWkMonth1 = Month(dWkdate1)
	sDate1 = dWkYear1 & dWkMonth1
'１ヶ月前
	dWkdate2 = DateAdd("m",-1,dWkdate1)	'
	dWkYear2 = Year(dWkdate2)
	dWkMonth2 = Month(dWkdate2)
	sDate2 = dWkYear2 & dWkMonth2
'２ヶ月前
	dWkdate3 = DateAdd("m",-2,dWkdate1)	'
	dWkYear3 = Year(dWkdate3)
	dWkMonth3 = Month(dWkdate3)
	sDate3 = dWkYear3 & dWkMonth3
'３ヶ月前
	dWkdate4 = DateAdd("m",-3,dWkdate1)	'
	dWkYear4 = Year(dWkdate4)
	dWkMonth4 = Month(dWkdate4)
	sDate4 = dWkYear4 & dWkMonth4

End Function

'シャーシ利用回数の取得
Function GetStartEnd(conn, rsd, sGrpID, sWorkDate, iMonthStar,  sStartDate, sEndDate)
	'conn		[ I ]DB接続Object
	'rsd		[I/O]レコードセットObject
	'sGrpID		[ I ]グループＩＤ
	'sWorkDate	[ I ]選択年月(YYYYMM)
	'iMonthStar	[ I ]月度開始日(dd)
	'戻り値		[ O ]なし

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
		'再変換
			sStartDate	=	mid(sStartDate,1,4) & mid(sStartDate,6,2) & mid(sStartDate,9,2)
			sEndDate	=	mid(sEndDate,1,4) & mid(sEndDate,6,2) & mid(sEndDate,9,2)
		case else
			sMonth		=	mid(sWorkDate,1,4) & "/"  & mid(sWorkDate,5,2)	'当月取得
			sMonth		=	DateAdd("m",-1,cDate(sMonth))					'前月取得

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
