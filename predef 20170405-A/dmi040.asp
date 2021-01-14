<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi040.asp				_/
'_/	Function	:事前実搬出作業番号通知画面		_/
'_/	Date		:2003/05/26				_/
'_/	Code By		:SEIKO Electric.Co 大重			_/
'_/	Modify		:C-002	2003/07/29	備考欄追加	_/
'_/	Modify		:3th	2003/01/31	3次変更	_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTPコンテンツタイプ設定
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->
<%
'セッションの有効性をチェック
  CheckLoginH

'サーバ日付の取得
  dim DayTime, YY,Yotei
  getDayTime DayTime

'ユーザデータ所得
  dim USER, sUN, Utype
  USER   = UCase(Session.Contents("userid"))
  sUN    = Session.Contents("sUN")
  Utype  = Session.Contents("UType")

'データ取得
  dim SakuNo,Flag,Num,CONnumA(),BLnum,CMPcd(5),Rmon,Rday,Rnissu
  dim CONsize,CONtype,CONhite,CONtear,HedId,HFrom,Hto
  dim Rhou,Nomon,Noday,Nohou,Nomin,NonyuDate						'3th add
  dim param,i,j,Way,Mord,WkContrlNo,Rval,RnissuA
  dim UpFlag,strNum,ret,ErrerM
  dim SendUser
  ret = true
  SakuNo = Request("SakuNo")
  Flag= Request("flag")
  UpFlag = Request("UpFlag")
  Num = Request("num")
  ReDim CONnumA(Num)
  i=0
  For Each param In Request.Form
    If Left(param, 6) = "CONnum" Then
      If param <> "CONnum" Then
        i = Mid(param,7)
'CW-308        CONnumA(i) = "'" & Request.Form(param) & "'"
        CONnumA(i) = Request.Form(param)
      Else
        CONnumA(0) = Request.Form(param)
      End If
    ElseIf Left(param,5) = "CMPcd" Then
      j = Right(param,1)
      CMPcd(j) = Request.Form(param)
    End If
  Next
'3th add start
  dim tmpCMPcd,tmpCONnum
  tmpCMPcd=CMPcd
  tmpCONnum=CONnumA
'3th add end
  CONtype = Request("CONtype")
  CONhite = Request("CONhite")
  CONtear = Request("CONtear")
  HedId   = Request("HedId")
  HFrom   = Request("HFrom")
  Hto     = Request("HTo")
  Rmon    = Right("00" & Request("Rmon") ,2)
  Rday    = Right("00" & Request("Rday") ,2)
'3th add start
  Rhou    = Right("00" & Request("Rhou") ,2)
  Nomon   = Right("00" & Request("Nomon") ,2)
  Noday   = Right("00" & Request("Noday") ,2)
  Nohou   = Right("00" & Request("Nohou") ,2)
  '2008/01/31 Add S G.Ariola
  Nomin   = Right("00" & Request("Nomin") ,2)
  '2008/01/31 Add S G.Ariola
'3th add end
  Rnissu  = Request("Rnissu")

  Way   =Array("","指定あり","指定なし","一覧から選択","ＢＬ番号")

'エラートラップ開始
  on error resume next
'DB接続
  dim ObjConn, ObjRS, StrSQL
  ConnDBH ObjConn, ObjRS

'CW-014 add strart
  '輸入コンテナテーブル搬出チェック
  If Flag=4 Then
    strNum="'"& Request("BLnum") &"'"
  Else
    strNum="'"& Request("CONnum") &"'"
  End If
  checkImportContComp ObjConn, ObjRS,strNum, Flag, ret
  If ret Then
'CW-014 add end

'データ整形
    dim FullName,RFlag
    RFlag=0
    BLnum   = Request("BLnum")
'CW-308    CONnumA(0) = "'" & CONnumA(0) &"'"
    For i = 0 to Num -1
      CONnumA(i) = "'" & CONnumA(i) &"'"
    Next

    If Flag = "1" Then
      BLnum = "Null"
      CONsize = "'" & Request("CONsize") &"'"
    Else
      BLnum = "'" & BLnum & "'"
      CONsize = "Null"
      If Flag = "4" Then
        CONnumA(0) = "Null"
      End If
    End If

   '元請陸運業者名取得
    FullName= "Null"
    If UpFlag<2 Then
'CW-040      If CMPcd(0) <> "" Then
      If CMPcd(1) <> "" Then
        StrSQL = "SELECT FullName FROM mUsers WHERE mUsers.HeadCompanyCode='" & CMPcd(1) &"'"
        ObjRS.Open StrSQL, ObjConn
        FullName = "'" & ObjRS("FullName") & "'"
        ObjRS.close
      End If
      if err <> 0 then
        DisConnDBH ObjConn, ObjRS	'DB切断
        jampErrerP "1","b10"&(2+Flag),"03","実搬出：データ登録","102","元請陸運業者名取得に失敗<BR>"&StrSQL
      end if
    End If
    If HedId = "" Then
      HedId   = "Null"
    Else
      HedId = "'" & HedId & "'"
    End If

    For i=1 To 4
      If CMPcd(i) = "" Then
        CMPcd(i) = "Null"
      Else
        If CMPcd(i) = Session.Contents("COMPcd") Then
          RFlag=1
        End If
        CMPcd(i) = "'" & CMPcd(i) & "'"
      End If
    Next

    RnissuA = Array("未入力","当日","2 日後","3 日後","4 日後","5 日後","5 日以上","リフトオフ")
    Rval = 0
    For i=0 To 6
      IF RnissuA(i) = Rnissu Then
        Rval= i
      End If
    Next
    '作業予定日の年度を決定
    If DayTime(1) > Rmon Then	'来年
      YY = DayTime(0) +1
    ElseIf DayTime(1) = Rmon AND DayTime(2) > Rday Then	'CW-043
      YY = DayTime(0) +1				'CW-043
    Else
      YY = DayTime(0)
    End If
    If Rmon = "00" Or Rday = "00" Then
      Yotei= "Null"
    Else
'3th chage      Yotei= "'" & YY &"/"& Rmon &"'"
      Yotei= "'" & YY &"/"& Rmon &"/"& Rday &" "& Rhou &":00'"
    End If

'3th add Start
    If DayTime(1) > Nomon Then	'来年
      YY = DayTime(0) +1
    ElseIf DayTime(1) = Nomon AND DayTime(2) > Noday Then
      YY = DayTime(0) +1
    Else
      YY = DayTime(0)
    End If
    If Nomon = "00" Or Noday = "00" Then
      NonyuDate= "Null"
    Else
	'2008/01/31 Edit S G.Ariola
	'NonyuDate= "'" & YY &"/"& Nomon &"/"& Noday &" "& Nohou &":00'"
      NonyuDate= "'" & YY &"/"& Nomon &"/"& Noday &" "& Nohou &":"& Nomin &"'"
	'2008/01/31 Edit E G.Ariola
    End If
'3th add End

    If SakuNo = "" Then '初期登録
      WriteLogH "b10"&(2+Flag), "実搬出事前情報一覧("&Way(Flag)&")","03",""
      Mord = 0
    '登録重複チェック
      If Flag=4 Then
        strNum= BLnum
      Else
        strNum= CONnumA(0)
      End If
      checkComInfo  ObjConn, ObjRS,strNum,"1", Flag, SakuNo, ret
      If ret Then
       '港運コード取得
        dim OpeCode
'CW-041        If Flag =1 Then
        OpeCode="Null"
'CW-041        Else
        If Flag <>1 Then								'CW-041 
          StrSQL = "SELECT Count(BL.OpeCode) AS Num FROM BL WHERE BL.BLNo="& BLnum	'CW-041
          ObjRS.Open StrSQL, ObjConn							'CW-041

          If ObjRS("Num") <> 0 Then							'CW-041
            ObjRS.close									'CW-041
            StrSQL = "SELECT BL.OpeCode FROM BL WHERE BL.BLNo="& BLnum
            ObjRS.Open StrSQL, ObjConn
            OpeCode = Trim(ObjRS("OpeCode"))
            OpeCode = "'" & OpeCode & "'"
          End If				'CW-041
          ObjRS.close
        End If
        
        if err <> 0 then
          DisConnDBH ObjConn, ObjRS	'DB切断
          jampErrerP "1","b10"&(2+Flag),"03","実搬出：データ登録","102","港運コード取得に失敗<BR>"&StrSQL
        end if
      '作業管理番号採番
        getWkContrlNo ObjConn, ObjRS, sUN, WkContrlNo


      '作業番号採番
'3th Change Start
'3th        StrSQL = "SELECT Count(WkNo) AS Num FROM hITWkNo where Status='3'"	'CW-042
'3th        ObjRS.Open StrSQL, ObjConn						'CW-042
'3th        If ObjRS("Num") <> 0 Then
'3th          ObjRS.close
'3th          StrSQL = "SELECT WkNo FROM hITWkNo where Status='3'"
'3th          ObjRS.Open StrSQL, ObjConn
'3th          SakuNo=ObjRS("WkNo")
'3th          ObjRS.close
'3th          StrSQL = "UPDATE hITWkNo SET Status='2' WHERE WkNo ='"& SakuNo &"'"
'3th          ObjConn.Execute(StrSQL)
'3th          if err <> 0 then
'3th            Set ObjRS = Nothing
'3th            jampErrerPDB ObjConn, "1","b10"&(2+Flag),"03","実搬出：データ登録","104","作業番号取得に失敗<BR>"&StrSQL
'3th          end if
'3th        Else
'3th          ObjRS.close
'3th'CW-042        If err <> 0 then
'3th'CW-042          err=0
'3th          StrSQL = "SELECT CurrentVal FROM mAutoNumber WHERE TypeCode='11'"
'3th          ObjRS.Open StrSQL, ObjConn
'3th          if err <> 0 then
'3th            ObjRS.Close
'3th            Set ObjRS = Nothing
'3th            jampErrerPDB ObjConn,"1","b10"&(2+Flag),"03","実搬出：データ登録","102","作業番号取得に失敗<BR>"&StrSQL
'3th          end if
'3th          SakuNo = ObjRS("CurrentVal")+1
'3th          ObjRS.close
'3th          StrSQL = "UPDATE mAutoNumber SET CurrentVal = "& SakuNo &", UpdtTime='"& now() &"',"&_
'3th                   "UpdtPgCd='PREDEF01',UpdtTmnl='"& USER &"' WHERE TypeCode='11'"
'3th          ObjConn.Execute(StrSQL)
'3th          if err <> 0 then
'3th            Set ObjRS = Nothing
'3th            jampErrerPDB ObjConn,"1","b10"&(2+Flag),"03","実搬出：データ登録","104","作業番号取得に失敗<BR>"&StrSQL
'3th          end if
'3th          SakuNo = Right("0000" & Hex(SakuNo),5)
'3th          StrSQL = "Insert Into hITWkNo (WkNo,UpdtTime,UpdtPgCd,UpdtTmnl,Status) values ('" &_
'3th                    SakuNo &"','"& Now() &"','PREDEF01','"& USER &"','2')"
'3th          ObjConn.Execute(StrSQL)
'3th          if err <> 0 then
'3th            Set ObjRS = Nothing
'3th            jampErrerPDB ObjConn,"1","b10"&(2+Flag),"03","実搬出：データ登録","103","作業番号取得に失敗<BR>"&StrSQL
'3th          end if
'3th        End If
        getWkNo ObjConn, ObjRS, USER, SakuNo
        if err <> 0 then
          Set ObjRS = Nothing
          jampErrerPDB ObjConn,"1","b10"&(2+Flag),"03","実搬出：データ登録","103","作業番号取得に失敗<BR>"&StrSQL
        end if
'3th Change End
    'データ登録
        StrSQL = "Insert Into hITCommonInfo (WkContrlNo,UpdtTime,UpdtPgCd,UpdtTmnl,Status," &_
                 "Process,WkType,FullOutType,InPutDate,UpdtUserCode,WkNo,ContNo,BLNo,OpeCode,ContSize," &_
                 "RegisterType,RegisterName,RegisterCode,TruckerSubCode1," &_
                 "HeadID,WorkDate,TruckerName,DeliverTo,ReturnDateStr," &_
                 "ReturnDateVal,Comment1,Comment2,GoodsName,DeliverTo1,DeliverTo2,DeliverDate, "&_
                 "TruckerSubName1)"&_
                 "values ('"& WkContrlNo &"','"& Now() &"','PREDEF01','"& USER &"',"&_
                 "'0','R','1','"& Flag &"','"& Now() &"','"& USER &"','"& SakuNo &"',"& CONnumA(0) &","&_
                  BLnum &","& OpeCode&","& CONsize &",'"& Utype &"','"& sUN &"','"& CMPcd(0) &"',"& CMPcd(1) &","&_
                  HedId &","& Yotei &","& FullName &",'"& Hto &"','"&Rnissu &"','"& Rval &"'"&_
                  ",'"& Request("Comment1") &"','"& Request("Comment2") &"'"&_
                  ",'"& Request("HinName") &"','"& Request("Nonyu1") &"','"& Request("Nonyu2") &"',"&NonyuDate &_
                  ",'"& Request("TruckerSubName") & "')"
		SendUser = CMPcd(1)
'C-002 ADD These Lines : ,Comment1,Comment2,Comment3
'                      :,'"& Request("Comment1") &"','"& Request("Comment2") &"','"& Request("Comment3") &"'
'3th del Comment3
'3th add GoodsName,DeliverTo1,DeliverTo2,DeliverDate
'3th add ,'"& Request("HinName") &"','"& Request("Nonyu1") &"','"& Request("Nonyu2") &"',"&NonyuDate&"
        ObjConn.Execute(StrSQL)
        if err <> 0 then
          Set ObjRS = Nothing
          jampErrerPDB ObjConn,"1","b10"&(2+Flag),"03","実搬出：データ登録","103","SQL:<BR>"&StrSQL
        end if
    '照会テーブル登録
        StrSQL = "Insert Into hITReference (WkContrlNo, UpdtTime, UpdtPgCd,UpdtTmnl," &_
                 "TruckerFlag1,TruckerFlag2,TruckerFlag3,TruckerFlag4)" &_
                 "values ('"& WkContrlNo &"','"& Now() &"','PREDEF01','"& USER &"'," &_
                 "'"&RFlag&"','0','0','0')"
        ObjConn.Execute(StrSQL)
        if err <> 0 then
          Set ObjRS = Nothing
          jampErrerPDB ObjConn,"1","b10"&(2+Flag),"03","実搬出：データ登録","104","SQL:<BR>"&StrSQL
        end if
        If Flag = 3 Then
          For i=0 To Num-1
            StrSQL = "Insert Into hITFullOutSelect (WkContrlNo,ContNo,UpdtTime,UpdtPgCd,UpdtTmnl) " &_
                   "values ('"& WkContrlNo &"',"& CONnumA(i) &",'"& Now() &"','PREDEF01','"& USER &"')"
            ObjConn.Execute(StrSQL)
            if err <> 0 then
              Set ObjRS = Nothing
              jampErrerPDB ObjConn,"1","b10"&(2+Flag),"03","実搬出：データ登録","104","SQL:<BR>"&StrSQL
            end if
          Next
        End If
      Else
'Chang 20050303 STAT fro Fourth Recon By SEIKO N.Oosige
'        ErrerM="指定の作業は画面操作中に他者によって作業番号「" & SakuNo & "」で登録されました。"
        ErrerM="指定の作業は画面操作中に他者によって登録されました。"
'Chang 20050303 END
      End If
    Else                '更新
      Mord = Request("Mord")
      WriteLogH "b10"&(2+Flag), "実搬出事前情報一覧("&Way(Flag)&")","14",""
'CW-004	ADD START ↓↓↓↓↓↓↓
     '完了・更新チェック
      If UpFlag <>5 Then
        StrSQL="SELECT ITC.WorkCompleteDate, ITR.TruckerFlag"& UpFlag &" AS Flag "&_
               "FROM hITCommonInfo AS ITC INNER JOIN hITReference AS ITR ON ITC.WkContrlNo = ITR.WkContrlNo "&_
               "Where WkNo='"& SakuNo &"' AND Process='R' AND WkType='1'"
      Else
        StrSQL="SELECT WorkCompleteDate FROM hITCommonInfo " &_
               "Where WkNo='"& SakuNo &"' AND Process='R' AND WkType='1'"
      End If
      ObjRS.Open StrSQL, ObjConn
      if err <> 0 then
        ObjRS.Close
        Set ObjRS = Nothing
        jampErrerPDB ObjConn,"1","b10"&(2+Flag),"14","実搬出：データ登録","101","SQL:<BR>"&StrSQL
      end if
      If NOT IsNull(ObjRS("WorkCompleteDate")) Then 
        ret=false
        ErrerM="指定の作業は画面操作中に作業が完了したため、更新はキャンセルされました。"
      End If
     'チェック
      If UpFlag <>5 Then
        If Trim(ObjRS("Flag"))=1 Then 
          ret=false
          ErrerM="指定の作業は画面操作中に指示先に受諾されたため、更新はキャンセルされました。"
        End If
      End If
      ObjRS.close
      If ret Then
'CW-004	End ADD ↑↑↑↑↑↑↑
        If Mord <> 2 Then
        'データ更新
          dim tmpStr
          If FullName <> "Null" Then
            FullName=",TruckerName="& FullName &" "
          Else
            FullName=""
          End If
          If UpFlag = 5 Then
            tmpStr = " "
          Else
            tmpStr=" TruckerSubCode"& UpFlag &"="& CMPcd(UpFlag) &","
            SendUser = CMPcd(UpFlag)
          End If

          tmpStr = tmpStr & " TruckerSubName"& (UpFlag) & "='" & Request("TruckerSubName") &"',"

          StrSQL = "UPDATE hITCommonInfo SET UpdtTime='"& Now() &"', UpdtPgCd='PREDEF01'," &_
                   "UpdtTmnl='"& USER &"', Status='0', Process='R', " &_
                   "UpdtUserCode='"& USER &"', "& tmpStr &_
                   "HeadID="& HedId &", WorkDate="& Yotei &", DeliverTo='"& Hto &"', " &_
                   "Comment1='"& Request("Comment1") &"',Comment2='"& Request("Comment2") &"'," &_
                   "ReturnDateStr='"& Rnissu &"', ReturnDateVal='"& Rval &"' "& FullName &"," &_
                   "GoodsName='"& Request("HinName") &"',DeliverTo1='"& Request("Nonyu1") &"'," &_ 
                   "DeliverTo2='"& Request("Nonyu2") &"',DeliverDate="&NonyuDate&" " &_
                   "Where WkNo='"& SakuNo &"' AND Process='R' AND WkType='1'"
                   
'C-002 ADD This Line : "Comment1='"& Request("Comment1") &"',Comment2='"& Request("Comment2") &"',Comment3='"& Request("Comment3") &"', "&_
'3th del Comment3
'3th add "GoodsName='"& Request("HinName") &"',DeliverTo1='"& Request("Nonyu1") &"',"
'3th add "DeliverTo2='"& Request("Nonyu2") &"',DeliverDate="&NonyuDate&"
          ObjConn.Execute(StrSQL)
          if err <> 0 then
            Set ObjRS = Nothing
            jampErrerPDB ObjConn,"1","b10"&(2+Flag),"14","実搬出：データ登録","104","SQL:<BR>"&StrSQL
          end if
          If UpFlag = 5 Then
            tmpStr = " "
          Else
            If UpFlag = 1 AND Mid(CMPcd(1),2,2) = UCase(Session.Contents("COMPcd")) Then 
              tmpStr = ", TruckerFlag1=1 "
            Else
              tmpStr = ", TruckerFlag"& UpFlag &"=0 "
            End If
          End If
          UpFlag = UpFlag-1
          If UpFlag = 0 Then
            tmpStr = tmpStr&" "
          Else
            tmpStr = tmpStr&", TruckerFlag"& UpFlag &"=1 "
          End If
       '参照フラグ更新
          StrSQL = "UPDATE hITReference SET UpdtTime='"& Now() &"', UpdtPgCd='PREDEF01'," &_
                   "UpdtTmnl='"& USER &"'"&tmpStr&_
                   "WHERE WkContrlNo IN (select WkContrlNo From hITCommonInfo "&_
                   "WHERE WkNo='"& SakuNo &"' AND Process='R' AND WkType='1')"
          ObjConn.Execute(StrSQL)
          if err <> 0 then
            Set ObjRS = Nothing
            jampErrerPDB ObjConn,"1","b10"&(2+Flag),"14","実搬出：データ登録","104","SQL:<BR>"&StrSQL
          end if
        Else
        'ヘッダID更新
          If UpFlag=5 Then
            tmpStr=""
          Else
            tmpStr=", TruckerSubCode"& UpFlag &"=Null"
          End If
          StrSQL = "UPDATE hITCommonInfo SET UpdtTime='"& Now() &"', UpdtPgCd='PREDEF01'," &_
                   "UpdtTmnl='"& USER &"', Status='0', Process='R', " &_
                   "UpdtUserCode='"& USER &"'"& tmpStr &", HeadID=Null " &_
                   "Where WkNo='"& SakuNo &"' AND Process='R' AND WkType='1'"
          ObjConn.Execute(StrSQL)
          if err <> 0 then
            Set ObjRS = Nothing
            jampErrerPDB ObjConn,"1","b10"&(2+Flag),"15","実搬出：保留","104","SQL:<BR>"&StrSQL
          end if

         '参照フラグ更新
          StrSQL = "UPDATE hITReference SET UpdtTime='"& Now() &"', UpdtPgCd='PREDEF01'," &_
                   "UpdtTmnl='"& USER &"', TruckerFlag"& UpFlag-1 &"=2 "&_
                   "WHERE WkContrlNo IN (select WkContrlNo From hITCommonInfo "&_
                   "WHERE WkNo='"& SakuNo &"' AND Process='R' AND WkType='1')"
          ObjConn.Execute(StrSQL)
          if err <> 0 then
            Set ObjRS = Nothing
            jampErrerPDB ObjConn,"1","b10"&(2+Flag),"15","実搬出：保留","104","SQL:<BR>"&StrSQL
          end if
        End If
      End If			'CW-004
    End If
'CW-014 add start

' 2009/03/10 R.Shibuta Add-S
'データ取得
	Dim Email1, Email2, Email3, Email4, Email5
	Dim UserName,ComInterval,rc

	'''通信間隔取得
	StrSQL = "SELECT ComInterval FROM mParam WHERE Seq = '1'"

	ObjRS.Open StrSQL, ObjConn
	if err <> 0 then
	'''DB切断
		DisConnDBH ObjConn, ObjRS
		jampErrerPDB ObjConn,"1","b10"&(2+Flag),"16","実搬出：メール送信","104","SQL:<BR>"&StrSQL
	end if

	ComInterval = ObjRS("ComInterval")
	ObjRS.Close

	if SendUser <> "" then
	''作業発生配信情報の取得
		StrSQL = "SELECT T.*, "
		StrSQL = StrSQL & "CASE WHEN U.NameAbrev IS NULL THEN U.FullName ELSE U.NameAbrev END AS USERNAME "
		StrSQL = StrSQL & "FROM mUsers U, "
		StrSQL = StrSQL & "(SELECT T.* FROM TargetOperation T, mUsers U WHERE T.UserCode = U.UserCode "
		StrSQL = StrSQL & "AND U.HeadCompanyCode =" & SendUser & ") T "
		StrSQL = StrSQL & "WHERE U.UserCode = '" & USER & "'"
		
		ObjRS.Open StrSQL, ObjConn
		if err <> 0 then
	'''DB切断
			DisConnDBH ObjConn, ObjRS
			jampErrerPDB ObjConn,"1","b10"&(2+Flag),"16","実搬出：メール送信","104","SQL:<BR>"&StrSQL
		end if

		Dim svName, mailTo, mailFrom, attachedFiles, ObjMail
		Dim mailFlag1, mailFlag2, mailFlag3, mailFlag4, mailFlag5
		Dim mailSubject, mailBody,WorkName
		Dim SendTime, UpdateSendTime
		Dim fp, fobj, tfile

	'''SMTPサーバ名の設定
		svName   = "slitdns2.hits-h.com"
		attachedFiles = ""
		mailFlag1 = 0
		mailFlag2 = 0
		mailFlag3 = 0
		mailFlag4 = 0
		mailFlag5 = 0
	'''メール送信元アドレスの設定
		mailFrom = "mrhits@hits-h.com"
		mailTo = ""
		rc = ""

		if Trim(ObjRS("Email1")) <> "" AND ObjRS("FlagDelResults1") = "1" then
			mailTo = mailTo & Trim(ObjRS("Email1"))
			mailFlag1 = 1
		else
			mailFlag1 = 0
		end if

		if Trim(ObjRS("Email2")) <> "" AND ObjRS("FlagDelResults2") = "1" then
			if mailFlag1 = 1 then
				mailTo = mailTo & vbtab & Trim(ObjRS("Email2"))
			else
				mailTo = mailTo & Trim(ObjRS("Email2"))
			end if
				mailFlag2 = 1
		else
			mailFlag2 = 0
		end if

		if Trim(ObjRS("Email3")) <> "" AND ObjRS("FlagDelResults3") = "1" then
			if mailFlag1 = 1 or mailFlag2 = 1 then
				mailTo = mailTo & vbtab & Trim(ObjRS("Email3"))
			else
				mailTo = mailTo & Trim(ObjRS("Email3"))
			end if
			mailFlag3 = 1
		else
			mailFlag3 = 0
		end if

		if Trim(ObjRS("Email4")) <> "" AND ObjRS("FlagDelResults4") = "1" then
			if mailFlag1 = 1 or mailFlag2 = 1 or mailFlag3 = 1 then
				mailTo = mailTo & vbtab & Trim(ObjRS("Email4"))
			else
				mailTo = mailTo & Trim(ObjRS("Email4"))
			end if
			mailFlag4 = 1
		else
			mailFlag4 = 0
		end if

		if Trim(ObjRS("Email5")) <> "" AND ObjRS("FlagDelResults5") = "1" then
			if mailFlag1 = 1 or mailFlag2 = 1 or mailFlag3 = 1 or mailFlag4 = 1 then
				mailTo = mailTo & vbtab & Trim(ObjRS("Email5"))
			else
				mailTo = mailTo & Trim(ObjRS("Email5"))
			end if
			mailFlag5 = 1
		else
			mailFlag5 = 0
		end if

		Set ObjMail = Server.CreateObject("BASP21")

		mailSubject = "HiTS 作業依頼"
		mailBody = "実搬出作業" & "発生 (" & Trim(ObjRS("USERNAME")) & "様より)" & vbCrLf & vbCrLf
		mailBody = mailBody & "実搬出作業" & "が発生しました。" & vbCrLf
		mailBody = mailBody & "詳しくはHiTSの事前情報登録の画面をご参照下さい。"
			
		'メール送信時刻から現在の時刻が通信間隔以上の場合はメールを送信する。

		if Trim(mailTo) <> "" Then
			if ObjRS("DelResultsDate") < DateAdd("n",(ComInterval * -1), Now()) OR IsNull(ObjRS("DelResultsDate")) = True then
				rc=ObjMail.Sendmail(svName, mailTo, mailFrom, mailSubject, mailBody, attachedFiles)
				sendTime=Now
			end if

			If rc = "" Then
				'''メール送信日付の更新を行う。
				StrSQL = "UPDATE TargetOperation SET UpdtTime='" & Now() & "', UpdtPgCd='dmi040',"
				StrSQL = StrSQL & " UpdtTmnl='" & USER & "',"&  "DelResultsDate='" & Now() & "'"
				StrSQL = StrSQL &"WHERE UserCode = '" & Trim(ObjRS("UserCode")) & "'"

				ObjConn.Execute(StrSQL)
				if err <> 0 then
					Set ObjRS = Nothing
					jumpErrorPDB ObjConn,"1","c104","14","実搬出：メール送信","104","SQL:<BR>"&StrSQL
				end if
			else
				fp = Server.MapPath("./mailerror") & "\error.txt"
				set fobj = Server.CreateObject("Scripting.FileSystemObject")
					if rc<>"" then
						if fobj.FileExists(fp) = True then
							set tfile = fobj.OpenTextFile(fp,8)
						else
							set tfile = fobj.CreateTextFile(fp,True,False)
						end if
						tfile.WriteLine sendTime & " " & rc
						tfile.Close
						ErrerM = "メール送信に失敗しました。<BR>"
						ret = 1
					end if
			end if
		else

		end if
	end if
' 2009/03/10 R.Shibuta Add-E
  Else
    ErrerM="指定のコンテナは画面操作中に搬出作業が完了しました。<BR>"&_
           "このため登録・更新処理はキャンセルされます。"
  End If
'CW-014 add end

'DB接続解除
  DisConnDBH ObjConn, ObjRS
'エラートラップ解除
  on error goto 0

'コンテナ番号受渡しメソッド
Sub Set_CONnum
  For i = 1 to Num -1
    Response.Write "       <INPUT type=hidden name='CONnum" & i & "' value='" & tmpCONnum(i) & "'>" & vbCrLf
  Next
End Sub

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>作業番号発行</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
function CloseWin(){
  try{
<% If Mord=0 Then %>
    window.opener.parent.List.location.href="./dmo010F.asp"
<% Else %>
    window.opener.parent.DList.location.href="./dmo010L.asp"
    window.opener.parent.Top.location.href="./dmo010T.asp"
<% End If %>
  }catch(e){
  }
  window.close();
}
//指示書印刷調整画面へ
function GoPrint(){
  window.resizeTo(500,700);
  target=document.dmi040F;
  target.submit();
}
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!-------------作業番号発行画面--------------------------->
<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%">
  <TR align=center valign=bottom height=50>
    <TD>
<% If ret Then
     'If Mord=0 Then '2010/06/07 M.Marquez Del
       Response.Write "      <B>作業番号発行</B></TD></TR>"&vbCrLf&"  <TR>"&vbCrLf
       Response.Write "    <TD>"&vbCrLf&"作業番号は「" & SakuNo & "」です。"
     '2010/06/07 M.Marquez Del-S
     'Else
     '  Response.Write "  <TD> 更新しました。<BR>画面は自動的に閉じられます。"
     '  Response.Write "    <SCRIPT language=JavaScript>"&vbCrLf&"      CloseWin();"&vbCrLf&"    </SCRIPT>"     
     'End If
     '2010/06/07 M.Marquez Del-E  
   Else
     Response.Write "      <DIV class=alert>"&ErreRM&"</DIV>"
   End If
%>
   </TD></TR>
  <TR><TD align=center valign=bottom height=100>
    <FORM name="dmi040F" action="./dmo091.asp" method="POST">
<%'2010/06/07 M.Marquez Upd-S
 'If ret AND Mord=0 Then 
 If ret Then 
 '2010/06/07 M.Marquez Upd-E%>
      <INPUT type=hidden name="UpFlag"  value="<%=Request("UpFlag")%>">
      <INPUT type=hidden name="CONnum"  value="<%=tmpCONnum(0)%>">
      <INPUT type=hidden name="BLnum"   value="<%=Request("BLnum")%>" >
      <INPUT type=hidden name="CONsize" value="<%=Request("CONsize")%>">
      <INPUT type=hidden name="CONtype" value="<%=Request("CONtype")%>">
      <INPUT type=hidden name="CONhite" value="<%=Request("CONhite")%>">
      <INPUT type=hidden name="CONtear" value="<%=Request("CONtear")%>">
      <INPUT type=hidden name="CMPcd0"  value="<%=tmpCMPcd(0)%>">
      <INPUT type=hidden name="CMPcd1"  value="<%=tmpCMPcd(1)%>">
      <INPUT type=hidden name="CMPcd2"  value="<%=tmpCMPcd(2)%>">
      <INPUT type=hidden name="CMPcd3"  value="<%=tmpCMPcd(3)%>">
      <INPUT type=hidden name="CMPcd4"  value="<%=tmpCMPcd(4)%>">
      <INPUT type=hidden name="Rmon"    value="<%=Request("Rmon")%>">
      <INPUT type=hidden name="Rday"    value="<%=Request("Rday")%>">
      <INPUT type=hidden name="Rnissu"  value="<%=Request("Rnissu")%>">
      <INPUT type=hidden name="HFrom"   value="<%=Request("HFrom")%>">
      <INPUT type=hidden name="flag"    value="<%=Request("flag")%>" >
      <INPUT type=hidden name="num"     value="<%=Request("num")%>" >
      <INPUT type=hidden name="Comment1" value="<%=Request("Comment1")%>" >
      <INPUT type=hidden name="Comment2" value="<%=Request("Comment2")%>" >
      <INPUT type=hidden name="Rhou"     value="<%=Request("Rhou")%>">
      <INPUT type=hidden name="shipFact" value="<%=Request("shipFact")%>" >
      <INPUT type=hidden name="shipName" value="<%=Request("shipName")%>" >
      <INPUT type=hidden name="HinName"  value="<%=Request("HinName")%>" >
      <INPUT type=hidden name="Nonyu1"   value="<%=Request("Nonyu1")%>" >
      <INPUT type=hidden name="Nonyu2"   value="<%=Request("Nonyu2")%>" >
      <INPUT type=hidden name="RPlace"   value="<%=Request("RPlace")%>" >
      <INPUT type=hidden name="Nomon"    value="<%=Request("Nomon")%>">
      <INPUT type=hidden name="Noday"    value="<%=Request("Noday")%>">
      <INPUT type=hidden name="Nohou"    value="<%=Request("Nohou")%>">
	  <!-- 2008/01/31 Add S G.Ariola -->
	  <INPUT type=hidden name="Nomin"    value="<%=Request("Nomin")%>">
	  <!-- 2008/01/31 Add E G.Ariola -->
      <INPUT type=hidden name="SakuNo"  value="<%=SakuNo%>">
      <INPUT type=hidden name="HedId"   value="<%=Request("HedId")%>">
      <INPUT type=hidden name="HTo"     value="<%=Request("HTo")%>">
  <% IF Num > 1 Then call Set_CONnum End If%>
      <INPUT type=button value="指示書印刷" onClick="GoPrint()">
<% End If %>
       <INPUT type=button value="閉じる" onClick="CloseWin()">
    </FORM>
    </TD>
  </TR>
</TABLE>
<!-------------画面終わり--------------------------->
</BODY></HTML>
