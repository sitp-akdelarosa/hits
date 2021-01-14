<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi015.asp				_/
'_/	Function	:実搬出入力情報取得			_/
'_/	Date		:2003/05/26				_/
'_/	Code By		:SEIKO Electric.Co 大重			_/
'_/	Modify		:C-002	2003/07/29	備考欄追加	_/
'_/	Modify		:3th	2003/01/31	3次変更	_/
'_/	Modify		:T.Okui 2017/03/07	データ取得条件変更 
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

'データ取得
  dim hCd,sUN,Utype,User,CONnum,Flag,BLnum,SakuNo,Mord
  dim Num,CONsize,CONtype,CONhite,CONtear,CMPcd,HFrom,CONnumA,i
  dim strNum, UpFlag,ret1,ret2,ret3
  dim targetNo,HinName,shipFact,shipName,Nonyu1,Nonyu2,NonyuDate(3),RPlace,Rhou '3th add
  dim TruckerSubName,TruckerName
'  NonyuDate = Array("","","")
  CONnum = Request("CONnum")
  Flag   = Request("flag")
'3th del  SakuNo = Request("SakuNo")
  targetNo = Request("targetNo")
  BLnum  = Request("BLnum")
  hCd    = Session.Contents("COMPcd")
  sUN    = Session.Contents("sUN")
  Utype  = Session.Contents("UType")
  User   = Session.Contents("userid")
  ret1   = true
  ret2   = true
  ret3   = true
  dim Comment1,Comment2,Comment3		'C-002 ADD

'del 3th'コンテナ詳細・一覧画面への遷移制御フラグの取得
'del 3th  If Request("InfoFlag") = "" Then
'del 3th    InfoFlag=0
'del 3th  Else
'del 3th    InfoFlag=Request("InfoFlag")
'del 3th  End If
'エラートラップ開始
  on error resume next
'DB接続
  dim ObjConn, ObjRS, StrSQL
  ConnDBH ObjConn, ObjRS

'遷移元振分け
'3th  If SakuNo = ""  Then
  If targetNo = ""  Then
    '初期登録
    Mord=0
    UpFlag=1
    '事前登録チェック
    If Flag=4 Then
      strNum="'"& BLnum &"'"
    Else
      strNum="'"& CONnum &"'"
    End If
    '輸入コンテナテーブルのコンテナ存在チェック
    checkImportCont ObjConn, ObjRS,strNum, Flag, ret1
'CW-014 add strart
    '輸入コンテナテーブル搬出チェック
    If ret1 Then
      checkImportContComp ObjConn, ObjRS,strNum, Flag, ret3
    End If
'CW-014 add end
    'IT共通テーブルの登録重複チェック
    checkComInfo ObjConn, ObjRS,strNum, "1",Flag, SakuNo, ret2

    dim tmpStr,Way
    Way   =Array("","指定あり","指定なし","一覧から選択","ＢＬ番号")
    If ret1 AND ret2 Then
      tmpStr=",入力内容の正誤:0(正しい)"
    Else
      tmpStr=",入力内容の正誤:1(誤り)"
    End If
    WriteLogH "b10"&(2+Flag), "実搬出事前情報入力("&Way(Flag)&")", "00",strNum&tmpStr
    If ret1 AND ret3 AND ret2 Then
    'コンテナデータ所得
      Select Case Flag
        Case "1"		'指定有り
          'DBよりデータ所得
'CW-050          StrSQL = "select ContType,ContSize,ContHeight,ContWeight,RecTerminal " &_
'CW-050                   "from Container INNER join BL on BL.VslCode=Container.VslCode " &_ 
'CW-050                   "and  BL.VoyCtrl=Container.VoyCtrl where Container.ContNo='" & CONnum &"'"
          StrSQL = "SELECT CON.ContType, CON.ContSize, CON.ContHeight, CON.ContWeight, BL.RecTerminal, "&_
                   "CON.ShipLine, mV.FullName, INC.ReturnPlace "&_
                   "FROM ((Container AS CON INNER JOIN ImportCont AS INC ON (CON.ContNo = INC.ContNo) "&_
                   "AND (CON.VoyCtrl = INC.VoyCtrl) AND (CON.VslCode = INC.VslCode)) "&_
                   "LEFT JOIN BL ON (INC.BLNo = BL.BLNo) AND (INC.VoyCtrl = BL.VoyCtrl) "&_
                   "AND (INC.VslCode = BL.VslCode)) "&_
                   "LEFT JOIN mVessel AS mV On (CON.VslCode=mV.VslCode) "&_
                   "WHERE CON.ContNo='" & CONnum &"' ORDER BY CON.UpdtTime DESC"
'3th add mV.ShipLine, mV.FullName,INC.ReturnPlace
'3th add LEFT JOIN mVessel AS mV On (INC.VslCode=mV.VslCode)
'CW-037 Chage mV.ShipLine -> CON.ShipLine

          ObjRS.Open StrSQL, ObjConn
          if err <> 0 then
            DisConnDBH ObjConn, ObjRS	'DB切断
            jampErrerP "1","b103","01","実搬出:データ取得","102","SQL:<BR>"&strSQL
          end if
          Num=1

        Case "2"		'指定なし
          'DBよりデータ所得
          '対象件数取得
          StrSQL = "SELECT Count(INC2.ContNo) AS NUM "&_
                   "FROM ImportCont AS INC1 INNER JOIN ImportCont AS INC2 ON "&_
                   "(INC1.VoyCtrl = INC2.VoyCtrl) AND (INC1.VslCode = INC2.VslCode) AND (INC1.BLNo = INC2.BLNo) "&_
                   "WHERE INC1.ContNo='" & CONnum & "'"
'CW-050          StrSQL = "SELECT Count(ImportCont_1.ContNo) AS NUM FROM ImportCont " &_
'CW-050                   "INNER JOIN ImportCont AS ImportCont_1 ON ImportCont.BLNo = ImportCont_1.BLNo " &_
'CW-050                   " WHERE ImportCont.ContNo='" & CONnum & "'"
          ObjRS.Open StrSQL, ObjConn
          if err <> 0 then
            DisConnDBH ObjConn, ObjRS	'DB切断
            jampErrerP "1","b104","01","実搬出:データ件数取得","101","SQL:<BR>"&strSQL
          end if
          Num=ObjRS("NUM")
          ObjRS.Close

          '対象コンテナ番号一覧取得
          StrSQL = "SELECT INC2.ContNo, INC1.BLNo "&_
                   "FROM ImportCont AS INC1 INNER JOIN ImportCont AS INC2 "&_
                   "ON (INC1.VoyCtrl = INC2.VoyCtrl) AND (INC1.VslCode = INC2.VslCode) AND (INC1.BLNo = INC2.BLNo) "&_
                   "WHERE INC1.ContNo='" & CONnum & "'"
'CW-050          StrSQL = "SELECT ImportCont_1.ContNo, ImportCont.BLNo FROM ImportCont " &_
'CW-050                   "INNER JOIN ImportCont AS ImportCont_1 ON ImportCont.BLNo = ImportCont_1.BLNo " &_
'CW-050                   " WHERE ImportCont.ContNo='" & CONnum & "' ORDER BY ImportCont.ContNo ASC"
          ObjRS.Open StrSQL, ObjConn
          if err <> 0 then
            DisConnDBH ObjConn, ObjRS	'DB切断
            jampErrerP "1","b104","01","実搬出:データ一覧取得","102","SQL:<BR>"&strSQL
          end if
'CW-050          BLnum=ObjRS("BLNo")
          ReDim CONnumA(Num)
          CONnumA(0)=CONnum
          i=1
          Do Until ObjRS.EOF
            If CONnum <> ObjRS("ContNo") Then 
              CONnumA(i)=ObjRS("ContNo")
              i=i+1
            End If
           ObjRS.MoveNext
          Loop
          ObjRS.Close
          '代表コンテナのデータ取得
          StrSQL = "SELECT CON.ContType, CON.ContSize, CON.ContHeight, CON.ContWeight, INC.BLNo, BL.RecTerminal, "&_
                   "BL.ShipLine, mV.FullName, INC.ReturnPlace "&_
                   "FROM ((Container AS CON INNER JOIN ImportCont AS INC ON (CON.ContNo = INC.ContNo) "&_
                   "AND (CON.VoyCtrl = INC.VoyCtrl) AND (CON.VslCode = INC.VslCode)) "&_
                   "LEFT JOIN BL ON (INC.BLNo = BL.BLNo) AND (INC.VoyCtrl = BL.VoyCtrl) "&_
                   "AND (INC.VslCode = BL.VslCode)) "&_
                   "LEFT JOIN mVessel AS mV On BL.VslCode=mV.VslCode "&_
                   "WHERE CON.ContNo='" & CONnum &"' ORDER BY BL.UpdtTime DESC"
'CW-050          StrSQL = "select ContType,ContSize,ContHeight,ContWeight,RecTerminal " &_
'CW-050                   "from Container INNER join BL on BL.VslCode=Container.VslCode " &_ 
'CW-050                   "and  BL.VoyCtrl=Container.VoyCtrl where Container.ContNo='" & CONnum &"'"
'3th add mV.ShipLine, mV.FullName,INC.ReturnPlace
'3th add LEFT JOIN mVessel AS mV On (INC.VslCode=mV.VslCode)
'CW307 Change mV.ShipLine -> BL.ShipLine
          ObjRS.Open StrSQL, ObjConn
          if err <> 0 then
            DisConnDBH ObjConn, ObjRS	'DB切断
            jampErrerP "1","b104","01","実搬出:データ取得","102","SQL:<BR>"&strSQL
          end if
          BLnum=ObjRS("BLNo")

        Case "3"		'一覧より
          'DBよりデータ所得
          '対象件数取得
          '2017/03/07 T.Okui Upd-S　　条件変更
          'StrSQL = "SELECT Count(IC2.ContNo) AS NUM FROM ImportCont AS IC1 " &_
          '         "INNER JOIN ImportCont AS IC2 ON (IC1.VoyCtrl = IC2.VoyCtrl) AND (IC1.VslCode = IC2.VslCode) "&_
          '         "AND (IC1.BLNo = IC2.BLNo) "&_
          '         "WHERE IC1.ContNo='" & CONnum & "' " &_
          '         "AND IC2.ContNo NOT IN (SELECT ITC.ContNo from hITCommonInfo AS ITC "&_
          '         "WHERE ITC.ContNo IS NOT Null AND ITC.Process='R') "&_
          '         "AND IC2.ContNo NOT IN (SELECT ITF.ContNo FROM hITFullOutSelect AS ITF "&_
          '         "INNER JOIN hITCommonInfo AS ITC2 ON ITF.WkContrlNo = ITC2.WkContrlNo "&_
          '         "WHERE ITC2.Process='R' AND WorkCompleteDate Is Null ) "
          StrSQL = "SELECT Count(IC2.ContNo) AS NUM FROM ImportCont AS IC1 " &_
                   "INNER JOIN ImportCont AS IC2 ON (IC1.VoyCtrl = IC2.VoyCtrl) AND (IC1.VslCode = IC2.VslCode) "&_
                   "AND (IC1.BLNo = IC2.BLNo) "&_
                   "WHERE IC1.ContNo='" & CONnum & "' " &_
                   "AND IC2.ContNo NOT IN (SELECT ITC.ContNo from hITCommonInfo AS ITC "&_
                   "WHERE ITC.ContNo IS NOT Null AND ITC.Process='R' AND WorkCompleteDate Is Null ) "&_
                   "AND IC2.ContNo NOT IN (SELECT ITF.ContNo FROM hITFullOutSelect AS ITF "&_
                   "INNER JOIN hITCommonInfo AS ITC2 ON ITF.WkContrlNo = ITC2.WkContrlNo "&_
                   "WHERE ITC2.Process='R' AND WorkCompleteDate Is Null ) "
           
'CW-050 "INNER JOIN ImportCont AS IC2 ON IC1.BLNo = IC2.BLNo " &_
'20030911 ADD This Item:"AND WorkCompleteDate Is Null " &_
          ObjRS.Open StrSQL, ObjConn
          if err <> 0 then
            DisConnDBH ObjConn, ObjRS	'DB切断
            jampErrerP "1","b105","01","実搬出:データ件数取得","101","SQL:<BR>"&strSQL
          end if
          Num=ObjRS("NUM")
          ObjRS.Close

          '対象コンテナ番号一覧取得
          'StrSQL = "SELECT IC2.ContNo, IC1.BLNo FROM ImportCont AS IC1 " &_
          '         "INNER JOIN ImportCont AS IC2 ON (IC1.VoyCtrl = IC2.VoyCtrl) AND (IC1.VslCode = IC2.VslCode) "&_
          '         "AND (IC1.BLNo = IC2.BLNo) "&_
          '         "WHERE IC1.ContNo='" & CONnum & "' " &_
          '         "AND IC2.ContNo NOT IN (SELECT ITC.ContNo from hITCommonInfo AS ITC "&_
          '         "WHERE ITC.ContNo IS NOT Null AND ITC.Process='R') "&_
          '         "AND IC2.ContNo NOT IN (SELECT ITF.ContNo FROM hITFullOutSelect AS ITF "&_
          '         "INNER JOIN hITCommonInfo AS ITC2 ON ITF.WkContrlNo = ITC2.WkContrlNo "&_
          '         "WHERE ITC2.Process='R' AND WorkCompleteDate Is Null ) "&_
          '         "ORDER BY IC2.ContNo ASC"
          StrSQL = "SELECT IC2.ContNo, IC1.BLNo FROM ImportCont AS IC1 " &_
                   "INNER JOIN ImportCont AS IC2 ON (IC1.VoyCtrl = IC2.VoyCtrl) AND (IC1.VslCode = IC2.VslCode) "&_
                   "AND (IC1.BLNo = IC2.BLNo) "&_
                   "WHERE IC1.ContNo='" & CONnum & "' " &_
                   "AND IC2.ContNo NOT IN (SELECT ITC.ContNo from hITCommonInfo AS ITC "&_
                   "WHERE ITC.ContNo IS NOT Null AND ITC.Process='R' AND WorkCompleteDate Is Null ) "&_
                   "AND IC2.ContNo NOT IN (SELECT ITF.ContNo FROM hITFullOutSelect AS ITF "&_
                   "INNER JOIN hITCommonInfo AS ITC2 ON ITF.WkContrlNo = ITC2.WkContrlNo "&_
                   "WHERE ITC2.Process='R' AND WorkCompleteDate Is Null ) "&_
                   "ORDER BY IC2.ContNo ASC"
          '2017/03/07 T.Okui Upd-E
'CW-050 "INNER JOIN ImportCont AS IC2 ON IC1.BLNo = IC2.BLNo " &_
'20030911 ADD This Item:"AND WorkCompleteDate Is Null " &_
          ObjRS.Open StrSQL, ObjConn
          if err <> 0 then
            DisConnDBH ObjConn, ObjRS	'DB切断
            jampErrerP "1","b105","01","実搬出:データ一覧取得","102","SQL:<BR>"&strSQL
          end if
'CW-050          BLnum=ObjRS("BLNo")
          ReDim CONnumA(Num)
          CONnumA(0)=CONnum
          i=1
          Do Until ObjRS.EOF
            If CONnum <> Trim(ObjRS("ContNo")) Then 
              CONnumA(i)=ObjRS("ContNo")
              i=i+1
            End If
            ObjRS.MoveNext
          Loop
          ObjRS.Close
          '代表コンテナのデータ取得
          StrSQL = "SELECT CON.ContType, CON.ContSize, CON.ContHeight, CON.ContWeight, INC.BLNo, BL.RecTerminal, "&_
                   "BL.ShipLine, mV.FullName,INC.ReturnPlace " &_
                   "FROM ((Container AS CON INNER JOIN ImportCont AS INC ON (CON.ContNo = INC.ContNo) "&_
                   "AND (CON.VoyCtrl = INC.VoyCtrl) AND (CON.VslCode = INC.VslCode)) "&_
                   "LEFT JOIN BL ON (INC.VslCode = BL.VslCode) AND (INC.VoyCtrl = BL.VoyCtrl) "&_
                   "AND (INC.BLNo = BL.BLNo)) "&_
                   "LEFT JOIN mVessel AS mV On (BL.VslCode=mV.VslCode) "&_
                   "WHERE CON.ContNo='" & CONnum &"' ORDER BY BL.UpdtTime DESC"
'CW-050          StrSQL = "select ContType,ContSize,ContHeight,ContWeight,RecTerminal " &_
'CW-050                   "from Container INNER join BL on BL.VslCode=Container.VslCode " &_ 
'CW-050                   "and  BL.VoyCtrl=Container.VoyCtrl where Container.ContNo='" & CONnum &"'"
'3th add mV.ShipLine, mV.FullName,INC.ReturnPlace
'3th add LEFT JOIN mVessel AS mV On (INC.VslCode=mV.VslCode)
'CW307 Change mV.ShipLine -> BL.ShipLine
          ObjRS.Open StrSQL, ObjConn
          if err <> 0 then
            DisConnDBH ObjConn, ObjRS	'DB切断
            jampErrerP "1","b105","01","実搬出:データ取得","102","SQL:<BR>"&strSQL
          end if
          BLnum=ObjRS("BLNo")	'CW-050

        Case "4"		'BL番号で
          'DBよりデータ所得
          '代表コンテナのデータ取得
          StrSQL = "SELECT BL.RecTerminal, INC.ReturnPlace, BL.ShipLine, mV.FullName "&_
                   "FROM (BL INNER JOIN ImportCont AS INC ON (BL.BLNo = INC.BLNo) "&_
                   "AND (BL.VoyCtrl = INC.VoyCtrl) AND (BL.VslCode = INC.VslCode)) "&_
                   "LEFT JOIN mVessel AS mV ON BL.VslCode = mV.VslCode "&_
                   "WHERE BL.BLNo='" & BLnum & "' ORDER BY BL.UpdtTime DESC"
'3th add mV.ShipLine, mV.FullName,INC.ReturnPlace
'3th add INNER JOIN ImportCont AS INC ON (BL.BLNo = INC.BLNo) AND (BL.VoyCtrl = INC.VoyCtrl) AND (BL.VslCode = INC.VslCode))
'3th add LEFT JOIN mVessel AS mV On (INC.VslCode=mV.VslCode)
'CW307 Change mV.ShipLine -> BL.ShipLine
          ObjRS.Open StrSQL, ObjConn
          if err <> 0 then
            DisConnDBH ObjConn, ObjRS	'DB切断
            jampErrerP "1","b106","01","実搬出:データ取得","102","SQL:<BR>"&strSQL
          end if
          Num=1
          CONnum = ""
      End Select
'CW-035      CMPcd   =Array(User,"","","","")
'2009/07/07 Add-S Fujiyama BL無しの処理
     If ObjRS.EOF <> true Then
'2009/07/07 Add-E Fujiyama
      CMPcd   =Array(Ucase(User),"","","","")
      Rmon = Null
      Rday = Null
      Rnissu = "未入力"
'C-002 ADD START
      Comment1  = ""
      Comment2  = ""
'3th del    Comment3  = ""
'C-002 ADD END
'3th add start
      Rhou = Null
      HFrom = Trim(ObjRS("RecTerminal"))
      shipFact = Trim(ObjRS("ShipLine"))		'船社
      shipName = Trim(ObjRS("FullName"))		'船名
      Nonyu1   = ""		'納入先1
      Nonyu2   = ""		'納入先2
      HinName  = ""		'品名
      NonyuDate(0) = Null
      NonyuDate(1) = Null
      NonyuDate(2) = Null
	  '2008/01/31 Add S G.Ariola
	  NonyuDate(3) = Null
	  '2008/01/31 Add E G.Ariola
      RPlace = Trim(ObjRS("ReturnPlace"))
'2009/07/07 Add-S Fujiyama BL無しの処理
     Else
      ret1=False
     End If
'2009/07/07 Add-E Fujiyama
'3th add End
    Else
      Num=0
    End If
    if err <> 0 then
      DisConnDBH ObjConn, ObjRS	'DB切断
      jampErrerP "1","b101","99","実搬出:データ編集","200", StrSQL & "<P>"
    end if
  Else

    '更新(作業番号から検索)
    dim WkCNo,HedId,HTo,Rmon,Rday,Rnissu,UpUser,TruckerFlag,compFlag
    
'3th add start
    dim DtTbl
    DtTbl = Split(Request("Datatbl"&targetNo), ",", -1, 1)
    HTo   = DtTbl(0)			'搬出先
    SakuNo= DtTbl(3)			'作業番号
    Flag  = DtTbl(4)			'指定フラグ
    If Flag <> 4 Then
      CONnum = DtTbl(5)			'コンテナ番号
    End IF
    Rnissu   = DtTbl(7)			'返却予定日数
    BLnum    = DtTbl(11)		'BL番号
    shipFact = DtTbl(15)		'船社
    shipName = DtTbl(16)		'船名
    HFrom = DtTbl(18)			'CY

    Comment1 = DtTbl(22)		'備考1
    Comment2 = DtTbl(23)		'備考2
    Nonyu1   = DtTbl(24)		'納入先1
'3th add End
    Num=0
    Mord=1

    If Flag = 4 Then		'BL番号で搬出元検索
      StrSQL = "SELECT ITC.WkContrlNo, ITC.HeadID, ITC.WorkCompleteDate, "&_
               "ITC.WorkDate, ITC.UpdtUserCode, ITC.RegisterCode, "&_
               "ITC.TruckerSubCode1, ITC.TruckerSubCode2, ITC.TruckerSubCode3, ITC.TruckerSubCode4, "&_
               "ITC.GoodsName, ITC.DeliverTo2, ITC.DeliverDate, INC.ReturnPlace, "&_
               "ITR.TruckerFlag1, ITR.TruckerFlag2, ITR.TruckerFlag3, ITR.TruckerFlag4, "&_
               "ITC.TruckerSubName1, ITC.TruckerSubName2, ITC.TruckerSubName3, ITC.TruckerSubName4, ITC.TruckerSubName5, "&_
               "T1.Trucked AS Trucked1, T2.Trucked AS Trucked2, T3.Trucked AS Trucked3, T4.Trucked AS Trucked4 "&_
               "FROM (hITCommonInfo AS ITC LEFT JOIN ImportCont AS INC ON ITC.BLNo = INC.BLNo) "&_
               "INNER JOIN hITReference AS ITR ON ITC.WkContrlNo = ITR.WkContrlNo "&_
               "LEFT JOIN mTrucker T1 ON (ITC.TruckerSubCode1 = T1.HeadCompanyCode) "&_
               "LEFT JOIN mTrucker T2 ON (ITC.TruckerSubCode2 = T2.HeadCompanyCode) "&_
               "LEFT JOIN mTrucker T3 ON (ITC.TruckerSubCode3 = T3.HeadCompanyCode) "&_
               "LEFT JOIN mTrucker T4 ON (ITC.TruckerSubCode4 = T4.HeadCompanyCode) "&_
               "WHERE ITC.BLNo='"& BLnum &"' AND ITC.WkNo='"& SakuNo &"' AND ITC.Process='R' AND ITC.WkType='1' "
'3th del ITC.ContNo, ITC.BLNo,ITC.DeliverTo,ITC.ReturnDateStr,ITC.Comment1, ITC.Comment2, ITC.Comment3,BL.RecTerminal
'3th del LEFT JOIN BL ON ITC.BLNo = BL.BLNo
'3th add ITC.GoodsName,ITC.DeliverTo2,ITC.DeliverDate,INC.ReturnPlace
'3th add LEFT JOIN ImportCont AS INC ON ITC.BLNo = INC.BLNo
'20040301 ADD ITC.BLNo='"& BLnum &"' AND
    Else			'コンテナ番号で搬出元検索：
      StrSQL = "SELECT ITC.WkContrlNo, ITC.HeadID, ITC.WorkDate, "&_
               "ITC.UpdtUserCode, ITC.WorkCompleteDate, ITC.RegisterCode, "&_
               "ITC.TruckerSubCode1, ITC.TruckerSubCode2, ITC.TruckerSubCode3, ITC.TruckerSubCode4, "&_
               "ITC.GoodsName,ITC.DeliverTo2,ITC.DeliverDate,INC.ReturnPlace, "&_
               "ITR.TruckerFlag1, ITR.TruckerFlag2, ITR.TruckerFlag3, "&_
               "ITR.TruckerFlag4, Cnt.ContSize, Cnt.ContType, Cnt.ContHeight, Cnt.ContWeight, "&_
               "ITC.TruckerSubName1, ITC.TruckerSubName2, ITC.TruckerSubName3, ITC.TruckerSubName4, ITC.TruckerSubName5, "&_
               "T1.Trucked AS Trucked1, T2.Trucked AS Trucked2, T3.Trucked AS Trucked3, T4.Trucked AS Trucked4 "&_
               "FROM ((hITCommonInfo AS ITC INNER JOIN Container AS Cnt ON ITC.ContNo = Cnt.ContNo) "&_
               "INNER JOIN ImportCont AS INC ON Cnt.ContNo = INC.ContNo AND Cnt.VslCode=INC.VslCode AND Cnt.VoyCtrl=INC.VoyCtrl) "&_
               "INNER JOIN hITReference AS ITR ON ITC.WkContrlNo=ITR.WkContrlNo "&_
               "LEFT JOIN mTrucker T1 ON (ITC.TruckerSubCode1 = T1.HeadCompanyCode) "&_
               "LEFT JOIN mTrucker T2 ON (ITC.TruckerSubCode2 = T2.HeadCompanyCode) "&_
               "LEFT JOIN mTrucker T3 ON (ITC.TruckerSubCode3 = T3.HeadCompanyCode) "&_
               "LEFT JOIN mTrucker T4 ON (ITC.TruckerSubCode4 = T4.HeadCompanyCode) "&_
               "WHERE ITC.ContNo='"& CONnum &"' AND ITC.WkNo='"& SakuNo &"' AND ITC.Process='R' AND ITC.WkType='1'" &_
               "ORDER BY Cnt.UpdtTime DESC"
'CW-038"INNER JOIN (ImportCont INNER JOIN BL ON (ImportCont.VslCode = BL.VslCode) AND "&_
'C-002 ADD This Line 4 each SQL : "ITC.Comment1, ITC.Comment2, ITC.Comment3, "&_
'3th del ITC.ContNo, ITC.BLNo,ITC.DeliverTo,ITC.ReturnDateStr,ITC.Comment1, ITC.Comment2, ITC.Comment3,BL.RecTerminal
'3th del LEFT JOIN BL ON ITC.BLNo = BL.BLNo
'3th add ITC.GoodsName,ITC.DeliverTo2,ITC.DeliverDate,INC.ReturnPlace
'3th add (LEFT JOIN ImportCont AS INC ON Cnt.ContNo = INC.ContNo AND Cnt.VslCode=INC.VslCode AND Cnt.VoyCtrl=INC.VoyCtrl )
'20040301 ADD ITC.ContNo='"& CONnum &"' AND
    End If
    ObjRS.Open StrSQL, ObjConn
    if err <> 0 then
      DisConnDBH ObjConn, ObjRS	'DB切断
      jampErrerP "1","b101","99","実搬出:詳細用データ取得","101","SQL:<BR>"&strSQL
    end if
    WkCNo  = ObjRS("WkContrlNo")
'3th del    CONnum = ObjRS("ContNo")
'3th del    BLnum  = ObjRS("BLNo")
    HedId  = ObjRS("HeadID")
    CMPcd  = Array("","","","","")
    CMPcd(0)  = Trim(ObjRS("RegisterCode"))
    CMPcd(1)  = Trim(ObjRS("TruckerSubCode1"))
    CMPcd(2)  = Trim(ObjRS("TruckerSubCode2"))
    CMPcd(3)  = Trim(ObjRS("TruckerSubCode3"))
    CMPcd(4)  = Trim(ObjRS("TruckerSubCode4"))
    UpUser    = Trim(ObjRS("UpdtUserCode"))
    compFlag  = isNull(ObjRS("WorkCompleteDate"))
'3th del'C-002 ADD START
'3th del    Comment1  = Trim(ObjRS("Comment1"))
'3th del    Comment2  = Trim(ObjRS("Comment2"))
'3th del    Comment3  = Trim(ObjRS("Comment3"))
'3th del'C-002 ADD END


'ログインユーザによって会社コード表示制御
      chengeCompCd CMPcd, UpFlag
      If UpFlag <> 5 Then
        TruckerFlag= Trim(ObjRS("TruckerFlag"&UpFlag))
      Else
        TruckerFlag = 0
      End If

'ログインユーザによってヘッドID表示制御
    IF TruckerFlag = 1 Then 
      HedId  = "*****"
    End If

'2009/08/04 Tanaka Upd-S    
''2009/03/10 R.Shibuta Add-S
''ログインユーザを元に担当者名を選択
'	Select Case User
'		Case Trim(ObjRS("RegisterCode"))
'			TruckerSubName = ObjRS("TruckerSubName1")
'		Case Trim(ObjRS("Trucked1"))
'			TruckerSubName = ObjRS("TruckerSubName2")
'		Case Trim(ObjRS("Trucked2"))
'			TruckerSubName = ObjRS("TruckerSubName3")
'		Case Trim(ObjRS("Trucked3"))
'			TruckerSubName = ObjRS("TruckerSubName4")
'		Case Trim(ObjRS("Trucked4"))
'			TruckerSubName = ObjRS("TruckerSubName5")
'		Case Else
'			TruckerSubName = ""
'	End Select 
''2009/03/10 R.Shibuta Add-E
	Select Case User
		Case Trim(ObjRS("RegisterCode"))
			TruckerSubName = ObjRS("TruckerSubName1")
			TruckerName = ObjRS("TruckerSubName1")
		Case Trim(ObjRS("Trucked1"))
			TruckerSubName = ObjRS("TruckerSubName1")
			TruckerName = ObjRS("TruckerSubName2")
		Case Trim(ObjRS("Trucked2"))
			TruckerSubName = ObjRS("TruckerSubName2")
			TruckerName = ObjRS("TruckerSubName3")
		Case Trim(ObjRS("Trucked3"))
			TruckerSubName = ObjRS("TruckerSubName3")
			TruckerName = ObjRS("TruckerSubName4")
		Case Trim(ObjRS("Trucked4"))
			TruckerSubName = ObjRS("TruckerSubName4")
			TruckerName = ObjRS("TruckerSubName5")
		Case Else
			TruckerSubName = ""
	End Select 


'2009/08/04 Tanaka Upd-E
	
'3th del    HTo    = ObjRS("DeliverTo")
'CW-018    dim TmpA
'CW-018    TmpA   = Split(ObjRS("WorkDate"), "/", -1, 1)
'CW-018    If ObjRS("WorkDate") = "1900/01/01" Then	'日付がNullであった場合
    Dim TmpA,TmpB
    If IsNull(ObjRS("WorkDate")) Or ObjRS("WorkDate") = "1900/01/01" Then	'日付がNullであった場合	'CW-018
      Rmon   = Null
      Rday   = Null
      Rhou   = Null
    Else
'3th chage    dim TmpA						'CW-018
'3th chage    TmpA   = Split(ObjRS("WorkDate"), "/", -1, 1)	'CW-018
'3th chage    Rmon   = TmpA(1)
'3th chage    Rday   = TmpA(2)
      TmpA   = Split(ObjRS("WorkDate"), " ", -1, 1)
      TmpB   = Split(TmpA(0), "/", -1, 1)
      Rmon   = TmpB(1)
      Rday   = TmpB(2)
      If UBound(TmpA) = 0 Then
        Rhou = Null
      Else
        TmpB   = Split(TmpA(1), ":", -1, 1)
        Rhou   = TmpB(0)
      End IF
    End If
'3th del    Rnissu = ObjRS("ReturnDateStr")
'3th add start
    HinName = Trim(ObjRS("GoodsName"))
    Nonyu2  = Trim(ObjRS("DeliverTo2"))
    If IsNull(ObjRS("DeliverDate")) Or ObjRS("DeliverDate") = "1900/01/01" Then	'日付がNullであった場合	'CW-018
    Else
      TmpA         = Split(ObjRS("DeliverDate"), " ", -1, 1)
      TmpB         = Split(TmpA(0), "/", -1, 1)
      NonyuDate(0) = TmpB(1)
      NonyuDate(1) = TmpB(2)
'WC-310      TmpB         = Split(TmpA(1), ":", -1, 1)
'WC-310      NonyuDate(2) = TmpB(0)
      If UBound(TmpA) = 0 Then
        NonyuDate(2) = Null
      Else
        TmpB   = Split(TmpA(1), ":", -1, 1)
        NonyuDate(2)   = TmpB(0)
		'2008/01/31 Add S G.Ariola
		NonyuDate(3)   = TmpB(1)
		'2008/01/31 Add E G.Ariola
      End IF
    End If
    RPlace = Trim(ObjRS("ReturnPlace"))
'3th add End
  End If

'データ設定
  If ret1 AND ret3 AND ret2 Then 	'CW-029 ADD
    If Flag<>4 Then	'CW-022 ADD
      CONsize = ObjRS("ContSize")
      CONtype = ObjRS("ContType")
      CONhite = ObjRS("ContHeight")
      CONtear = ObjRS("ContWeight")*100   'Modified 2003.7.26 コンテナテーブルがもつ重量の単位は100kg
    End If		'CW-022 ADD
'3th del    HFrom   = ObjRS("RecTerminal")
    ObjRS.close
  End If  	'CW-029 ADD
  if err <> 0 then
    DisConnDBH ObjConn, ObjRS	'DB切断
    jampErrerP "1","b101","99","実搬出:データ編集","200",""
  end if

'コンテナ詳細用
  If Mord=1 AND Flag=3 Then
    if err <> 0 then
      err=0
    end if
   '対象件数取得
    StrSQL = "SELECT Count(ContNo) AS NUM FROM hITFullOutSelect WHERE WkContrlNo="& WkCNo
    ObjRS.Open StrSQL, ObjConn
    if err <> 0 then
      DisConnDBH ObjConn, ObjRS	'DB切断
      jampErrerP "1","b101","99","実搬出:詳細用データ取得","101","SQL:<BR>"&strSQL
    end if
    Num=ObjRS("NUM")
    ObjRS.Close
   '対象コンテナ番号一覧取得
    StrSQL = "SELECT ContNo FROM hITFullOutSelect WHERE WkContrlNo="&WkCNo
    ObjRS.Open StrSQL, ObjConn
    if err <> 0 then
      DisConnDBH ObjConn, ObjRS	'DB切断
      jampErrerP "1","b101","99","実搬出:詳細用データ取得","102","SQL:<BR>"&strSQL
    end if
    ReDim CONnumA(Num)
    CONnumA(0)=Trim(CONnum)
    i=1
    Do Until ObjRS.EOF
      If CONnumA(0) <> Trim(ObjRS("ContNo")) Then 
        CONnumA(i)=Trim(ObjRS("ContNo"))
        i=i+1
      End If
      ObjRS.MoveNext
    Loop
    ObjRS.close
  End If

'CW-016 add strart
  If Mord=1 AND compFlag Then
   '輸入コンテナテーブル搬出チェック
    If Flag=4 Then
      strNum="'"& BLnum &"'"
    Else
      strNum="'"& CONnum &"'"
    End If
    checkImportContComp ObjConn, ObjRS,strNum, Flag, compFlag
    ObjRS.close
  End If
'CW-016 add end

'DB接続解除
  DisConnDBH ObjConn, ObjRS
'エラートラップ解除
  on error goto 0

'コンテナ番号受渡しメソッド
Sub Set_CONnum
  For i = 1 to Num-1
    Response.Write "  <INPUT type=hidden name='CONnum" & i & "' value='" & Trim(CONnumA(i)) & "'>" & vbCrLf
  Next
End Sub
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>データ検索中</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
n=<%=Num%>;
//指定なしの場合コンテナ数を表示し、確認を行う
<%IF  Flag = "2" and Mord = "0" AND ret1 AND ret3 AND ret2 Then%>
  flag = confirm('全部で'+n+'口ですがよろしいですか？');
  if(flag==false)
    window.history.back();
<%End If%>

function GoNext(){
<% IF ret1 AND ret3 AND ret2 Then %>
  mord=<%=Mord%>;
  flag=<%=Flag%>;
<%'del 3th  infFlag=<%=InfoFlag%>;
  target=document.dmi015F;
  if(mord==0){
    if(flag==3)
      target.action="./dmi020.asp";
    else
      target.action="./dmi021.asp";
  } else {
<%'del 3th    if(infFlag==9){
'del 3th        ConInfo(target,flag,1);
'del 3th        return;
'del 3th    } else {%>
      target.action="./dmo020.asp";
<%'del 3th    }%>
  }
  target.submit();
<%Else%>
  window.resizeTo(500,500);
<%End If%>
}

// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY onLoad="GoNext();">
<% IF ret1 AND ret3 AND ret2 Then %>
<!-------------DB検索用画面--------------------------->
<FORM name="dmi015F" method="POST">
<P>データ検索中<BR>
しばらくお待ちください</P>
  <INPUT type=hidden name="UpFlag"  value="<%=UpFlag%>">
  <INPUT type=hidden name="CONnum"  value="<%=Trim(CONnum)%>">
  <INPUT type=hidden name="BLnum"   value="<%=Trim(BLnum)%>" >
  <INPUT type=hidden name="CONsize" value="<%=Trim(CONsize)%>">
  <INPUT type=hidden name="CONtype" value="<%=Trim(CONtype)%>">
  <INPUT type=hidden name="CONhite" value="<%=Trim(CONhite)%>">
  <INPUT type=hidden name="CONtear" value="<%=Trim(CONtear)%>">
  <INPUT type=hidden name="CMPcd0"  value="<%=Trim(CMPcd(0))%>">
  <INPUT type=hidden name="CMPcd1"  value="<%=Trim(CMPcd(1))%>">
  <INPUT type=hidden name="CMPcd2"  value="<%=Trim(CMPcd(2))%>">
  <INPUT type=hidden name="CMPcd3"  value="<%=Trim(CMPcd(3))%>">
  <INPUT type=hidden name="CMPcd4"  value="<%=Trim(CMPcd(4))%>">
 <%' 2009/03/10 R.Shibuta Add-S %>
  <INPUT type=hidden name="TruckerSubName" value="<%=Trim(TruckerSubName)%>">
 <%' 2009/03/10 R.Shibuta Add-E %>
 <%' 2009/08/04 Tanaka Add-S %>
  <INPUT type=hidden name="TruckerName" value="<%=Trim(TruckerName)%>">
 <%' 2009/08/04 Tanaka Add-E %>
  <INPUT type=hidden name="Rmon"    value="<%=Rmon%>">
  <INPUT type=hidden name="Rday"    value="<%=Rday%>">
  <INPUT type=hidden name="Rnissu"  value="<%=Trim(Rnissu)%>">
  <INPUT type=hidden name="HFrom"   value="<%=Trim(HFrom)%>">
  <INPUT type=hidden name="flag"    value="<%=Flag%>" >
  <INPUT type=hidden name="num"     value="<%=Num%>" >
<%'C-002 ADD START%>
  <INPUT type=hidden name="Comment1" value="<%=Comment1%>" >
  <INPUT type=hidden name="Comment2" value="<%=Comment2%>" >
<%'3th add   <INPUT type=hidden name="Comment3" value="<%=Comment3% >" > %>
  <INPUT type=hidden name="Rhou"     value="<%=Rhou%>">
  <INPUT type=hidden name="shipFact" value="<%=shipFact%>" >
  <INPUT type=hidden name="shipName" value="<%=shipName%>" >
  <INPUT type=hidden name="HinName"  value="<%=HinName%>" >
  <INPUT type=hidden name="Nonyu1"   value="<%=Nonyu1%>" >
  <INPUT type=hidden name="Nonyu2"   value="<%=Nonyu2%>" >
  <INPUT type=hidden name="RPlace"   value="<%=RPlace%>" >
  <INPUT type=hidden name="Nomon"    value="<%=NonyuDate(0)%>">
  <INPUT type=hidden name="Noday"    value="<%=NonyuDate(1)%>">
  <INPUT type=hidden name="Nohou"    value="<%=NonyuDate(2)%>">
  <INPUT type=hidden name="Nomin"    value="<%=NonyuDate(3)%>">
<%'3th add End %>
<%'C-002 ADD END%>
<% If Num > 1 Then call Set_CONnum End If%>
<% If Mord = 1 Then %>
  <INPUT type=hidden name="SakuNo"  value="<%=SakuNo%>">
  <INPUT type=hidden name="UpUser"  value="<%=UpUser%>">
  <INPUT type=hidden name="HedId"   value="<%=Trim(HedId)%>">
  <INPUT type=hidden name="HTo"     value="<%=Trim(HTo)%>">
  <INPUT type=hidden name="WkCNo"     value="<%=WkCNo%>">
  <INPUT type=hidden name="TruckerFlag" value="<%=TruckerFlag%>">
  <INPUT type=hidden name="compFlag" value=<%=compFlag%>>
<% Else 'ADD 20050303 STAT fro Fourth Recon By SEIKO N.Oosige%>
  <INPUT type=hidden name="compFlag" value="false">
<% End If%>
</TABLE>
 <INPUT type=submit OnClick="GoNext()">
</FORM>
<!-------------画面終わり--------------------------->
<%Else%>
<!-------------エラー画面--------------------------->
<CENTER>
<FORM name="dmi015F" method="POST">
<DIV class=alert>
  <%If ret1=false Then%>
    <P>指定されたコンテナ又はBL番号「<%=strNum%>」は<BR>
       システムに登録されていません。<BR>
       入力の間違いがないか番号を確認してください。</P>
  <%ElseIf ret2=false Then%>
    <P>指定されたコンテナ又はBL番号「<%=strNum%>」は<BR>
       既に登録されています。</P>
  <%Else%>
    <P>指定されたコンテナ又はBL番号「<%=strNum%>」は<BR>
       既に搬出作業を終了しています。</P>
  <%End If%>
</DIV>
<P><INPUT type=submit value="閉じる" onClick="window.close()"></P>
</FORM>
</CENTER>
<%End If%>
</BODY></HTML>
