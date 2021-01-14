<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmo380.asp				_/
'_/	Function	:実搬出情報一覧CSVファイルダウンロード	_/
'_/	Date		:2003/08/07				_/
'_/	Code By		:SEIKO Electric.Co 大重			_/
'_/	Modify		:3th   2004/01/31	3次対応	_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<!--#include File="./Common.inc"-->
<%
'セッションの有効性をチェック
  CheckLoginH
  WriteLogH "b408", "実搬出事前情報CSVファイルダウンロード","01",""

'CSVデータ取得
  dim Num,preDtTbl,i,j
  dim abspage
  Get_Data Num,preDtTbl
'3th ADD Start ↓↓↓↓↓↓↓
'表示データ配列の準備
  dim DtTbl  
  ReDim DtTbl(Num)

'ユーザデータ取得
  dim USER, COMPcd
  USER   = UCase(Session.Contents("userid"))
  COMPcd = Session.Contents("COMPcd")
'検索条件取得
  dim StrSQL,strWhere
  strWhere= Request("strWhere")
  abspage= Request("absPage")  
  
'エラートラップ開始
  on error resume next
'DB接続
  dim ObjConn, ObjRS, ErrerM
  ConnDBH ObjConn, ObjRS

  '追加情報の取得＆展開データの取得
  Dim Flag,tmp
  StrSQL = "SELECT ITC.WkContrlNo, ITC.HeadID, CYV.Voyage, CYV.DPort, CYV.DelivPlace, CYV.ContType, "&_
           "CYV.Material, CYV.TareWeight, CYV.CustOK, CYV.SealNo, CYV.ContWeight, CYV.CustClear, "&_
           "CYV.OvHeight, CYV.OvWidthL, CYV.OvWidthR, CYV.OvLengthF, CYV.OvLengthA, CYV.Operator, "&_
           "EPC.RHO, CASE WHEN mP.FullName IS Null Then Bok.PlaceRec Else mP.FullName END AS PlaceDel, BOK.LPort "&_
           "FROM ((hITCommonInfo AS ITC INNER JOIN CYVanInfo AS CYV ON ITC.WkNo = CYV.WkNo "&_
           "AND ITC.ContNo=CYV.ContNo) "&_
           "LEFT JOIN ExportCont AS EPC ON CYV.BookNo = EPC.BookNo AND CYV.ContNo = EPC.ContNo) "&_
           "LEFT JOIN Booking AS BOK ON EPC.VslCode = BOK.VslCode AND EPC.VoyCtrl = BOK.VoyCtrl AND EPC.BookNo = BOK.BookNo "&_
           "LEFT JOIN mPort AS mP ON Bok.PlaceRec = mP.PortCode "&_
           "WHERE WkType='3' AND (ITC.RegisterCode='"& USER &"' "&_
           "OR ITC.TruckerSubCode1='"& COMPcd &"' OR ITC.TruckerSubCode2='"& COMPcd &"' "&_
           "OR ITC.TruckerSubCode3='"& COMPcd &"' OR ITC.TruckerSubCode4='"& COMPcd &"') AND Process='R' " &_
           strWhere &_
           "ORDER BY isnull(ITC.WorkDate,DATEADD(Year,100,getdate())),ITC.InputDate ASC"
'CW-325 Change INNER->Left
'20040227 Change Bok.PlaceDel->CASE WHEN mP.FullName IS Null Then Bok.PlaceRec Else mP.FullName END AS PlaceDel,
'20040227 ADD LEFT JOIN mPort AS mP ON Bok.PlaceRec = mP.PortCode

  ObjRS.PageSize = 200
  ObjRS.CacheSize = 200
  ObjRS.CursorLocation = 3	
  ObjRS.Open StrSQL, ObjConn    
  ObjRS.AbsolutePage = abspage
  

  if err <> 0 then
    DisConnDBH ObjConn, ObjRS	'DB切断
    jampErrerP "1","b108","01","実搬出：CSVファイルダウンロード","101","SQL:<BR>"&StrSQL
  end if
  i=1
  j=0
  Flag=0
'2009/10/09 Upd-S Tanaka
'  Do Until ObjRS.EOF
''2009/02/25 Del-S G.Ariola  
'      'If preDtTbl(i)(21)=Trim(ObjRS("WkContrlNo")) then
''2009/02/25 Del-E G.Ariola	  
'        tmp=preDtTbl(i)
'        ReDim Preserve tmp(42)
'        If tmp(8)="Yes" Then
'          tmp(22)="*****"					'ヘッドＩＤ非表示
'        Else
'          tmp(22)=Trim(ObjRS("HeadID"))		'ヘッドＩＤ
'        End If
'        tmp(23)=Trim(ObjRS("Voyage"))		'次航
'        tmp(24)=Trim(ObjRS("PlaceDel"))		'荷受地
'        tmp(25)=Trim(ObjRS("LPort"))		'積港
'        tmp(26)=Trim(ObjRS("DPort"))		'揚港
'        tmp(27)=Trim(ObjRS("DelivPlace"))	'荷渡地
'        tmp(28)=ObjRS("ContType")			'タイプ
'        tmp(29)=Trim(ObjRS("Material"))		'材質
'        tmp(30)=Trim(ObjRS("TareWeight"))	'テアウェイト
'        tmp(31)=Trim(ObjRS("CustOK"))		'丸関
'        tmp(32)=Trim(ObjRS("SealNo"))		'シール番号
'        tmp(33)=Trim(ObjRS("ContWeight"))	'グロスウェイト
'        tmp(34)=Trim(ObjRS("CustClear"))	'通関
'        tmp(35)=Trim(ObjRS("RHO"))			'RHO
'        tmp(36)=Trim(ObjRS("OvHeight"))		'O/H
'        tmp(37)=Trim(ObjRS("OvWidthL"))		'O/WL
'        tmp(38)=Trim(ObjRS("OvWidthR"))		'O/WR
'        tmp(39)=Trim(ObjRS("OvLengthF"))	'O/LF
'        tmp(40)=Trim(ObjRS("OvLengthA"))	'O/LA
'        tmp(41)=Trim(ObjRS("Operator"))		'オペレータ
'        DtTbl(j)=tmp
'        j=j+1
'        i=i+1
'        ObjRS.MoveNext
''2009/02/25 Del-S G.Ariola		
'      'Else
'      '  ObjRS.MoveNext
'      'End If
''2009/02/25 Del-E G.Ariola	  
'  Loop
'◆◆◆◆◆◆◆◆
  For i=1 To Ubound(preDtTbl)
    ObjRS.MoveFirst
    Do Until ObjRS.EOF
      
      If preDtTbl(i)(21)=Trim(ObjRS("WkContrlNo")) then
        tmp=preDtTbl(i)
        ReDim Preserve tmp(42)
        If tmp(8)="Yes" Then
          tmp(22)="*****"					'ヘッドＩＤ非表示
        Else
          tmp(22)=Trim(ObjRS("HeadID"))		'ヘッドＩＤ
        End If
        tmp(23)=Trim(ObjRS("Voyage"))		'次航
        tmp(24)=Trim(ObjRS("PlaceDel"))		'荷受地
        tmp(25)=Trim(ObjRS("LPort"))		'積港
        tmp(26)=Trim(ObjRS("DPort"))		'揚港
        tmp(27)=Trim(ObjRS("DelivPlace"))	'荷渡地
        tmp(28)=ObjRS("ContType")			'タイプ
        tmp(29)=Trim(ObjRS("Material"))		'材質
        tmp(30)=Trim(ObjRS("TareWeight"))	'テアウェイト
        tmp(31)=Trim(ObjRS("CustOK"))		'丸関
        tmp(32)=Trim(ObjRS("SealNo"))		'シール番号
        tmp(33)=Trim(ObjRS("ContWeight"))	'グロスウェイト
        tmp(34)=Trim(ObjRS("CustClear"))	'通関
        tmp(35)=Trim(ObjRS("RHO"))			'RHO
        tmp(36)=Trim(ObjRS("OvHeight"))		'O/H
        tmp(37)=Trim(ObjRS("OvWidthL"))		'O/WL
        tmp(38)=Trim(ObjRS("OvWidthR"))		'O/WR
        tmp(39)=Trim(ObjRS("OvLengthF"))	'O/LF
        tmp(40)=Trim(ObjRS("OvLengthA"))	'O/LA
        tmp(41)=Trim(ObjRS("Operator"))		'オペレータ
        DtTbl(j)=tmp
        j=j+1
        Exit Do
      End If
      ObjRS.MoveNext
    Loop
  Next
'2009/10/09 Upd-E Tanaka
  ObjRS.close
'DB接続解除
  DisConnDBH ObjConn, ObjRS
'エラートラップ解除
  on error goto 0

'3th ADD END ↑↑↑↑↑↑↑

' ファイルのダウンロード
  Response.ContentType="application/octet-stream"
  Response.AddHeader "Content-Disposition","attachment; filename=output.csv"
'3th chage    Response.Write "搬入予定日,指示元,指示元への回答,作業番号,ブッキング番号,コンテナ番号,船社,船名,"
'3th chage    Response.Write "サイズ,ハイト,搬入元,ＣＹ,ＣＹカット日,完了日時,指示先,指示先回答,備考１,備考２,備考３"
'3th chage    Response.Write Chr(13) & Chr(10)
'3th chage    For j=0 To Num-1
'3th chage      Response.Write Trim(DtTbl(j)(1))&","&Trim(DtTbl(j)(2))&","&DtTbl(j)(10)&","&Trim(DtTbl(j)(3))&","
'3th chage      Response.Write Trim(DtTbl(j)(4))&","&Trim(DtTbl(j)(5))&","&DtTbl(j)(11)&","&Trim(DtTbl(j)(12))&","
'3th chage      Response.Write Trim(DtTbl(j)(13))&","&Trim(DtTbl(j)(14))&","&DtTbl(j)(15)&","&Trim(DtTbl(j)(16))&","
'3th chage      Response.Write Trim(DtTbl(j)(17))&","&Trim(DtTbl(j)(6))&","&DtTbl(j)(7)&","&Trim(DtTbl(j)(8))&","
'3th chage      Response.Write Trim(DtTbl(j)(18))&","&DtTbl(j)(19)&","&Trim(DtTbl(j)(20))
'3th chage      Response.Write Chr(13) & Chr(10)
'3th chage    Next
    Response.Write "搬入予定日,指示元,指示元への回答,作業番号,ヘッドID,ブッキング番号,コンテナ番号,"
    Response.Write "船社,船名,次航,荷受地,積港,揚港,荷渡地,サイズ,タイプ,ハイト,材質,テアウェイト,"
    Response.Write "丸関,シール番号,グロスウェイト,搬入元,通関,ＲＨＯ,Ｏ/Ｈ,Ｏ/ＷＬ,Ｏ/ＷＲ,Ｏ/ＬＦ,Ｏ/ＬＡ,"
    Response.Write "オペレータ,ＣＹ,ＣＹカット日,完了日時,指示先,指示先回答,備考１,備考２,備考３"
    Response.Write Chr(13) & Chr(10)
    For j=0 To Num-1
      Response.Write Trim(DtTbl(j)(1))&","&Trim(DtTbl(j)(2))&","&DtTbl(j)(10)&","&Trim(DtTbl(j)(3))&","
      Response.Write Trim(DtTbl(j)(22))&","&DtTbl(j)(4)&","&Trim(DtTbl(j)(5))&","
      Response.Write Trim(DtTbl(j)(11))&","&Trim(DtTbl(j)(12))&","&DtTbl(j)(23)&","&Trim(DtTbl(j)(24))&","
      Response.Write Trim(DtTbl(j)(25))&","&Trim(DtTbl(j)(26))&","&Trim(DtTbl(j)(27))&","
      Response.Write Trim(DtTbl(j)(13))&","&Trim(DtTbl(j)(28))&","&DtTbl(j)(14)&","&Trim(DtTbl(j)(29))&","
      Response.Write Trim(DtTbl(j)(30))&","&Trim(DtTbl(j)(31))&","&DtTbl(j)(32)&","&Trim(DtTbl(j)(33))&","
      Response.Write Trim(DtTbl(j)(15))&","&Trim(DtTbl(j)(34))&","&DtTbl(j)(35)&","&Trim(DtTbl(j)(36))&","
      Response.Write Trim(DtTbl(j)(37))&","&Trim(DtTbl(j)(38))&","&DtTbl(j)(39)&","&Trim(DtTbl(j)(40))&","
      Response.Write Trim(DtTbl(j)(41))&","&Trim(DtTbl(j)(16))&","&DtTbl(j)(17)&","
      Response.Write Trim(DtTbl(j)(6))&","&Trim(DtTbl(j)(7))&","&Trim(DtTbl(j)(8))&","
      Response.Write Trim(DtTbl(j)(18))&","&DtTbl(j)(19)&","&Trim(DtTbl(j)(20))
      Response.Write Chr(13) & Chr(10)
    Next
 
%>
