<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi215.asp				_/
'_/	Function	:事前空搬出入力情報取得機能	_/
'_/	Date		:2004/01/31				_/
'_/	Code By		:SEIKO Electric.Co 大重			_/
'_/	Modify		:2017/05/09 行数を１０行に変更					_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTPコンテンツタイプ設定
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->
<!--#include File="CommonFunc.inc"-->
<%
'セッションの有効性をチェック
  CheckLoginH
'データ所得
  dim BookNo, COMPcd0,COMPcd1,Mord, ret, ErrerM
  dim shipFact,shipName,delivTo,Terminal,CYCut,Continfo
  dim TFlag,VanTime,VanPlace1,VanPlace2,GoodsName,Comment1,Comment2
  dim i,tmpDate,BookNoM
  dim TruckerSubName
  dim VslCode, VoyCtrl						'2016/08/22 H.Yoshikawa Add（新規登録部分から移動）
  dim TEL, Mail, MailFlag					'2016/08/22 H.Yoshikawa Add
  dim ExVoyage								'2016/10/17 H.Yoshikawa Add
  Const RowNum = 10							'2017/05/09 H.Yoshikawa Add

  Redim Continfo(RowNum)					'2017/05/09 H.Yoshikawa Upd(5⇒RowNum)
  For i=0 To RowNum							'2017/05/09 H.Yoshikawa Upd(5⇒RowNum)
    '2016/08/22 H.Yoshikawa Upd Start
    'Continfo(i)= Array("","","","","","")
    Continfo(i)= Array("","","","","","","","","","","","","","")		'2017/05/10 H.Yoshikawa Upd(要素数追加)
    '2016/08/22 H.Yoshikawa Upd End
  Next
  VanTime=Array("","","","")
  BookNo = Trim(Request("BookNo"))
  Mord    = Request("Mord")

  'add-s h.matsuda 2006/03/06
  dim ShipLine,VslCodeX,VoyCtrlX,ShoriMode
	ShoriMode = Trim(Request("ShoriMode"))
	ShipLine = Trim(Request("ShipLine"))
	VslCodeX=""
	VoyCtrlX=""
  'add-e h.matsuda 2006/03/06


  ret = true
  ErrerM = ""
'エラートラップ開始
  on error resume next
'DB接続
  dim ObjConn, ObjRS, StrSQL
  ConnDBH ObjConn, ObjRS
  
  If Mord=0 Then '新規登録
    COMPcd0 = UCase(Session.Contents("userid"))
    COMPcd1 = ""
  'ブックの存在チェック
    dim cmpNum
    StrSQL = "SELECT Count(Bok.BookNo) AS numB FROM Booking AS Bok WHERE Bok.BookNo='"& BookNo &"' "
'2006/03/06 add-s h.matsuda
	if ShipLine<>"" and ShoriMode<>"" then
		strSQL=strSQL & " AND BOK.shipline='"& ShipLine &"'"
	End If
'2006/03/06 add-s h.matsuda
    ObjRS.Open StrSQL, ObjConn
    if err <> 0 then
      DisConnDBH ObjConn, ObjRS
      jampErrerP "1","b303","01","空搬出：ブッキング番号存在チェック","101","SQL:<BR>"&strSQL
    end if
    cmpNum=ObjRS("numB")
    ObjRS.close
    If cmpNum<1 Then
      ret=false
      ErrerM="<P>指定されたブッキング番号「"&BookNo&"」は<BR>システムに登録されていません。<BR>"&_
             "入力の間違いがないか番号を確認してください。</P>"
    End If


'2006/03/06 add-s h.matsuda
    If ShipLine<>"" and ShoriMode<>"" Then
		StrSQL = "select a.vslcode,a.voyctrl,a.bookno,maxtime,shipline "
		StrSQL = StrSQL & "from "
		StrSQL = StrSQL & "(select distinct "
		StrSQL = StrSQL & "BookNo,VslCode ,VoyCtrl,shipline,"
		StrSQL = StrSQL & "max(UpdtTime) maxTime from Booking "
		StrSQL = StrSQL & "group by BookNo ,VslCode ,VoyCtrl,shipline) as a "
		StrSQL = StrSQL & "where "
		StrSQL = StrSQL & "bookno='" & BookNo & "' "
		StrSQL = StrSQL & "and ShipLine='" & ShipLine & "'"
		StrSQL = StrSQL & "order by maxtime desc"
		ObjRS.Open StrSQL, ObjConn
		
		if not objrs.eof then
			VslCodeX=trim(ObjRS("vslcode"))
			VoyCtrlX=trim(ObjRS("voyctrl"))
		end if
		ObjRS.close
		else
    End If
'2006/03/06 add-e h.matsuda


    If ret Then
'2006/03/06 mod-s h.matsuda
'      BookAs ObjConn, ObjRS, BookNo,COMPcd0,ret
	if ShipLine<>"" then
      BookAs2 ObjConn, ObjRS, BookNo,COMPcd0,ret,ShipLine
	else
      BookAs ObjConn, ObjRS, BookNo,COMPcd0,ret
	end if
'2006/03/06 mod-e h.matsuda
      If Not ret Then
        '2016/08/25 H.Yoshikawa Upd Start
        'ErrerM="<P>指定されたブッキング番号「"&BookNo&"」は<BR>"&_
        '       "別の登録者によってすでに登録されているため、<BR>登録できません。</P>"
        ErrerM="<P>指定されたブッキング番号「"&BookNo&"」は<BR>すでに同一ユーザで登録されています。<BR>"&_
               " 空バンピック事前情報 、もしくは作業一覧より<BR>更新してください。</P>"
        '2016/08/25 H.Yoshikawa Upd End
      End If
    End If
    If ret Then
    'ブックの搬出完了チェック
     StrSQL = "SELECT Count(EXC.BookNo) AS numB, Count(Pic.Qty) AS numQ "&_
              "FROM ExportCont AS EXC INNER JOIN Pickup AS Pic ON (EXC.VslCode = Pic.VslCode) "&_
              "AND (EXC.VoyCtrl = Pic.VoyCtrl) AND (EXC.BookNo = Pic.BookNo) AND (EXC.PickPlace=Pic.PickPlace) "&_
              "AND (EXC.VslCode = Pic.VslCode) "&_
              "WHERE EXC.BookNo='"& BookNo &"' AND EmpDelTime IS NOT NULL"
'2006/03/06 add-s h.matsuda
    If ShipLine<>"" and ShoriMode<>"" Then
		StrSQL = StrSQL & " AND (EXC.VslCode = '" & VslCodeX & "')"
		StrSQL = StrSQL & " AND (EXC.VoyCtrl = '" & VoyCtrlX & "')"
    End If
'2006/03/06 add-e h.matsuda
     ObjRS.Open StrSQL, ObjConn
     if err <> 0 then
       DisConnDBH ObjConn, ObjRS
       jampErrerP "1","b303","01","空搬出：搬出完了チェック","101","SQL:<BR>"&strSQL
     end if
     cmpNum=ObjRS("numB")
     If ObjRS("numQ")<>0 Then
       ObjRS.close
       StrSQL = "SELECT Pic.Qty "&_
                "FROM ExportCont AS EXC INNER JOIN Pickup AS Pic ON (EXC.VslCode = Pic.VslCode) "&_
                "AND (EXC.VoyCtrl = Pic.VoyCtrl) AND (EXC.BookNo = Pic.BookNo) AND (EXC.PickPlace=Pic.PickPlace) "&_
                "AND (EXC.VslCode = Pic.VslCode) "&_
                "WHERE EXC.BookNo='"& BookNo &"' GROUP BY Pic.Qty"
'2006/03/06 add-s h.matsuda(SQL文を再構築)
	    If ShipLine<>"" and ShoriMode<>"" Then
			StrSQL =replace(strsql,"GROUP BY"," AND (EXC.VslCode = '" & VslCodeX & "')"&_
					" AND (EXC.VoyCtrl = '" & VoyCtrlX & "') GROUP BY")
		End If
'2006/03/06 add-e h.matsuda
       ObjRS.Open StrSQL, ObjConn
       if err <> 0 then
         DisConnDBH ObjConn, ObjRS
         jampErrerP "1","b303","01","空搬出：搬出完了チェック","102","SQL:<BR>"&strSQL
       end if
       If cmpNum = ObjRS("Qty") Then
         ErrerM="<DIV class=alert><注意>指定のブッキング番号は搬出が完了しています。</DIV>"
       End If
     End If
     ObjRS.close
      '情報取得
      StrSQL = "SELECT Bok.RecTerminal, Bok.VslCode, Bok.VoyCtrl, "&_
               "CASE WHEN mV.FullName IS Null Then Bok.VslCode Else mV.FullName END AS shipName, "&_
               "CASE WHEN mS.FullName IS Null Then Bok.ShipLine Else mS.FullName END AS shipfact, "&_
               "CASE WHEN mP.FullName IS Null Then Bok.DPort Else mP.FullName END AS delivTo, "&_
               "Pic.ContSize, Pic.ContType, Pic.ContHeight, Pic.PickPlace, Pic.Material, VSC.CYCut "&_
               "FROM ((((Booking AS Bok LEFT JOIN mVessel AS mV ON Bok.VslCode = mV.VslCode) "&_
               "LEFT JOIN mShipLine AS mS ON Bok.ShipLine = mS.ShipLine) "&_
               "LEFT JOIN mPort AS mP ON Bok.DPort = mP.PortCode) "&_
               "LEFT JOIN Pickup AS Pic ON (Bok.BookNo = Pic.BookNo) AND (Bok.VoyCtrl = Pic.VoyCtrl) AND (Bok.VslCode = Pic.VslCode)) "&_
               "LEFT JOIN VslSchedule AS VSC ON (Bok.VoyCtrl = VSC.VoyCtrl) AND (Bok.VslCode = VSC.VslCode) "&_
               "WHERE Bok.BookNo='"& BookNo &"' ORDER BY Bok.UpdtTime DESC"
'2006/03/06 add-s h.matsuda(SQL文を再構築)
	    If ShipLine<>"" and ShoriMode<>"" Then
			StrSQL =replace(strsql,"ORDER BY"," AND (Bok.shipline = '" & ShipLine & "')"&_
					" ORDER BY")
		End If
'2006/03/06 add-e h.matsuda

'CW-315 ADD Bok.VslCode, Bok.VoyCtrl,
'20040227C Change Bok.DelivPlace = mP.PortCode -> Bok.DPort = mP.PortCode
'20040227C Change mP.FullName AS delivTo -> CASE WHEN mP.FullName IS Null Then Bok.DPort Else  mP.FullName END AS delivTo, 

      ObjRS.Open StrSQL, ObjConn
      if err <> 0 then
        DisConnDBH ObjConn, ObjRS
        jampErrerP "1","b302","01","空搬出：情報取得","102","SQL:<BR>"&strSQL
      end if
      shipFact = Trim(ObjRS("shipFact"))		'船社
'      shipName = Trim(ObjRS("shipName"))		'船名							'2016/08/22 H.Yoshikawa Del
      delivTo  = Trim(ObjRS("delivTo"))		'仕向地
      Terminal = Trim(ObjRS("RecTerminal"))	'搬入先ＣＹ
      If Not IsNull(ObjRS("CYCut")) Then
        CYCut    = Left(ObjRS("CYCut"),10)			'CYカット日
        tmpDate  = Split(CYCut, "/", -1, 1)
        CYCut    = tmpDate(0) & "年" & tmpDate(1) & "月" & tmpDate(2) & "日"
      End if
'CW-315 ADD Start
      'Dim VslCode, VoyCtrl						'2016/08/22 H.Yoshikawa Del（プログラムの最初の方へ移動）
      'VslCode =Trim(ObjRS("VslCode"))			'2016/08/22 H.Yoshikawa Del
      'VoyCtrl=Trim(ObjRS("VoyCtrl"))			'2016/08/22 H.Yoshikawa Del
'CW-315 ADD End
'2016/08/22 H.Yoshikawa Del Start
'      i=0
'      Do Until ObjRS.EOF OR i=5
''CW-315 Change        If shipFact = Trim(ObjRS("shipFact")) AND shipName = Trim(ObjRS("shipName")) Then
'        If VslCode = Trim(ObjRS("VslCode")) AND VoyCtrl = Trim(ObjRS("VoyCtrl")) Then
'          Continfo(i)(0)= Trim(ObjRS("ContSize"))			'サイズ
'          Continfo(i)(1)= Trim(ObjRS("ContType"))			'タイプ
'          Continfo(i)(2)= Trim(ObjRS("ContHeight"))		'高さ
'          Continfo(i)(3)= Trim(ObjRS("Material"))		'材質
'          Continfo(i)(4)= Trim(ObjRS("PickPlace"))	'ピック場所
'          i=i+1
'          ObjRS.MoveNext
'        Else
'          i=5
'        End If
'      Loop
'2016/08/22 H.Yoshikawa Del End
      ObjRS.close
    End If
    TFlag     =""
    VanPlace1 =""
    VanPlace2 =""
    GoodsName =""
    Comment1  =""
    Comment2  =""
    BookNoM  = BookNo
    '2016/08/22 H.Yoshikawa Add Start
    TEL = ""
    Mail = ""
    MailFlag = "0"
    TruckerSubName = ""
	StrSQL = "select * from mUsers "
	StrSQL = StrSQL & "where UserCode = '" & UCase(Session.Contents("userid")) & "' "
	ObjRS.Open StrSQL, ObjConn
	if not ObjRS.eof then
		TEL=trim(ObjRS("TelNo"))
		Mail=trim(ObjRS("MailAddress"))
		TruckerSubName = trim(ObjRS("TTName"))
	end if
	ObjRS.close

    '2016/08/22 H.Yoshikawa Add End
    
    dim tmpstr
    If ret Then
      tmpstr=",入力内容の正誤:0(正しい)"
    Else
      tmpstr=",入力内容の正誤:1(誤り)"
    End If
    WriteLogH "b302", "空搬出事前情報入力","02",BookNo&tmpstr
  Else		'更新

'2006/03/06 add-s h.matsuda
    If ShipLine<>"" and ShoriMode<>"" Then
		StrSQL = "		   select a.vslcode,a.voyctrl,a.bookno,maxtime,shipline	"
		StrSQL = StrSQL & "from													"
		StrSQL = StrSQL & "(select distinct										"
		StrSQL = StrSQL & "BookNo,VslCode ,VoyCtrl,shipline,					"
		StrSQL = StrSQL & "max(UpdtTime) maxTime from Booking					"
		StrSQL = StrSQL & "group by BookNo ,VslCode ,VoyCtrl,shipline) as a		"
		StrSQL = StrSQL & "where												"
		StrSQL = StrSQL & "bookno='" & BookNo & "'								"
		StrSQL = StrSQL & "and ShipLine='" & ShipLine & "'						"
		StrSQL = StrSQL & "order by maxtime desc								"
		ObjRS.Open StrSQL, ObjConn
		
		if not objrs.eof then
			VslCodeX=trim(ObjRS("vslcode"))
			VoyCtrlX=trim(ObjRS("voyctrl"))
		end if
		ObjRS.close
		else
    End If
'2006/03/06 add-e h.matsuda

    dim tmpTimeA,tmpTimeB
    COMPcd0 = Request("COMPcd0")
    COMPcd1 = Request("COMPcd1")

   '情報取得
    '2016/08/22 H.Yoshikawa Upd Start
    'StrSQL = "SELECT TruckerFlag,ContSize1,ContType1,ContHeight1,ContMaterial1,PickPlace1,Qty1, "&_
    '         "ContSize2,ContType2,ContHeight2,ContMaterial2,PickPlace2,Qty2, "&_
    '         "ContSize3,ContType3,ContHeight3,ContMaterial3,PickPlace3,Qty3, "&_
    '         "ContSize4,ContType4,ContHeight4,ContMaterial4,PickPlace4,Qty4, "&_
    '         "ContSize5,ContType5,ContHeight5,ContMaterial5,PickPlace5,Qty5, "&_
    '         "VanTime,VanPlace1,VanPlace2,GoodsName,Comment1,Comment2, "&_
    '         "TruckerSubName "&_
    '         "FROM BookingAssign "&_
    '         "WHERE BookNo='"& BookNo &"' AND SenderCode='"& COMPcd0 &"' AND TruckerCode='"& COMPcd1 &"'"
    StrSQL = ""
    StrSQL = StrSQL & "SELECT BA.*, CASE WHEN mV.FullName IS Null Then BA.VslCode Else mV.FullName END AS shipName "
    StrSQL = StrSQL & "      ,SC.LdVoyage "																					'2016/10/17 H.Yoshikawa Add
    StrSQL = StrSQL & "  FROM BookingAssign BA "
    StrSQL = StrSQL & "  LEFT JOIN mVessel AS mV ON BA.VslCode = mV.VslCode "
    StrSQL = StrSQL & "  LEFT JOIN VslSchedule SC ON BA.VslCode = SC.VslCode AND BA.Voyage = SC.VoyCtrl "					'2016/10/17 H.Yoshikawa Add
    StrSQL = StrSQL & "WHERE BA.BookNo='"& BookNo &"' AND BA.SenderCode='"& COMPcd0 &"' AND BA.TruckerCode='"& COMPcd1 &"' "
    StrSQL = StrSQL & "ORDER BY BA.Seq "
    '2016/08/22 H.Yoshikawa Upd End
    ObjRS.Open StrSQL, ObjConn
    if err <> 0 then
      DisConnDBH ObjConn, ObjRS
      jampErrerP "1","b302","12","空搬出：情報取得","102","SQL:<BR>"&strSQL
    end if

    TFlag         = ObjRS("TruckerFlag")
  '2016/08/22 H.Yoshikawa Upd Start
    'Continfo(0)(0)= Trim(ObjRS("ContSize1"))
    'Continfo(0)(1)= Trim(ObjRS("ContType1"))
    'Continfo(0)(2)= Trim(ObjRS("ContHeight1"))
    'Continfo(0)(3)= Trim(ObjRS("ContMaterial1"))
    'Continfo(0)(4)= Trim(ObjRS("PickPlace1"))
    'Continfo(0)(5)= Trim(ObjRS("Qty1"))
    'Continfo(1)(0)= Trim(ObjRS("ContSize2"))
    'Continfo(1)(1)= Trim(ObjRS("ContType2"))
    'Continfo(1)(2)= Trim(ObjRS("ContHeight2"))
    'Continfo(1)(3)= Trim(ObjRS("ContMaterial2"))
    'Continfo(1)(4)= Trim(ObjRS("PickPlace2"))
    'Continfo(1)(5)= Trim(ObjRS("Qty2"))
    'Continfo(2)(0)= Trim(ObjRS("ContSize3"))
    'Continfo(2)(1)= Trim(ObjRS("ContType3"))
    'Continfo(2)(2)= Trim(ObjRS("ContHeight3"))
    'Continfo(2)(3)= Trim(ObjRS("ContMaterial3"))
    'Continfo(2)(4)= Trim(ObjRS("PickPlace3"))
    'Continfo(2)(5)= Trim(ObjRS("Qty3"))
    'Continfo(3)(0)= Trim(ObjRS("ContSize4"))
    'Continfo(3)(1)= Trim(ObjRS("ContType4"))
    'Continfo(3)(2)= Trim(ObjRS("ContHeight4"))
    'Continfo(3)(3)= Trim(ObjRS("ContMaterial4"))
    'Continfo(3)(4)= Trim(ObjRS("PickPlace4"))
    'Continfo(3)(5)= Trim(ObjRS("Qty4"))
    'Continfo(4)(0)= Trim(ObjRS("ContSize5"))
    'Continfo(4)(1)= Trim(ObjRS("ContType5"))
    'Continfo(4)(2)= Trim(ObjRS("ContHeight5"))
    'Continfo(4)(3)= Trim(ObjRS("ContMaterial5"))
    'Continfo(4)(4)= Trim(ObjRS("PickPlace5"))
    'Continfo(4)(5)= Trim(ObjRS("Qty5"))
    VanPlace1 = Trim(ObjRS("VanPlace1"))
    VanPlace2 = Trim(ObjRS("VanPlace2"))
    If Trim(ObjRS("VanTime")) <> "" Then
      tmpTimeA  = Split(ObjRS("VanTime"), " ", -1, 1)
      tmpTimeB  = Split(tmpTimeA(0), "/", -1, 1)
      VanTime(0)= tmpTimeB(1)
      VanTime(1)= tmpTimeB(2)
      If UBound(tmpTimeA)>0 then		'CW-318
        tmpTimeB  = Split(tmpTimeA(1), ":", -1, 1)
        VanTime(2)= tmpTimeB(0)
        VanTime(3)= tmpTimeB(1)
      End If		'CW-318
    End If
    GoodsName = Trim(ObjRS("GoodsName"))
    Comment1  = Trim(ObjRS("Comment1"))
    Comment2  = Trim(ObjRS("Comment2"))
' 2009/03/10 R.Shibuta Add-S
    TruckerSubName = Trim(ObjRS("TruckerSubName"))
' 2009/03/10 R.Shibuta Add-E

    VslCode =Trim(ObjRS("VslCode"))
    VoyCtrl=Trim(ObjRS("Voyage"))
    ExVoyage=Trim(ObjRS("LdVoyage"))												'2016/10/17 H.Yoshikawa Add
    TEL=Trim(ObjRS("TEL"))
    Mail=Trim(ObjRS("Mail"))
    MailFlag=Trim(ObjRS("MailFlag"))
    shipName = Trim(ObjRS("shipName"))		'船名
    
    dim idx
    dim PickHM
    Do Until ObjRS.EOF
      if gfTrim(ObjRS("Seq")) = "" then
      	idx = 0
      else
      	'2017/05/10 H.Yoshikawa Upd Start
      	'idx = CInt(gfTrim(ObjRS("Seq"))) - 1
      	if gfTrim(ObjRS("Seq")) = "A" then
      		idx = 9
      	else
      		idx = CInt(gfTrim(ObjRS("Seq"))) - 1
      	end if
      	'2017/05/10 H.Yoshikawa Upd End
      end if
      if idx > RowNum - 1 then							'2017/05/09 H.Yoshikawa Upd(4⇒RowNum - 1)
         Exit Do
      end if
      Continfo(idx)(0)= gfTrim(ObjRS("ContSize1"))
      Continfo(idx)(1)= gfTrim(ObjRS("ContType1"))
      Continfo(idx)(2)= gfTrim(ObjRS("ContHeight1"))
      Continfo(idx)(3)= Replace(gfTrim(ObjRS("SetTemp")), "C", "")
      Continfo(idx)(4)= gfTrim(ObjRS("Pcool"))
      Continfo(idx)(5)= gfTrim(ObjRS("Ventilation"))
      Continfo(idx)(6)= gfTrim(ObjRS("PickDate"))
      if gfTrim(ObjRS("PickHM")) = "" then
        Continfo(idx)(7)= ""
        Continfo(idx)(8)= ""
      else
      	PickHM = Split(gfTrim(ObjRS("PickHM")), ":")
        Continfo(idx)(7)= PickHM(0)
        Continfo(idx)(8)= PickHM(1)
      end if
      Continfo(idx)(9)= gfTrim(ObjRS("Qty1"))
      Continfo(idx)(10)= gfTrim(ObjRS("OutFlag"))
      Continfo(idx)(11)= gfTrim(ObjRS("PickPlace1"))
      Continfo(idx)(12)= gfTrim(ObjRS("Terminal"))
      Continfo(idx)(13)= gfTrim(ObjRS("Process"))
      ObjRS.MoveNext
    Loop
  '2016/08/22 H.Yoshikawa Upd End

    ObjRS.close
    if err <> 0 then
      DisConnDBH ObjConn, ObjRS
      jampErrerP "1","b303","12","空搬出：データ取得","200","SQL:<BR>"&UBound(tmpTimeA)
    end if

    'ブックの存在チェック
    StrSQL = "SELECT Count(Bok.BookNo) AS numB FROM Booking AS Bok WHERE Bok.BookNo='"& BookNo &"' "
    ObjRS.Open StrSQL, ObjConn

    if err <> 0 then
      DisConnDBH ObjConn, ObjRS
      jampErrerP "1","b303","12","空搬出：ブッキング番号存在チェック","101","SQL:<BR>"&strSQL
    end if

    If ObjRS("numB")>0 Then
      ObjRS.close
      '情報取得
      StrSQL = "SELECT Bok.RecTerminal, "&_
               "CASE WHEN mV.FullName IS Null Then Bok.VslCode Else mV.FullName END AS shipName, "&_
               "CASE WHEN mS.FullName IS Null Then Bok.ShipLine Else mS.FullName END AS shipfact, "&_
               "CASE WHEN mP.FullName IS Null Then Bok.DPort Else  mP.FullName END AS delivTo, VSC.CYCut "&_
               "FROM (((Booking AS Bok LEFT JOIN mVessel AS mV ON Bok.VslCode = mV.VslCode) "&_
               "LEFT JOIN mShipLine AS mS ON Bok.ShipLine = mS.ShipLine) "&_
               "LEFT JOIN mPort AS mP ON Bok.DPort = mP.PortCode) "&_
               "LEFT JOIN VslSchedule AS VSC ON (Bok.VoyCtrl = VSC.VoyCtrl) AND (Bok.VslCode = VSC.VslCode) "&_
               "WHERE Bok.BookNo='"& BookNo &"' ORDER BY Bok.UpdtTime DESC"
'2006/03/06 add-s h.matsuda(SQL文を再構築)
	    If ShipLine<>"" and ShoriMode<>"" Then
			StrSQL =replace(strsql,"ORDER BY"," AND (Bok.shipline = '" & ShipLine & "')"&_
					" ORDER BY")
		End If
'2006/03/06 add-e h.matsuda

'20040227C Change Bok.DelivPlace = mP.PortCode -> Bok.DPort = mP.PortCode
'20040227C Change mP.FullName AS delivTo -> CASE WHEN mP.FullName IS Null Then Bok.DPort Else  mP.FullName END AS delivTo, 
      ObjRS.Open StrSQL, ObjConn
      if err <> 0 then
        DisConnDBH ObjConn, ObjRS
        jampErrerP "1","b302","12","空搬出：情報取得","102","SQL:<BR>"&strSQL
      end if
      shipFact = Trim(ObjRS("shipFact"))		'船社
      'shipName = Trim(ObjRS("shipName"))		'船名							'2016/08/30 H.Yoshikawa Del
      delivTo  = Trim(ObjRS("delivTo"))		'仕向地
      Terminal = Trim(ObjRS("RecTerminal"))	'搬入先ＣＹ
      If Not IsNull(ObjRS("CYCut")) Then
        CYCut    = Left(ObjRS("CYCut"),10)			'CYカット日
        tmpDate  = Split(CYCut, "/", -1, 1)
        CYCut    = tmpDate(0) & "年" & tmpDate(1) & "月" & tmpDate(2) & "日"
      Else
        CYCut    = ""
      End If
      BookNoM  = BookNo
    Else
      shipFact = ""
      shipName = ""
      delivTo  = ""
      Terminal = ""
      CYCut    = ""
      BookNoM   = "ブッキング番号が削除されています"
    End If
    ObjRS.close
  End If
'DB接続解除
  DisConnDBH ObjConn, ObjRS
'エラートラップ解除
  on error goto 0

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>空搬出情報入力確認</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<% IF ret Then %>
<SCRIPT language=JavaScript>
<!--
//登録
function GoNext(){
  target=document.dmi215F;
  <% If Mord=0 Then %>
  target.action="./dmi220.asp";
  <% Else %>
  target.action="./dmo220.asp";
  <% End If%>
  target.submit();
}
// -->
</SCRIPT>
<!-------------DB検索用画面--------------------------->
<BODY onLoad="GoNext();">
<FORM name="dmi215F" method="POST">
  <INPUT type=hidden name="BookNo"   value="<%=BookNo%>">
  <INPUT type=hidden name="BookNoM"   value="<%=BookNoM%>">
  <INPUT type=hidden name="COMPcd0"  value="<%=COMPcd0%>">
  <INPUT type=hidden name="COMPcd1"  value="<%=COMPcd1%>">
  <INPUT type=hidden name="oldCOMPcd1" value="<%=COMPcd1%>">
  <INPUT type=hidden name="Mord"     value="<%=Mord%>">
  <INPUT type=hidden name="CompF"    value="<%=Request("CompF")%>" >
  <INPUT type=hidden name="shipFact" value="<%=shipFact%>">
  <INPUT type=hidden name="shipName" value="<%=shipName%>">
  <INPUT type=hidden name="delivTo"  value="<%=delivTo%>">
  <INPUT type=hidden name="Terminal"  value="<%=Terminal%>">
  <INPUT type=hidden name="CYCut"    value="<%=CYCut%>">

  <INPUT type=hidden name="shipline" value="<%=shipline%>">

  <% For i=0 To RowNum - 1%>						<!-- 2017/05/09 H.Yoshikawa Upd(4⇒RowNum - 1) -->
  <INPUT type=hidden name="ContSize<%=i%>"   value="<%=Continfo(i)(0)%>">
  <INPUT type=hidden name="ContType<%=i%>"   value="<%=Continfo(i)(1)%>">
  <INPUT type=hidden name="ContHeight<%=i%>" value="<%=Continfo(i)(2)%>">
<% '2016/08/22 H.Yoshikawa Upd Start %>
  <!--INPUT type=hidden name="Material<%=i%>"   value="<%=Continfo(i)(3)%>" -->
  <!--INPUT type=hidden name="PickPlace<%=i%>"  value="<%=Continfo(i)(4)%>" -->
  <!--INPUT type=hidden name="PickNum<%=i%>"    value="<%=Continfo(i)(5)%>" -->
  <INPUT type=hidden name="SetTemp<%=i%>"     value="<%=Continfo(i)(3)%>">
  <INPUT type=hidden name="Pcool<%=i%>"       value="<%=Continfo(i)(4)%>">
  <INPUT type=hidden name="Ventilation<%=i%>" value="<%=Continfo(i)(5)%>">
  <INPUT type=hidden name="PickDate<%=i%>"    value="<%=Continfo(i)(6)%>">
  <INPUT type=hidden name="PickHour<%=i%>"    value="<%=Continfo(i)(7)%>">
  <INPUT type=hidden name="PickMinute<%=i%>"  value="<%=Continfo(i)(8)%>">
  <INPUT type=hidden name="PickNum<%=i%>"    value="<%=Continfo(i)(9)%>">
  <INPUT type=hidden name="OutFlag<%=i%>"    value="<%=Continfo(i)(10)%>">
  <INPUT type=hidden name="PickPlace<%=i%>"  value="<%=Continfo(i)(11)%>">
  <INPUT type=hidden name="Terminal<%=i%>"   value="<%=Continfo(i)(12)%>">
  <% if i = 0 then %>
  <INPUT type=hidden name="UpdFlag<%=i%>"    value="1">
  <% else %>
  <INPUT type=hidden name="UpdFlag<%=i%>"    value="0">
  <% end if %>
<% '2016/08/22 H.Yoshikawa Upd End %>
<% '2017/05/10 H.Yoshikawa Upd Start %>
  <% if Continfo(i)(13) = "D" then %>
  <INPUT type=hidden name="DelFlag<%=i%>"    value="1">
  <% else %>
  <INPUT type=hidden name="DelFlag<%=i%>"    value="0">
  <% end if %>
<% '2017/05/10 H.Yoshikawa Upd End %>
<% '2016/10/27 H.Yoshikawa Upd Start %>
  <INPUT type=hidden name="Bef_ContSize<%=i%>"    value="<%=Continfo(i)(0)%>">
  <INPUT type=hidden name="Bef_ContType<%=i%>"    value="<%=Continfo(i)(1)%>">
  <INPUT type=hidden name="Bef_ContHeight<%=i%>"  value="<%=Continfo(i)(2)%>">
  <INPUT type=hidden name="Bef_SetTemp<%=i%>"     value="<%=Continfo(i)(3)%>">
  <INPUT type=hidden name="Bef_Pcool<%=i%>"       value="<%=Continfo(i)(4)%>">
  <INPUT type=hidden name="Bef_Ventilation<%=i%>" value="<%=Continfo(i)(5)%>">
  <INPUT type=hidden name="Bef_PickDate<%=i%>"    value="<%=Continfo(i)(6)%>">
  <INPUT type=hidden name="Bef_PickHour<%=i%>"    value="<%=Continfo(i)(7)%>">
  <INPUT type=hidden name="Bef_PickMinute<%=i%>"  value="<%=Continfo(i)(8)%>">
  <INPUT type=hidden name="Bef_PickNum<%=i%>"     value="<%=Continfo(i)(9)%>">
  <INPUT type=hidden name="Bef_OutFlag<%=i%>"     value="<%=Continfo(i)(10)%>">
  <INPUT type=hidden name="Bef_PickPlace<%=i%>"   value="<%=Continfo(i)(11)%>">
  <INPUT type=hidden name="Bef_Terminal<%=i%>"    value="<%=Continfo(i)(12)%>">
<% '2016/10/27 H.Yoshikawa Upd End %>
  <% Next %>
  <INPUT type=hidden name="TFlag"     value="<%=TFlag%>">
  <INPUT type=hidden name="vanMon"    value="<%=VanTime(0)%>">
  <INPUT type=hidden name="vanDay"    value="<%=VanTime(1)%>">
  <INPUT type=hidden name="vanHou"    value="<%=VanTime(2)%>">
  <INPUT type=hidden name="vanMin"    value="<%=VanTime(3)%>">
  <INPUT type=hidden name="vanPlace2" value="<%=VanPlace2%>">
  <INPUT type=hidden name="vanPlace1" value="<%=VanPlace1%>">
  <INPUT type=hidden name="goodsName" value="<%=GoodsName%>">
  <INPUT type=hidden name="Comment1"  value="<%=Comment1%>">
  <INPUT type=hidden name="Comment2"  value="<%=Comment2%>">
  <INPUT type=hidden name="ErrerM"  value="<%=ErrerM%>">
<!-- 2009/03/10 R.Shibuta Add-S -->
  <INPUT type=hidden name="TruckerSubName" value="<%=TruckerSubName%>">
<!-- 2009/03/10 R.Shibuta Add-E -->
<% '2016/08/22 H.Yoshikawa Upd Start %>
  <INPUT type=hidden name="VslCode" value="<%=VslCode%>">
  <INPUT type=hidden name="VoyCtrl" value="<%=VoyCtrl%>">
  <INPUT type=hidden name="ExVoyage" value="<%=ExVoyage%>">				<!-- 2016/10/17 H.Yoshikawa Add -->
  <INPUT type=hidden name="TEL" value="<%=TEL%>">
  <INPUT type=hidden name="Mail" value="<%=Mail%>">
  <INPUT type=hidden name="MailFlag" value="<%=MailFlag%>">
<% '2016/08/22 H.Yoshikawa Upd End %>

</FORM>
<!-------------画面終わり--------------------------->
<%Else%>
<!-------------エラー画面--------------------------->
<SCRIPT language=JavaScript>
<!--
window.resizeTo(600,400);
// -->
</SCRIPT>
<BODY>
<CENTER>
<DIV class=alert>
  <%=ErrerM%>
</DIV>
<P><INPUT type=submit value="閉じる" onClick="window.close()"></P>

</CENTER>
<%End If%>

</BODY></HTML>
