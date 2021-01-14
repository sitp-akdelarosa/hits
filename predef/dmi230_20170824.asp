<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi230.asp				_/
'_/	Function	:事前空搬出入力確認画面			_/
'_/	Date		:2003/05/28				_/
'_/	Code By		:SEIKO Electric.Co 大重			_/
'_/	Modify		:C-002	2003/08/06	備考欄追加	_/
'_/	Modify		:3th	2003/01/31	3次全面改修	_/
'_/	Modify		:2017/05/09			行数を１０行に変更	_/
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
  dim BookNo, COMPcd0, COMPcd1,Mord, ret, ErrerM,i
  dim WkOutFlag,Pcool, OutStyle							'2016/08/25 H.Yoshikawa Add
  dim PickPlace(), Terminal()							'2016/09/07 H.Yoshikawa Add			2017/05/09 H.Yoshikawa Upd(4 ⇒ なし)
  dim WarningM											'2016/10/27 H.Yoshikawa Add

  Const RowNum = 10										'2017/05/09 H.Yoshikawa Add
  Redim PickPlace(RowNum-1)								'2017/05/09 H.Yoshikawa Add
  Redim Terminal(RowNum-1)								'2017/05/09 H.Yoshikawa Add

  BookNo = Trim(Request("BookNo"))
  COMPcd0 = Request("COMPcd0")
  COMPcd1 = Request("COMPcd1")
  Mord    = Request("Mord")
  ret = true
  ErrerM = ""
  WarningM = ""											'2016/10/27 H.Yoshikawa Add
'エラートラップ開始
  on error resume next
'DB接続
  dim ObjConn, ObjRS, StrSQL
  ConnDBH ObjConn, ObjRS

'ブックの重複登録チェック
  dim strCodes,dummy1, dummy2
  If Mord=0 OR (Mord=1 AND COMPcd1 <> Request("oldCOMPcd1")) Then
'2006/03/06 mod-s h.matsuda(SQL文を再構築)
'    checkSPBook ObjConn, ObjRS, BookNo,COMPcd0,COMPcd1,strCodes,dummy1, dummy2, ret
    checkSPBook2 ObjConn, ObjRS, BookNo,COMPcd0,COMPcd1,strCodes,dummy1, dummy2, ret
'2006/03/06 mod-e h.matsuda
    If Not ret Then
      ErrerM="指定したブッキングNoは指示先「"& Left(strCodes,Len(strCodes)-1) &"」で既に登録されています。"
    End If
  End If
  if err <> 0 then
    DisConnDBH ObjConn, ObjRS	'DB切断
    jampErrerP "2","b303","01","ブッキング指示テーブル","101","SQL：<BR>"&StrSQL
  end if
  If (ret) Then
   'ヘッドIDのチェック
    dim CMPcd
    'CW-327 Change
    'CMPcd = Array("",COMPcd1,"","","")
    CMPcd = Array("",Trim(COMPcd1),"","","")
    checkHdCd ObjConn, ObjRS, CMPcd, ret
    If (ret) Then
    Else
      ErrerM="指定された会社コードは存在しません。"
    End If
  End If

'ブックの搬出完了チェック
  If ret Then
    dim cmpNum
    StrSQL = "SELECT Count(EXC.BookNo) AS numB, Count(Pic.Qty) AS numQ "&_
             "FROM ExportCont AS EXC INNER JOIN Pickup AS Pic ON (EXC.VslCode = Pic.VslCode) "&_
             "AND (EXC.VoyCtrl = Pic.VoyCtrl) AND (EXC.BookNo = Pic.BookNo) "&_
             "WHERE EXC.BookNo='"& BookNo &"' AND EmpDelTime IS NOT NULL"
    ObjRS.Open StrSQL, ObjConn
    if err <> 0 then
      DisConnDBH ObjConn, ObjRS
      jampErrerP "2","b303","01","空搬出：搬出完了チェック","101","SQL:<BR>"&strSQL
    end if
    cmpNum=ObjRS("numB")
    If ObjRS("numQ")<>0 Then
      ObjRS.close
      StrSQL = "SELECT Pic.Qty "&_
               "FROM ExportCont AS EXC INNER JOIN Pickup AS Pic ON (EXC.VslCode = Pic.VslCode) "&_
               "AND (EXC.VoyCtrl = Pic.VoyCtrl) AND (EXC.BookNo = Pic.BookNo) "&_
               "WHERE EXC.BookNo='"& BookNo &"' GROUP BY Pic.Qty"
      ObjRS.Open StrSQL, ObjConn
      if err <> 0 then
        DisConnDBH ObjConn, ObjRS
        jampErrerP "2","b303","01","空搬出：搬出完了チェック","101","SQL:<BR>"&strSQL
      end if
      If cmpNum = ObjRS("Qty") Then
        WarningM="<注意>指定のブッキング番号は搬出が完了しています。<BR>"
      End If
    End If
    ObjRS.close
  End If
  
  '2016/09/07 H.Yoshikawa Add Start
  If ret Then
 	dim OutNum, OdrNum			'2016/10/26 H.Yoshikawa Add
  	dim SizeChk					'2017/05/10 H.Yoshikawa Add
  	
 	''本数チェック
  	For i=0 To RowNum-1			'2017/05/09 H.Yoshikawa Upd(4⇒RowNum-1)
  		PickPlace(i) = gfTrim(Request("PickPlace" & i))
  		Terminal(i) = gfTrim(Request("Terminal" & i))
		if gfTrim(Request("UpdFlag" & i)) = "1" then
			'2016/10/12 H.Yoshikawa Add Start （属性が同じものの本数を合算する）
			Dim Sz, Tp, Ht, Qty, j
			Sz = gfTrim(Request("ContSize" & i))
			Tp = gfTrim(Request("ContType" & i))
			Ht = gfTrim(Request("ContHeight" & i))
			Qty = CInt(Request("PickNum" & i))
			for j=0 To RowNum-1						'2017/05/09 H.Yoshikawa Upd(4⇒RowNum-1)
				if i<>j then
					if gfTrim(Request("DelFlag" & i)) <> "1" then	'2017/05/10 H.Yoshikawa Add
						if gfTrim(Request("ContSize" & j)) = Sz and gfTrim(Request("ContType" & j)) = Tp and gfTrim(Request("ContHeight" & j)) = Ht then
							Qty = Qty + CInt(Request("PickNum" & j))
						end if 
					end if											'2017/05/10 H.Yoshikawa Add
				end if
			next
			'2016/10/12 H.Yoshikawa Add End
			
			'ピックアップ場所取得
			StrSQL = "SELECT * FROM Pickup  "
			StrSQL = StrSQL & "WHERE VslCode    = '" & gfSQLEncode(Request("VslCode")) & "'"
			StrSQL = StrSQL & "  AND VoyCtrl    = '" & gfSQLEncode(Request("VoyCtrl")) & "'"
			StrSQL = StrSQL & "  AND BookNo     = '" & gfSQLEncode(BookNo) & "'"
			StrSQL = StrSQL & "  AND ContSize   = '" & gfSQLEncode(Request("ContSize" & i)) & "'"
			StrSQL = StrSQL & "  AND ContType   = '" & gfSQLEncode(Request("ContType" & i)) & "'"
			StrSQL = StrSQL & "  AND ContHeight = '" & gfSQLEncode(Request("ContHeight" & i)) & "'"
			StrSQL = StrSQL & " ORDER BY Qty desc "
		    ObjRS.Open StrSQL, ObjConn
		    if err <> 0 then
		      DisConnDBH ObjConn, ObjRS
		      jampErrerP "2","b303","01","空搬出：本数チェック","101","SQL:<BR>"&strSQL & "<BR>" & err.description
		    end if
			if ObjRS.eof then
				WarningM=WarningM & "<注意>空搬出オーダーが登録されていません。（" & i + 1 & "行目）<BR>"
				PickPlace(i) = ""
				Terminal(i) = ""
			else
				'2016/10/26 H.Yoshikawa Del Start
				'if Qty > CInt(ObjRS("Qty")) then
				'	ret = false
				'	ErrerM=ErrerM & "入力された本数が、空搬出オーダー本数を超えています。（" & i + 1 & "行目）<BR>"
				'end if
				'2016/10/26 H.Yoshikawa Del Start
				PickPlace(i) = gfTrim(ObjRS("PickPlace"))
				Terminal(i) = gfTrim(ObjRS("Terminal"))
			end if
			ObjRS.close

			'2016/10/26 H.Yoshikawa Add Start
			'他ユーザ登録の予約本数を加算
			StrSQL = "SELECT ISNULL(Sum(Qty1), 0) as NumCont FROM BookingAssign "
			StrSQL = StrSQL & "WHERE VslCode    = '" & gfSQLEncode(Request("VslCode")) & "'"
			StrSQL = StrSQL & "  AND Voyage     = '" & gfSQLEncode(Request("VoyCtrl")) & "'"
			StrSQL = StrSQL & "  AND BookNo     = '" & gfSQLEncode(BookNo) & "'"
			StrSQL = StrSQL & "  AND ContSize1   = '" & gfSQLEncode(Request("ContSize" & i)) & "'"
			StrSQL = StrSQL & "  AND ContType1   = '" & gfSQLEncode(Request("ContType" & i)) & "'"
			StrSQL = StrSQL & "  AND ContHeight1 = '" & gfSQLEncode(Request("ContHeight" & i)) & "'"
			StrSQL = StrSQL & "  AND SenderCode <> '" & gfSQLEncode(COMPcd0) & "'"
			StrSQL = StrSQL & "  AND Process     = 'R'"
			ObjRS.Open StrSQL, ObjConn
			if err <> 0 then
				DisConnDBH ObjConn, ObjRS
				jampErrerP "1","b303","01","空搬出：別ユーザ予約本数取得","101","SQL:<BR>"&strSQL
			end if
			if not ObjRS.eof then
				Qty = Qty + CInt(ObjRS("NumCont"))
			end if
			ObjRS.close

			
			'同一属性のオーダー本数を取得
			StrSQL = "SELECT ISNULL(Sum(Qty), 0) as NumQty FROM PickUp "
			StrSQL = StrSQL & "WHERE VslCode    = '" & gfSQLEncode(Request("VslCode")) & "'"
			StrSQL = StrSQL & "  AND VoyCtrl    = '" & gfSQLEncode(Request("VoyCtrl")) & "'"
			StrSQL = StrSQL & "  AND BookNo     = '" & gfSQLEncode(BookNo) & "'"
			StrSQL = StrSQL & "  AND ContSize   = '" & gfSQLEncode(Request("ContSize" & i)) & "'"
			StrSQL = StrSQL & "  AND ContType   = '" & gfSQLEncode(Request("ContType" & i)) & "'"
			StrSQL = StrSQL & "  AND ContHeight = '" & gfSQLEncode(Request("ContHeight" & i)) & "'"
			ObjRS.Open StrSQL, ObjConn
			if err <> 0 then
				DisConnDBH ObjConn, ObjRS
				jampErrerP "1","b303","01","空搬出：オーダー本数取得","101","SQL:<BR>"&strSQL
			end if
			if not ObjRS.eof then
				OdrNum=CInt(ObjRS("NumQty"))
			end if
			ObjRS.close
			if OdrNum > 0 then
				if Qty > OdrNum then
					ret = false
					ErrerM=ErrerM & "入力された属性の本数合計が、空搬出オーダー本数を超えています。（" & i + 1 & "行目）<BR>"
				end if
			end if

			'同一属性の搬出済み本数を取得
			if Qty < CInt(Request("OutNum" & i)) then
				ret = false
				ErrerM=ErrerM & "入力された属性の本数合計が、搬出済み本数を下回っています。（" & i + 1 & "行目）<BR>"
			end if
			'2016/10/26 H.Yoshikawa Add End
			
			'2017/05/10 H.Yoshikawa Add Start
			'サイズ／タイプ／ハイトの組合せがマスタに存在するかチェック
			SizeChk = 0
			StrSQL = "SELECT Count(*) AS CNT FROM ViewkMSizeTypeHeight "
			StrSQL = StrSQL & "WHERE ContSize   = '" & gfSQLEncode(Request("ContSize" & i)) & "'"
			StrSQL = StrSQL & "  AND ContType   = '" & gfSQLEncode(Request("ContType" & i)) & "'"
			StrSQL = StrSQL & "  AND ContHeight = '" & gfSQLEncode(Request("ContHeight" & i)) & "'"
			ObjRS.Open StrSQL, ObjConn
			if err <> 0 then
				DisConnDBH ObjConn, ObjRS
				jampErrerP "1","b303","01","空搬出：サイズタイプハイトマスタ取得","101","SQL:<BR>"&strSQL
			end if
			if not ObjRS.eof then
				SizeChk=CInt(ObjRS("CNT"))
			end if
			ObjRS.close
			if SizeChk <= 0 then
				ret = false
				ErrerM=ErrerM & "入力された属性がサイズタイプハイトマスタに登録されていません。（" & i + 1 & "行目）<BR>"
			end if
			'2017/05/10 H.Yoshikawa Add End
		end if
		
		'2017/05/10 H.Yoshikawa Add Start
		if gfTrim(Request("DelFlag" & i)) = "1" then
			if CInt(Request("OutNum" & i)) > 0 then
				ret = false
				ErrerM=ErrerM & "既に搬出済みのコンテナがあるため、行削除できません。（" & i + 1 & "行目）<BR>"
			end if
		end if
		'2017/05/10 H.Yoshikawa Add End
	Next
  End If
  '2016/09/07 H.Yoshikawa Add End
  
  '2016/10/27 H.Yoshikawa Add Start
  if ErrerM <> "" then
  	ErrerM = ErrerM & "<BR>「戻る」ボタンを押下し、再入力してください。"
  end if
  '2016/10/27 H.Yoshikawa Add End
  
'DB接続解除
  DisConnDBH ObjConn, ObjRS
'エラートラップ解除
  on error goto 0

  dim tmpstr
  If ret Then
    tmpstr=",入力内容の正誤:0(正しい)"
  Else
    tmpstr=",入力内容の正誤:1(誤り)"
  End If
  If Request("Mord")=0 Then
    WriteLogH "b302", "空搬出事前情報入力","02",BookNo&"/"&COMPcd1&tmpstr
  Else
    WriteLogH "b302", "空搬出事前情報入力","13",BookNo&"/"&COMPcd1&tmpstr
  End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>空バンピック情報入力確認</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--

function setParam(target){
//  window.resizeTo(500,260);
  bgset(target);
}

//登録
function GoEntry(printFlag){
  target=document.dmi230F;
  target.SijiF.value=printFlag
  target.action="./dmi240.asp";
  target.submit();
}
//戻る
function GoBackT(){
  target=document.dmi230F;
  target.action="./dmi220.asp";
  target.submit();
}
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="setParam(document.dmi230F)">
<!-------------空搬出情報入力確認画面--------------------------->
<FORM name="dmi230F" method="POST">
<TABLE border=0 cellPadding=3 cellSpacing=0 width="100%">
  <TR>
    <TD colspan=2>
      <B>空バンピック情報入力確認</B></TD></TR>
  <TR>
    <TD><DIV class=bgb>ブッキングＮｏ．</DIV></TD>
    <TD><INPUT type=text name="BookNoM" value="<%=Request("BookNoM")%>" readOnly size=40>
        <INPUT type=hidden name="BookNo" value="<%=Request("BookNo")%>"></TD></TR>
  <TR>
    <TD><DIV class=bgb>船社</DIV></TD>
    <TD><INPUT type=text name="shipFact" value="<%=Request("shipFact")%>" readOnly size=40></TD></TR>
  <TR>
    <TD><DIV class=bgb>*船名</DIV></TD>
    <TD><INPUT type=text name="shipName" value="<%=Request("shipName")%>" readOnly size=40>
    	<INPUT type=hidden name="VslCode" value="<%=Request("VslCode")%>">							<!-- 2016/08/23 H.Yoshikawa Add -->
    </TD></TR>
  <TR>
  	<!-- 2016/08/23 H.Yoshikawa Upd Start -->
    <!-- <TD><DIV class=bgb>仕向地</DIV></TD>
    <TD><INPUT type=text name="delivTo" value="<%=Request("delivTo")%>" readOnly size=40></TD></TR> -->
    <TD><DIV class=bgb>*Voyage</DIV></TD>
    <TD><INPUT type=hidden name="delivTo" value="<%=Request("delivTo")%>">
    	<INPUT type=text name="ExVoyage" value="<%=Request("ExVoyage")%>" readOnly size=12>			<!-- 2016/08/23 H.Yoshikawa Add -->
    	<INPUT type=hidden name="VoyCtrl" value="<%=Request("VoyCtrl")%>">							<!-- 2016/10/17 H.Yoshikawa Upd(text⇒hidden) -->
    </TD></TR>
  	<!-- 2016/08/23 H.Yoshikawa Upd End -->
  <TR>
    <TD><DIV class=bgb>会社コード(陸運)</DIV></TD>
    <TD><INPUT type=text name="COMPcd1" value="<%=COMPcd1%>" size=5  readOnly>
        <INPUT type=hidden name="oldCOMPcd1" value="<%=Request("oldCOMPcd1")%>"></TD></TR>
  <TR>
    <TD><DIV class=bgb>属性と本数</DIV></TD>
    <TD></TD></TR>
  <TR>
    <TD colspan=2>
    <TABLE border=0 cellPadding=0 cellSpacing=0 width=900 align=center>
    <!-- 2016/08/23 H.Yoshikawa Upd Start -->
    <!-- <TR><TD></TD><TD>サイズ</TD><TD>タイプ</TD><TD>高さ</TD><TD>材質</TD><TD>ピック場所</TD><TD></TD><TD>本数</TD></TR> -->
    <TR>
    	<TD></TD>
    	<TD>*サイズ</TD>
    	<TD>*タイプ</TD>
    	<TD>*高さ</TD>
    	<TD>設定温度</TD>
    	<TD>プレクール</TD>
    	<TD>ベンチレーション</TD>
    	<TD>*ピック予定日時(時間はﾌﾟﾚｸｰﾙ時のみ必須)</TD>
    	<TD>　*本数</TD>
    	<TD>搬出可否</TD>
    	<TD>ピックアップ場所</TD>
    	<TD>変更</TD>
    	<TD>行削除</TD>									<!-- 2017/05/10 H.Yoshikwawa Add -->
    </TR>
    <!-- 2016/08/23 H.Yoshikawa Upd End -->
<% For i=0 To RowNum-1%>						<!-- 2017/05/09 H.Yoshikawa Upd(4⇒RowNum-1) -->
      <TR><TD>(<%=i+1%>)</TD>
          <TD><INPUT type=text name="ContSize<%=i%>"   value="<%=Request("ContSize"&i)%>" size=4  readOnly></TD>
          <TD><INPUT type=text name="ContType<%=i%>"   value="<%=Request("ContType"&i)%>" size=4  readOnly></TD>
          <TD><INPUT type=text name="ContHeight<%=i%>" value="<%=Request("ContHeight"&i)%>" size=4  readOnly></TD>
      <!-- 2016/08/23 H.Yoshikawa Upd Start
          <TD><INPUT type=text name="Material<%=i%>"   value="<%=Request("Material"&i)%>"   size=4  readOnly></TD>
          <TD><INPUT type=text name="PickPlace<%=i%>"  value="<%=Request("PickPlace"&i)%>"  size=25 readOnly></TD>
          <TD>・・・</TD>
          <TD><INPUT type=text name="PickNum<%=i%>" value="<%=Request("PickNum"&i)%>" size=4  readOnly></TD></TR> -->
          <TD><INPUT type=text name="SetTemp<%=i%>"  value="<%=Request("SetTemp"&i)%>" size=8 readOnly>℃</TD>
          <TD>
          	<%	if gfTrim(Request("Pcool"&i)) = "" then 
          			Pcool = gfTrim(Request("Bef_Pcool"&i))
          	 	else
          	 		Pcool = gfTrim(Request("Pcool"&i))
          	 	end if
          	%>
          	  <select disabled>
				<option value="0"></option>
				<option value="1" <% if Pcool = "1" then %>selected<% end if %> >有</option>
			  </select>
              <INPUT type=hidden name="Pcool<%=i%>"  value="<%=Pcool%>" size=5 readOnly>
          </TD>
          <TD><INPUT type=text name="Ventilation<%=i%>"  value="<%=Request("Ventilation"&i)%>" size=5 readOnly>%（開口）</TD>
          <TD>
              <INPUT type=text name="PickDate<%=i%>"  value="<%=Request("PickDate"&i)%>" size=15 readOnly>
              <INPUT type=text name="PickHour<%=i%>"  value="<%=Request("PickHour"&i)%>" size=4 readOnly>時
              <INPUT type=text name="PickMinute<%=i%>"  value="<%=Request("PickMinute"&i)%>" size=4 readOnly>分
          </TD>
          <TD>…<INPUT type=text name="PickNum<%=i%>" value="<%=Request("PickNum"&i)%>" size=4 readOnly></TD>
          <% OutStyle=""
             select case Trim(Request("OutFlag"&i))
               case "0"
                 WkOutFlag = "確認中"
               case "1"
                 WkOutFlag = "可"
               case "9"
                 WkOutFlag = "不可"
                 OutStyle="color:red;"
               case else
                 WkOutFlag = ""
             end select
          %>
          <TD style="<%=OutStyle%>"><INPUT type=hidden name="OutFlag<%=i%>"  value="<%=Request("OutFlag"&i)%>" ><%=WkOutFlag %></TD>
          <TD><INPUT type=hidden name="PickPlace<%=i%>"  value="<%=PickPlace(i)%>"><%=gfHTMLEncode(PickPlace(i))%>
              <INPUT type=hidden name="Terminal<%=i%>"  value="<%=Terminal(i)%>">
          </TD>
          <TD><INPUT type=checkbox value="1" disabled <% if Request("UpdFlag"&i) = "1" then%> checked <% end if %>>
              <INPUT type=hidden name="UpdFlag<%=i%>" value="<%=Request("UpdFlag"&i)%>">
          </TD>
		  <% '2017/05/10 H.Yoshikawa Upd Start %>
          <TD><INPUT type=checkbox value="1" disabled <% if Request("DelFlag"&i) = "1" then%> checked <% end if %>>
              <INPUT type=hidden name="DelFlag<%=i%>" value="<%=Request("DelFlag"&i)%>">
          </TD>
		  <% '2017/05/10 H.Yoshikawa Upd End %>
			<% '2016/10/27 H.Yoshikawa Upd Start %>
			<INPUT type=hidden name="OutNum<%=i%>" value="<%=Request("OutNum"&i)%>">  <!-- 2016/10/26 H.Yoshikawa Add -->
			<INPUT type=hidden name="Bef_ContSize<%=i%>"    value="<%=Request("Bef_ContSize"&i)%>">
			<INPUT type=hidden name="Bef_ContType<%=i%>"    value="<%=Request("Bef_ContType"&i)%>">
			<INPUT type=hidden name="Bef_ContHeight<%=i%>"  value="<%=Request("Bef_ContHeight"&i)%>">
			<INPUT type=hidden name="Bef_SetTemp<%=i%>"     value="<%=Request("Bef_SetTemp"&i)%>">
			<INPUT type=hidden name="Bef_Pcool<%=i%>"       value="<%=Request("Bef_Pcool"&i)%>">
			<INPUT type=hidden name="Bef_Ventilation<%=i%>" value="<%=Request("Bef_Ventilation"&i)%>">
			<INPUT type=hidden name="Bef_PickDate<%=i%>"    value="<%=Request("Bef_PickDate"&i)%>">
			<INPUT type=hidden name="Bef_PickHour<%=i%>"    value="<%=Request("Bef_PickHour"&i)%>">
			<INPUT type=hidden name="Bef_PickMinute<%=i%>"  value="<%=Request("Bef_PickMinute"&i)%>">
			<INPUT type=hidden name="Bef_PickNum<%=i%>"     value="<%=Request("Bef_PickNum"&i)%>">
			<INPUT type=hidden name="Bef_OutFlag<%=i%>"   value="<%=Request("Bef_OutFlag"&i)%>">
			<INPUT type=hidden name="Bef_PickPlace<%=i%>"   value="<%=Request("Bef_PickPlace"&i)%>">
			<INPUT type=hidden name="Bef_Terminal<%=i%>"    value="<%=Request("Bef_Terminal"&i)%>">
			<% '2016/10/27 H.Yoshikawa Upd End %>
	  </TR>
      <!-- 2016/08/23 H.Yoshikawa Upd End -->
<% Next %>
    </TABLE>
    </TD></TR>
  <TR>
    <TD><DIV class=bgb>バン詰め日時</DIV></TD>
    <TD><INPUT type=text name="vanMon" value="<%=Request("vanMon")%>" size=3  readOnly>月
        <INPUT type=text name="vanDay" value="<%=Request("vanDay")%>" size=3  readOnly>日
        <INPUT type=text name="vanHou" value="<%=Request("vanHou")%>" size=3  readOnly>時
        <INPUT type=text name="vanMin" value="<%=Request("vanMin")%>" size=3  readOnly>分
        </TD></TR>
  <TR>
    <TD><DIV class=bgb>バン詰め場所１</DIV></TD>
    <TD><INPUT type=text name="vanPlace1" value="<%=Request("vanPlace1")%>" size=73  readOnly></TD></TR>
  <TR>
    <TD><DIV class=bgb>バン詰め場所２</DIV></TD>
    <TD><INPUT type=text name="vanPlace2" value="<%=Request("vanPlace2")%>" size=73  readOnly></TD></TR>
  <TR>
    <TD><DIV class=bgb>品名</DIV></TD>
    <TD><INPUT type=text name="goodsName" value="<%=Request("goodsName")%>" size=30  readOnly></TD></TR>
  <TR>
    <TD><DIV class=bgb>搬入先ＣＹ．ＣＹカット日</DIV></TD>
    <TD><INPUT type=text name="Terminal" value="<%=Request("Terminal")%>" readOnly>
        <INPUT type=text name="CYCut" value="<%=Request("CYCut")%>" readOnly></TD></TR>
  <TR>
    <TD><DIV class=bgb>備考１</DIV></TD>
    <TD><INPUT type=text name="Comment1" value="<%=Request("Comment1")%>" size=73  readOnly></TD></TR>
  <TR>
    <TD><DIV class=bgb>備考２</DIV></TD>
    <TD><INPUT type=text name="Comment2" value="<%=Request("Comment2")%>" size=73  readOnly></TD></TR>
    
  <TR>
<!-- 2009/03/10 R.Shibuta Add-S -->
  	<TD><DIV class=bgy>*登録担当者</DIV></TD>
 	<TD><INPUT type=text name="TruckerSubName" readOnly = "readOnly" value="<%=Request("TruckerSubName")%>" maxlength=16></TD></TR>
<!-- 2009/03/10 R.Shibuta Add-E -->
<!-- 2016/08/23 H.Yoshikawa Add Start -->
  <TR>
  	<TD><DIV class=bgy>*電話番号</DIV></TD>
 	<TD><INPUT type=text name="Tel" value="<%=Request("Tel")%>"  readonly></TD></TR>
  <TR>
  	<TD><DIV class=bgy>*メールアドレス</DIV></TD>
 	<TD><INPUT type=text name="Mail" value="<%=Request("Mail")%>" readonly size=60>
 		<INPUT type=checkbox value="1" <% if Request("MailFlag") = "1" then %>checked <% end if %> disabled>
 		搬出可否状態変更時にメールを受け取る
 		<INPUT type=hidden name="MailFlag" value="<%=Request("MailFlag")%>">
 	</TD></TR>
<!-- 2016/08/23 H.Yoshikawa Add End -->
  <TR>
    <TD colspan=2 align=center>
      <INPUT type=hidden name=Mord value="<%=Request("Mord")%>" >
      <INPUT type=hidden name=COMPcd0 value="<%=COMPcd0%>" >
      <INPUT type=hidden name=Res value="<%=Request("Res")%>" >
      <INPUT type=hidden name=SijiF value="" ><P><BR></P>
      <INPUT type=hidden name=shipline value="<%=Request("shipline")%>" ><%'add h.matsuda%>
<%'2016/08/30 H.Yoshikawa Add Start%>
       <INPUT type=hidden name=compFlag value="<%=Request("compFlag")%>" >
<%'2016/08/30 H.Yoshikawa Add End%>
<% IF ret Then %>
	<% if WarningM <> "" then %>
       <P><DIV class=alert><%=WarningM%></DIV></P>
	<% end if %>
       <INPUT type=button value="確定" onClick="GoEntry('No')">
<% Else %>
       <P><DIV class=alert><%=ErrerM%></DIV></P>
<% End If %>
       <INPUT type=button value="戻る" onClick="GoBackT()">
<% IF Mord=0  AND ret Then %>
       <P><INPUT type=button value="確定＆指示書印刷" onClick="GoEntry('Yes')"></P>
<% End If %>

    </TD></TR>

</TABLE>
</FORM>
<!-------------画面終わり--------------------------->
</BODY></HTML>
