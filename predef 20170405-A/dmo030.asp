<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmo030.asp				_/
'_/	Function	:実搬出情報一覧展開画面			_/
'_/	Date		:2003/07/23				_/
'_/	Code By		:SEIKO Electric.Co 大重			_/
'_/	Modify		:C-001 2003/07/29	CSV出力対応	_/
'_/			:C-002 2003/07/29	備考欄対応	_/
'_/			:3th   2004/01/31	3次対応：HTML修正_/
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
  WriteLogH "b109", "実搬出事前情報展開","01",""

'サーバ日時の取得
  dim DayTime,day
  getDayTime DayTime
  day = DayTime(0) & "年" & DayTime(1) & "月" & DayTime(2) & "日" &_
        DayTime(3) & "時" & DayTime(4) & "分現在の情報"
'INIファイルより設定値を取得
  dim param(2)
  getIni param

'データ取得
  dim No,Num, preDtTbl,DtTbl,Siji,i,j
  Siji  =Array("","指定あり","指定なし","一覧","ＢＬ")
  No=Request("targetNo")
  ReDim preDtTbl(1)

  preDtTbl(0)=Split(Request("Datatbl0"), ",", -1, 1)
  preDtTbl(1)=Split(Request("Datatbl"&No), ",", -1, 1)

  Num=1
  ReDim DtTbl(Num)
  DtTbl(0)=preDtTbl(0)
  DtTbl(0)(5)="コンテナ番号"
  'エラートラップ開始
    on error resume next
  'DB接続
    dim ObjConn, ObjRS, StrSQL
    ConnDBH ObjConn, ObjRS
  '展開データ生成
    j=1
'3th     If preDtTbl(1)(4)="1" Then		'指示あり
    Select Case preDtTbl(1)(4)
      Case "1"			'指示あり
        DtTbl(1)=preDtTbl(1)
        DtTbl(1)(11)="　"
      Case "2" 			'指定なし
        '対象取得
        StrSQL = "SELECT Cnt.ContNo,Cnt.ContSize, INC2.ReturnTime, INC2.CYDelTime, "&_
                 "NULLIF('-',LEFT(DateDiff(day,getdate(),DateAdd(day,"&preDtTbl(1)(21)-param(0)+1&" ,INC2.CYDelTime))*("&preDtTbl(1)(21)&"%6),1)) AS ReturnArrert "&_
                 "From (ImportCont AS INC1 INNER JOIN ImportCont AS INC2 ON "&_
                 "(INC1.VoyCtrl = INC2.VoyCtrl) AND (INC1.VslCode = INC2.VslCode) AND (INC1.BLNo = INC2.BLNo)) "&_
                 "INNER JOIN Container AS Cnt "&_
                 "ON INC2.ContNo=Cnt.ContNo AND INC2.VslCode=Cnt.VslCode AND INC2.VoyCtrl=Cnt.VoyCtrl "&_
                 "WHERE INC1.ContNo='" & preDtTbl(1)(5) & "' AND INC1.BLNo= '"& preDtTbl(1)(11) &"' " &_
                 "ORDER BY INC2.ContNo ASC, INC2.UpdtTime DESC"
        ObjRS.Open StrSQL, ObjConn
        if err <> 0 then
          DisConnDBH ObjConn, ObjRS	'DB切断
          jampErrerP "1","b109","01","実搬出：展開画面","101","SQL:<BR>"&StrSQL
        end if
        j=2
        Do Until ObjRS.EOF
          If preDtTbl(1)(5) = Trim(ObjRS("ContNo")) Then
            DtTbl(1)=preDtTbl(1)
            DtTbl(1)(12)=ObjRS("ReturnArrert")
            If IsNull(ObjRS("ReturnTime")) Then
              DtTbl(1)(8)="未"
            Else
              DtTbl(1)(8)="済"
            End If
            DtTbl(1)(17)=Trim(ObjRS("ContSize"))
            DtTbl(1)(20)=Trim(ObjRS("CYDelTime"))
          Else
            ReDim Preserve DtTbl(j)
            DtTbl(j)=preDtTbl(1)
            DtTbl(j)(5)=Trim(ObjRS("ContNo"))
            DtTbl(j)(12)=ObjRS("ReturnArrert")
            If IsNull(ObjRS("ReturnTime")) Then
              DtTbl(j)(8)="未"
            Else
              DtTbl(j)(8)="済"
            End If
            DtTbl(j)(17)=Trim(ObjRS("ContSize"))
            DtTbl(j)(20)=Trim(ObjRS("CYDelTime"))
            j=j+1
          End If
          ObjRS.MoveNext
      Loop
      ObjRS.close
      Num=j-1
'3th    ElseIf preDtTbl(1)(4)="3" Then	'一覧
    Case "3" 			'一覧
        '対象件数取得
        StrSQL = "SELECT count(ITF.ContNo) AS CNUM FROM "&_
                 "(hITCommonInfo ITC INNER JOIN hITFullOutSelect ITF ON ITC.WkContrlNo = ITF.WkContrlNo) "&_
                 "INNER JOIN ImportCont IPC ON ITF.ContNo =IPC.ContNo AND ITC.BLNo = IPC.BLNo "&_
                 "WHERE ITC.ContNo='"&preDtTbl(1)(5)&"' AND ITC.WkNo='"&preDtTbl(1)(3)&"' AND Process='R' AND ITC.WkType='1'"
'ADD 20030908 This Line:AND Process='R' AND ITC.WkType='1'
'ADD 20030911 This Item:AND ITC.WkNo='"&preDtTbl(1)(3)&"'
        ObjRS.Open StrSQL, ObjConn
        if err <> 0 then
          DisConnDBH ObjConn, ObjRS	'DB切断
          jampErrerP "1","b109","01","実搬出：展開画面","101","SQL:<BR>"&StrSQL
        end if
        Num = Num + ObjRS("CNUM")-1
        ObjRS.close
        ReDim Preserve DtTbl(Num)
        'データ取得
        StrSQL = "SELECT ITF.ContNo, IPC.ReturnTime, IPC.CYDelTime, "&_
                 "NULLIF('-',LEFT(DateDiff(day,getdate(),DateAdd(day,"&preDtTbl(1)(21)-param(0)+1&" ,IPC.CYDelTime))*("&preDtTbl(1)(21)&"%6),1)) AS ReturnArrert "&_
                 "FROM (hITCommonInfo ITC INNER JOIN hITFullOutSelect ITF ON ITC.WkContrlNo = ITF.WkContrlNo) "&_
                 "INNER JOIN ImportCont IPC ON ITF.ContNo =IPC.ContNo AND ITC.BLNo = IPC.BLNo "&_
                 "WHERE ITC.ContNo='"&preDtTbl(1)(5)&"' AND ITC.WkNo='"&preDtTbl(1)(3)&"' AND Process='R' AND ITC.WkType='1'"
'ADD 20030908 This Line:AND Process='R' AND ITC.WkType='1'
'ADD 20030911 This Item:AND ITC.WkNo='"&preDtTbl(1)(3)&"'
        ObjRS.Open StrSQL, ObjConn
        if err <> 0 then
          DisConnDBH ObjConn, ObjRS	'DB切断
          jampErrerP "1","b109","01","実搬出：展開画面","101","SQL:<BR>"&StrSQL
        end if
        Do Until ObjRS.EOF
          DtTbl(j)=preDtTbl(1)
          DtTbl(j)(5)=Trim(ObjRS("ContNo"))
          DtTbl(j)(12)=ObjRS("ReturnArrert")
          If IsNull(ObjRS("ReturnTime")) Then
            DtTbl(j)(8)="未"
          Else
            DtTbl(j)(8)="済"
          End If
          ObjRS.MoveNext
          j=j+1
        Loop
        ObjRS.close
'3th      ElseIf preDtTbl(1)(4)="2" Or preDtTbl(1)(4)="4" Then	'指定なし,BL
      Case "4"			'BL
        '対象件数取得
        dim VslCode,VoyCtrl
        '対象BL選定
        StrSQL = "SELECT INC.VslCode, INC.VoyCtrl "&_
                 "From ImportCont AS INC  "&_
                 "Where INC.BLNo= '"& preDtTbl(1)(11) &"' ORDER BY INC.UpdtTime DESC"
        ObjRS.Open StrSQL, ObjConn
        if err <> 0 then
          DisConnDBH ObjConn, ObjRS	'DB切断
          jampErrerP "1","b109","01","実搬出：展開画面","101","SQL:<BR>"&StrSQL
        end if
        VslCode=Trim(ObjRS("VslCode"))
        VoyCtrl=Trim(ObjRS("VoyCtrl"))
        ObjRS.close

        StrSQL = "SELECT count(ContNo) AS CNUM FROM ImportCont WHERE BLNo='"&preDtTbl(1)(11)&"' "&_
                 "AND VoyCtrl =" & VoyCtrl & " AND VslCode= '"& VslCode &"' "
        ObjRS.Open StrSQL, ObjConn
        if err <> 0 then
          DisConnDBH ObjConn, ObjRS	'DB切断
          jampErrerP "1","b109","01","実搬出：展開画面","101","SQL:<BR>"&StrSQL
        end if
        Num = Num + ObjRS("CNUM")-1
        ObjRS.close
        ReDim Preserve DtTbl(Num)
        'データ取得
        StrSQL = "SELECT ContNo, ReturnTime, CYDelTime, "&_
                 "NULLIF('-',LEFT(DateDiff(day,getdate(),DateAdd(day,"&preDtTbl(1)(21)-param(0)+1&" ,CYDelTime))*("&preDtTbl(1)(21)&"%6),1)) AS ReturnArrert "&_
                 "FROM ImportCont WHERE BLNo='"&preDtTbl(1)(11)&"' "&_
                 "AND VoyCtrl =" & VoyCtrl & " AND VslCode= '"& VslCode &"' "
        ObjRS.Open StrSQL, ObjConn
        if err <> 0 then
          DisConnDBH ObjConn, ObjRS	'DB切断
          jampErrerP "1","b109","01","実搬出：展開画面","101","SQL:<BR>"&StrSQL
        end if
        Do Until ObjRS.EOF
          DtTbl(j)=preDtTbl(1)
          DtTbl(j)(5)=Trim(ObjRS("ContNo"))
          DtTbl(j)(12)=ObjRS("ReturnArrert")
          If IsNull(ObjRS("ReturnTime")) Then
            DtTbl(j)(8)="未"
          Else
            DtTbl(j)(8)="済"
          End If
          ObjRS.MoveNext
          j=j+1
        Loop
        ObjRS.close
      Case Else
          jampErrerP "1","b109","01","実搬出：展開画面","101","SQL:<BR>"&StrSQL
      End Select
  'DB接続解除
    DisConnDBH ObjConn, ObjRS
  'エラートラップ解除
    on error goto 0

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>実搬出事前情報展開</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
window.focus();
// -->
</SCRIPT>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY>
<!-------------実搬出情報展開画面--------------------------->
<TABLE border="0" cellPadding="0" cellSpacing="0" width="100%">
   <TR>
     <TD align="right" bgColor="#000099" height="25" colspan="3">
       <IMG src="Image/logo_hits_ver2.gif" height="25" width="300"></TD>
   </TR>
   <TR height="48">
       <TD width="506" align=center><FONT size=+1><B>事前情報入力<B></FONT></TD>
       <TD width="20%"><B>実搬出作業</B></TD>
       <TD nowrap><%=day%></TD></TR>
</TABLE>
<HR>
<CENTER>
<TABLE border="1" cellPadding="3" cellSpacing="0" cols="<%=Num+1%>">
<%If Num<>0 Then%> 
  <% If DtTbl(1)(14)<>"　" Then %>
  <TR class=bga>
    <TH nowrap><%=DtTbl(0)(1)%></TH><TH nowrap><%=DtTbl(0)(2)%></TH>
    <TH nowrap>指示元<BR>への回答</TH>
    <TH nowrap><%=DtTbl(0)(3)%></TH><TH nowrap><%=DtTbl(0)(4)%></TH><TH nowrap><%=DtTbl(0)(5)%></TH>
    <TH nowrap><%=DtTbl(0)(15)%></TH><TH nowrap><%=DtTbl(0)(16)%></TH><TH nowrap><%=DtTbl(0)(17)%></TH>
    <TH nowrap><%=DtTbl(0)(18)%></TH><TH nowrap><%=DtTbl(0)(19)%></TH><TH nowrap><%=DtTbl(0)(24)%></TH>
    <!--<TH nowrap><%'=DtTbl(0)(6)%></TH>--><!-- Commented 2003.9.4 -->
    <TH nowrap><%=DtTbl(0)(7)%></TH>
    <TH nowrap><%=DtTbl(0)(8)%></TH><TH nowrap><%=DtTbl(0)(9)%></TH><TH nowrap><%=DtTbl(0)(10)%></TH>
    <TH nowrap><%=DtTbl(0)(22)%></TH><TH nowrap><%=DtTbl(0)(23)%></TH>
  </TR>
    <% For j=1 to Num %>
      <% If DtTbl(j)(12) = "-" Or DtTbl(j)(8) = "済" Then %>
  <TR class=bgw>
      <% Else %>
  <TR class=bgarrt>  
      <% End If%>
    <TD nowrap><%=DtTbl(j)(1)%><BR></TD><TD nowrap><%=DtTbl(j)(2)%></TD>
    <TD nowrap><%=DtTbl(j)(14)%></TD><TD nowrap><%=DtTbl(j)(3)%></TD>
    <TD nowrap><%=Siji(DtTbl(j)(4))%></TD><TD nowrap><%=DtTbl(j)(5)%></TD>
<%'C-001    <TD nowrap>< %=DtTbl(j)(15)% ></TD><TD nowrap>< %=DtTbl(j)(16)% ></TD><TD nowrap>< %=DtTbl(j)(17)% ></TD> %>
    <TD nowrap><%=DtTbl(j)(15)%><BR></TD><TD nowrap><%=DtTbl(j)(16)%><BR></TD><TD nowrap><%=DtTbl(j)(17)%><BR></TD>
    <TD nowrap><%=DtTbl(j)(18)%><BR></TD><TD nowrap><%=DtTbl(j)(19)%><BR></TD><TD nowrap><%=DtTbl(j)(24)%><BR></TD>
    <!--<TD nowrap><%'=DtTbl(j)(6)%></TD>--><!-- Commented 2003.9.4 -->
    <TD nowrap><%=DtTbl(j)(7)%></TD><TD nowrap><%=DtTbl(j)(8)%></TD>
    <TD nowrap><%=DtTbl(j)(9)%><BR></TD><TD nowrap><%=DtTbl(j)(10)%><BR></TD>
    <TD nowrap><%=DtTbl(j)(22)%><BR></TD><TD nowrap><%=DtTbl(j)(23)%><BR></TD>
  </TR>
    <% Next %>
  <% Else %>
  <TR class=bga>
    <TH nowrap><%=DtTbl(0)(1)%></TH><TH nowrap><%=DtTbl(0)(2)%></TH>
    <TH nowrap><%=DtTbl(0)(3)%></TH><TH nowrap><%=DtTbl(0)(4)%></TH><TH nowrap><%=DtTbl(0)(5)%></TH>
    <TH nowrap><%=DtTbl(0)(15)%></TH><TH nowrap><%=DtTbl(0)(16)%></TH><TH nowrap><%=DtTbl(0)(17)%></TH>
    <TH nowrap><%=DtTbl(0)(18)%></TH><TH nowrap><%=DtTbl(0)(19)%></TH><TH nowrap><%=DtTbl(0)(24)%></TH>
    <!--<TH nowrap><%'=DtTbl(0)(6)%></TH>--><!-- Commented 2003.9.4 -->
    <TH nowrap><%=DtTbl(0)(7)%></TH>
    <TH nowrap><%=DtTbl(0)(8)%></TH><TH nowrap><%=DtTbl(0)(9)%></TH><TH nowrap><%=DtTbl(0)(10)%></TH>
    <TH nowrap><%=DtTbl(0)(22)%></TH><TH nowrap><%=DtTbl(0)(23)%></TH>
  </TR>
    <% For j=1 to Num %>
      <% If DtTbl(j)(12) = "-" Or DtTbl(j)(8) = "済" Then %>
  <TR class=bgw>
      <% Else %>
  <TR class=bgarrt>  
      <% End If%>
    <TD nowrap><%=DtTbl(j)(1)%><BR></TD><TD nowrap><%=DtTbl(j)(2)%></TD>
    <TD nowrap><%=DtTbl(j)(3)%></TD><TD nowrap><%=Siji(DtTbl(j)(4))%></TD><TD nowrap><%=DtTbl(j)(5)%></TD>
<%'C-001    <TD nowrap>< %=DtTbl(j)(15)% ></TD><TD nowrap>< %=DtTbl(j)(16)% ></TD><TD nowrap>< %=DtTbl(j)(17)% ></TD> --%>
    <TD nowrap><%=DtTbl(j)(15)%><BR></TD><TD nowrap><%=DtTbl(j)(16)%><BR></TD><TD nowrap><%=DtTbl(j)(17)%><BR></TD>
    <TD nowrap><%=DtTbl(j)(18)%><BR></TD><TD nowrap><%=DtTbl(j)(19)%><BR></TD><TD nowrap><%=DtTbl(j)(24)%><BR></TD>
    <!--<TD nowrap><%'=DtTbl(j)(6)%></TD>--><!-- Commented 2003.9.4 -->
    <TD nowrap><%=DtTbl(j)(7)%></TD><TD nowrap><%=DtTbl(j)(8)%></TD>
    <TD nowrap><%=DtTbl(j)(9)%><BR></TD><TD nowrap><%=DtTbl(j)(10)%><BR></TD>
    <TD nowrap><%=DtTbl(j)(22)%><BR></TD><TD nowrap><%=DtTbl(j)(23)%><BR></TD>
  </TR>
    <% Next %>
  <% End If %>
<% Else %>
  <TR class=bgw><TD nowrap>作業案件はありません</TD></TR>
<% End If %>
</TABLE>
</CENTER>
<!-------------画面終わり--------------------------->
</BODY></HTML>
