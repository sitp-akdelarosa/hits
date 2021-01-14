<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi290.asp				_/
'_/	Function	:削除処理				_/
'_/	Date		:2003/05/28				_/
'_/	Code By		:SEIKO Electric.Co 大重			_/
'_/	Modify		:3th	2003/01/31	3次	_/
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
  WriteLogH "b302", "空搬出事前情報入力","14",""

'サーバ日付の取得
  dim DayTime, YY,Yotei
  getDayTime DayTime

'データ所得
  dim BookNo,COMPcd0
  BookNo = Request("BookNo")
  COMPcd0 = Request("COMPcd0")

'ユーザデータ所得
  dim USER
  USER   = UCase(Session.Contents("userid"))
'エラートラップ開始
  on error resume next
'DB接続
  dim ObjConn, ObjRS, StrSQL
  ConnDBH ObjConn, ObjRS
'CW-009	ADD START ↓↓↓↓↓↓↓
  dim ret, ErrerM,compNum
  ret=true
'20030912　Del START ↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓
 '完了チェック
'  StrSQL = "SELECT Count(BookNo) AS numC FROM ExportCont "&_
'           "WHERE BookNo='"& BookNo &"' AND EmpDelTime IS NOT NULL"
'  ObjRS.Open StrSQL, ObjConn
'  if err <> 0 then
'    ObjRS.Close
'    Set ObjRS = Nothing
'    jampErrerPDB ObjConn,"1","b302","15","空搬出：データ削除","101","SQL:<BR>"&StrSQL
'  end if
'  compNum= ObjRS("numC")
'  ObjRS.close
'  If compNum <> 0 Then
'    StrSQL = "SELECT PIC.Qty FROM SPBookInfo AS SPB "&_
'             "INNER JOIN Pickup AS PIC ON SPB.BookNo = PIC.BookNo "&_
'             "WHERE SPB.BookNo='"& BookNo &"'"
'    ObjRS.Open StrSQL, ObjConn
'    if err <> 0 then
'      Set ObjRS = Nothing
'      jampErrerPDB ObjConn,"1","b302","15","空搬出：データ削除","101","SQL:<BR>"&StrSQL
'    end if
'    If Trim(ObjRS("Qty")) = compNum Then
'      ret=false
'      ErrerM="指定の作業は画面操作中に作業が完了したため、削除処理はキャンセルされました。"
'    End If
'    ObjRS.close
'  End If
'20030912　Del END ↑↑↑↑↑↑↑↑↑↑↑↑↑↑
  If ret Then
'CW-009	End ADD ↑↑↑↑↑↑↑
'3th ADD START
    StrSQL = "UPDATE BookingAssign SET UpdtTime='"& Now() &"', UpdtPgCd='PREDEF01', "&_
             "UpdtTmnl='"& USER &"', Process='D' "&_
             "WHERE BookNo='"& BookNo &"' AND SenderCode='"& COMPcd0 &"' AND TruckerCode='"& Request("oldCOMPcd1") &"'"
    ObjConn.Execute(StrSQL)
    if err <> 0 then
      Set ObjRS = Nothing
      jampErrerPDB ObjConn,"1","b302","15","空搬出：データ削除","104","SQL:<BR>"&StrSQL
    end if
    StrSQL = "SELECT count(BookNo) AS Num FROM BookingAssign "&_
             "WHERE BookNo='"& BookNo &"' AND Process='R'"
    ObjRS.Open StrSQL, ObjConn
    if err <> 0 then
      Set ObjRS = Nothing
      jampErrerPDB ObjConn,"1","b302","15","空搬出：データ削除","101","SQL:<BR>"&StrSQL
    end if
    If ObjRS("num") = 0 Then
      ObjRS.close
'3th ADD END
    '削除
      StrSQL = "UPDATE SPBookInfo SET UpdtTime='"& Now() &"', UpdtPgCd='PREDEF01', "&_
               "UpdtTmnl='"& USER &"', Status='0', Process='D' "&_
               "WHERE BookNo='"& BookNo &"' AND SenderCode='"& COMPcd0 &"' AND Process='R' "
      ObjConn.Execute(StrSQL)
      if err <> 0 then
        Set ObjRS = Nothing
        jampErrerPDB ObjConn,"1","b302","15","空搬出：データ削除","104","SQL:<BR>"&StrSQL
      end if
    ElseIf ObjRS("num") = 1 Then
      ObjRS.close
      Dim TruckerCode,TruckerName
      StrSQL = "SELECT TruckerCode,TruckerName FROM BookingAssign "&_
               "WHERE BookNo='"& BookNo &"' AND Process='R'"
      ObjRS.Open StrSQL, ObjConn
      if err <> 0 then
        Set ObjRS = Nothing
        jampErrerPDB ObjConn,"1","b302","15","空搬出：データ削除","101","SQL:<BR>"&StrSQL
      end if
      TruckerCode=Trim(ObjRS("TruckerCode"))
      TruckerName=Trim(ObjRS("TruckerName"))
      ObjRS.close
      StrSQL = "UPDATE SPBookInfo SET UpdtTime='"& Now() &"', UpdtPgCd='PREDEF01', "&_
               "UpdtTmnl='"& USER &"', Status='0', TruckerCode='"&TruckerCode&"', "&_
               "TruckerName='"&TruckerName&"' "&_
               "WHERE BookNo='"& BookNo &"' AND SenderCode='"& COMPcd0 &"' AND Process='R' "
      ObjConn.Execute(StrSQL)
      if err <> 0 then
        Set ObjRS = Nothing
        jampErrerPDB ObjConn,"1","b302","15","空搬出：データ削除","104","SQL:<BR>"&StrSQL
      end if
    End If		'3th ADD
  End If		'CW-008
'DB接続解除
  DisConnDBH ObjConn, ObjRS
'エラートラップ解除
  on error goto 0


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>削除処理</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!-------------削除処理--------------------------->
<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%">
  <TR align=center><TD>
<% If ret Then %>
    削除処理中です。<BR>しばらくお待ちください。<P>画面は自動的に閉じられます。
    <SCRIPT language=JavaScript>
      try{
        window.opener.parent.DList.location.href="./dmo210L.asp"
      }catch(e){}
      window.close();
    </SCRIPT>
<% Else %>
    <DIV class=alert><%=ErrerM%></DIV><BR>
    <INPUT type=button value="閉じる" onClick="window.close()">
<% End If%>
  </TD></TR>
</TABLE>
<!-------------画面終わり--------------------------->
</BODY></HTML>
