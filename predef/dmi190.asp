<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi190.asp				_/
'_/	Function	:削除処理				_/
'_/	Date		:2003/05/28				_/
'_/	Code By		:SEIKO Electric.Co 大重			_/
'_/	Modify		:					_/
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
  WriteLogH "b202", "空搬入事前情報入力","15",""

'サーバ日付の取得
  dim DayTime
  getDayTime DayTime

'データ取得
  dim SakuNo,Num, WkCNo, userid
  userid = UCase(Session.Contents("userid"))
  WkCNo = Request("WkCNo")
'エラートラップ開始
  on error resume next
'DB接続
  dim ObjConn, ObjRS, StrSQL
  ConnDBH ObjConn, ObjRS
'CW-008	ADD START ↓↓↓↓↓↓↓
  dim ret, ErrerM
  ret=true
'20030912　Del START ↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓
 '完了チェック
'  StrSQL="SELECT WorkCompleteDate FROM hITCommonInfo " &_
'         "Where WkContrlNo="& WkCNo &" AND Process='R' AND WkType='2'"
'  ObjRS.Open StrSQL, ObjConn
'  if err <> 0 then
'    ObjRS.Close
'    Set ObjRS = Nothing
'    jampErrerPDB ObjConn, ObjRS, "1","b102","15","空搬入：データ削除","101","SQL:<BR>"&StrSQL
'  end if
'  If NOT IsNull(ObjRS("WorkCompleteDate")) Then 
'    ret=false
'    ErrerM="指定の作業は画面操作中に作業が完了したため、削除処理はキャンセルされました。"
'  End If
'  ObjRS.close
'20030912　Del END ↑↑↑↑↑↑↑↑↑↑↑↑↑↑

  If ret Then
'CW-008	End ADD ↑↑↑↑↑↑↑
    StrSQL = "UPDATE hITCommonInfo SET UpdtTime='"& Now() &"', UpdtPgCd='PREDEF01'," &_
             "UpdtTmnl='"& userid &"', UpdtUserCode='"& userid &"', Status='0', Process='D' " &_
             "Where WkContrlNo="& WkCNo &" AND Process='R' AND WkType='2'"
    ObjConn.Execute(StrSQL)
    if err <> 0 then
      Set ObjRS = Nothing
      jampErrerPDB ObjConn,"1","b202","15","空搬入：データ削除","104","SQL:<BR>"&StrSQL
    end if
'  StrSQL = "DELETE FROM hITReference Where WkContrlNo="& WkCNo 
'    ObjConn.Execute(StrSQL)
'    if err <> 0 then
'      jampErrerPDB ObjConn,"1","b202","15","空搬入：データ削除","105","SQL:<BR>"&StrSQL
'    end if
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
      window.opener.parent.DList.location.href="./dmo110L.asp"
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
