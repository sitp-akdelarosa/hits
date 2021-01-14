<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi090.asp				_/
'_/	Function	:削除処理				_/
'_/	Date		:2003/05/27				_/
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

'サーバ日付の取得
  dim DayTime
  getDayTime DayTime

'データ取得
  dim SakuNo,Flag,Num, WkCNo, userid,Way
  userid = UCase(Session.Contents("userid"))
  SakuNo = Request("SakuNo")
  Flag= Request("flag")
  WkCNo = Request("WkCNo")
  Way   =Array("","指定あり","指定なし","一覧から選択","ＢＬ番号")
  WriteLogH "b10"&(2+Flag), "実搬出事前情報一覧("&Way(Flag)&")","15",""

'エラートラップ開始
  on error resume next
'DB接続
  dim ObjConn, ObjRS, StrSQL
  ConnDBH ObjConn, ObjRS
'CW-007	ADD START ↓↓↓↓↓↓↓
  dim ret, ErrerM,strNum
  ret=true
'20030912　Del START ↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓
'CW-044 ADD START
 '輸入コンテナテーブル搬出完了チェック
'  If Flag=4 Then
'    strNum="'"& Request("BLnum") &"'"
'  Else
'    strNum="'"& Request("CONnum") &"'"
'  End If
'  checkImportContComp ObjConn, ObjRS,strNum, Flag, ret
'  If Not ret Then
'    ret=false
'    ErrerM="指定の作業は画面操作中に搬出されたため、削除処理はキャンセルされました。"
'  Else
'CW-044 ADD END
   '完了チェック
'    StrSQL="SELECT WorkCompleteDate FROM hITCommonInfo " &_
'           "Where WkContrlNo="& WkCNo &" AND Process='R' AND WkType='1'"
'    ObjRS.Open StrSQL, ObjConn
'Response.Write StrSQL &"<P>"
'    if err <> 0 then
'      ObjRS.Close
'      Set ObjRS = Nothing
'      jampErrerPDB ObjConn,"1","b100","15","実搬出：データ削除","101","SQL:<BR>"&StrSQL
'    end if
'    If NOT IsNull(ObjRS("WorkCompleteDate")) Then 
'      ret=false
'      ErrerM="指定の作業は画面操作中に作業が完了したため、削除処理はキャンセルされました。"
'    End If
'    ObjRS.close
'  End If 				'CW-044 ADD
'20030912　Del END ↑↑↑↑↑↑↑↑↑↑↑↑↑↑
  If ret Then
'CW-007	End ADD ↑↑↑↑↑↑↑
    'IT共通テーブルの更新
    StrSQL = "UPDATE hITCommonInfo SET UpdtTime='"& Now() &"', UpdtPgCd='PREDEF01'," &_
             "UpdtTmnl='"& userid &"', UpdtUserCode='"& userid &"', Status='0', Process='D' " &_
             "Where WkContrlNo="& WkCNo &" AND Process='R' AND WkType='1'"
    ObjConn.Execute(StrSQL)
    if err <> 0 then
      Set ObjRS = Nothing
      jampErrerPDB ObjConn,"1","b100","15","実搬出:削除","104","SQL:<BR>"&strSQL
    end if
  '作業番号の開放
  StrSQL = "UPDATE hITWkNo SET UpdtTime='"& Now() &"', UpdtPgCd='PREDEF01'," &_
           "UpdtTmnl='"& userid &"', Status='4' " &_
           "Where WkNo='"& SakuNo &"'"
    ObjConn.Execute(StrSQL)
    if err <> 0 then
      Set ObjRS = Nothing
      jampErrerPDB ObjConn,"1","b100","15","実搬出:削除","104","SQL:<BR>"&strSQL
    end if
'  StrSQL = "DELETE FROM hITReference Where WkContrlNo="& WkCNo 
'    ObjConn.Execute(StrSQL)
'    if err <> 0 then
'      Set ObjRS = Nothing
'      jampErrerPDB ObjConn,"1","b100","15","実搬出:削除","105","SQL:<BR>"&strSQL
'    end if
  End If		'CW-007
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
        window.opener.parent.DList.location.href="./dmo010L.asp"
     }catch(e){
     }
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
