<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:error.asp				_/
'_/	Function	:エラー画面				_/
'_/	Date			:2004/01/05				_/
'_/	Code By		:aspLand HARA			_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTPコンテンツタイプ設定
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<%
'エラー情報取得
  Dim ObjConn, ObjRS,WinFlag,dispId,wkID,wkName,errorCd,etc
  WinFlag= Session.Contents("WinFlag")
  dispId = Session.Contents("dispId")
  wkID   =  Session.Contents("wkID")
  wkName =  Session.Contents("wkName")
  errorCd=  Session.Contents("errorCd")
  etc    =  Session.Contents("etc")
'セッションクリア
  Session.Contents.Remove("WinFlag")
  Session.Contents.Remove("dispId")
  Session.Contents.Remove("wkID")
  Session.Contents.Remove("wkName")
  Session.Contents.Remove("errorCd")
  Session.Contents.Remove("etc")

'エラーメッセージ取得
  Dim ErrorM1,ErrorM2
  Dim ObjFSO,ObjTS,tmpStr,tmp
  Set ObjFSO = Server.CreateObject("Scripting.FileSystemObject")
  Set ObjTS = ObjFSO.OpenTextFile(Server.Mappath("./INI/SendStatusError.ini"),1,false)
  '--- ファイルデータの読込み ---
  Do Until ObjTS.AtEndofStream
    tmpStr = ObjTS.ReadLine
    If Left(tmpStr,3) = errorCd Then
      tmp=Split(tmpStr,":",3,1)
      ErrorM1 = tmp(1)
      ErrorM2 = tmp(2)
      Exit Do
    End If
  Loop
  ObjTS.Close
  Set ObjTS = Nothing
  Set ObjFSO = Nothing

'ボタン表示制御
  Dim Button
  If WinFlag = 0 Then
    Button="'ログイン画面に戻る' onClick='submit()'"
  ElseIf WinFlag = 1 Then
    Button="'閉じる' onClick='window.close()'"
  Else
    Button="'戻る' onClick='window.history.back()'"
  End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>エラー</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!-------------エラー画面--------------------------->
<TABLE border=0 cellPadding=3 cellSpacing=3 width="90%" align=center>
  <TR><TD colspan=2><DIV class=alert>エラー</DIV></TD></TR>
  <TR><TD>エラー画面ID：作業ID</TD><TD>：<%=dispId%>：<%=wkId%></TD></TR>
  <TR><TD>作業名</TD><TD>：<%=wkName%></TD></TR>
  <TR><TD>エラーコード</TD><TD>：<%=errorCd%></TD></TR>
  <TR><TD>メッセージ</TD><TD>：<%=ErrorM1%><BR></TD></TR>
  <TR><TD>対処</TD><TD>：<%=ErrorM2%><BR></TD></TR>
  <TR><TD colspan=2><%=etc%></TD></TR>
  <TR><TD colspan=2 align=center>
        <FORM action="../Userchk.asp" target="_top">
          <INPUT type=hidden name="link" value="SendStatus/sst000F.asp">
          <INPUT type=button value=<%=Button%>>
        </FORM>
      </TD></TR>
</TABLE>
<!-------------画面終わり--------------------------->
</BODY></HTML>
