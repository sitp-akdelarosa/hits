<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmo900.asp				_/
'_/	Function	:実搬出入力情報取得			_/
'_/	Date		:2003/12/17				_/
'_/	Code By		:SEIKO Electric.Co 大重			_/
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
  dim CONnum,Flag,BLnum,SakuNo
  dim inPutStr,strNums
  CONnum = Request("CONnum")
  Flag   = Request("flag")
  SakuNo = Request("SakuNo")

'エラートラップ開始
  on error resume next
'DB接続
  dim ObjConn, ObjRS, StrSQL
  ConnDBH ObjConn, ObjRS

  Select Case Flag
    Case "1"		'指定有り
      inPutStr="<INPUT type=hidden name='cntnrno' value='"& CONnum &"'>"
	Case "2"		'指定なし
      StrSQL = "SELECT ITC.BLNo FROM hITCommonInfo AS ITC " &_
               "WHERE ITC.ContNo='"& CONnum &"' AND ITC.WkNo='"& SakuNo &"' AND ITC.Process='R' AND ITC.WkType='1'"
      ObjRS.Open StrSQL, ObjConn
	  inPutStr="<INPUT type=hidden name='blno' value='"& Trim(ObjRS("BLNo")) &"'>"
      ObjRS.close
      if err <> 0 then
        DisConnDBH ObjConn, ObjRS	'DB切断
        jampErrerP "1","b101","99","実搬出:詳細用データ取得","102","SQL:<BR>"&strSQL
      end if
	Case "3"		'一覧から選択
		strNums=CONnum
	   '対象コンテナ番号一覧取得
      StrSQL = "SELECT ITF.ContNo FROM hITCommonInfo AS ITC " &_
               "LEFT JOIN hITFullOutSelect AS ITF ON ITC.WkContrlNo = ITF.WkContrlNo " &_
               "WHERE ITC.ContNo='"& CONnum &"' AND ITC.WkNo='"& SakuNo &"' AND ITC.Process='R' AND ITC.WkType='1'"
	    ObjRS.Open StrSQL, ObjConn
	    Do Until ObjRS.EOF
	      If CONnum <> Trim(ObjRS("ContNo")) Then 
	        strNums = strNums & "," & Trim(ObjRS("ContNo"))
	      End If
	      ObjRS.MoveNext
	    Loop
	    ObjRS.close
	    if err <> 0 then
	      DisConnDBH ObjConn, ObjRS	'DB切断
	      jampErrerP "1","b101","99","実搬出:詳細用データ取得","102","SQL:<BR>"&strSQL
	    end if
        inPutStr="<INPUT type=hidden name='cntnrno' value='"& strNums &"'>"
	Case "4"		'BL
	  inPutStr="<INPUT type=hidden name='blno' value='"& CONnum &"'>"
  End Select

  if Flag=1 Then
	Session.Contents("route") = "輸入コンテナ情報照会（作業選択） "
  Else
	Session.Contents("route") = "Top > 輸入コンテナ情報照会（作業選択） "
  End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>転送中</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT language=JavaScript>
<!--
function opnewin(){
  window.focus();
  document.dmi900F.submit();
}
// -->
</SCRIPT>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY onLoad="opnewin()">
<P>転送中...しばらくお待ちください。</P>
<FORM action="../impcntnr.asp" name="dmi900F">
<%= inPutStr %>
</FORM>
</BODY></HTML>

