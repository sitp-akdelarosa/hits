<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi820.asp				_/
'_/	Function	:事前空搬出CSV入力取込・登録		_/
'_/	Date		:2003/05/30				_/
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
  WriteLogH "b302", "空搬出事前情報入力","06",""

'ユーザデータ所得
  dim USER,COMPcd,tFlag
  USER   = UCase(Session.Contents("userid"))
  COMPcd = UCase(Session.Contents("COMPcd"))

'ファイルよりデータを取得する
  dim aryBinary,strType,nPos,nAsc,sPos,count
  dim strFile  'ファイル拡張子
  dim dataA    'データ

  ' バイナリデータを取得
  aryBinary = Request.BinaryRead(Request.TotalBytes)
  nPos=1
  count=0
  Do
    '一行ずつ読込み
    strType = ""
    'On Error Resume Next
    Do
      ' コンテンツを取得
      nAsc = MidB(aryBinary,nPos,1)
      nAsc = AscB(nAsc)
      If (&h81 <= nAsc And nAsc <= &h9F) Or (&hE0 <= nAsc And nAsc <= &hEF) Then
        strType = strType & Chr(nAsc*256+AscB(MidB(aryBinary,nPos+1,1)))
        nPos = nPos + 1
      Else
        strType = strType & Chr(nAsc)
      End If
      If Right(strType,4) = vbCrLf & vbCrLf Then
        Exit Do
      End If
      nPos = nPos + 1
    Loop While nPos < UBound(aryBinary)
    If nPos = UBound(aryBinary) Then
      Exit Do
    End If
    If count=0 Then
      strFile = Mid(strType,InStr(LCase(strType),"filename=")+9)
      strFile = Mid(strFile,2,InStr(Mid(LCase(strFile),2),"""")-1)
      strFile = Mid(strFile,InStrRev(strFile,".")+1)
    ElseIf count=1 Then
      dataA = Split(strType, vbCrLf , -1, 1)
    End If
    count=count+1
'Response.Write strType & "<P>"
  Loop While nPos < UBound(aryBinary)

  dim ret,tmpA,ret2
  ret = true
  ret2=0
'  If strFile <> "csv" Then
'    ret=false
'    ret2=0
'    tmpA = Array("-","-")	'CW-026 ADD
'  ElseIf InStr(1,dataA(0),",",1) = 0 Then	'CW-027 ADD
  If InStr(1,dataA(0),",",1) = 0 Then	'CW-027 ADD
    ret=false					'CW-027 ADD
    ret2=1					'CW-027 ADD
    tmpA = Array("-","-")			'CW-027 ADD
  Else
    If Left(dataA(0),1)=Chr(10) OR Left(dataA(0),1)=Chr(13)Then
      dataA(0) = Mid(dataA(0),2)
    End If
    'エラートラップ開始
    on error resume next
    'DB接続
    dim ObjConn, ObjRS, StrSQL
    ConnDBH ObjConn, ObjRS

    dim i,CMPcd,FullName,PFlag
    CMPcd = Array("","","","","")

    For i = 0 to UBound(dataA)

    'データチェック
'CW-053      If tmpA(0)= "" Then	'ファイルの終了
      If Trim(dataA(i))= "" Then	'ファイルの終了
        objConn.CommitTrans
        Exit For
      End If
      tmpA = Split(UCase(Trim(dataA(i))), ",", 3, 1)
      checkStr tmpA(0), ret		'ブッキング番号のチェック
      If Not ret Then
	ret2=2
        errerF ObjConn, ObjRS, ret
        Exit For
      End If
'      If tmpA(1)="" Then
'        ret2=3
'        errerF ObjConn, ObjRS, ret
'        Exit For
'      End If
      If tmpA(1)<>"" Then
       'ヘッド会社コードのチェック
        CMPcd(1) = tmpA(1)
        checkHdCd ObjConn, ObjRS, CMPcd, ret
        If Not ret Then
          ret2=4
          errerF ObjConn, ObjRS, ret
          Exit For
        End If
        if err <> 0 then
          ObjRS.Close
          Set ObjRS = Nothing
          jampErrerPDB ObjConn,"1","b302","06","空搬出：CSVデータ登録","102","ヘッドIDチェックに失敗<BR>"&StrSQL
        end if
      End If
    'ブックの重複登録チェック
      checkSPBook ObjConn, ObjRS, tmpA(0), PFlag, ret
      If Not ret Then
        ret2=5
        errerF ObjConn, ObjRS, ret
        Exit For
      End If
      If tmpA(1)<>"" Then		'20031112 add
    '元請陸運業者名取得
        StrSQL = "SELECT FullName FROM mUsers WHERE mUsers.HeadCompanyCode='" & tmpA(1) &"'"
        ObjRS.Open StrSQL, ObjConn
        if err <> 0 then
          ObjRS.Close
          Set ObjRS = Nothing
          jampErrerPDB ObjConn,"1","b302","06","空搬出：CSVデータ登録","102","元請陸運業者名取得に失敗<BR>"&StrSQL
        end if
        FullName = ObjRS("FullName")
        ObjRS.close
      End If 				'20031112 add
    '登録
'CW-052 ADD Start
      If tmpA(1) = COMPcd Then 
        tFlag=1
      Else
        tFlag=0
      End If
'CW-052 ADD END

      If PFlag="0" Then
        StrSQL = "Insert Into SPBookInfo (BookNo, SenderCode, UpdtTime, UpdtPgCd, UpdtTmnl, Status,"&_
                 " Process, InputDate, TruckerCode, TruckerFlag, TruckerName ) "&_
                 "values ('"& tmpA(0) &"','"& USER &"','"& Now() &"','PREDEF01','"& USER &"','0',"&_
                 "'R','"& Now() &"','"& tmpA(1) &"','"& tFlag &"','"& FullName &"')"
        ObjConn.Execute(StrSQL)
        if err <> 0 then
          Set ObjRS = Nothing
          jampErrerPDB ObjConn,"1","b302","06","空搬出：CSVデータ登録","103","データ登録に失敗<BR>"&StrSQL
        end if
      Else
        StrSQL = "UPDATE SPBookInfo SET SenderCode='"& USER &"', UpdtTime='"& Now() &"', UpdtPgCd='PREDEF01', "&_
                 "UpdtTmnl='"& USER &"', Status='0', Process='R', InputDate='"& Now() &"', "&_
                 "TruckerCode='"& tmpA(1) &"', TruckerFlag='"& tFlag &"', TruckerName='"& FullName &"' "&_
                 "WHERE BookNo='"& tmpA(0) &"' "
        ObjConn.Execute(StrSQL)
        if err <> 0 then
          Set ObjRS = Nothing
          jampErrerPDB ObjConn,"1","b302","06","空搬出：CSVデータ登録","104","データ登録に失敗<BR>"&StrSQL
        end if
      End If
    Next
  'DB接続解除
    DisConnDBH ObjConn, ObjRS
  'エラートラップ解除
    on error goto 0
  End If

  dim tmpstr
  If ret Then
    tmpstr=",入力内容の正誤:0(正しい)"
  Else
    tmpstr=",入力内容の正誤:1(誤り)"
  End If
  WriteLogH "b302", "空搬出事前情報入力","06",tmpA(0)&"/"&tmpA(1)&tmpstr

Function checkStr(str, ret)
  dim checkChr,i,checkF
  checkChr="0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ- /"
'CW-054  If Len(str) = 0 Then
  If Len(str) = 0 Or Len(str) > 21 Then
      ret = false
      Exit Function
  End If 
  For i= 1 To Len(str)
    If InStr(1,checkChr,Mid(str,i,1),1) = 0 Then
      ret = false
      Exit Function
    End If
  Next
End Function

Function errerF(ObjRS, StrSQL, ret)
  ObjConn.RollbackTrans	'ロールバック
  ret = false
End Function
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>事前空搬出CSV入力</TITLE>
<SCRIPT language=JavaScript>
<!--
window.resizeTo(600,400);
// -->
</SCRIPT>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY>
<!-------------事前空搬出CSV入力--------------------------->
<P><B>事前空搬出CSV入力処理</B></P>
<CENTER>
<%If ret Then %>
  <P><%=i%>件登録しました</P>
  <INPUT type=button onClick="window.close()" value="閉じる">
<% Else %>
<P><DIV class=alert>エラー<P>
  <% Select Case ret2
       Case "0" %>
      指定されたファイルの拡張子がCSVではありません。
  <%   Case "1" %>
      指定されたファイルはNullまたはフォーマットが不正です。
  <%   Case "2" %>
      <%=i+1%>番目のブッキング番号が不正です。<BR>
      Nullまたは21桁以上か不正な文字含まれています。<BR>
      修正後もう一度やり直してください。
  <%   Case "3" %>
      <%=i+1%>番目の会社コードが指定されていません。<BR>修正後もう一度やり直してください。
  <%   Case "4" %>
      <%=i+1%>番目の会社コードは存在しません。<BR>修正後もう一度やり直してください。
  <%   Case "5" %>
      <%=i+1%>番目のブッキング番号は既に登録されています。<BR>修正後もう一度やり直してください。
  <%   End Select %>
</DIV></P>
  <INPUT type=button onClick="window.history.back();" value="戻る">
<% End If%>
</CENTER>
</BODY>
</HTML>
