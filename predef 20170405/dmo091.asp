<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmo091.asp				_/
'_/	Function	:事前実搬出指示書印刷調整画面		_/
'_/	Date		:2004/01/31				_/
'_/	Code By		:SEIKO Electric.Co 大重			_/
'_/	Modify		:								_/
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
'ログ出力
  WriteLogH "b109", "実搬出指示書印刷", "01",""

'サーバ日時の取得
  dim DayTime,day
  getDayTime DayTime
  day = DayTime(0) & "年" & DayTime(1) & "月" & DayTime(2) & "日" 

'前画面からのデータ取得
  dim Flag,UpFlag,Num,CONnumA(),CMPcd(5),HedId,RDate,NoDate
  dim param,i,j,Way
  dim YY,Rmon,Rday,Rhou,Nomon,Noday,Nohou,Nomin,NonyuDate
  Way   =Array("","指定あり","指定なし","一覧から選択","ＢＬ番号")
  Flag= Request("flag")
  Num = Request("num")
  UpFlag=Request("UpFlag")
  If Request("HedId")= "" OR Request("HedId") = Null Then
    HedId="　　　　　　　　　　　　　　"
  Else
    HedId=Request("HedId")
  End IF

'日の整形
  Rmon    = Right("00" & Request("Rmon") ,2)
  Rday    = Right("00" & Request("Rday") ,2)
  If Request("Rhou") = "" OR Request("Rhou") = Null Then
    Rhou =""
  Else
'Chang 20050303 STAT fro Fourth Recon By SEIKO N.Oosige
'    Rhou = Right("00" & Request("Rhou") ,2) & ":00"
    Rhou = Right("00" & Request("Rhou") ,2) & "時"
'Chang 20050303 END
  End If
  Nomon   = Right("00" & Request("Nomon") ,2)
  Noday   = Right("00" & Request("Noday") ,2)
  If Request("Nohou") = "" OR Request("Nohou") = Null Then
    Nohou = ""
  Else
'Chang 20050303 STAT fro Fourth Recon By SEIKO N.Oosige
'    Nohou = Right("00" & Request("Nohou") ,2) & ":00"
    Nohou = Right("00" & Request("Nohou") ,2) & "時"
'Chang 20050303 END
  End If
'2008/01/31 Add S G.Ariola  
  If Request("Nomin") = "" OR Request("Nomin") = Null Then
    Nomin = ""
  Else
    Nomin = Right("00" & Request("Nomin") ,2) & "分"
  End If
'2008/01/31 Add E G.Ariola  

  If DayTime(1) > Rmon Then	'来年
    YY = DayTime(0) +1
  ElseIf DayTime(1) = Rmon AND DayTime(2) > Rday Then
    YY = DayTime(0) +1
  Else
    YY = DayTime(0)
  End If
  If Rmon = "00" Or Rday = "00" Then
    RDate= ""
  Else
    RDate= YY &"年"& Rmon &"月"& Rday &"日　"& Rhou
  End If

  If DayTime(1) > Nomon Then	'来年
    YY = DayTime(0) +1
  ElseIf DayTime(1) = Nomon AND DayTime(2) > Noday Then
    YY = DayTime(0) +1
  Else
    YY = DayTime(0)
  End If
  If Nomon = "00" Or Noday = "00" Then
    NoDate= ""
  Else
  '2008/01/31 Edit S G.Ariola
    'NoDate= YY &"年"& Nomon &"月"& Noday &"日　"& Nohou
    NoDate= YY &"年"& Nomon &"月"& Noday &"日　"& Nohou &""& Nomin
  '2008/01/31 Edit S G.Ariola
  End If
  
  ReDim CONnumA(Num)
  i=0
  For Each param In Request.Form
    If Left(param, 6) = "CONnum" Then
      CONnumA(i) = Request.Form(param)
      i=i+1
    ElseIf Left(param,5) = "CMPcd" Then
      j = Right(param,1)
      CMPcd(j) = Request.Form(param)
    End If
  Next
'セッションからユーザ名称を取得
  Dim SjManN
  SjManN = Session.Contents("LinUN")

'DBからのデータ取得
  'エラートラップ開始
  on error resume next
  'DB接続
  dim ObjConn, ObjRS, StrSQL
  ConnDBH ObjConn, ObjRS

  '作業者名称取得
  Dim WkManN
  If CMPcd(UpFlag)="" OR CMPcd(UpFlag)=Null Then
    WkManN=SjManN
  Else
    StrSQL = "Select FullName From mUsers Where HeadCompanyCode='" & CMPcd(UpFlag) &"'"
    ObjRS.Open StrSQL, ObjConn
    if err <> 0 then
      DisConnDBH ObjConn, ObjRS	'DB切断
      jampErrerP "1","b109","01","実搬出指示書印刷調整・作業者名取得","102","SQL:<BR>"&strSQL
    end if
    WkManN= Trim(ObjRS("FullName"))
    ObjRS.close
  End If
'指示者電話番号取得
  dim USER,TelNo
  USER       = Session.Contents("userid")
  StrSQL = "select TelNo from mUsers where UserCode='" & USER &"'"
  ObjRS.Open StrSQL, ObjConn
  if err <> 0 then
    DisConnDBH ObjConn, ObjRS	'DB切断
    jampErrerP "1","b109","01","実搬出指示書印刷調整・指示者電話番号取得","102","SQL:<BR>"&strSQL
  end if
  TelNo = Trim(ObjRS("TelNo"))
  ObjRS.close
  If TelNo<>"" Then
    TelNo="（電話番号："&TelNo&"）"
  End If

'コンテナデータ取得
  Dim ConInfo
  ReDIm ConInfo(Num)
  Select Case Flag
    Case "1"			'指示あり
      Num=1
      ConInfo(0)=Array(CONnumA(0),Request("CONsize"),Request("CONtype"),Request("CONhite"),Request("CONtear"))
    Case "2" 			'指定なし
      '対象取得
      StrSQL = "SELECT Cnt.ContNo,Cnt.ContSize, Cnt.ContType, Cnt.ContHeight, Cnt.ContWeight "&_
               "From (ImportCont AS INC1 INNER JOIN ImportCont AS INC2 ON "&_
               "(INC1.VoyCtrl = INC2.VoyCtrl) AND (INC1.VslCode = INC2.VslCode) AND (INC1.BLNo = INC2.BLNo)) "&_
               "INNER JOIN Container AS Cnt "&_
               "ON INC2.ContNo=Cnt.ContNo AND INC2.VslCode=Cnt.VslCode AND INC2.VoyCtrl=Cnt.VoyCtrl "&_
               "WHERE INC1.ContNo='" & CONnumA(0) & "' AND INC1.BLNo= '"& Request("BLnum") &"' " &_
               "ORDER BY INC2.ContNo ASC, INC2.UpdtTime DESC"
      ObjRS.Open StrSQL, ObjConn
      if err <> 0 then
        DisConnDBH ObjConn, ObjRS	'DB切断
        jampErrerP "1","b109","01","実搬出指示書印刷調整・個別情報取得","102","SQL:<BR>"&strSQL
      end if
      i=1
      ConInfo(0)=Array("","","","","")
      Do Until ObjRS.EOF
        If CONnumA(0) = Trim(ObjRS("ContNo")) Then
          If ConInfo(0)(0)<>Trim(ObjRS("ContNo")) Then 
            ConInfo(0)(0)=Trim(ObjRS("ContNo"))
            ConInfo(0)(1)=Trim(ObjRS("ContSize"))
            ConInfo(0)(2)=Trim(ObjRS("ContType"))
            ConInfo(0)(3)=Trim(ObjRS("ContHeight"))
            ConInfo(0)(4)=Trim(ObjRS("ContWeight"))*100
          End If
        Else
          If ConInfo(i-1)(0)<>Trim(ObjRS("ContNo")) Then
          ReDim Preserve ConInfo(i)
            ConInfo(i)=Array("","","","","")
            ConInfo(i)(0)=Trim(ObjRS("ContNo"))
            ConInfo(i)(1)=Trim(ObjRS("ContSize"))
            ConInfo(i)(2)=Trim(ObjRS("ContType"))
            ConInfo(i)(3)=Trim(ObjRS("ContHeight"))
            ConInfo(i)(4)=Trim(ObjRS("ContWeight"))*100
            i=i+1
          End If
        End If
        ObjRS.MoveNext
      Loop
      ObjRS.close
      Num=i
    Case "3" 			'一覧
      Dim strConNums
      strConNums="'"& CONnumA(0) &"'"
      For i = 1 to Num-1
        strConNums=strConNums &",'"& CONnumA(i) &"'"
      Next
      '対象件数取得
      StrSQL = "SELECT Cnt.ContNo,Cnt.ContSize, Cnt.ContType, Cnt.ContHeight, Cnt.ContWeight "&_
               "From Container AS Cnt Where Cnt.ContNo In("& strConNums &") " &_
               "ORDER BY Cnt.ContNo ASC, Cnt.UpdtTime DESC"
      ObjRS.Open StrSQL, ObjConn
      if err <> 0 then
        DisConnDBH ObjConn, ObjRS	'DB切断
        jampErrerP "1","b109","01","実搬出指示書印刷調整・個別情報取得","102","SQL:<BR>"&strSQL
      end if
      i=1
      ConInfo(0)=Array("","","","","")
      Do Until ObjRS.EOF
        If CONnumA(0) = Trim(ObjRS("ContNo")) Then
          If ConInfo(0)(0)<>Trim(ObjRS("ContNo")) Then 
            ConInfo(0)(0)=Trim(ObjRS("ContNo"))
            ConInfo(0)(1)=Trim(ObjRS("ContSize"))
            ConInfo(0)(2)=Trim(ObjRS("ContType"))
            ConInfo(0)(3)=Trim(ObjRS("ContHeight"))
            ConInfo(0)(4)=Trim(ObjRS("ContWeight"))*100
          End If
        Else
          If ConInfo(i-1)(0)<>Trim(ObjRS("ContNo")) Then
            ConInfo(i)=Array("","","","","")
            ConInfo(i)(0)=Trim(ObjRS("ContNo"))
            ConInfo(i)(1)=Trim(ObjRS("ContSize"))
            ConInfo(i)(2)=Trim(ObjRS("ContType"))
            ConInfo(i)(3)=Trim(ObjRS("ContHeight"))
            ConInfo(i)(4)=Trim(ObjRS("ContWeight"))*100
            i=i+1
          End If
        End If
        ObjRS.MoveNext
      Loop
      ObjRS.close
    Case "4" 			'指定なし又はBL
      dim VslCode,VoyCtrl
      '対象BL選定
      StrSQL = "SELECT INC.VslCode, INC.VoyCtrl "&_
               "From ImportCont AS INC  "&_
               "Where INC.BLNo= '"& Request("BLnum") &"' ORDER BY INC.UpdtTime DESC"
      ObjRS.Open StrSQL, ObjConn
      if err <> 0 then
        DisConnDBH ObjConn, ObjRS	'DB切断
        jampErrerP "1","b109","01","実搬出指示書印刷調整・個別情報取得","102","SQL:<BR>"&strSQL
      end if
      VslCode=Trim(ObjRS("VslCode"))
      VoyCtrl=Trim(ObjRS("VoyCtrl"))
      ObjRS.close
      '対象データ取得
      StrSQL = "SELECT Cnt.ContNo,Cnt.ContSize, Cnt.ContType, Cnt.ContHeight, Cnt.ContWeight "&_
               "From ImportCont AS INC INNER JOIN Container AS Cnt "&_
               "ON INC.ContNo=Cnt.ContNo AND INC.VslCode=Cnt.VslCode AND INC.VoyCtrl=Cnt.VoyCtrl "&_
               "Where INC.BLNo= '"& Request("BLnum") &"' AND INC.VslCode= '"& VslCode &"' AND INC.VoyCtrl= '"& VoyCtrl &"' " &_
               "ORDER BY INC.ContNo ASC, INC.UpdtTime DESC"
      ObjRS.Open StrSQL, ObjConn
      if err <> 0 then
        DisConnDBH ObjConn, ObjRS	'DB切断
        jampErrerP "1","b109","01","実搬出指示書印刷調整・個別情報取得","102","SQL:<BR>"&strSQL
      end if
      If Flag="4" Then	'BL
        CONnumA(0)=Trim(ObjRS("ContNo"))
      End If
      i=1
      ConInfo(0)=Array("","","","","")
      Do Until ObjRS.EOF
        If CONnumA(0) = Trim(ObjRS("ContNo")) Then
          If ConInfo(0)(0)<>Trim(ObjRS("ContNo")) Then 
            ConInfo(0)(0)=Trim(ObjRS("ContNo"))
            ConInfo(0)(1)=Trim(ObjRS("ContSize"))
            ConInfo(0)(2)=Trim(ObjRS("ContType"))
            ConInfo(0)(3)=Trim(ObjRS("ContHeight"))
            ConInfo(0)(4)=Trim(ObjRS("ContWeight"))*100
          End If
        Else
          If ConInfo(i-1)(0)<>Trim(ObjRS("ContNo")) Then
          ReDim Preserve ConInfo(i)
            ConInfo(i)=Array("","","","","")
            ConInfo(i)(0)=Trim(ObjRS("ContNo"))
            ConInfo(i)(1)=Trim(ObjRS("ContSize"))
            ConInfo(i)(2)=Trim(ObjRS("ContType"))
            ConInfo(i)(3)=Trim(ObjRS("ContHeight"))
            ConInfo(i)(4)=Trim(ObjRS("ContWeight"))*100
            i=i+1
          End If
        End If
        ObjRS.MoveNext
      Loop
      ObjRS.close
      Num=i
  End Select

  'DB接続解除
  DisConnDBH ObjConn, ObjRS
  'エラートラップ解除
  on error goto 0

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>指示書印刷調整</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
window.focus();
<% If Num>1 Then %>
//全てのチェックを外す
function clearCeck(){
  target=document.dmo091F;
  len=target.elements.length;
  for(i=0;i<len;i++){
    if(target.elements[i].type=="checkbox")
      target.elements[i].checked=false;
  }
}
<% End If %>
//指示書印刷画面へ
function GoNext(){
  target=document.dmo091F;
<% If Num>1 Then %>
  len=target.elements.length;
  checkFlag=0;
  checkedstrs="";
  for(i=0;i<len;i++){
    if(target.elements[i].type=="checkbox")
      if(target.elements[i].checked==true){
        checkFlag++;
        checkedstrs=checkedstrs+target.elements[i].name+",";
      }
  }
  if(checkFlag==0){
    alert("どれか一つのコンテナには必ずチェックを付けてください");
    return;
  }else{
    target.checkNum.value=checkFlag;
    target.checkeds.value=checkedstrs;
  }
<% End If %>
  newWin = window.open("", "Print2", "width=650,height=700,left=30,top=10,resizable=yes,scrollbars=yes,menubar=yes,top=0");
  target.target="Print2";

  target.submit();
}
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY>
<!-------------実搬出指示書印刷調整画面--------------------------->
<FORM name="dmo091F" method="POST" action="./dmo092.asp";>
<CENTER><B class=titleB>実搬出指示書</B></CENTER>
<DIV class=right>作成&nbsp;<%=day%></DIV>
<INPUT type=hidden name="day" value="<%=day%>">
<TABLE border=0 cellPadding=1 cellSpacing=0>
  <TR>
    <TD>作業番号</TD><TD>＝<%=Request("SakuNo")%></TD><TD></TD></TR>
  <TR>
    <TD valign=top>指示者</TD><TD valign=top>＝<%=SjManN%></TD>
    <TD>（担当者：　　　　　　　　　　　　　　　）<BR>
        <%=TelNo%></TD></TR>
  <TR>
    <TD>作業者</TD><TD>＝<INPUT type=text name="WkManN" value="<%=WkManN%>"></TD>
    <TD>(ヘッドＩＤ＝<%=HedId%>）</TD></TR>
  <TR>
    <TD>指定方法</TD><TD>＝<%=Way(Flag)%></TD><TD></TD></TR>
</TABLE><P>
<INPUT type=hidden name="SakuNo" value="<%=Request("SakuNo")%>">
<INPUT type=hidden name="SjManN" value="<%=SjManN%>">
<INPUT type=hidden name="HedId" value="<%=HedId%>">
<INPUT type=hidden name="Way" value="<%=Way(Flag)%>">
<INPUT type=hidden name="TelNo" value="<%=TelNo%>">
<TABLE border=0 cellPadding=1 cellSpacing=0>
  <TR><TH rowspan=4 valign=top>１．</TH>
    <TD><B>コンテナ情報</B>&nbsp;</TD><TD></TD></TR>
  <TR>
    <TD>（船社）</TD><TD><%=Request("shipFact")%></TD></TR>
  <TR>
    <TD>（船名）</TD><TD><%=Request("shipName")%></TD></TR>
  <TR>
    <TD>（品名）</TD><TD><%=Request("HinName")%></TD></TR>
</TABLE><P>
<INPUT type=hidden name="shipFact" value="<%=Request("shipFact")%>">
<INPUT type=hidden name="shipName" value="<%=Request("shipName")%>">
<INPUT type=hidden name="HinName"  value="<%=Request("HinName")%>">
<TABLE border=0 cellPadding=1 cellSpacing=0>
  <TR><TH rowspan=6 valign=top>２．</TH>
    <TD><B>搬出情報</B></TD><TD></TD></TR>
  <TR>
    <TD>（ＣＹ）</TD><TD><%=Request("Hfrom")%></TD></TR>
  <TR>
    <TD>（搬出予定日時）&nbsp;</TD><TD><%=RDate%></TD></TR>
  <TR>
    <TD>（納入先１）</TD><TD><%=Request("Nonyu1")%></TD></TR>
  <TR>
    <TD>（納入先２）</TD><TD><%=Request("Nonyu2")%></TD></TR>
  <TR>
    <TD>（納入日時分）</TD><TD><%=NoDate%></TD></TR>
</TABLE><P>
<INPUT type=hidden name="Hfrom"  value="<%=Request("Hfrom")%>">
<INPUT type=hidden name="RDate"  value="<%=RDate%>">
<INPUT type=hidden name="Nonyu1" value="<%=Request("Nonyu1")%>">
<INPUT type=hidden name="Nonyu2" value="<%=Request("Nonyu2")%>">
<INPUT type=hidden name="NoDate" value="<%=NoDate%>">
<TABLE border=0 cellPadding=1 cellSpacing=0>
  <TR><TH rowspan=3 valign=top>３．</TH>
    <TD><B>空コン返却情報</B></TD><TD></TD></TR>
  <TR>
    <TD>（返却先）</TD><TD><%=Request("RPlace")%></TD></TR>
  <TR>
    <TD>（返却予定日数）&nbsp;</TD><TD><%=Request("Rnissu")%></TD></TR>
</TABLE><P>
<INPUT type=hidden name="RPlace"  value="<%=Request("RPlace")%>">
<INPUT type=hidden name="Rnissu"  value="<%=Request("Rnissu")%>">
<TABLE border=0 cellPadding=1 cellSpacing=0>
  <TR><TH rowspan=3 valign=top>４．</TH>
    <TD><B>備考</B></TD><TD></TD></TR>
  <TR>
    <TD>（備考１）&nbsp;</TD><TD><%=Request("Comment1")%></TD></TR>
  <TR>
    <TD>（備考２）</TD><TD><%=Request("Comment2")%></TD></TR>
</TABLE><P>
<INPUT type=hidden name="Comment1"  value="<%=Request("Comment1")%>">
<INPUT type=hidden name="Comment2"  value="<%=Request("Comment2")%>">
<TABLE border=0 cellPadding=1 cellSpacing=0>
  <TR><TH rowspan=<%=Num+3%> valign=top>５．</TH>
    <TD colspan=7><B>コンテナ番号</B></TD></TR>
  <TR><TD width=20></TD><TD></TD><TD>&nbsp;コンテナ番号&nbsp;</TD><TD>&nbsp;サイズ&nbsp;</TD>
      <TD>&nbsp;タイプ&nbsp;</TD><TD>&nbsp;高さ&nbsp;</TD><TD>&nbsp;グロス&nbsp;</TD>
  <TR align=center><TD></TD>
    <TD><% If Num>1 Then Response.Write "<INPUT type='checkbox' name=No0 checked>" Else Response.Write "　" End If %></TD>
    <TD><%=ConInfo(0)(0)%></TD><TD><%=ConInfo(0)(1)%>'</TD><TD><%=ConInfo(0)(2)%></TD>
    <TD><%=ConInfo(0)(3)%></TD><TD><%=ConInfo(0)(4)%>kg</TD></TR>
<% For i=1 To Num-1 %>
  <TR align=center><TD></TD>
    <TD><% If Num>1 Then Response.Write "<INPUT type='checkbox' name=No"&i&" checked>" Else Response.Write "　" End If %></TD>
    <TD><%=ConInfo(i)(0)%></TD><TD><%=ConInfo(i)(1)%>'</TD><TD><%=ConInfo(i)(2)%></TD>
    <TD><%=ConInfo(i)(3)%></TD><TD><%=ConInfo(i)(4)%>kg</TD></TR>
<%Next%>
  <TR><TD colspan=7>
<% If Num>1 Then %>
  <A HREF="JavaScript:clearCeck()">すべてのチェックを外す</A>
<% End If %>
      </TD></TR>
</TABLE><P>
<INPUT type=hidden name="checkNum"  value="">
<INPUT type=hidden name="checkeds"  value="">
<%Set_Data Num-1,ConInfo%>
<CENTER>
  <INPUT type=button value="ＯＫ" onClick="GoNext()">
  <INPUT type=button value="閉じる" onClick="window.close()">
</CENTER>
</FORM>
<!-------------画面終わり--------------------------->
</BODY></HTML>
