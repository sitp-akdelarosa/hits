<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' Tempファイル属性のチェック
    CheckTempFile "IMPORT", "impentry.asp"

    ' 指定引数の取得
    Dim iLineNo
    iLineNo = CInt(Request.QueryString("line"))
    Dim iReturn
    iReturn = Session.Contents("dispreturn")

    ' 表示モードの取得
    Dim bDispMode          ' true=コンテナ検索 / false=BL検索
    If Session.Contents("findkind")="Cntnr" Then
        bDispMode = true
    Else
        bDispMode = false
    End If

'☆☆☆ Add by nics 2010.02.02
    Dim USER
	USER    = Session.Contents("userid")
'☆☆☆ end of Add by nics 2010.02.02

    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' 表示ファイルの取得
    Dim strFileName
    strFileName = Session.Contents("tempfile")
    If strFileName="" Then
        ' セッションが切れているとき
        Response.Redirect "impentry.asp"             '輸入コンテナ照会トップ
        Response.End
    End If
    strFileName="./temp/" & strFileName

    ' 表示ファイルのOpen
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

    ' 詳細表示行のデータの取得
    LineNo=0
    Do While Not ti.AtEndOfStream
        anyTmp=Split(ti.ReadLine,",")
        LineNo=LineNo+1
        If iLineNo=LineNo Then
           Exit Do
        End If
    Loop
    bSingle = false                    '単独検索結果フラグ
    If iLineNo=1 And LineNo=1 Then
        '単独検索結果かどうかチェックする
        if ti.AtEndOfStream Then
            bSingle = true
        End If
    End If
    ti.Close

    ' 輸入コンテナ照会詳細
	WriteLog fs, "2006","輸入コンテナ照会-単独コンテナ", "00", anyTmp(1) & ","

    Session.Contents("dispcntnr")=anyTmp(1)     ' 表示コンテナNo.を記憶
%>

<html>
<head>
<title></title>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<link href="./index.css" rel="stylesheet" type="text/css">
<SCRIPT language="javascript" type="text/javascript" src="./index.js"></SCRIPT>
<SCRIPT Language="JavaScript">
<!--
function Submit(formName){
    document.forms[formName].submit();
}
// -->
<%
    DispMenuJava
%>
</SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
function winOpen(winName,url,W,H){
  var WinD11=window.open(url,winName,'scrollbars=yes,resizable=yes,width='+W+',height='+H+'');
  WinD11.focus();
  WinD11.document.close();
}
</Script>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#0000ff" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" vlink="#0000ff">
<!-------------ここから詳細画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/impdetailt.gif" width="506" height="73"></td>
          <td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
        </tr>
        <tr>
          <td align="right" width="100%" height="48">
<%
' Added and Commented by seiko-denki 2003.07.17
	DisplayCodeListButton
'    DispMenu
'	Dim strScriptName,strRoute
'	strScriptName = Request.ServerVariables("SCRIPT_NAME")
'	strRoute = SetRoute(strScriptName)
'	Session.Contents("route") = strRoute
' End of Addition by seiko-denki 2003.07.17
%>
          </td>
        </tr>
      </table>
      <center>
<!-- commented by seiko-denki 2003.07.17
		<table width=95% cellpadding="0" cellspacing="0" border="0">
		  <tr>
			<td align="right">
			  <font color="#333333" size="-1">
				<%=strRoute%>
			  </font>
			</td>
		  </tr>
		</table>
end of comment by seiko-denki 2003.07.07 -->

<!-- mod by nics 2009.02.09 -->
<!--		<table width=95% cellpadding=3>-->
		<table width=95% cellpadding=0>
<!-- end of mod by nics 2009.02.09 -->
			<tr>
				<td align=right>
					<font color="#224599">
					&nbsp;&nbsp;<%=GetUpdateTime(fs)%>
					</font>
				</td>
			</tr>
		</table>

      <table>
        <tr>
          <td>
            <table border=1 cellpadding="3" cellspacing="1">
              <tr> 
<% ' BL No
    If Not bDispMode Then
        Response.Write "<td bgcolor='#003399' background='gif/tableback.gif' nowrap><font color='#FFFFFF'><b>BL No</b></font></td>"
        Response.Write "<td bgcolor='#FFFFFF' nowrap>" & anyTmp(0) & "</td>"
    End If
%>
                <td bgcolor="#003399" background="gif/tableback.gif" nowrap><font color="#FFFFFF"><b>コンテナNo</b></font></td>
                <td bgcolor="#FFFFFF" nowrap>
<% ' コンテナNo.
    Response.Write anyTmp(1)
%>
                </td>
              </tr>
            </table>
<!-- mod by nics 2009.02.09 -->
<!--			<BR>-->
			<font size="-2"><BR></font>
<!-- end of mod by nics 2009.02.09 -->
<!---------------基本情報------------------------------------------- commented by nics 2009.02.09 -->
<!---------------基本情報--------------------------------------------->
<!-- commented by nics 2009.02.09
            <table>
              <tr>
                <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>基本情報</b></td>
                <td><img src="gif/hr.gif"></td>
              </tr>
            </table>
end of comment by nics 2009.02.09 -->
            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center" bgcolor="#FFCC33"> 
<!-- add by nics 2009.02.09 -->
                <td rowspan="5" nowrap bgcolor="#6495ED">&nbsp;基本情報&nbsp;</td>
<!-- end of add by nics 2009.02.09 -->
<!-- commented by nics 2009.02.09
                <td valign="top" nowrap>項目</td>
end of comment by nics 2009.02.09 -->
                <td nowrap bgcolor="#FFCC33">サイズ</td>
<!-- mod by nics 2009.02.09 -->
<!-- Add-S MES Aoyagi 2010.11.23 -->
		<td nowrap bgcolor="#FFCC33">タイプ</td>
<!-- Add-E MES Aoyagi 2010.11.23 -->
<!--                <td nowrap bgcolor="#FFCC33">高さ<font size="-1"><sup>(※4)</sup></font></td>-->
                <td nowrap bgcolor="#FFCC33">高さ<font size="-1"><sup>(※1)</sup></font></td>
<!-- end of mod by nics 2009.02.09 -->
                <td nowrap bgcolor="#FFCC33">リーファ</td>
                <td nowrap bgcolor="#FFCC33">総重量(t)</td>
<!-- mod by nics 2009.02.09 -->
<!--                <td valign="top" nowrap>危険品<font size="-1"><sup>(※5)</sup></font></td>-->
                <td valign="top" nowrap>危険品<font size="-1"><sup>(※2)</sup></font></td>
<!-- end of mod by nics 2009.02.09 -->
<!-- mod by nics 2009.02.09 -->
<!--                <td nowrap bgcolor="#FFCC33">搬出ターミナル</td>-->
                <td nowrap bgcolor="#FFCC33">搬出ターミナル(蔵置場所コード)</td>
<!-- end of mod by nics 2009.02.09 -->
<!-- add by nics 2009.02.09 -->
<!-- del-s 2009/11/01 とりあえず消すことに -->
                <td nowrap bgcolor="#FFCC33">本船担当オペレータ</td>
<!--del-e 2009/11/01 とりあえず消すことに -->
<!-- end of add by nics 2009.02.09 -->
                <td nowrap bgcolor="#FFCC33">ストックヤード利用</td>
                <td nowrap bgcolor="#FFCC33">返却場所</td>
              </tr>
              <tr align="center"> 
<!-- commented by nics 2009.02.09
                <td bgcolor="#FFFFCC" nowrap>情報</td>
end of comment by nics 2009.02.09 -->
                <td align="center" nowrap>
<% ' サイズ
    If anyTmp(23)<>"" Then
        Response.Write anyTmp(23)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" nowrap>

<!-- Add-S MES Aoyagi 2010.11.23 コンテナタイプも表示 -->
<% ' タイプ
    If anyTmp(46)<>"" Then
        Response.Write anyTmp(46)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" nowrap>
<!-- Add-E MES Aoyagi 2010.11.23 コンテナタイプも表示 -->

<% ' 高さ
    If anyTmp(24)<>"" Then
        Response.Write anyTmp(24)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' リーファー
    If anyTmp(25)="R" Then
        Response.Write "○"
    ElseIf anyTmp(25)<>"" Then
        Response.Write "－"
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' 総重量
    If anyTmp(26)<>"" And anyTmp(26)<>"0" Then
        dWeight=anyTmp(26) / 10
        Response.Write dWeight
    Else
        Response.Write "－"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' 危険物
    If anyTmp(27)="H" Then
        Response.Write "○"
    ElseIf anyTmp(27)<>"" Then
        Response.Write "－"
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- commented by nics 2009.02.09
                <td align="center" nowrap>
<% ' 搬出ターミナル
    If anyTmp(5)<>"" Then
        Response.Write anyTmp(5)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
end of comment by nics 2009.02.09 -->
<!-- add by nics 2009.02.09 -->
                <td align="center" nowrap>
<% ' 搬出ターミナル(蔵置場所コード)
    strDisp = "<br>"
    If anyTmp(5) <> "" Then
        strDisp = anyTmp(5)
'' 2009/11/01 del-s 蔵置場所コードをとりあえず消すことに
        If anyTmp(43) <> "" Then
            strDisp = strDisp & "(" & anyTmp(43) & ")"
        End If
'' 2009/11/01 del-e 蔵置場所コードをとりあえず消すことに
    End If
    Response.Write strDisp
%>
                </td>
<!-- del-s 2009/11/01 とりあえず消すことに -->
                <td align="center" nowrap>
<% ' 担当オペレータ
    If anyTmp(45)<>"" Then
        Response.Write anyTmp(45)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- del-s 2009/11/01 とりあえず消すことに -->
<!-- end of add by nics 2009.02.09 -->
                <td align="center" nowrap>
<% ' ストックヤード利用 $追加
    If anyTmp(35)>="1" And anyTmp(35)<="4" Then
        Response.Write "○"
    Else
        Response.Write "×"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' 返却場所
    If anyTmp(10)<>"" Then
        Response.Write anyTmp(10)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
              </tr>
            </table>
<!-- mod by nics 2009.02.09 -->
<!--            <table border="0" cellspacing="1" cellpadding="3">-->
            <table border="0" cellspacing="0" cellpadding="0">
<!-- end of mod by nics 2009.02.09 -->
              <tr> 
                <td width="15">&nbsp;</td>
<!-- mod by nics 2009.02.09 -->
<!--                <td nowrap><font color="#000000" size="-1">（※4）96=HC</td>-->
                <td nowrap><font color="#000000" size="-1">（※1）96=HC</td>
<!-- end of mod by nics 2009.02.09 -->
                <td width="15">&nbsp;</td>
<!-- mod by nics 2009.02.09 -->
<!--                <td nowrap><font color="#000000" size="-1">（※5）消防法に関わる危険品の有無</td>-->
                <td nowrap><font color="#000000" size="-1">（※2）消防法に関わる危険品の有無</td>
<!-- end of mod by nics 2009.02.09 -->
              </tr>
            </table>
<!-- commented by nics 2009.02.09
            <br>
end of comment by nics 2009.02.09 -->
<!---------------本船情報------------------------------------------- commented by nics 2009.02.09 -->
<!---------------本船情報--------------------------------------------->
<!-- commented by nics 2009.02.09
            <table>
              <tr> 
                <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>本船情報&nbsp;&nbsp;</b></td>
                <td><img src="gif/hr.gif"></td>
              </tr>
            </table>
end of comment by nics 2009.02.09 -->
            <table border=1 cellpadding="3" cellspacing="1">
<!-- mod by nics 2009.02.09 -->
<!--              <tr> -->
              <tr align="center" bgcolor="#FFCC33"> 
<!-- end of mod by nics 2009.02.09 -->
<!-- add by nics 2009.02.09 -->
                <td rowspan="2" nowrap bgcolor="#6495ED">&nbsp;本船情報&nbsp;</td>
<!-- end of add by nics 2009.02.09 -->
                <td bgcolor="#FFCC33" nowrap><font color="#000000">船社</font></td>
<!-- add by nics 2009.02.09 -->
                <td bgcolor="#FFCC33" nowrap><font color="#000000">船名</font></td>
                <td bgcolor="#FFCC33" nowrap>Voyage No.<font color="#FFFFFF"><b> 
                </b></font></td>
                <td bgcolor="#FFCC33" nowrap>仕出港</td>
                <td bgcolor="#FFCC33" nowrap>前港</td>
              </tr>
              <tr align="center"> 
<!-- end of add by nics 2009.02.09 -->
                <td bgcolor="#FFFFFF" nowrap>
<% ' 船社
    If anyTmp(6)<>"" Then
        Response.Write anyTmp(6)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- commented by nics 2009.02.09
                <td bgcolor="#FFCC33" nowrap><font color="#000000">船名</font></td>
end of comment by nics 2009.02.09 -->
                <td bgcolor="#FFFFFF" nowrap>
<% ' 船名
    If anyTmp(7)<>"" Then
        Response.Write anyTmp(7)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- commented by nics 2009.02.09
                <td bgcolor="#FFCC33" nowrap>Voyage No.<font color="#FFFFFF"><b> 
                </b></font></td>
end of comment by nics 2009.02.09 -->
                <td bgcolor="#FFFFFF" nowrap>
<% ' 次航
    If anyTmp(8)<>"" Then
        Response.Write anyTmp(8)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- commented by nics 2009.02.09
                <td bgcolor="#FFCC33" nowrap>仕出港</td>
end of comment by nics 2009.02.09 -->
                <td bgcolor="#FFFFFF" nowrap>
<% ' 仕出港
    If anyTmp(9)<>"" Then
        Response.Write anyTmp(9)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- commented by nics 2009.02.09
                <td bgcolor="#FFCC33" nowrap>前港</td>
end of comment by nics 2009.02.09 -->
                <td bgcolor="#FFFFFF" nowrap>
<% ' 前港
    If anyTmp(38)<>"" Then
        Response.Write anyTmp(38)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
              </tr>
            </table>
<!-- mod by nics 2009.02.09 -->
<!--            <br>-->
            <font size="-1"><br></font>
<!-- end of mod by nics 2009.02.09 -->
<!---------------位置情報------------------------------------------- commented by nics 2009.02.09 -->
            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center"> 
                <td rowspan="3" nowrap bgcolor="#6495ED">&nbsp;位置情報&nbsp;</td>
                <td nowrap bgcolor="#FFCC33" colspan="2">仕出港<font size="-1"><sup>(※3)</sup></font></td>
                <td nowrap bgcolor="#FFCC33">前港<font size="-1"><sup>(※4)</sup></font></td>
              </tr>
              <tr align=middle>
                <td noWrap bgColor=#ffff99>中国</td>
                <td noWrap bgColor=#ffff99>東南アジア</td>
                <td noWrap bgColor=#ffffcc>離岸完了</td>
              </tr>
              <tr align="center"> 
                <td align="center"><table border="0" cellspacing="5">
                    <tr>
                      <td nowrap align="center"><a href="javascript:Submit('Form1')" class="splinkG" onClick="javascript:winOpen('win1','./cct/index.html',560,500)">&nbsp;赤湾&nbsp;</a></td>
                      <td nowrap align="center"><a href="gaibuif/expcntnr.asp?cntnrno=<%Response.Write anyTmp(1)%>&portcode=HUANG" class="splinkY" onClick="">&nbsp;黄埔&nbsp;</a></td>
                      <td nowrap align="center"><a href="gaibuif/expcntnr.asp?cntnrno=<%Response.Write anyTmp(1)%>&portcode=QINGD" class="splinkB" onClick="">&nbsp;青島&nbsp;</a></td>
                      <td nowrap align="center"><a href="#" class="dummylink" onClick="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</a></td>
                      </tr>
                    <tr>

<!-- 2015/11/30 Add-S MES Aoyagi -->
                      <td nowrap align="center"><a href="http://iport.sctcn.com/en-us/" target="_blank" class="splinkG" onClick="javascript:winOpen('win1','./sct/index.htm',560,500)">&nbsp;蛇口&nbsp;</a></td>
<!-- 2015/11/30 Add-E MES Aoyagi -->

<!-- 2015/11/30 Del-S MES Aoyagi
                      <td nowrap align="center"><a href="javascript:Submit('queryForm')" class="splinkG" onClick="javascript:winOpen('win1','./sct/index.html',560,500)">&nbsp;蛇口&nbsp;</a></td>
2015/11/30 Del-E MES Aoyagi -->
                      <td nowrap align="center"><a href="gaibuif/expcntnr.asp?cntnrno=<%Response.Write anyTmp(1)%>&portcode=NANSH" class="splinkY" onClick="">&nbsp;南沙&nbsp;</a></td>
                      <td nowrap align="center"><a href="#" class="dummylink" onClick="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</a></td>
                      <td nowrap align="center"><a href="#" class="dummylink" onClick="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</a></td>
                    </tr>
                </table></td>
                <td align="center"><table border="0" cellspacing="5">
<!--
                    <tr>
                      <td nowrap align="center"><a href="gaibuif/expcntnr.asp?cntnrno=<%Response.Write anyTmp(1)%>&portcode=TWTPE" class="splinkR" onClick="">&nbsp;台北&nbsp;</a></td>
                      <td nowrap align="center"><a href="#" class="dummylink" onClick="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</a></td>
                    </tr>
-->
                    <tr>
                      <td nowrap align="center"><a href="gaibuif/expcntnr.asp?cntnrno=<%Response.Write anyTmp(1)%>&portcode=THBKK" class="splinkLG" onClick="">&nbsp;バンコク&nbsp;</a></td>
                    </tr>
                </table></td>
                <td align="center" rowspan="2" nowrap>
<% ' 前港 - 離岸完了 $追加
    Response.Write DispDateTimeCell(anyTmp(37),11)
%>
                </td>
              </tr>
            </table>
<!-- mod by nics 2009.02.09 -->
<!--            <table border="0" cellspacing="2" cellpadding="1">-->
            <table border="0" cellspacing="0" cellpadding="0">
<!-- end of mod by nics 2009.02.09 -->
              <tr> 
                <td width="15">&nbsp;</td>
<!-- mod by nics 2009.02.09 -->
<!--                <td><font color="#000000" size="-1">（※1）仕出港・前港の時刻は、現地時間です。</font></td>-->
                <td><font color="#000000" size="-1">（※3）ボタンをクリックすると当該港での位置情報等が表示されます（現地時間表示）。&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;（※4）前港の時刻は、現地時間です。</font></td>
<!-- end of mod by nics 2009.02.09 -->
              </tr>
            </table>
<!---------------位置情報------------------------------------------- commented by nics 2009.02.09 -->
<!-- commented by nics 2009.02.09
            <table>
              <tr>
                <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>位置情報　</b></td>
                <td><img src="gif/hr.gif"></td>
              </tr>
            </table>
end of comment by nics 2009.02.09 -->
            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center"> 
<!-- add by nics 2009.02.09 -->
                <td rowspan="5" nowrap bgcolor="#6495ED">&nbsp;位置情報&nbsp;</td>
<!-- end of add by nics 2009.02.09 -->
<!-- commented by nics 2009.02.09
                <td nowrap align="center" bgcolor="#FFCC33">場所</td>
end of comment by nics 2009.02.09 -->
                <td colspan="4" nowrap bgcolor="#FFCC33">ターミナル</td>
                <td nowrap bgcolor="#FFCC33">ストックヤード</td>
                <td colspan="3" nowrap bgcolor="#FFCC33">陸上輸送</td>
              </tr>
              <tr align="center"> 
                <td nowrap colspan="2" bgcolor="#FFFFCC">着岸</td>
                <td nowrap colspan="2" bgcolor="#FFFFCC">ヤード</td>
                <td nowrap bgcolor="#FFFFCC">搬出完了</td>
                <td nowrap bgcolor="#FFFFCC">倉庫到着</td>
                <td nowrap rowspan="2" bgcolor="#FFFFCC">デバン<BR>完了</td>
                <td nowrap rowspan="2" bgcolor="#FFFFCC">空コン<BR>返却完了</td>
              </tr>
              <tr align="center" bgcolor="#FFFFCC">
                <td nowrap>計画</td>
                <td nowrap>予定／完了</td>
                <td nowrap>搬入完了</td>
                <td nowrap>搬出完了</td>
                <td nowrap>予約／完了</td>
                <td nowrap>指示／完了</td>
              </tr>
              <tr align="center"> 
                <td align="center" rowspan="2" nowrap>
<% ' ターミナル － 着岸スケジュール
    If anyTmp(31)<>"" Then
        Response.Write "<font color='#0000FF'>"
    End If
    Response.Write DispDateTimeCell(anyTmp(31),5)
    If anyTmp(31)<>"" Then
        Response.Write "</font>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' ターミナル - 着岸予定
    If anyTmp(2)<>"" Then
        bLate = false
        If anyTmp(3)<>"" Then
            If anyTmp(2)<anyTmp(3) Then
                bLate = true
            End If
        End If
        If anyTmp(31)<>"" Then
            If Left(anyTmp(31),10)<Left(anyTmp(2),10) Then
                bLate = true
            End If
        End If
        If bLate Then
            Response.Write "<font color='#FF0000'>"
        Else
            Response.Write "<font color='#0000FF'>"
        End If
        Response.Write DispDateTimeCell(anyTmp(2),11)
        Response.Write "</font>"
    Else
        Response.Write DispDateTimeCell(anyTmp(2),11)
    End If
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' ターミナル - ヤード搬入(確認)完了
    Response.Write DispDateTimeCell(anyTmp(12),11)
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' ターミナル - ヤード搬出完了
    Response.Write DispDateTimeCell(anyTmp(13),11)
%>
                </td>
                <td align="center" nowrap>
<% ' ストックヤード - 搬出予約 $追加
    sTemp=DispReserveCell(anyTmp(35),anyTmp(36),sColor)
    Response.Write sColor
    Response.Write sTemp
    If sColor<>"" Then
        Response.Write "</font>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' 陸上運送 - 倉庫到着スケジュール
    If anyTmp(34)<>"" Then
        If anyTmp(34)<anyTmp(14) Then
            Response.Write "<font color='#FF0000'>"
        Else
            Response.Write "<font color='#0000FF'>"
        End If
    End If
    Response.Write DispDateTimeCell(anyTmp(34),11)
    If anyTmp(34)<>"" Then
        Response.Write "</font>"
    End If
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' 陸上輸送 - デバン完了
    Response.Write DispDateTimeCell(anyTmp(15),11)
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' 陸上輸送 - 空コン返却完了
    Response.Write DispDateTimeCell(anyTmp(16),11)
%>
                </td>
              </tr>
              <tr>
                <td align="center" nowrap>
<% ' ターミナル - 着岸完了
    Response.Write DispDateTimeCell(anyTmp(3),11)
%>
                </td>
                <td align="center" nowrap>
<% ' ストックヤード - 搬出完了
    Response.Write DispDateTimeCell(anyTmp(30),11)
%>
                </td>
                <td align="center" nowrap>
<% ' 陸上輸送 - 倉庫到着完了
    Response.Write DispDateTimeCell(anyTmp(14),11)
%>
                </td>
              </tr>
            </table>
            <br>
<!---------------手続き及び搬入確認--------------------------------- commented by nics 2009.02.09 -->
<!-----手続情報---------------->
<!-- commented by nics 2009.02.09
            <table>
              <tr>
                <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>手続き及びターミナル搬出可否情報</b></td>
                <td><img src="gif/hr.gif"></td>
              </tr>
            </table>
            <br>
end of comment by nics 2009.02.09 -->
            <table border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td>
				  <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center"> 
<!-- add by nics 2009.02.09 -->
                <td rowspan="5" nowrap bgcolor="#6495ED">手続き及び<br>ターミナル<br>搬出可否<br>情報</td>
<!-- end of add by nics 2009.02.09 -->
<!-- commented by nics 2009.02.09
                <td rowspan="3" nowrap bgcolor="#FFCC33">項目</td>
end of comment by nics 2009.02.09 -->
<!-- mod by nics 2009.02.09 -->
<!--                <td colspan="4" nowrap bgcolor="#FFCC33">行政手続き</td>-->
                <td colspan="6" nowrap bgcolor="#FFCC33">行政手続き</td>
<!-- end of mod by nics 2009.02.09 -->
                <td rowspan="3" nowrap bgcolor="#FFCC33">商取引<br>
                  DO発行</td>
<!-- mod by nics 2010.02.02 
                <td rowspan="3" nowrap bgcolor="#FFCC33">フリー<br>
                  タイム</td>-->
<%'	ＢＬ番号指定あるいは事前情報入力のみ表示する
	    If Not bDispMode  or USER <> "" Then
			Response.Write "<td rowspan='3' nowrap bgcolor='#FFCC33'>"
			Response.Write "フリー"
			Response.Write "<br>"
			Response.Write "タイム"
			Response.Write "</td>"
		End If
%>
<!-- end of mod by nics 2010.02.02 -->
                <td rowspan="3" nowrap bgcolor="#FFCC33">ターミナル<br>
                  搬出可否</td>
<%'HiTS ver2 ADD by SEIKO n.Ooshige 2003/06/26%>
<!-- mod by nics 2010.02.02 
                <td rowspan="3" nowrap bgcolor="#FFCC33">ディテンション<br>フリータイム</td>-->
<%'	ＢＬ番号指定あるいは事前情報入力のみ表示する
	    If Not bDispMode  or USER <> "" Then
			Response.Write "<td rowspan='3' nowrap bgcolor='#FFCC33'>"
			Response.Write "ディテンション"
			Response.Write "<br>"
			Response.Write "フリータイム"
			Response.Write "</td>"
		End If
%>
<!-- end of mod by nics 2010.02.02 -->
              </tr>
              <tr> 
                <td align="center" nowrap bgcolor="#FFFFCC">搬入確認時刻</td>
                <td align="center" nowrap bgcolor="#FFFFCC" rowspan="2">動植物</td>
                <td align="center" nowrap bgcolor="#FFFFCC" rowspan="2">個別搬入</td>
<!-- mod by nics 2009.02.09 -->
<!--                <td align="left" nowrap bgcolor="#FFFFCC" rowspan="2">通関 /<BR>保税輸送<font size="-1"><sup>(※2)</sup></font></td>-->
                <td align="left" nowrap bgcolor="#FFFFCC" rowspan="2">通関 /<BR>保税輸送<font size="-1"><sup>(※5)</sup></font></td>
<!-- end of mod by nics 2009.02.09 -->
<!-- add by nics 2009.02.09 -->
                <td align="center" nowrap bgcolor="#FFFFCC" colspan="2">X線検査</td>
<!-- end of add by nics 2009.02.09 -->
              </tr>
              <tr> 
                <td align="center" nowrap bgcolor="#FFFFCC">予定／完了</td>
<!-- add by nics 2009.02.09 -->
                <td align="center" nowrap bgcolor="#FFFFCC">有無</td>
                <td align="center" nowrap bgcolor="#FFFFCC">CY返却</td>
<!-- end of add by nics 2009.02.09 -->
              </tr>
              <tr align="center"> 
<!-- commented by nics 2009.02.09
                <td bgcolor="#FFFFCC" rowspan="2" nowrap>情報</td>
end of comment by nics 2009.02.09 -->
                <td align="center" nowrap>
<% ' 搬入確認予定時刻
    If anyTmp(32)<>"" Then
        If anyTmp(18)<>"" Then
            If anyTmp(32)<anyTmp(18) Then
                Response.Write "<font color='#FF0000'>"
            Else
                Response.Write "<font color='#0000FF'>"
            End If
        Else
            Response.Write "<font color='#0000FF'>"
        End If
        Response.Write DispDateTimeCell(anyTmp(32),11)
        Response.Write "</font>"
    Else
        Response.Write DispDateTimeCell(anyTmp(32),11)
    End If
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' 動植物
    If anyTmp(17)="S" Then
        Response.Write "×"
    ElseIf anyTmp(17)="C" Then
        Response.Write "○"
    Else
        Response.Write "－"
    End If
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' 個別搬入
    If anyTmp(33)<>"" Then
        Response.Write "○"
    Else
        Response.Write "－"
    End If
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' 通関／保税輸送
    If anyTmp(19)<>"" Then
        If anyTmp(19)="O" Or anyTmp(19)="T" Then
            Response.Write "<a href='#"
            Response.Write iLineNo
            Response.Write "' onClick=""winOpen('win1','impdetail-h.asp?line="
            Response.Write iLineNo
            Response.Write "',150,150)"">○</a>"
        Else
            Response.Write "○"
        End If
    Else
        Response.Write "×"
    End If
%>
                </td>
<!-- add start by nics 2009.02.09  -->
                <td align="center" rowspan="2" nowrap>
<% ' X線有無
    If anyTmp(41)<>"" Then
        Response.Write anyTmp(41)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' X線CY返却
    If anyTmp(42)<>"" Then
        Response.Write anyTmp(42)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- add end by nics 2009.02.09  -->
                <td align="center" rowspan="2" nowrap>
<% ' 商取引ＤＯ発行
    If anyTmp(21)<>"Y" Then
        Response.Write "×"
    Else
        Response.Write "○"
    End If
%>
                </td>
<!-- mod by nics 2010.02.02 
                <td align="center" rowspan="2" nowrap>-->
<% ' フリータイム ＢＬ番号指定あるいは事前情報入力のみ表示する
    If Not bDispMode  or USER <> "" Then
	    Response.Write "<td align='center' rowspan='2' nowrap>"
'☆☆☆ Mod_S  by nics 2009.02.09
'    If anyTmp(22)<>"" Then
'        If anyTmp(22)<DispDateTime(Now,10) Then
'            Response.Write "<font color='#FF0000'>"
'        Else
'            Response.Write "<font color='#000000'>"
'        End If
'        Response.Write DispDateTimeCell(anyTmp(22),5)
'        Response.Write "</font>"
'    Else
'        Response.Write DispDateTimeCell(anyTmp(22),5)
'    End If
'☆☆☆
    ' anyTmp(13) ← CY搬出日時[yyyy/mm/dd hh:nn]
    ' anyTmp(22) ← フリータイム(フリータイム延長日付)[yyyy/mm/dd]
	    strDisp = DispDateTimeCell(anyTmp(22),5)
	    strColor = "#000000"    ' 黒
	    ' 搬出日時が設定されている場合
	    If anyTmp(13) <> "" Then
	        ' CY搬出日時＜システム日付の場合
	        If Left(anyTmp(13),10) < DispDateTime(Now,10) Then
	            strDisp = "－"
	        End If
	    ' 搬出日時が設定されていない場合
	    Else
	        ' フリータイムが設定されている場合
	        If IsDate(anyTmp(22)) Then
	            ' フリータイム≦システム日付の場合
	            If anyTmp(22) <= DispDateTime(Now,10) Then
	                strColor = "#FF0000"    ' 赤
	            ' (フリータイム－２日)≦システム日付の場合
	            ElseIf DispDateTime(DateAdd("d", -2, cDate(anyTmp(22))),10) <= DispDateTime(Now,10) Then
	                strColor = "#FFA500"    ' 黄
	            End If
	        End If
	    End If
	    Response.Write "<font color='" & strColor & "'>"
	    Response.Write strDisp
	    Response.Write "</font>"
'☆☆☆ Mod_E  by nics 2009.02.09
'Add by nics 2010.02.02 
	    Response.Write "</td>"
'end of Add by nics 2010.02.02 
	End If
'end of Mod by nics 2010.02.02 
%>
<!-- del by nics 2010.02.02 
                </td>
	 end of del by nics 2010.02.02 -->
                <td align="center" rowspan="2" nowrap>
<% ' ターミナル搬出可否
    If anyTmp(4)="Y" Then
        Response.Write "○"
    ElseIf anyTmp(4)="S" Then
        Response.Write "済"
    Else
        Response.Write "×"
    End If
%>
                </td>
<%'HiTS ver2 ADD by SEIKO n.Ooshige 2003/06/26%>
<!-- mod by nics 2010.02.02 
                <td align="center" rowspan="2" nowrap>-->
<% ' ディテンションフリータイム ＢＬ番号指定あるいは事前情報入力のみ表示する
'☆☆☆ Mod_S  by nics 2009.02.09
'    Response.Write anyTmp(39)
'☆☆☆
    ' anyTmp(39) ← ディテンションフリータイム
    ' anyTmp(16) ← 空バン返却日時[yyyy/mm/dd hh:nn]
    ' anyTmp(44) ← 空バン返却予定日[yyyy/mm/dd]
    If Not bDispMode  or USER <> "" Then
	    Response.Write "<td align='center' rowspan='2' nowrap>"
	    strDisp = anyTmp(39)
	    strColor = "#000000"    ' 黒
	    ' 空バン返却日時が設定されている場合
	    If anyTmp(16) <> "" Then
	        ' 空バン返却日時＜システム日付の場合
	        If Left(anyTmp(16),10) < DispDateTime(Now,10) Then
	            strDisp = "－"
	        End If
	    ' 空バン返却日時が設定されていない場合
	    Else
	        ' 空バン返却予定日時が設定されている場合
	        If IsDate(anyTmp(44)) Then
	            ' 空バン返却予定日≦システム日付の場合
	            If anyTmp(44) <= DispDateTime(Now,10) Then
	                strColor = "#FF0000"    ' 赤
	            ' (空バン返却予定日－2日)≦システム日付の場合
	            ElseIf DispDateTime(DateAdd("d", -2, cDate(anyTmp(44))),10) <= DispDateTime(Now,10) Then
	                strColor = "#FFA500"    ' 黄
	            End If
	        End If
	    End If
	    Response.Write "<font color='" & strColor & "'>"
	    Response.Write strDisp
	    Response.Write "</font>"
'☆☆☆ Mod_E  by nics 2009.02.09
'Add by nics 2010.02.02 
	    Response.Write "　</td>"
'end of Add by nics 2010.02.02 
	End If
'end of Mod by nics 2010.02.02 
%>
<!-- del by nics 2010.02.02 
              　</td>
	 end of del by nics 2010.02.02 -->
              </tr>
              <tr>
                <td align="center" nowrap>
<% ' 搬入確認完了時刻
    Response.Write DispDateTimeCell(anyTmp(18),5)
%>
                </td>
              </tr>
            </table>
			</td>
<!-- commented by nics 2009.02.09
                <td>&nbsp;</td>
                <td valign="top"><table border="1" cellpadding=" 3" cellspacing="1" bgcolor="#FFFFFF">
                  <tr>
                    <td align="center" nowrap bgcolor="#FFCC33">仕出港内位置情報<font size="-1"><sup>(※3)</sup></font></td>
                  </tr>
                  <tr>
                    <td align="center"><table border="0" cellspacing="5">
                        <tr>
                          <td nowrap><a href="javascript:Submit('Form1')" class="splink" onClick="javascript:winOpen('win1','./cct/index.html',560,500)">&nbsp;赤湾&nbsp;</a></td>
                          </tr>
                        <tr>
                          <td><a href="javascript:Submit('queryForm')" class="splink" onClick="javascript:winOpen('win1','./sct/index.html',560,500)">&nbsp;蛇口&nbsp;</a></td>
                          </tr>
                    </table></td>
                  </tr>
                </table></td>
end of comment by nics 2009.02.09 -->
              </tr>
            </table>
<!-- mod by nics 2009.02.09 -->
<!--            <table border="0" cellspacing="2" cellpadding="1">-->
            <table border="0" cellspacing="0" cellpadding="0">
<!-- end of mod by nics 2009.02.09 -->
              <tr> 
                <td width="15">&nbsp;</td>
<!-- mod by nics 2009.02.09 -->
<!--                <td nowrap>（※2）○をクリックすると保税輸送期間が表示されます。<br>
                  （※3）仕出港が枠内に表示されている場合、ボタンをクリックすると当該港での位置情報等が表示されます。
                </td>-->
                <td nowrap><font color="#000000" size="-1">（※5）保税輸送の場合、○をクリックすると保税輸送期間が表示されます。</td>
<!-- end of mod by nics 2009.02.09 -->
              </tr>
            </table>
<!-- commented by nics 2009.02.09
            <br>
end of comment by nics 2009.02.09 -->
<form>
      <input type=button value='表示データの更新' OnClick="JavaScript:window.location.href='impreload.asp?request=impdetail.asp'">
</form>
<!-- 2015/11/30 Del-S MES Aoyagi
<form name="queryForm" method="get" target="_blank" action="http://iport.sctcn.com/portal/page/portal/PG_IPort/Tab_OI/">
    <input type="hidden" name="p_parametertype" value="ContainerInfo">
    <input type="hidden" name="p_parametervalue" value="<%=anyTmp(1)%>">
2015/11/30 Del-E MES Aoyagi-->
<!-- 2010/09/13 URL修正
<form name="queryForm" method="post" target="_blank" action="http://oi.sctcn.com/Default.aspx?Action=Nav&Content=CONTAINER%20INFO.%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20&sm=CONTAINER%20INFO.">
    <input type="hidden" name="data" value="<%=anyTmp(1)%>">		
    <input type="hidden" name="OrgMenu" value="">
    <input type="hidden" name="targetPage" value="CONTAINER_INFO">
    <input type="hidden" name="nav" value="CONTAINER INFO.                         ">
-->
</form>

<!-- 2014/1/8 DEL-S MES aoyagi
<form name="Form1" method="post" action="http://www.cwcct.com/cct/conhis/con_his_infoE.aspx" id="Form1" target="_blank">
    <input type="hidden" name="Image1.x" value="0" />
    <input type="hidden" name="Image1.y" value="0" />
    <input type="hidden" name="__EVENTTARGET" value="" />
    <input type="hidden" name="__EVENTARGUMENT" value="" /> 
    <input type="hidden" name="__VIEWSTATE" value="dDwtMzMwNTk0MTUxOztsPEltYWdlMTs+Po9koK7lFKyndTfCh4n1g7KjLvsH" />
    <input type="hidden" name="cont_no" id="cont_no" value="<%=anyTmp(1)%>" />
    <input type="hidden" name="wyex" value="wyE" />
 2014/1/8 DEL-E MES aoyagi -->
<!-- 2014/1/8 ADD-S MES aoyagi -->
<form name="Form1" method="post" action="http://uport.cwcct.com/Portal/Ship/EN/Public/Pub_cntr_history_show.aspx" id="Form1" target="_blank">
    <input type="hidden" name="txtContainer_no" id="txtContainer_no" value="<%=anyTmp(1)%>" />
    <input type="hidden" name="rdoDisplay" id="rdoHTML" value="HTML" />
<!-- 2014/1/8 ADD-E MES aoyagi -->

</form>

<%
    ' 検索画面から直接飛んできたときは表示する
    If bSingle Then
        Response.Write "<form action='impcsvout.asp'>"
        Response.Write "<center>"
        Response.Write "<input type='submit' name='submit' value='CSVファイル出力'>　"
        Response.Write "<a href='help05.asp'>CSVファイル出力とは？</a>"
        Response.Write "</center>"
        Response.Write "</form>"
    End If
%>
          </td>
        </tr>
      </table>
      </center>
    </td>
  </tr>
  <tr>
    <td valign="bottom"> 
<%
    DispMenuBar
%>
    </td>
  </tr>
</table>
<!-------------詳細画面終わり--------------------------->
<%
    ' 検索画面から直接飛んできたとき
    If bSingle Then
        DispMenuBarBack "impentry.asp"
    Else
        If iReturn=1 Then
            DispMenuBarBack "implist1.asp"
        ElseIf iReturn=2 Then
            DispMenuBarBack "implist2.asp"
        ElseIf iReturn=3 Then
            DispMenuBarBack "implist3.asp"
        Else
            DispMenuBarBack "implist.asp"
        End If
    End If
%>
</body>
</html>
