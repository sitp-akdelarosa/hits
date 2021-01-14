<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' Tempファイル属性のチェック
    CheckTempFile "MSIMPORT", "impentry.asp"

    ' 指定引数の取得
    Dim iLineNo
    iLineNo = CInt(Request.QueryString("line"))
    Dim iReturn
    iReturn = Session.Contents("dispreturn")

    ' ユーザ種類をチェックする
    strUserKind=Session.Contents("userkind")

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
    ti.Close

    ' 輸入コンテナ照会詳細
    WriteLog fs, "2108","輸入コンテナ照会-コンテナ詳細", "00", anyTmp(1) & ","

    Session.Contents("dispcntnr")=anyTmp(1)     ' 表示コンテナNo.を記憶
%>

<html>
<head>
<title></title>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
<%
    DispMenuJava
%>
</SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
function winOpen(winName,url,W,H){
  var WinD11=window.open(url,winName,'scrollbars=auto,resizable=yes,width='+W+',height='+H+'');
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
' Added and Commented by seiko-denki 2003.07.18
	DisplayCodeListButton
'    DispMenu
'	Dim strScriptName,strRoute
'	strScriptName = Request.ServerVariables("SCRIPT_NAME")
'	strRoute = SetRoute(strScriptName)
'	Session.Contents("route") = strRoute
' End of Addition by seiko-denki 2003.07.18
%>
          </td>
        </tr>
      </table>
      <center>
<!-- commented by seiko-denki 2003.07.18
		<table width=95% cellpadding="0" cellspacing="0" border="0">
		  <tr>
			<td align="right">
			  <font color="#333333" size="-1">
				<%=strRoute%>
			  </font>
			</td>
		  </tr>
		</table>
End of comment by seiko-denki 2003.07.18 -->

		<table width=95% cellpadding=3>
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
          <td>　<br>
            <table border=1 cellpadding="3" cellspacing="1">
              <tr> 
<% ' BL No
    Response.Write "<td bgcolor='#003399' background='gif/tableback.gif' nowrap><font color='#FFFFFF'><b>BL No</b></font></td>"
    Response.Write "<td bgcolor='#FFFFFF' nowrap>" & anyTmp(0) & "</td>"
%>
                <td bgcolor="#003399" background="gif/tableback.gif" nowrap><font color="#FFFFFF"><b>コンテナNo</b></font></td>
                <td bgcolor="#FFFFFF" nowrap>
<% ' コンテナNo.
    Response.Write anyTmp(1)
%>
                </td>
              </tr>
            </table>
<BR>
            <table>
              <tr>
                <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>位置情報　</b></td>
                <td><img src="gif/hr.gif"></td>
              </tr>
            </table>

            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center"> 
                <td nowrap align="center" bgcolor="#FFCC33">場所</td>
                <td nowrap bgcolor="#FFCC33">仕出港<font size="-1"><sup>(※1)</sup></font></td>
                <td nowrap bgcolor="#FFCC33">前港<font size="-1"><sup>(※1)</sup></font></td>
                <td colspan="4" nowrap bgcolor="#FFCC33">ターミナル</td>
                <td nowrap bgcolor="#FFCC33">ストックヤード</td>
                <td colspan="3" nowrap bgcolor="#FFCC33">陸上輸送</td>
              </tr>
              <tr align="center"> 
                <td nowrap rowspan="2" bgcolor="#FFFFCC">工程</td>
                <td nowrap rowspan="2" bgcolor="#FFFFCC">離岸完了</td>
                <td nowrap rowspan="2" bgcolor="#FFFFCC">離岸完了</td>
                <td nowrap colspan="2" bgcolor="#FFFFCC">着岸</td>
                <td nowrap colspan="2" bgcolor="#FFFFCC">ヤード</td>
                <td nowrap bgcolor="#FFFFCC">搬出完了</td>
                <td nowrap bgcolor="#FFFFCC">倉庫到着</td>
<% If anyTmp(54)<>"" And strUserKind="陸運" Then %>
                <td nowrap rowspan="2" bgcolor="#FFFFCC">デバン完了<font size="-1"><sup>(※2)</sup></font></td>
<% Else %>
                <td nowrap rowspan="2" bgcolor="#FFFFCC">デバン完了</td>
<% End If %>
                <td nowrap rowspan="2" bgcolor="#FFFFCC">空コン<BR>返却完了</td>
              </tr>
              <tr align="center" bgcolor="#FFFFCC">
                <td nowrap>計画</td>
                <td nowrap>予定／完了</td>
                <td nowrap>搬入完了</td>
                <td nowrap>搬出完了</td>
                <td nowrap>予約／完了</td>
<% If anyTmp(54)<>"" And strUserKind="陸運" Then %>
                <td nowrap>指示／完了<font size="-1"><sup>(※2)</sup></font></td>
<% Else %>
                <td nowrap>指示／完了</td>
<% End If %>
              </tr>
              <tr align="center"> 
                <td bgcolor="#FFFFCC" rowspan="2" nowrap>時刻</td>
                <td align="center" rowspan="2" nowrap>
<% ' 仕出港 - 離岸完了
    Response.Write DispDateTimeCell(anyTmp(41),11)
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' 前港 - 離岸完了 $追加
    Response.Write DispDateTimeCell(anyTmp(67),11)
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' ターミナル － 着岸スケジュール
    If anyTmp(61)<>"" Then
        Response.Write "<font color='#0000FF'>"
    End If
    Response.Write DispDateTimeCell(anyTmp(61),5)
    If anyTmp(61)<>"" Then
        Response.Write "</font>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' ターミナル - 着岸予定
    If anyTmp(32)<>"" Then
        bLate = false
        If anyTmp(33)<>"" Then
            If anyTmp(32)<anyTmp(33) Then
                bLate = true
            End If
        End If
        If anyTmp(61)<>"" Then
            If Left(anyTmp(61),10)<Left(anyTmp(32),10) Then
                bLate = true
            End If
        End If
        If bLate Then
            Response.Write "<font color='#FF0000'>"
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
<% ' ターミナル - ヤード搬入(確認)完了
    Response.Write DispDateTimeCell(anyTmp(42),11)
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' ターミナル - ヤード搬出完了
    Response.Write DispDateTimeCell(anyTmp(43),11)
%>
                </td>
                <td align="center" nowrap>
<% ' ストックヤード - 搬出予約 $追加
    sTemp=DispReserveCell(anyTmp(65),anyTmp(66),sColor)
    Response.Write sColor
    Response.Write sTemp
    If sColor<>"" Then
        Response.Write "</font>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' 陸上運送 - 倉庫到着スケジュール
    If anyTmp(64)<>"" Then
        strTemp=anyTmp(64)
    Else
        strTemp=anyTmp(13)
    End If
    If strTemp<>"" Then
        If strTemp<anyTmp(44) Then
            Response.Write "<font color='#FF0000'>"
        Else
            Response.Write "<font color='#0000FF'>"
        End If
    End If
    Response.Write DispDateTimeCell(strTemp,11)
    If strTemp<>"" Then
        Response.Write "</font>"
    End If
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' 陸上輸送 - デバン完了
    If anyTmp(54)<>"" And strUserKind="陸運" Then
        Response.Write "<a href='ms-impinput.asp?kind=2&line=" & LineNo & "&request=ms-impdetail.asp'>"
    End If
    strTemp = DispDateTimeCell(anyTmp(45),11)
    If Left(strTemp,1)="<" And anyTmp(54)<>"" Then
        Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"
    Else
        Response.Write strTemp
    End If
    If anyTmp(54)<>"" And strUserKind="陸運" Then
        Response.Write "</a>"
    End If
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' 陸上輸送 - 空コン返却完了
    Response.Write DispDateTimeCell(anyTmp(46),11)
%>
                </td>
              </tr>
              <tr>
                <td align="center" nowrap>
<% ' ターミナル - 着岸完了
    Response.Write DispDateTimeCell(anyTmp(33),11)
%>
                </td>
                <td align="center" nowrap>
<% ' ストックヤード - 搬出完了
    Response.Write DispDateTimeCell(anyTmp(60),11)
%>
                </td>
                <td align="center" nowrap>
<% ' 陸上輸送 - 倉庫到着完了
    If anyTmp(54)<>"" And strUserKind="陸運" Then
        Response.Write "<a href='ms-impinput.asp?kind=1&line=" & LineNo & "&request=ms-impdetail.asp'>"
    End If
    strTemp = DispDateTimeCell(anyTmp(44),11)
    If Left(strTemp,1)="<" And anyTmp(54)<>"" Then
        Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"
    Else
        Response.Write strTemp
    End If
    If anyTmp(54)<>"" And strUserKind="陸運" Then
        Response.Write "</a>"
    End If
%>
                </td>
              </tr>
            </table>
        <table border="0" cellspacing="2" cellpadding="1">
          <tr> 
            <td width="15"><BR></td>
            <td><font color="#000000" size="-1">（※1）仕向港、前港の時刻は、現地時間です。</font></td>
<% If anyTmp(54)<>"" And strUserKind="陸運" Then %>
            <td width="15"><BR></td>
            <td><font color="#000000" size="-1">（※2）クリックで完了時刻入力画面へ</font></td>
<% End If %>
          </tr>
        </table>
            <br>
<!-----手続情報---------------->
            <table>
              <tr>
                <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>手続き及びターミナル搬出可否情報</b></td>
                <td><img src="gif/hr.gif"></td>
              </tr>
            </table>
            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center"> 
                <td rowspan="3" nowrap bgcolor="#FFCC33">項目</td>
                <td colspan="4" nowrap bgcolor="#FFCC33">行政手続き</td>
                <td rowspan="3" nowrap bgcolor="#FFCC33">商取引<br>
                  DO発行</td>
                <td rowspan="3" nowrap bgcolor="#FFCC33">フリー<br>
                  タイム</td>
                <td rowspan="3" nowrap bgcolor="#FFCC33">ターミナル<br>
                  搬出可否</td>
              </tr>
              <tr> 
                <td align="center" nowrap bgcolor="#FFFFCC">搬入確認時刻</td>
                <td align="center" nowrap bgcolor="#FFFFCC" rowspan="2">動植物</td>
                <td align="center" nowrap bgcolor="#FFFFCC" rowspan="2">個別搬入</td>
                <td align="center" nowrap bgcolor="#FFFFCC" rowspan="2">通関/<br>保税輸送</td>
              </tr>
              <tr> 
                <td align="center" nowrap bgcolor="#FFFFCC">予定／完了</td>
              </tr>
              <tr align="center"> 
                <td bgcolor="#FFFFCC" rowspan="2" nowrap>情報</td>
                <td align="center" nowrap>
<% ' 搬入確認予定時刻
    If anyTmp(62)<>"" Then
        If anyTmp(48)<>"" Then
            If Left(anyTmp(62),10)<Left(anyTmp(48),10) Then
                Response.Write "<font color='#FF0000'>"
            Else
                Response.Write "<font color='#0000FF'>"
            End If
        Else
            Response.Write "<font color='#0000FF'>"
        End If
        Response.Write DispDateTimeCell(anyTmp(62),11)
        Response.Write "</font>"
    Else
        Response.Write DispDateTimeCell(anyTmp(62),11)
    End If
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' 動植物
    If anyTmp(47)="S" Then
        Response.Write "×"
    ElseIf anyTmp(47)="C" Then
        Response.Write "○"
    Else
        Response.Write "－"
    End If
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' 個別搬入
    If anyTmp(63)<>"" Then
        Response.Write "○"
    Else
        Response.Write "－"
    End If
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' 通関／保税輸送
    If anyTmp(49)<>"" Then
        If anyTmp(49)="O" Or anyTmp(49)="T" Then
            Response.Write "<a href='#"
            Response.Write iLineNo
            Response.Write "' onClick=""winOpen('win1','ms-impdetail-h.asp?line="
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
                <td align="center" rowspan="2" nowrap>
<% ' 商取引ＤＯ発行
    If anyTmp(51)<>"Y" Then
        Response.Write "×"
    Else
        Response.Write "○"
    End If
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' フリータイム
    If anyTmp(52)<>"" Then
        If anyTmp(52)<DispDateTime(Now,10) Then
            Response.Write "<font color='#FF0000'>"
        Else
            Response.Write "<font color='#000000'>"
        End If
        Response.Write DispDateTimeCell(anyTmp(52),5)
        Response.Write "</font>"
    Else
        Response.Write DispDateTimeCell(anyTmp(52),5)
    End If
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' ターミナル搬出可否
    If anyTmp(34)="Y" Then
        Response.Write "○"
    ElseIf anyTmp(34)="S" Then
        Response.Write "済"
    Else
        Response.Write "×"
    End If
%>
                </td>
              </tr>
              <tr>
                <td align="center" nowrap>
<% ' 搬入確認完了時刻
    Response.Write DispDateTimeCell(anyTmp(48),5)
%>
                </td>
              </tr>
            </table>
            <br>
<!---------------基本情報--------------------------------------------->
            <table>
              <tr>
                <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>基本情報</b></td>
                <td><img src="gif/hr.gif"></td>
              </tr>
            </table>
            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center" bgcolor="#FFCC33"> 
                <td valign="top" nowrap>項目</td>
                <td nowrap bgcolor="#FFCC33">サイズ</td>
<%
	Dim iSupNum
	If anyTmp(54)<>"" And strUserKind="陸運" Then
		iSupNum = 3
	Else
		iSupNum = 2
	End If
%>
                <td nowrap bgcolor="#FFCC33">高さ<font size="-1"><sup>(※<%=iSupNum%>)</sup></font></td>
                <td nowrap bgcolor="#FFCC33">リーファー</td>
                <td nowrap bgcolor="#FFCC33">総重量(t)</td>
                <td valign="top" nowrap>危険物<font size="-1"><sup>(※<%=iSupNum+1%>)</sup></font></td>
                <td nowrap bgcolor="#FFCC33">搬出ターミナル</td>
                <td nowrap bgcolor="#FFCC33">ストックヤード利用</td>
                <td nowrap bgcolor="#FFCC33">返却場所</td>
              </tr>
              <tr align="center"> 
                <td bgcolor="#FFFFCC" nowrap>情報</td>
                <td align="center" nowrap>
<% ' サイズ
    If anyTmp(53)<>"" Then
        Response.Write anyTmp(53)
    Else
        If anyTmp(10)<>"" Then
            Response.Write "<font color='#0000FF'>" & anyTmp(10) & "</font>"
        Else
            Response.Write "<br>"
        End If
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' 高さ
    If anyTmp(54)<>"" Then
        Response.Write anyTmp(54)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' リーファー
    If anyTmp(55)<>"" Then
        If anyTmp(55)="R" Then
            Response.Write "○"
        Else
            Response.Write "－"
        End If
    Else
        Response.Write "<font color='#0000FF'>"
        If anyTmp(11)<>"" Then
            If anyTmp(11)<>"RF" Then
                Response.Write "－"
            Else
                Response.Write "○"
            End If
        Else
            Response.Write "<br>"
        End If
        Response.Write "</font>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' 総重量
    If anyTmp(56)<>"" And anyTmp(56)<>"0" Then
        dWeight=anyTmp(56) / 10
        Response.Write dWeight
    Else
        Response.Write "－"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' 危険物
    If anyTmp(57)<>"" Then
        If anyTmp(57)<>"H" Then
            Response.Write "－"
        Else
            Response.Write "○"
        End If
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' 搬出ターミナル
    If anyTmp(35)<>"" Then
        Response.Write anyTmp(35)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' ストックヤード利用 $追加
    If anyTmp(65)>="1" And anyTmp(65)<="4" Then
        Response.Write "○"
    Else
        Response.Write "×"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' 返却場所
    If anyTmp(40)<>"" Then
        Response.Write anyTmp(40)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
              </tr>
            </table>
        <table border="0" cellspacing="2" cellpadding="1">
          <tr> 
            <td width="15"><BR></td>
            <td><font color="#000000" size="-1">（※<%=iSupNum%>) 96=HC</font></td>
            <td width="15"><BR></td>
            <td><font color="#000000" size="-1">（※<%=iSupNum+1%>）消防法に関わる危険物の有無</font></td>
          </tr>
        </table>
            <br>
<!---------------本船情報--------------------------------------------->
            <table>
              <tr> 
                <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>本船情報&nbsp;&nbsp;</b></td>
                <td><img src="gif/hr.gif"></td>
              </tr>
            </table>
            <table border=1 cellpadding="3" cellspacing="1">
              <tr> 
                <td bgcolor="#FFCC33" nowrap><font color="#000000">船社</font></td>
                <td bgcolor="#FFFFFF" nowrap>
<% ' 船社
    If anyTmp(36)<>"" Then
        Response.Write anyTmp(36)
    ElseIf anyTmp(21)<>"" Then
        Response.Write anyTmp(21)
    ElseIf anyTmp(15)<>"" Then
        Response.Write anyTmp(15)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td bgcolor="#FFCC33" nowrap><font color="#000000">船名</font></td>
                <td bgcolor="#FFFFFF" nowrap>
<% ' 船名
    If anyTmp(37)<>"" Then
        Response.Write anyTmp(37)
    ElseIf anyTmp(2)<>"" Then
        Response.Write anyTmp(2)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td bgcolor="#FFCC33" nowrap>Voyage No.<font color="#FFFFFF"><b> 
                </b></font></td>
                <td bgcolor="#FFFFFF" nowrap>
<% ' 次航
    If anyTmp(38)<>"" Then
        Response.Write anyTmp(38)
    ElseIf anyTmp(3)<>"" Then
        Response.Write anyTmp(3)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td bgcolor="#FFCC33" nowrap>仕出港</td>
                <td bgcolor="#FFFFFF" nowrap>
<% ' 仕出港
    If anyTmp(39)<>"" Then
        Response.Write anyTmp(39)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td bgcolor="#FFCC33" nowrap>前港</td>
                <td bgcolor="#FFFFFF" nowrap>
<% ' 前港
    If anyTmp(68)<>"" Then
        Response.Write anyTmp(68)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
              </tr>
            </table>
<form>
      <input type=button value='表示データの更新' OnClick="JavaScript:window.location.href='ms-impreload.asp?request=ms-impdetail.asp'">
</form>
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
    If iReturn=1 Then
        DispMenuBarBack "ms-implist1.asp"
    ElseIf iReturn=2 Then
        DispMenuBarBack "ms-implist2.asp"
    End If
%>
</body>
</html>
