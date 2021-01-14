<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' Tempファイル属性のチェック
    CheckTempFile "MSEXPORT", "expentry.asp"

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
        Response.Redirect "expentry.asp"             '輸出コンテナ照会トップ
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

    ' 輸出コンテナ照会詳細
    WriteLog fs, "1108","輸出コンテナ照会-コンテナ詳細", "00", anyTmp(1) & ","

    Session.Contents("dispexpctrl")=anyTmp(14)     ' 表示荷主管理番号を記憶
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
</head>
<body bgcolor="DEE1FF" text="#000000" link="#0000FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" vlink="#0000FF">
<!-------------ここから詳細画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/expdetailt.gif" width="506" height="73"></td>
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
<% ' Booking No
    Response.Write "<td bgcolor='#003399' background='gif/tableback.gif' nowrap><font color='#FFFFFF'><b>Booking No</b></font></td>"
    Response.Write "<td bgcolor='#FFFFFF' nowrap>" & anyTmp(0) & "</td>"
%>
                <td bgcolor="#003399" background="gif/tableback.gif" nowrap><font color="#FFFFFF"><b>コンテナNo.</b></font></td>
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
                <td nowrap><b>基本情報　</b></td>
                <td><img src="gif/hr.gif"></td>
              </tr>
            </table>
            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center" bgcolor="#FFCC33"> 
                <td colspan="4" nowrap>空コンテナ</td>
                <td colspan="5" nowrap bgcolor="#FFCC33">バンニング後コンテナ</td>
                <td bgcolor="#FFCC33" nowrap colspan="2">搬入受付期間</td>
              </tr>
              <tr align="center" bgcolor="#FFFF99"> 
                <td nowrap><font color="#000000">空コン受取場所</font></td>
                <td nowrap><font color="#000000">サイズ</font></td>
                <td nowrap>高さ<font size="-1"><sup>(※1)</sup></font></td>
                <td nowrap><font color="#000000">リーファー</font></td>
                <td nowrap><font color="#000000">シールNo.</font></td>
                <td nowrap><font color="#000000">貨物重量(t)</font></td>
                <td nowrap><font color="#000000">総重量(t)</font></td>
                <td nowrap><font color="#000000">危険品</font></td>
                <td nowrap><font color="#000000">搬入ターミナル名</font></td>
                <td nowrap><font color="#000000">オープン日</font></td>
                <td nowrap>クローズ日</td>
              </tr>
              <tr> 
                <td nowrap align="center">
<% ' 空コン受取場所
    If anyTmp(32)<>"" Then
        Response.Write anyTmp(32)
    ElseIf anyTmp(20)<>"" Then
        Response.Write "<font color='#0000FF'>" & anyTmp(20) & "</font>"
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' サイズ
    If anyTmp(33)<>"" Then
        Response.Write anyTmp(33)
    ElseIf anyTmp(10)<>"" Then
        Response.Write "<font color='#0000FF'>" & anyTmp(10) & "</font>"
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' 高さ
    If anyTmp(34)<>"" Then
        Response.Write anyTmp(34)
    ElseIf anyTmp(12)<>"" Then
        Response.Write "<font color='#0000FF'>" & anyTmp(12) & "</font>"
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' リーファー
    If anyTmp(35)<>"" Then
        If anyTmp(35)="R" Then
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
<% ' シールNo.
    If anyTmp(37)<>"" Then
        Response.Write anyTmp(37)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' 貨物重量 $追加
    If anyTmp(57)<>"" And anyTmp(57)<>"0" Then
        dWeight=anyTmp(57) / 10
        Response.Write dWeight
    Else
        Response.Write "－"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' 総重量
    If anyTmp(38)<>"" And anyTmp(38)<>"0" Then
        dWeight=anyTmp(38) / 10
        Response.Write dWeight
    Else
        Response.Write "－"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' 危険品
    If anyTmp(61)="H" Then
        Response.Write "○"
    ElseIf anyTmp(61)<>"" Then
        Response.Write "－"
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' 搬入ターミナル名
    If anyTmp(36)<>"" Then
        Response.Write anyTmp(36)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' CYオープン
    Response.Write DispDateTimeCell(anyTmp(39),5)
%>
                </td>
                <td align="center" nowrap>
<% ' CYクローズ
    Response.Write DispDateTimeCell(anyTmp(40),5)
%>
                </td>
              </tr>
            </table>
            <table border="0" cellspacing="2" cellpadding="1">
              <tr> 
                <td width="15">&nbsp;</td>
                <td><font color="#000000" size="-1">(※1)96=HC</font></td>
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
              <tr align="center" bgcolor="#FFCC33"> 
                <td nowrap>場所</td>
                <td colspan="3" nowrap>陸上輸送</td>
                <td nowrap bgcolor="#FFCC33">ストックヤード</td>
                <td colspan="4" nowrap bgcolor="#FFCC33">ターミナル</td>
                <td bgcolor="#FFCC33" nowrap>仕向港</font></td>
              </tr>
              <tr align="center" bgcolor="#FFFF99"> 
                <td nowrap rowspan="2"><font color="#000000">工程</font></td>
                <td nowrap rowspan="2"><font color="#000000">空コン受取</font></td>
                <td nowrap><font color="#000000">倉庫到着</font></td>
<% 
	Dim iSupNum
	If anyTmp(34)<>"" And strUserKind="陸運" Then
		iSupNum = 3
%>
                <td nowrap rowspan="2"><font color="#000000">バンニング</font><font size="-1"><sup>(※2)</sup></font></td>
<%
	Else
		iSupNum = 2
%>
                <td nowrap rowspan="2"><font color="#000000">バンニング</font></td>
<% End If %>
                <td nowrap><font color="#000000">搬入</font></td>
                <td nowrap><font color="#000000">CY搬入</font></td>
                <td nowrap rowspan="2"><font color="#000000">船積完了</font></td>
                <td nowrap colspan="2"><font color="#000000">離岸</font></td>
                <td nowrap><font color="#000000">着岸時刻</font><font size="-1"><sup>(※<%=iSupNum%>)</sup></font></td>
              </tr>
              <tr align="center" bgcolor="#FFFF99"> 
<% If anyTmp(34)<>"" And strUserKind="陸運" Then %>
                <td nowrap><font color="#000000">指示／完了</font><font size="-1"><sup>(※2)</sup></font></td>
<% Else %>
                <td nowrap><font color="#000000">指示／完了</font></td>
<% End If %>
                <td nowrap><font color="#000000">予約／完了</font></td>
                <td nowrap><font color="#000000">指示／完了</font></td>
                <td nowrap><font color="#000000">計画</font></td>
                <td nowrap><font color="#000000">予定／完了</font></td>
                <td nowrap><font color="#000000">予定／完了</font></td>
              </tr>
              <tr> 
                <td nowrap rowspan="2" bgcolor="#FFFFCC" align="center"><font color="#000000">時刻</font></td>
                <td rowspan="2" align="center" nowrap>
<% ' 陸上運送 - 空コン受取
    Response.Write DispDateTimeCell(anyTmp(46),11)
%>
                </td>
                <td align="center" nowrap>
<% ' 陸上運送 - 倉庫到着スケジュール
    If anyTmp(56)<>"" Then
        strTemp=anyTmp(56)
    Else
        strTemp=anyTmp(15)
    End If
    If strTemp<>"" Then
        If strTemp<anyTmp(47) Then
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
                <td rowspan="2" align="center" nowrap> 
<% ' 陸上運送 - バンニング
    If anyTmp(34)<>"" And strUserKind="陸運" Then
        Response.Write "<a href='ms-expinput.asp?kind=2&line=" & LineNo & "&request=ms-expdetail.asp'>"
    End If
    strTemp = DispDateTimeCell(anyTmp(48),11)
    If Left(strTemp,1)="<" And anyTmp(34)<>"" Then
        Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"
    Else
        Response.Write strTemp
    End If
    If anyTmp(34)<>"" And strUserKind="陸運" Then
        Response.Write "</a>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' ストックヤード - 搬入予約 $追加
    sTemp=DispReserveCell(anyTmp(58),anyTmp(59),sColor)
    Response.Write sColor
    Response.Write sTemp
    If sColor<>"" Then
        Response.Write "</font>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' ターミナル - CY搬入指示 $追加
    If anyTmp(60)<>"" Then
        strTemp=anyTmp(60)
    Else
        strTemp=anyTmp(16)
    End If
    If strTemp<>"" Then
        If Left(strTemp,10)<Left(anyTmp(49),10) Then
            Response.Write "<font color='#FF0000'>"
        Else
            Response.Write "<font color='#0000FF'>"
        End If
    End If
    Response.Write DispDateTimeCell(strTemp,5)
    If strTemp<>"" Then
        Response.Write "</font>"
    End If
%>
                </td>
                <td rowspan="2" align="center" nowrap> 
<% ' ターミナル - 船積完了
    Response.Write DispDateTimeCell(anyTmp(50),11)
%>
                </td>
                <td rowspan="2" align="center" nowrap>
<% ' ターミナル - 離岸スケジュール
    If anyTmp(55)<>"" Then
        Response.Write "<font color='#0000FF'>"
    End If
    Response.Write DispDateTimeCell(anyTmp(55),5)
    If anyTmp(55)<>"" Then
        Response.Write "</font>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' ターミナル - 離岸予定
    If anyTmp(45)<>"" Then
        bLate = false
        If anyTmp(51)<>"" Then
            If anyTmp(45)<anyTmp(51) Then
                bLate = true
            End If
        End If
        If anyTmp(55)<>"" Then
            If Left(anyTmp(55),10)<Left(anyTmp(45),10) Then
                bLate = true
            End If
        End If
        If bLate Then
            Response.Write "<font color='#FF0000'>"
        Else
            Response.Write "<font color='#0000FF'>"
        End If
        Response.Write DispDateTimeCell(anyTmp(45),11)
        Response.Write "</font>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' 仕向港 - 着岸予定
    If anyTmp(53)<>"" Then
        If anyTmp(52)<>"" Then
            If anyTmp(53)<anyTmp(52) Then
                Response.Write "<font color='#FF0000'>"
            Else
                Response.Write "<font color='#0000FF'>"
            End If
        Else
            Response.Write "<font color='#0000FF'>"
        End If
        Response.Write DispDateTimeCell(anyTmp(53),11)
        Response.Write "</font>"
    Else
        Response.Write DispDateTimeCell(anyTmp(53),11)
    End If
%>
                </td>
              </tr>
              <tr>
                <td align="center" nowrap>
<% ' 陸上運送 - 倉庫到着
    If anyTmp(34)<>"" And strUserKind="陸運" Then
        Response.Write "<a href='ms-expinput.asp?kind=1&line=" & LineNo & "&request=ms-expdetail.asp'>"
    End If
    strTemp = DispDateTimeCell(anyTmp(47),11)
    If Left(strTemp,1)="<" And anyTmp(34)<>"" Then
        Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"
    Else
        Response.Write strTemp
    End If
    If anyTmp(34)<>"" And strUserKind="陸運" Then
        Response.Write "</a>"
    End If
%>
                </td>
                <td align="center" nowrap> 
<% ' ストックヤード - 搬入完了
    Response.Write DispDateTimeCell(anyTmp(54),11)
%>
                </td>
                <td align="center" nowrap> 
<% ' ターミナル - CY搬入完了
    Response.Write DispDateTimeCell(anyTmp(49),11)
%>
                </td>
                <td align="center" nowrap>
<% ' ターミナル - 離岸完了
    Response.Write DispDateTimeCell(anyTmp(51),11)
%>
                </td>
                <td align="center" nowrap>
<% ' 仕向港 - 着岸完了
    Response.Write DispDateTimeCell(anyTmp(52),11)
%>
                </td>
              </tr>
            </table>
            <table border="0" cellspacing="2" cellpadding="1">
              <tr> 
<% If anyTmp(34)<>"" And strUserKind="陸運" Then %>
                <td width="15">&nbsp;</td>
                <td><font color="#000000" size="-1">（※2）クリックで完了時刻入力画面へ</font></td>
<% End If %>
                <td width="15">&nbsp;</td>
                <td><font color="#000000" size="-1">（※<%=iSupNum%>）仕向港の時刻は、現地時間です。</font></td>
              </tr>
            </table>
<BR>
            <table>
              <tr>
                <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>本船情報　</b></td>
                <td><img src="gif/hr.gif"></td>
              </tr>
            </table>
            <table border=1 cellpadding="3" cellspacing="1">
              <tr> 
                <td bgcolor="#FFCC33" nowrap>船社</td>
                <td bgcolor="#FFFFFF">
<% ' 船社
    If anyTmp(41)<>"" Then
        Response.Write anyTmp(41)
    ElseIf anyTmp(24)<>"" Then
        Response.Write anyTmp(24)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td bgcolor="#FFCC33" nowrap><font color="#000000">船名</font></td>
                <td bgcolor="#FFFFFF" nowrap>
<% ' 船名
    If anyTmp(42)<>"" Then
        Response.Write anyTmp(42)
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
    If anyTmp(43)<>"" Then
        Response.Write anyTmp(43)
    ElseIf anyTmp(3)<>"" Then
        Response.Write anyTmp(3)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td bgcolor="#FFCC33" nowrap>仕向港</td>
                <td bgcolor="#FFFFFF" nowrap>
<% ' 仕向港
    If anyTmp(44)<>"" Then
        Response.Write anyTmp(44)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
              </tr>
            </table>
            <br>
<form>
      <input type=button value='表示データの更新' OnClick="JavaScript:window.location.href='ms-expreload.asp?request=ms-expdetail.asp'">
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
        DispMenuBarBack "ms-explist1.asp"
    ElseIf iReturn=2 Then
        DispMenuBarBack "ms-explist2.asp"
    ElseIf iReturn=3 Then
        DispMenuBarBack "ms-explist3.asp"
    End If
%>
</body>
</html>
