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

'☆☆☆ Add by nics 2010.02.18
    Dim USER
	USER    = Session.Contents("userid")
'☆☆☆ end of Add by nics 2010.02.18

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
	WriteLog fs, "2306","輸入コンテナ照会-単独コンテナ", "00", anyTmp(1) & ","

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
<body bgcolor="DEE1FF" text="#000000" link="#0000ff" background="../gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" vlink="#0000ff">
<!-------------ここから詳細画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="../gif/impdetailt.gif" width="506" height="73"></td>
          <td height="25" bgcolor="000099" align="right"><img src="../gif/logo_hits_ver2.gif" width="300" height="25"></td>
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

<!-- mod by nics 2009.03.03 -->
<!--		<table width=95% cellpadding=3>-->
		<table width=95% cellpadding=0>
<!-- end of mod by nics 2009.03.03 -->
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
<!-- mod by nics 2009.03.03 -->
<!--          <td>　<br>-->
          <td>
<!-- end of mod by nics 2009.03.03 -->
            <table border=1 cellpadding="3" cellspacing="1">
              <tr> 
<% ' BL No
    If Not bDispMode Then
        Response.Write "<td bgcolor='#003399' background='../gif/tableback.gif' nowrap><font color='#FFFFFF'><b>BL No</b></font></td>"
        Response.Write "<td bgcolor='#FFFFFF' nowrap>" & anyTmp(0) & "</td>"
    End If
%>
                <td bgcolor="#003399" background="../gif/tableback.gif" nowrap><font color="#FFFFFF"><b>コンテナNo</b></font></td>
                <td bgcolor="#FFFFFF" nowrap>
<% ' コンテナNo.
    Response.Write anyTmp(1)
%>
                </td>
              </tr>
            </table>
<!-- mod by nics 2009.03.03 -->
<!--			<BR>-->
			<font size="-2"><BR></font>
<!-- end of mod by nics 2009.03.03 -->
<!---------------基本情報------------------------------------------- commented by nics 2009.03.03 -->
<!---------------基本情報--------------------------------------------->
<!-- commented by nics 2009.03.03
            <table>
              <tr>
                <td><img src="../gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>基本情報</b></td>
                <td><img src="../gif/hr.gif"></td>
              </tr>
            </table>
end of comment by nics 2009.03.03 -->
            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center" bgcolor="#FFCC33"> 
<!-- add by nics 2009.03.03 -->
                <td rowspan="5" nowrap bgcolor="#6495ED">&nbsp;基本情報&nbsp;</td>
<!-- end of add by nics 2009.03.03 -->
<!-- commented by nics 2009.03.03
                <td valign="top" nowrap>項目</td>
end of comment by nics 2009.03.03 -->
                <td nowrap bgcolor="#FFCC33">サイズ</td>
<!-- mod by nics 2009.03.03 -->
<!-- Add-S MES Aoyagi 2010.11.23 -->
		<td nowrap bgcolor="#FFCC33">タイプ</td>
<!-- Add-E MES Aoyagi 2010.11.23 -->
<!--                <td nowrap bgcolor="#FFCC33">高さ<font size="-1"><sup>(※3)</sup></font></td>-->
                <td nowrap bgcolor="#FFCC33">高さ<font size="-1"><sup>(※1)</sup></font></td>
<!-- end of mod by nics 2009.03.03 -->
                <td nowrap bgcolor="#FFCC33">リーファ</td>
                <td nowrap bgcolor="#FFCC33">総重量(t)</td>
<!-- mod by nics 2009.03.03 -->
<!--                <td valign="top" nowrap>危険品<font size="-1"><sup>(※4)</sup></font></td>-->
                <td valign="top" nowrap>危険品<font size="-1"><sup>(※2)</sup></font></td>
<!-- end of mod by nics 2009.03.03 -->
<!-- mod by nics 2009.03.03 -->
<!--                <td nowrap bgcolor="#FFCC33">搬出ターミナル</td>-->
                <td nowrap bgcolor="#FFCC33">搬出ターミナル(蔵置場所コード)</td>
<!-- end of mod by nics 2009.03.03 -->
<!-- add by nics 2009.03.03 -->
                <td nowrap bgcolor="#FFCC33">本船担当<br>オペレータ</td>
<!-- end of add by nics 2009.03.03 -->
                <td nowrap bgcolor="#FFCC33">ストックヤード利用</td>
                <td nowrap bgcolor="#FFCC33">返却場所</td>
              </tr>
              <tr align="center"> 
<!-- commented by nics 2009.03.03
                <td bgcolor="#FFFFCC" nowrap>情報</td>
end of comment by nics 2009.03.03 -->
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
    If anyTmp(44)<>"" Then
        Response.Write anyTmp(44)
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
'        dWeight=anyTmp(26) / 10
		dWeight=anyTmp(26)
		if anyTmp(26)<1000 then
            dWeight=anyTmp(26) / 10
		else
            dWeight=anyTmp(26) / 1000
		end if
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
                <td align="center" nowrap>
<% ' 搬出ターミナル
'    If anyTmp(5)<>"" Then
'        Response.Write anyTmp(5)
'    Else
'        Response.Write "<br>"
'    End If
%>
<!-- mod by nics 2009.03.03 -->
<!--				中央ふ頭-->
				中央ふ頭ＣＹ(6TK43)
<!-- end of mod by nics 2009.03.03 -->
                </td>
<!-- add by nics 2009.03.03 -->
                <td align="center" nowrap>
<% ' 担当オペレータ
        Response.Write "ジェネック"
%>
                </td>
<!-- end of add by nics 2009.03.03 -->
                <td align="center" nowrap>
<% ' ストックヤード利用 $追加
'2006/03/23
'    If anyTmp(35)>="1" And anyTmp(35)<="4" Then
'        Response.Write "○"
'    Else
        Response.Write "×"
'    End If
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
<!-- mod by nics 2009.03.03 -->
<!--            <table border="0" cellspacing="1" cellpadding="3">-->
            <table border="0" cellspacing="0" cellpadding="0">
<!-- end of mod by nics 2009.03.03 -->
              <tr> 
                <td width="15">&nbsp;</td>
<!-- mod by nics 2009.03.03 -->
<!--                <td nowrap><font color="#000000" size="-1">（※3）96=HC</td>-->
                <td nowrap><font color="#000000" size="-1">（※1）96=HC</td>
<!-- end of mod by nics 2009.03.03 -->
                <td width="15">&nbsp;</td>
<!-- mod by nics 2009.03.03 -->
<!--                <td nowrap><font color="#000000" size="-1">（※4）消防法に関わる危険品の有無</td>-->
                <td nowrap><font color="#000000" size="-1">（※2）消防法に関わる危険品の有無</td>
<!-- end of mod by nics 2009.03.03 -->
              </tr>
            </table>
<!-- commented by nics 2009.03.03
            <br>
end of comment by nics 2009.03.03 -->
<!---------------本船情報------------------------------------------- commented by nics 2009.03.03 -->
<!---------------本船情報--------------------------------------------->
<!-- commented by nics 2009.03.03
            <table>
              <tr> 
                <td><img src="../gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>本船情報&nbsp;&nbsp;</b></td>
                <td><img src="../gif/hr.gif"></td>
              </tr>
            </table>
end of comment by nics 2009.03.03 -->
            <table border=1 cellpadding="3" cellspacing="1">
<!-- mod by nics 2009.03.03 -->
<!--              <tr> -->
              <tr align="center" bgcolor="#FFCC33"> 
<!-- end of mod by nics 2009.03.03 -->
<!-- add by nics 2009.03.03 -->
                <td rowspan="2" nowrap bgcolor="#6495ED">&nbsp;本船情報&nbsp;</td>
<!-- end of add by nics 2009.03.03 -->
                <td bgcolor="#FFCC33" nowrap><font color="#000000">船社</font></td>
<!-- add by nics 2009.03.03 -->
                <td bgcolor="#FFCC33" nowrap><font color="#000000">船名</font></td>
                <td bgcolor="#FFCC33" nowrap>Voyage No.<font color="#FFFFFF"><b> 
                </b></font></td>
                <td bgcolor="#FFCC33" nowrap>仕出港</td>
                <td bgcolor="#FFCC33" nowrap>前港</td>
              </tr>
              <tr align="center"> 
<!-- end of add by nics 2009.03.03 -->
                <td bgcolor="#FFFFFF" nowrap>
<% ' 船社
    If anyTmp(6)<>"" Then
        Response.Write anyTmp(6)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- commented by nics 2009.03.03
                <td bgcolor="#FFCC33" nowrap><font color="#000000">船名</font></td>
end of comment by nics 2009.03.03 -->
                <td bgcolor="#FFFFFF" nowrap>
<% ' 船名
    If anyTmp(7)<>"" Then
        Response.Write anyTmp(7)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- commented by nics 2009.03.03
                <td bgcolor="#FFCC33" nowrap>Voyage No.<font color="#FFFFFF"><b> 
                </b></font></td>
end of comment by nics 2009.03.03 -->
                <td bgcolor="#FFFFFF" nowrap>
<% ' 次航
    If anyTmp(8)<>"" Then
        Response.Write anyTmp(8)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- commented by nics 2009.03.03
                <td bgcolor="#FFCC33" nowrap>仕出港</td>
end of comment by nics 2009.03.03 -->
                <td bgcolor="#FFFFFF" nowrap>
<% ' 仕出港
    If anyTmp(9)<>"" Then
        Response.Write anyTmp(9)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- commented by nics 2009.03.03
                <td bgcolor="#FFCC33" nowrap>前港</td>
end of comment by nics 2009.03.03 -->
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
<!-- mod by nics 2009.03.03 -->
<!--            <br>-->
            <font size="-1"><br></font>
<!-- end of mod by nics 2009.03.03 -->
<!---------------位置情報------------------------------------------- commented by nics 2009.03.03 -->
<!-- commented by nics 2009.03.03
            <table>
              <tr>
                <td><img src="../gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>位置情報　</b></td>
                <td><img src="../gif/hr.gif"></td>
              </tr>
            </table>
end of comment by nics 2009.03.03 -->
            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center"> 
<!-- add by nics 2009.03.03 -->
                <td rowspan="5" nowrap bgcolor="#6495ED">&nbsp;位置情報&nbsp;</td>
<!-- end of add by nics 2009.03.03 -->
<!-- commented by nics 2009.03.03
                <td nowrap align="center" bgcolor="#FFCC33">場所</td>
                <td nowrap bgcolor="#FFCC33">仕出港<font size="-1"><sup>(※1)</sup></font></td>
end of comment by nics 2009.03.03 -->
<!-- mod by nics 2009.03.03 -->
<!--                <td nowrap bgcolor="#FFCC33">前港<font size="-1"><sup>(※1)</sup></font></td>-->
                <td nowrap bgcolor="#FFCC33">前港<font size="-1"><sup>(※3)</sup></font></td>
<!-- end of mod by nics 2009.03.03 -->
                <td colspan="4" nowrap bgcolor="#FFCC33">ターミナル</td>
                <td nowrap bgcolor="#FFCC33">&nbsp;</td>
                <td colspan="3" nowrap bgcolor="#FFCC33">陸上輸送</td>
              </tr>
              <tr align="center"> 
<!-- commented by nics 2009.03.03
                <td nowrap rowspan="2" bgcolor="#FFFFCC">工程</td>
                <td nowrap rowspan="2" bgcolor="#FFFFCC">離岸完了</td>
end of comment by nics 2009.03.03 -->
                <td nowrap rowspan="2" bgcolor="#FFFFCC">離岸完了</td>
                <td nowrap colspan="2" bgcolor="#FFFFCC">入港</td>
                <td nowrap colspan="2" bgcolor="#FFFFCC">ヤード</td>
                <td nowrap bgcolor="#FFFFCC">搬出予約</td>
                <td nowrap bgcolor="#FFFFCC">倉庫到着</td>
                <td nowrap rowspan="2" bgcolor="#FFFFCC">デバン<BR>完了</td>
                <td nowrap rowspan="2" bgcolor="#FFFFCC">空コン<BR>返却完了</td>
              </tr>
              <tr align="center" bgcolor="#FFFFCC">
                <td nowrap>計画</td>
                <td nowrap>予定／完了</td>
                <td nowrap>搬入完了</td>
                <td nowrap>搬出完了</td>
                <td nowrap>予約／－</td>
                <td nowrap>指示／完了</td>
              </tr>
              <tr align="center"> 
<!-- commented by nics 2009.03.03
                <td bgcolor="#FFFFCC" rowspan="2" nowrap>時刻</td>
                <td align="center" rowspan="2" nowrap>
<% ' 仕出港 - 離岸完了
    Response.Write DispDateTimeCell(anyTmp(11),11)
%>
                </td>
end of comment by nics 2009.03.03 -->
                <td align="center" rowspan="2" nowrap>
<% ' 前港 - 離岸完了 $追加
    Response.Write DispDateTimeCell(anyTmp(37),11)
%>
                </td>
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
'2006/10/31
    sTemp=DispReserveCell(anyTmp(35),anyTmp(36),sColor)
'    sTemp=DispReserveCell("","",sColor)
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
'2006/03/23
'    Response.Write DispDateTimeCell(anyTmp(30),11)
    Response.Write DispDateTimeCell("",11)
%>
                </td>
                <td align="center" nowrap>
<% ' 陸上輸送 - 倉庫到着完了
    Response.Write DispDateTimeCell(anyTmp(14),11)
%>
                </td>
              </tr>
            </table>
<!-- mod by nics 2009.03.03 -->
<!--            <table border="0" cellspacing="2" cellpadding="1">-->
            <table border="0" cellspacing="0" cellpadding="0">
<!-- end of mod by nics 2009.03.03 -->
              <tr> 
                <td width="15">&nbsp;</td>
<!-- mod by nics 2009.03.03 -->
<!--                <td><font color="#000000" size="-1">（※1）仕出港・前港の時刻は、現地時間です。</font></td>-->
                <td><font color="#000000" size="-1">（※3）前港の時刻は、現地時間です。</font></td>
<!-- end of mod by nics 2009.03.03 -->
              </tr>
            </table>
<!-- commented by nics 2009.03.03
            <br>
end of comment by nics 2009.03.03 -->
<!---------------手続き及び搬入確認--------------------------------- commented by nics 2009.03.03 -->
<!-----手続情報---------------->
<!-- commented by nics 2009.03.03
            <table>
              <tr>
                <td><img src="../gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>手続き及びターミナル搬出可否情報</b></td>
                <td><img src="../gif/hr.gif"></td>
              </tr>
            </table>
end of comment by nics 2009.03.03 -->
            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center"> 
<!-- add by nics 2009.03.03 -->
                <td rowspan="5" nowrap bgcolor="#6495ED">手続き及び<br>ターミナル<br>搬出可否<br>情報</td>
<!-- end of add by nics 2009.03.03 -->
<!-- commented by nics 2009.03.03
                <td rowspan="3" nowrap bgcolor="#FFCC33">項目</td>
end of comment by nics 2009.03.03 -->
                <td colspan="4" nowrap bgcolor="#FFCC33">行政手続き</td>
                <td rowspan="3" nowrap bgcolor="#FFCC33">商取引<br>
                  DO発行</td>
<!-- mod by nics 2010.02.18 
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
<!-- end of mod by nics 2010.02.18 -->
                <td rowspan="3" nowrap bgcolor="#FFCC33">ターミナル<br>
                  搬出可否</td>
<%'HiTS ver2 ADD by SEIKO n.Ooshige 2003/06/26%>
                <td rowspan="3" nowrap bgcolor="#FFCC33">リマーク</td>
              </tr>
              <tr> 
                <td align="center" nowrap bgcolor="#FFFFCC">搬入確認時刻</td>
                <td align="center" nowrap bgcolor="#FFFFCC" rowspan="2">動植物</td>
                <td align="center" nowrap bgcolor="#FFFFCC" rowspan="2">個別搬入</td>
<!-- mod by nics 2009.03.03 -->
<!--                <td align="left" nowrap bgcolor="#FFFFCC" rowspan="2">通関 /<BR>保税輸送<font size="-1"><sup>(※2)</sup></font></td>-->
                <td align="left" nowrap bgcolor="#FFFFCC" rowspan="2">通関 /<BR>保税輸送<font size="-1"><sup>(※4)</sup></font></td>
<!-- end of mod by nics 2009.03.03 -->
              </tr>
              <tr> 
                <td align="center" nowrap bgcolor="#FFFFCC">予定／完了</td>
              </tr>
              <tr align="center"> 
<!-- commented by nics 2009.03.03
                <td bgcolor="#FFFFCC" rowspan="2" nowrap>情報</td>
end of comment by nics 2009.03.03 -->
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
                <td align="center" rowspan="2" nowrap>
<% ' 商取引ＤＯ発行
    If anyTmp(21)<>"Y" Then
        Response.Write "×"
    Else
        Response.Write "○"
    End If
%>
                </td>
<!-- mod by nics 2010.02.18 
                <td align="center" rowspan="2" nowrap>-->
<% ' フリータイム ＢＬ番号指定あるいは事前情報入力のみ表示する
    If Not bDispMode  or USER <> "" Then
	    Response.Write "<td align='center' rowspan='2' nowrap>"
	    If anyTmp(22)<>"" Then
	        If anyTmp(22)<DispDateTime(Now,10) Then
	            Response.Write "<font color='#FF0000'>"
	        Else
	            Response.Write "<font color='#000000'>"
	        End If
	        Response.Write DispDateTimeCell(anyTmp(22),5)
	        Response.Write "</font>"
	    Else
	        Response.Write DispDateTimeCell(anyTmp(22),5)
	    End If
'Add by nics 2010.02.18 
	    Response.Write "</td>"
'end of Add by nics 2010.02.18 
	End If
'end of Mod by nics 2010.02.18 
%>
<!-- del by nics 2010.02.18 
                </td>
	 end of del by nics 2010.02.18 -->
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
                <td align="center" rowspan="2" nowrap>
<% ' リマーク
    If anyTmp(35)="4" Then
	    Response.Write "夕積"
    End If
%>
              　</td>
              </tr>
              <tr>
                <td align="center" nowrap>
<% ' 搬入確認完了時刻
    Response.Write DispDateTimeCell(anyTmp(18),5)
%>
                </td>
              </tr>
            </table>
<!-- mod by nics 2009.03.03 -->
<!--            <table border="0" cellspacing="2" cellpadding="1">-->
            <table border="0" cellspacing="0" cellpadding="0">
<!-- end of mod by nics 2009.03.03 -->
              <tr> 
                <td width="15">&nbsp;</td>
<!-- mod by nics 2009.03.03 -->
<!--                <td><font color="#000000" size="-1">（※2）○をクリックすると保税輸送期間が表示されます。</font></td>-->
                <td><font color="#000000" size="-1">（※4）保税輸送の場合、○をクリックすると保税輸送期間が表示されます。</font></td>
<!-- end of mod by nics 2009.03.03 -->
              </tr>
            </table>
<!-- commented by nics 2009.03.03
            <br>
end of comment by nics 2009.03.03 -->
<form>
      <input type=button value='表示データの更新' OnClick="JavaScript:window.location.href='impreload.asp?request=impdetail.asp'">
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
