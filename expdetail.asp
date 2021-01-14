<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' Tempファイル属性のチェック
    CheckTempFile "EXPORT", "expentry.asp"

    ' 指定引数の取得
    Dim iLineNo
    iLineNo = CInt(Request.QueryString("line"))
    Dim iReturn
    iReturn = Session.Contents("dispreturn")

    ' 表示モードの取得
    Dim bDispMode          ' true=コンテナ検索 / false=ブッキング検索
    If Session.Contents("findkind")="Cntnr" Then
        bDispMode = true
    Else
        bDispMode = false
    End If

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
    bSingle = false                    '単独検索結果フラグ
    If iLineNo=1 And LineNo=1 Then
        '単独検索結果かどうかチェックする
        if ti.AtEndOfStream Then
            bSingle = true
        End If
    End If
    ti.Close

    ' 輸出コンテナ照会詳細
    WriteLog fs, "1007","輸出コンテナ照会-単独コンテナ","00", anyTmp(1) & ","

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
End of comment by seiko-denki 2003.07.17 -->
<!-- mod by nics 2009.02.12 -->
<!--		<table width=95% cellpadding=3>-->
		<table width=95% cellpadding=0>
<!-- end of mod by nics 2009.02.12 -->
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
<!-- mod by nics 2009.02.12 -->
<!--          <td>　<br>-->
          <td>
<!-- end of mod by nics 2009.02.12 -->
            <table border=1 cellpadding="3" cellspacing="1">
              <tr> 
<% ' Booking No
    If Not bDispMode Then
        Response.Write "<td bgcolor='#003399' background='gif/tableback.gif' nowrap><font color='#FFFFFF'><b>Booking No</b></font></td>"
        Response.Write "<td bgcolor='#FFFFFF' nowrap>" & anyTmp(0) & "</td>"
    End If
%>
                <td bgcolor="#003399" background="gif/tableback.gif" nowrap><font color="#FFFFFF"><b>コンテナNo.</b></font></td>
                <td bgcolor="#FFFFFF" nowrap>
<% ' コンテナNo.
    Response.Write anyTmp(1)
%>
                </td>
              </tr>
            </table>
<!-- mod by nics 2009.02.12 -->
<!--			<BR>-->
			<font size="-2"><BR></font>
<!-- end of mod by nics 2009.02.12 -->
<!---------------基本情報------------------------------------------- commented by nics 2009.02.12 -->
<!-- commented by nics 2009.02.12
            <table>
              <tr>
                <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>基本情報　</b></td>
                <td><img src="gif/hr.gif"></td>
              </tr>
            </table>
end of comment by nics 2009.02.12 -->
            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center" bgcolor="#FFCC33"> 
<!-- add by nics 2009.02.12 -->
                <td rowspan="3" nowrap bgcolor="#6495ED">&nbsp;基本情報&nbsp;</td>
<!-- end of add by nics 2009.02.12 -->
<!-- add by mes(2005/3/28) テアウェイト追加 -->
<!--                <td colspan="4" nowrap>空コンテナ</td>-->
<!--                <td colspan="5" nowrap>空コンテナ</td>-->
		<td colspan="6" nowrap>空コンテナ</td>
<!-- end mes -->
<!-- mod by nics 2009.02.12 -->
<!--                <td colspan="5" nowrap bgcolor="#FFCC33">バンニング後コンテナ</td>-->
                <td colspan="4" nowrap bgcolor="#FFCC33">バンニング後コンテナ</td>
<!-- end of mod by nics 2009.02.12 -->
<!-- commented by nics 2009.02.12
                <td bgcolor="#FFCC33" nowrap colspan="2">搬入受付期間</td>
end of comment by nics 2009.02.12 -->
<!-- add by nics 2009.02.12 -->
                <td rowspan="2" nowrap bgcolor="#FFCC33">搬入ターミナル<br>(蔵置場所コード)</td>
                <td rowspan="2" nowrap bgcolor="#FFCC33">本船担当<br>オペレータ</td>
<!-- end of add by nics 2009.02.12 -->
              </tr>
              <tr align="center" bgcolor="#FFFF99"> 
<!-- mod by nics 2009.02.12 -->
<!--                <td nowrap><font color="#000000">空コン受取場所</font></td>-->
                <td nowrap><font color="#000000">受取場所</font></td>
<!-- end of mod by nics 2009.02.12 -->
                <td nowrap><font color="#000000">サイズ</font></td>
<!-- Add-S MES Aoyagi 2010.11.23 -->
		<td nowrap><font color="#000000">タイプ</font></td>
<!-- Add-E MES Aoyagi 2010.11.23 -->
                <td nowrap>高さ<font size="-1"><sup>(※1)</sup></font></td>
<!-- add by mes(2005/3/28) テアウェイト追加 -->
                <td nowrap><font color="#000000">TW</font></td>
<!-- end mes -->
                <td nowrap><font color="#000000">リーファ</font></td>
                <td nowrap><font color="#000000">シールNo.</font></td>
                <td nowrap><font color="#000000">貨物重量(t)</font></td>
                <td nowrap><font color="#000000">総重量(t)</font></td>
<!-- mod by nics 2009.02.12 -->
<!--                <td nowrap><font color="#000000">危険品</font></td>-->
                <td nowrap><font color="#000000">危険品<font size="-1"><sup>(※2)</sup></font></font></td>
<!-- end of mod by nics 2009.02.12 -->
<!-- commented by nics 2009.02.12
                <td nowrap><font color="#000000">搬入ターミナル名</font></td>
                <td nowrap><font color="#000000">オープン日</font></td>
                <td nowrap>クローズ日</td>
end of comment by nics 2009.02.12 -->
              </tr>
              <tr> 
                <td nowrap align="center">
<% ' 空コン受取場所
    If anyTmp(2)<>"" Then
        Response.Write anyTmp(2)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' サイズ
    If anyTmp(3)<>"" Then
        Response.Write anyTmp(3)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" nowrap>

<!-- Add-S MES Aoyagi 2010.11.23 -->
<% ' タイプ
    If anyTmp(39)<>"" Then
        Response.Write anyTmp(39)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" nowrap>
<!-- Add-E MES Aoyagi 2010.11.23 -->

<% ' 高さ
    If anyTmp(4)<>"" Then
        Response.Write anyTmp(4)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- add by mes(2005/3/28) テアウェイト追加 -->
                <td align="center" nowrap>
<% ' テアウェイト
    If anyTmp(32)<>"" And anyTmp(32)>0 Then
    	If anyTmp(32)<100 then
	        dWeight=anyTmp(32) * 100
	    Else
	        dWeight=anyTmp(32)
	    End If
        Response.Write dWeight
    Else
        Response.Write "－"
    End If
%>
                </td>
<!-- end mes -->
                <td align="center" nowrap>
<% ' リーファー
    If anyTmp(5)="R" Then
        Response.Write "○"
    ElseIf anyTmp(5)<>"" Then
        Response.Write "－"
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' シールNo.
    If anyTmp(7)<>"" Then
        Response.Write anyTmp(7)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' 貨物重量 $追加
    If anyTmp(27)<>"" And anyTmp(27)<>"0" Then
        dWeight=anyTmp(27) / 10
        Response.Write dWeight
    Else
        Response.Write "－"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' 総重量
    If anyTmp(8)<>"" And anyTmp(8)<>"0" Then
        dWeight=anyTmp(8) / 10
        Response.Write dWeight
    Else
        Response.Write "－"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' 危険品
    If anyTmp(31)="H" Then
        Response.Write "○"
    ElseIf anyTmp(31)<>"" Then
        Response.Write "－"
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- commented by nics 2009.02.12
                <td align="center" nowrap>
<% ' 搬入ターミナル名
    If anyTmp(6)<>"" Then
        Response.Write anyTmp(6)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' CYオープン
    Response.Write DispDateTimeCell(anyTmp(9),5)
%>
                </td>
                <td align="center" nowrap>
<% ' CYクローズ
    Response.Write DispDateTimeCell(anyTmp(10),5)
%>
                </td>
end of comment by nics 2009.02.12 -->
<!-- add by nics 2009.02.12 -->
                <td align="center" nowrap>
<% ' 搬入ターミナル(蔵置場所コード)
    strDisp = "<br>"
    If anyTmp(6) <> "" Then
        strDisp = anyTmp(6)
        If anyTmp(36) <> "" Then
            strDisp = strDisp & "(" & anyTmp(36) & ")"
        End If
    End If
    Response.Write strDisp
%>
                </td>
                <td align="center" nowrap>
<% ' 担当オペレータ
    If anyTmp(37)<>"" Then
        Response.Write anyTmp(37)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- end of add by nics 2009.02.12 -->
              </tr>
            </table>
<!-- mod by nics 2009.02.12 -->
<!--            <table border="0" cellspacing="2" cellpadding="1">-->
            <table border="0" cellspacing="0" cellpadding="0">
<!-- end of mod by nics 2009.02.12 -->
              <tr> 
                <td width="15">&nbsp;</td>
<!-- mod by nics 2009.02.12 -->
<!--                <td><font color="#000000" size="-1">(※1)96=HC</font></td>-->
                <td><font color="#000000" size="-1">(※1)96=HC &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; （※2）消防法に関わる危険品の有無</font></td>
<!-- end of mod by nics 2009.02.12 -->
              </tr>
            </table>
<!-- commented by nics 2009.02.09
            <BR>
end of comment by nics 2009.02.09 -->
<!---------------本船情報------------------------------------------- commented by nics 2009.02.12 -->
<!-- commented by nics 2009.02.12
            <table>
              <tr>
                <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>本船情報　</b></td>
                <td><img src="gif/hr.gif"></td>
              </tr>
            </table>
end of comment by nics 2009.02.12 -->
            <table border=1 cellpadding="3" cellspacing="1">
<!-- mod by nics 2009.02.12 -->
<!--              <tr> -->
              <tr align="center" bgcolor="#FFCC33"> 
<!-- end of mod by nics 2009.02.12 -->
<!-- add by nics 2009.02.12 -->
                <td rowspan="2" nowrap bgcolor="#6495ED">&nbsp;本船情報&nbsp;</td>
<!-- end of add by nics 2009.02.12 -->
                <td bgcolor="#FFCC33" nowrap>船社</td>
<!-- add by nics 2009.02.12 -->
                <td bgcolor="#FFCC33" nowrap><font color="#000000">船名</font></td>
                <td bgcolor="#FFCC33" nowrap>Voyage No.<font color="#FFFFFF"><b> 
                </b></font></td>
                <td bgcolor="#FFCC33" nowrap>仕向港</td>
              </tr> 
              <tr align="center"> 
<!-- end of add by nics 2009.02.12 -->
                <td bgcolor="#FFFFFF">
<% ' 船社
    If anyTmp(11)<>"" Then
        Response.Write anyTmp(11)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- commented by nics 2009.02.12
                <td bgcolor="#FFCC33" nowrap><font color="#000000">船名</font></td>
end of comment by nics 2009.02.12 -->
                <td bgcolor="#FFFFFF" nowrap>
<% ' 船名
    If anyTmp(12)<>"" Then
        Response.Write anyTmp(12)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- commented by nics 2009.02.12
                <td bgcolor="#FFCC33" nowrap>Voyage No.<font color="#FFFFFF"><b> 
                </b></font></td>
end of comment by nics 2009.02.12 -->
                <td bgcolor="#FFFFFF" nowrap>
<% ' 次航
    If anyTmp(13)<>"" Then
        Response.Write anyTmp(13)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- commented by nics 2009.02.12
                <td bgcolor="#FFCC33" nowrap>仕向港</td>
end of comment by nics 2009.02.12 -->
                <td bgcolor="#FFFFFF" nowrap>
<% ' 仕向港
    If anyTmp(14)<>"" Then
        Response.Write anyTmp(14)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
              </tr>
            </table>
<!-- mod by nics 2009.02.12 -->
<!--            <BR>-->
            <font size="-2"><BR></font>
<!-- end of mod by nics 2009.02.12 -->
<!---------------位置情報------------------------------------------- commented by nics 2009.02.12 -->
<!-- commented by nics 2009.02.12
            <table>
              <tr>
                <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>位置情報　</b></td>
                <td><img src="gif/hr.gif"></td>
              </tr>
            </table> 
            <br>
end of comment by nics 2009.02.12 -->
            <table border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
                  <tr align="center" bgcolor="#FFCC33">
<!-- add by nics 2009.02.12 -->
                    <td rowspan="5" nowrap bgcolor="#6495ED">&nbsp;位置情報&nbsp;</td>
<!-- end of add by nics 2009.02.12 -->
<!-- commented by nics 2009.02.12
                    <td nowrap>場所</td>
end of comment by nics 2009.02.12 -->
                    <td colspan="3" nowrap>陸上輸送</td>
                    <td nowrap bgcolor="#FFCC33">ストックヤード</td>
                    <td colspan="4" nowrap bgcolor="#FFCC33">ターミナル</td>
                  </tr>
                  <tr align="center" bgcolor="#FFFF99">
<!-- commented by nics 2009.02.12
                    <td nowrap rowspan="2"><font color="#000000">工程</font></td>
end of comment by nics 2009.02.12 -->
                    <td nowrap rowspan="2"><font color="#000000">空コン受取</font></td>
                    <td nowrap><font color="#000000">倉庫到着</font></td>
<!--  mod by mes 2013.8.29     <td nowrap rowspan="2"><font color="#000000">バンニング</font></td> -->
                    <td nowrap rowspan="2"><font color="#000000">倉庫出発</font></td>
                    <td nowrap><font color="#000000">搬入</font></td>
                    <td nowrap><font color="#000000">CY搬入</font></td>
                    <td nowrap rowspan="2"><font color="#000000">船積完了</font></td>
                    <td nowrap colspan="2"><font color="#000000">離岸</font></td>
<!-- commented by nics 2009.02.12
                    <td nowrap><font color="#000000">着岸時刻</font><font size="-1"><sup>(※3)</sup></font></td>
end of comment by nics 2009.02.12 -->
                  </tr>
                  <tr align="center" bgcolor="#FFFF99">
<!-- mod by nics 2009.02.12 -->
<!--                    <td nowrap><font color="#000000">指示<font size="-1"><sup>(※2)</sup></font>／完了</font></td>-->
                    <td nowrap><font color="#000000">指示／完了</font></td>
<!-- end of mod by nics 2009.02.12 -->
                    <td nowrap><font color="#000000">予約／完了</font></td>
                    <td nowrap><font color="#000000">指示／完了</font></td>
                    <td nowrap><font color="#000000">計画</font></td>
                    <td nowrap><font color="#000000">予定／完了</font></td>
<!-- commented by nics 2009.02.12
                    <td nowrap><font color="#000000">予定／完了</font></td>
end of comment by nics 2009.02.12 -->
                  </tr>
                  <tr>
<!-- commented by nics 2009.02.12
                    <td nowrap rowspan="2" bgcolor="#FFFFCC" align="center"><font color="#000000">時刻</font></td>
end of comment by nics 2009.02.12 -->
                    <td rowspan="2" align="center" nowrap><% ' 陸上運送 - 空コン受取
    Response.Write DispDateTimeCell(anyTmp(16),11)
%>
                    </td>
                    <td align="center" nowrap><% ' 陸上運送 - 倉庫到着スケジュール
    If anyTmp(26)<>"" Then
        If anyTmp(26)<anyTmp(17) Then
            Response.Write "<font color='#FF0000'>"
        Else
            Response.Write "<font color='#0000FF'>"
        End If
    End If
    Response.Write DispDateTimeCell(anyTmp(26),11)
    If anyTmp(26)<>"" Then
        Response.Write "</font>"
    End If
%>
                    </td>
                    <td rowspan="2" align="center" nowrap><% ' 陸上運送 - バンニング
    Response.Write DispDateTimeCell(anyTmp(18),11)
%>
                    </td>
                    <td align="center" nowrap><% ' ストックヤード - 搬入予約 $追加
    sTemp=DispReserveCell(anyTmp(28),anyTmp(29),sColor)
    Response.Write sColor
    Response.Write sTemp
    If sColor<>"" Then
        Response.Write "</font>"
    End If
%>
                    </td>
                    <td align="center" nowrap><% ' ターミナル - CY搬入指示 $追加
    If anyTmp(30)<>"" Then
        If Left(anyTmp(30),10)<Left(anyTmp(19),10) Then
            Response.Write "<font color='#FF0000'>"
        Else
            Response.Write "<font color='#0000FF'>"
        End If
    End If
    Response.Write DispDateTimeCell(anyTmp(30),5)
    If anyTmp(30)<>"" Then
        Response.Write "</font>"
    End If
%>
                    </td>
                    <td rowspan="2" align="center" nowrap><% ' ターミナル - 船積完了
    Response.Write DispDateTimeCell(anyTmp(20),11)
%>
                    </td>
                    <td rowspan="2" align="center" nowrap><% ' ターミナル - 離岸スケジュール
    If anyTmp(25)<>"" Then
        Response.Write "<font color='#0000FF'>"
    End If
    Response.Write DispDateTimeCell(anyTmp(25),5)
    If anyTmp(25)<>"" Then
        Response.Write "</font>"
    End If
%>
                    </td>
                    <td align="center" nowrap><% ' ターミナル - 離岸予定
    If anyTmp(15)<>"" Then
        bLate = false
        If anyTmp(21)<>"" Then
            If anyTmp(15)<anyTmp(21) Then
                bLate = true
            End If
        End If
        If anyTmp(25)<>"" Then
            If Left(anyTmp(25),10)<Left(anyTmp(15),10) Then
                bLate = true
            End If
        End If
        If bLate Then
            Response.Write "<font color='#FF0000'>"
        Else
            Response.Write "<font color='#0000FF'>"
        End If
        Response.Write DispDateTimeCell(anyTmp(15),11)
        Response.Write "</font>"
    Else
        Response.Write DispDateTimeCell(anyTmp(15),11)
    End If
%>
                    </td>
<!-- commented by nics 2009.02.12
                    <td align="center" nowrap><% ' 仕向港 - 着岸予定
    If anyTmp(23)<>"" Then
        If anyTmp(22)<>"" Then
            If anyTmp(23)<anyTmp(22) Then
                Response.Write "<font color='#FF0000'>"
            Else
                Response.Write "<font color='#0000FF'>"
            End If
        Else
            Response.Write "<font color='#0000FF'>"
        End If
        Response.Write DispDateTimeCell(anyTmp(23),11)
        Response.Write "</font>"
    Else
        Response.Write DispDateTimeCell(anyTmp(23),11)
    End If
%>
                    </td>
end of comment by nics 2009.02.12 -->
                  </tr>
                  <tr>
                    <td align="center" nowrap><% ' 陸上運送 - 倉庫到着
    Response.Write DispDateTimeCell(anyTmp(17),11)
%>
                    </td>
                    <td align="center" nowrap><% ' ストックヤード - 搬入完了
    Response.Write DispDateTimeCell(anyTmp(24),11)
%>
                    </td>
                    <td align="center" nowrap><% ' ターミナル - CY搬入完了
    Response.Write DispDateTimeCell(anyTmp(19),11)
%>
                    </td>
                    <td align="center" nowrap><% ' ターミナル - 離岸完了
    Response.Write DispDateTimeCell(anyTmp(21),11)
%>
                    </td>
<!-- commented by nics 2009.02.12
                    <td align="center" nowrap><% ' 仕向港 - 着岸完了
    Response.Write DispDateTimeCell(anyTmp(22),11)
%>
                    </td>
end of comment by nics 2009.02.12 -->
                  </tr>
                </table></td>
                <td>&nbsp;</td>
<!-- commented by nics 2009.02.12
                <td valign="top"><table border="1" cellpadding=" 3" cellspacing="1" bgcolor="#FFFFFF">
                  <tr>
                    <td align="center" nowrap bgcolor="#FFCC33">仕向港内位置情報<font size="-1"><sup>(※4)</sup></font></td>
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
end of comment by nics 2009.02.12 -->
              </tr>
            </table>
<BR>
<!---------------手続き及び搬入確認--------------------------------- commented by nics 2009.02.12 -->
<!-- add by nics 2009.02.12 -->
            <table border="0" cellspacing="0" cellpadding="0">
              <tr><td>
            <table border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td>
				  <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center"> 
                <td rowspan="4" nowrap bgcolor="#6495ED">手続き及び<br>搬入確認</td>
                <td bgcolor="#FFCC33" nowrap colspan="2">搬入受付期間</td>
<!-- mod by mes aoyagi 2010.5.13 -->
<!-- commented by nics 2010.02.02 -->
                <td colspan="2" nowrap bgcolor="#FFCC33">行政手続き</td>
<!-- end of comment by nics 2010.02.02 -->
<!-- mod by nics 2010.02.02 -->
<!--                <td colspan="3" nowrap bgcolor="#FFCC33">行政手続き</td> -->
<!-- end of mod by nics 2010.02.02 -->
<!-- end of mod by mes aoyagi 2010.5.13 -->
                <td rowspan="3" nowrap bgcolor="#FFCC33">ターミナル<br>搬入確認</td>
              </tr>
              <tr align="center" bgcolor="#FFFF99">
                <td rowspan="2" nowrap><font color="#000000">オープン日</font></td>
                <td rowspan="2" nowrap>クローズ日</td>
                <td colspan="2" nowrap>X線検査</td>
<!-- del by mes aoyagi 2010.05.13 -->
<!-- add by nics 2010.02.02 -->
<!--                <td rowspan="2" nowrap>通<br>関</td> -->
<!-- end of add by nics 2010.02.02 -->
<!-- del by mes aoyagi 2010.05.13 -->
              </tr>
              <tr align="center" bgcolor="#FFFF99">
                <td colspan="1" nowrap>有無</td>
                <td colspan="1" nowrap>CY返却</td>
              </tr>

              <tr> 
                <td align="center" nowrap>
<% ' CYオープン
    Response.Write DispDateTimeCell(anyTmp(9),5)
%>
                </td>
                <td align="center" nowrap>
<% ' CYクローズ
    Response.Write DispDateTimeCell(anyTmp(10),5)
%>
                </td>
                <td align="center" nowrap>
<% ' X線有無
        If anyTmp(33)<>"" Then
            Response.Write anyTmp(33)
        Else
            Response.Write "<br>"
        End If
%>
                </td>
                <td align="center" nowrap>
<% ' X線CY返却
        If anyTmp(34)<>"" Then
            Response.Write anyTmp(34)
        Else
            Response.Write "<br>"
        End If
%>
                </td>
<!-- del by mes aoyagi 2010.05.13-->
<!-- add by nics 2010.02.02 -->
<!--                <td align="center" nowrap> 
<% ' 通関
        If anyTmp(38)<>"" Then
            Response.Write anyTmp(38)
        Else
            Response.Write "<br>"
        End If
%>
               </td> -->
<!-- end of add by nics 2010.02.02 -->
<!-- del by mes aoyagi 2010.5.13 -->
                <td align="center" nowrap>
<% ' ターミナル搬入確認	
        If anyTmp(35)<>"" Then
            Response.Write anyTmp(35)
        Else
            Response.Write "<br>"
        End If
%>
                </td>
              </tr> 
            </table>
			</td>
              </tr>
            </table>
<!-- end of add by nics 2009.02.12 -->
<!-- mod-s by MES 2015/06/08 表示方法変更対応 -->
              </td>
              <td>&nbsp;</td>
            <td valign="top"><table border="1" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF">
              <tr>
                <td align="center" nowrap bgcolor="#6495ED" colspan="2">位置情報／仕向港<sup>(※3)</sup></td>
              </tr>
              <tr>
                <td align="center" nowrap bgcolor="#ffff99">中国</td>
                <td align="center" nowrap bgcolor="#ffff99">東南アジア</td>
              </tr>
              <tr>
                <td align="center"><table border="0" cellspacing="3">
                  <tr>
                    <td nowrap align="center"><a href="javascript:Submit('Form1')" class="splinkG" onClick="javascript:winOpen('win1','./cct/index.html',560,500)">&nbsp;赤湾&nbsp;</a></td>
                    <td nowrap align="center"><a href="gaibuif/impcntnr.asp?cntnrno=<%Response.Write anyTmp(1)%>&portcode=HUANG" class="splinkY" onClick="">&nbsp;黄埔&nbsp;</a></td>
                    <td nowrap align="center"><a href="gaibuif/impcntnr.asp?cntnrno=<%Response.Write anyTmp(1)%>&portcode=QINGD" class="splinkB" onClick="">&nbsp;青島&nbsp;</a></td>
                    <td nowrap align="center"><a href="#" class="dummylink" onClick="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</a></td>
                  </tr>
                  <tr>

<!-- 2015/11/30 Add-S MES Aoyagi -->
                    <td nowrap align="center"><a href="http://iport.sctcn.com/en-us" target="_blank" class="splinkG" onClick="javascript:winOpen('win1','./sct/index.htm',560,500)">&nbsp;蛇口&nbsp;</a>
<!-- 2015/11/30 Add-E MES Aoyagi -->

<!-- 2015/11/30 Del-S MES Aoyagi
                    <td nowrap align="center"><a href="javascript:Submit('queryForm')" class="splinkG" onClick="javascript:winOpen('win1','./sct/index.asp',560,500)">&nbsp;蛇口&nbsp;</a>
2015/11/30 Del-E MES Aoyagi -->
                    <td nowrap align="center"><a href="gaibuif/impcntnr.asp?cntnrno=<%Response.Write anyTmp(1)%>&portcode=NANSH" class="splinkY" onClick="">&nbsp;南沙&nbsp;</a></td>
                    <td nowrap align="center"><a href="#" class="dummylink" onClick="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</a></td>
                    <td nowrap align="center"><a href="#" class="dummylink" onClick="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</a></td>
                  </tr>
                </table></td>
                <td align="center"><table border="0" cellspacing="2">
<!--
                  <tr>
                    <td nowrap align="center"><a href="gaibuif/impcntnr.asp?cntnrno=<%Response.Write anyTmp(1)%>&portcode=TWTPE" class="splinkR" onClick="">&nbsp;台北&nbsp;</a></td>
                    <td nowrap align="center"><a href="#" class="dummylink" onClick="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</a></td>
                  </tr>
-->
                  <tr>
                    <td nowrap align="center"><a href="gaibuif/impcntnr.asp?cntnrno=<%Response.Write anyTmp(1)%>&portcode=THBKK" class="splinkLG" onClick="">&nbsp;バンコク&nbsp;</a></td>
                  </tr>
                </table></td>
              </tr>
            </table></td>
            <tr>
              <td></td><td></td>
              <td><font color="#000000" size="-1">（※3）ボタンをクリックすると当該港での位置情報等が表示されます（現地時間表示）。</font></td>
            </tr>
<!-- mod-e by MES 2015/06/08 表示方法変更対応 -->
              </tr>
            </table>

<form>
      <input type=button value='表示データの更新' OnClick="JavaScript:window.location.href='expreload.asp?request=expdetail.asp'">
</form>
<!-- 2015/11/30 Del-S MES Aoyagi
<!-- 2014/5/20 Mod-S MES aoyagi URL 
<form name="queryForm" method="get" target="_blank" action="http://iport.sctcn.com/portal/page/portal/PG_IPort/Tab_OI/">
    <input type="hidden" name="p_parametertype" value="ContainerInfo">
    <input type="hidden" name="p_parametervalue" value="<%=anyTmp(1)%>">
-->
<!--
<form name="queryForm" method="post" target="_blank" action="http://oi.sctcn.com/Default.aspx?Action=Nav&Content=CONTAINER%20INFO.%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20&sm=CONTAINER%20INFO.">
    <input type="hidden" name="data" value="<%=anyTmp(1)%>">		
    <input type="hidden" name="OrgMenu" value="">
    <input type="hidden" name="targetPage" value="CONTAINER_INFO">
    <input type="hidden" name="nav" value="CONTAINER INFO.                         ">
</form>
-->
<!-- 2014/5/20 Mod-E MES aoyagi URL -->
<!-- 2013/12/26 DEL-S MES aoyagi
<form name="Form1" method="post" action="http://www.cwcct.com/cct/conhis/con_his_info_show.aspx" id="Form1" target="_blank">
 2013/12/26 DEL-E MES aoyagi -->
<!-- 2013/12/26 ADD-S MES aoyagi -->
<form name="Form1" method="post" action="http://uport.cwcct.com/Portal/Ship/EN/Public/Pub_cntr_history_show.aspx" id="Form1" target="_blank">
<!-- 2013/12/26 ADD-E MES aoyagi -->
<!-- 2013/12/26 DEL-S MES aoyagi
    <input type="hidden" name="Image1.x" value="0" />
    <input type="hidden" name="Image1.y" value="0" />
 2013/12/26 DEL-E MES aoyagi -->
<!--
    <input type="hidden" name="__EVENTTARGET" value="" />
    <input type="hidden" name="__EVENTARGUMENT" value="" />
    <input type="hidden" name="__VIEWSTATE" value="dDwtMzMwNTk0MTUxOztsPEltYWdlMTs+Po9koK7lFKyndTfCh4n1g7KjLvsH" />
-->
<!-- 2013/12/26 DEL-S MES aoyagi
    <input type="hidden" name="cont_no" id="cont_no" value="<%=anyTmp(1)%>" />
    <input type="hidden" name="wyex" value="wyE" />
 2013/12/26 DEL-E MES aoyagi -->
<!-- 2013/12/26 ADD-S MES aoyagi -->
    <input type="hidden" name="txtContainer_no" id="txtContainer_no" value="<%=anyTmp(1)%>" />
    <input type="hidden" name="rdoDisplay" id="rdoHTML" value="HTML" />
<!-- 2013/12/26 ADD-E MES aoyagi -->

</form>
<%
    ' 検索画面から直接飛んできたときは表示する
    If bSingle Then
        Response.Write "<form action='expcsvout.asp'>"
        Response.Write "<center>"
        Response.Write "<input type='submit' name='submit' value='CSVファイル出力'>　"
        Response.Write "<a href='help03.asp'>CSVファイル出力とは？</a>"
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
        DispMenuBarBack "expentry.asp"
    Else
        If iReturn=1 Then
            DispMenuBarBack "explist1.asp"
        ElseIf iReturn=2 Then
            DispMenuBarBack "explist2.asp"
        ElseIf iReturn=3 Then
            DispMenuBarBack "explist3.asp"
        Else
            DispMenuBarBack "explist.asp"
        End If
    End If
%>
</body>
</html>
