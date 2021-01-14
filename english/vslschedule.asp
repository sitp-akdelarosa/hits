<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' セッションのチェック
    CheckLogin "vslentry.asp"

    ' 指定引数の取得
    Dim strVslCode
    Dim strVoyCtrl
    strVslCode = Request.QueryString("vslcode")
    strVoyCtrl = Request.QueryString("voyctrl")

    ' 本船動静情報レコードの取得
    ConnectSvr conn, rsd

    sql = "SELECT CurrentPort FROM sEnvironment"
    'SQLを発行してカレントポートを検索
    rsd.Open sql, conn, 0, 1, 1
    If Not rsd.EOF Then
        strPort = Trim(rsd("CurrentPort"))
    End If
    rsd.Close

    sql = "SELECT VslSchedule.VslCode, VslSchedule.VoyCtrl, VslSchedule.DsVoyage, VslSchedule.LdVoyage, " & _
          "VslSchedule.CYOpen, VslSchedule.CYCut, mVessel.FullName VslName, mShipLine.FullName LineName " & _
          "FROM VslSchedule, mVessel, mShipLine "
    sql = sql & " WHERE VslSchedule.VslCode='" & strVslCode & "' And VslSchedule.VoyCtrl='" & strVoyCtrl & "' And " & _
          "mVessel.VslCode=*VslSchedule.VslCode And mShipLine.ShipLine=*VslSchedule.ShipLine"
    'SQLを発行して本船動静一覧を検索
    rsd.Open sql, conn, 0, 1, 1
%>

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
<%
    DispMenuJava
%>
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------ここから画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/vslentryt.gif" width="506" height="73"></td>
          <td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
        </tr>
        <tr>
          <td align="right" width="100%" height="48"> 
<%
' Added and Commented by seiko-denki 2003.07.18
	DisplayCodeListButton
'    DispMenu
' End of Addition by seiko-denki 2003.07.18
%>
          </td>
        </tr>
      </table>
      <br>　
      <br>　
      <center>
      <table  width="100%">
        <tr>
          <td>  
          <br>
          <center>
          <table>
            <tr>
              <td>　<br>
                <table>
                  <tr>
                    <td><img src="gif/botan.gif" width="17" height="17"></td>
                    <td nowrap><b>本 船 情 報　</b></td>
                    <td><img src="gif/hr.gif"></td>
                  </tr>
                </table>
                <table border=1 cellpadding="3" cellspacing="1">
                  <tr> 
                    <td bgcolor="#003399" background="gif/tableback.gif" nowrap height="21"><font color="#FFFFFF"><b>船名</b></font></td>
                    <td bgcolor="#FFFFFF" nowrap>
<%
    Response.Write Trim(rsd("VslName"))
%>
                    </td>
                    <td bgcolor="#000099" background="gif/tableback.gif" nowrap height="21"><font color="#FFFFFF"><b>次航</b></font></td>
                    <td bgcolor="#FFFFFF" nowrap>
<%
    strVoyage = Trim(rsd("DsVoyage"))
    strTemp = Trim(rsd("LdVoyage"))
    If strVoyage<>strTemp Then
       If strVoyage<>"" Then
           If strTemp<>"" Then
               strVoyage = strVoyage & "/" & strTemp
           End If
       Else
           strVoyage = strTemp
       End If
    End If
    Response.Write strVoyage
%>
                    </td>
                    <td bgcolor="#000099" background="gif/tableback.gif" nowrap height="21"><font color="#FFFFFF"><b>コールサイン</b></font></td>
                    <td bgcolor="#FFFFFF" nowrap>
<%
    Response.Write Trim(rsd("VslCode"))
%>
                    </td>
                  </tr>
                </table>
                <table border=1 cellpadding="3" cellspacing="1">
                  <tr>
                    <td bgcolor="#000099" background="gif/tableback.gif" nowrap><font color="#FFFFFF"><b>船社</b></font></td>
                    <td bgcolor="#FFFFFF" nowrap>
<%
    Response.Write Trim(rsd("LineName"))
%>
                    </td>
                  </tr>
                </table>
                <br>
                <table>
                  <tr>
                    <td><img src="gif/botan.gif" width="17" height="17"></td>
                    <td nowrap><b>博多港情報　</b></td>
                    <td><img src="gif/hr.gif"></td>
                  </tr>
                </table>
                <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
                  <tr bgcolor="#FFCC33" align="center"> 
                    <td nowrap>着岸予定</td>
                    <td nowrap>着岸完了</td>
                    <td nowrap>離岸予定</td>
                    <td nowrap>離岸完了</td>
                    <td nowrap>CYオープン日</td>
                    <td nowrap>CYカット日</td>
                  </tr>
                  <tr>
<%
    strCYOpen = DispDateTime(rsd("CYOpen"),10)
    strCYCut = DispDateTime(rsd("CYCut"),10)
    rsd.Close

    sql = "SELECT ETA, TA, ETD, TD " & _
          "FROM VslPort "
    sql = sql & " WHERE VslCode='" & strVslCode & "' And VoyCtrl='" & strVoyCtrl & "' And " & _
          "PortCode='" & strPort & "'"
    'SQLを発行して本船寄港地を検索
    rsd.Open sql, conn, 0, 1, 1

    Response.Write "<td align='center'>"
    strTemp = DispDateTime(rsd("ETA"),0)
    If strTemp="" Then
        Response.Write "<hr width=80%" & ">"
    Else
        Response.Write strTemp
    End If
    Response.Write "</td>"
    Response.Write "<td align='center'>"
    strTemp = DispDateTime(rsd("TA"),0)
    If strTemp="" Then
        Response.Write "<hr width=80%" & ">"
    Else
        Response.Write strTemp
    End If
    Response.Write "</td>"
    Response.Write "<td align='center'>"
    strTemp = DispDateTime(rsd("ETD"),0)
    If strTemp="" Then
        Response.Write "<hr width=80%" & ">"
    Else
        Response.Write strTemp
    End If
    Response.Write "</td>"
    Response.Write "<td align='center'>"
    strTemp = DispDateTime(rsd("TD"),0)
    If strTemp="" Then
        Response.Write "<hr width=80%" & ">"
    Else
        Response.Write strTemp
    End If
    Response.Write "</td>"
    Response.Write "<td align='center'>"
    If strCYOpen="" Then
        Response.Write "<hr width=80%" & ">"
    Else
        Response.Write strCYOpen
    End If
    Response.Write "</td>"
    Response.Write "<td align='center'>"
    If strCYCut="" Then
        Response.Write "<hr width=80%" & ">"
    Else
        Response.Write strCYCut
    End If
    Response.Write "</td>"

    rsd.Close
%>
                  </tr>
                </table>
                <br>
                <table>
                  <tr>
                    <td><img src="gif/botan.gif" width="17" height="17"></td>
                    <td nowrap><b>寄港地情報　</b></td>
                    <td><img src="gif/hr.gif"></td>
                  </tr>
                </table>
                <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
                  <tr bgcolor="#FFCC33" align="center"> 
                    <td nowrap>寄港地</td>
                    <td nowrap>日時</td>
                    <td nowrap>状態</td>
                  </tr>
<%
    sql = "SELECT VslPort.PortCode, VslPort.ETA, VslPort.TA, VslPort.ETD, VslPort.TD, mPort.FullName " & _
          "FROM VslPort, mPort "
    sql = sql & " WHERE VslPort.VslCode='" & strVslCode & "' And VslPort.VoyCtrl='" & strVoyCtrl & "' And " & _
          "mPort.PortCode=*VslPort.PortCode ORDER BY VslPort.CallSeq"
    'SQLを発行して本船寄港地を検索
    rsd.Open sql, conn, 0, 1, 1
%>
<!-- ここからデータ繰り返し -->
<%
    Do While Not rsd.EOF
        strETA = DispDateTime(rsd("ETA"),0)
        strTA = DispDateTime(rsd("TA"),0)
        strETD = DispDateTime(rsd("ETD"),0)
        strTD = DispDateTime(rsd("TD"),0)
        If strTD<>"" Then
            strDate = strTD
            strStat = "離岸完了"
        ElseIf strETD<>"" Then
            strDate = strETD
            strStat = "離岸予定"
        ElseIf strTA<>"" Then
            strDate = strTA
            strStat = "着岸完了"
        ElseIf strETA<>"" Then
            strDate = strETA
            strStat = "着岸予定"
        Else
            strDate = "<hr width=80%" & ">"
            strStat = "<hr width=80%" & ">"
        End If
        Response.Write "<tr>"
        Response.Write "<td align='center'>" & Trim(rsd("FullName")) & "</td>"
        Response.Write "<td align='center'>" & strDate & "</td>"
        Response.Write "<td align='center'>" & strStat & "</td>"
        Response.Write "</tr>"

        rsd.MoveNext
    Loop
    rsd.Close
    conn.Close
%>
<!-- ここまで -->
                </table>
                <br>
              </td>
            </tr>
          </table>
          </center>
          </td>
        </tr>
      </table>
      <br>
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
<!-------------画面終わり--------------------------->
</body>
</html>

<%
    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' 本船情報照会
    WriteLog fs, "本船情報照会", "選択本船動静," & strVslCode & "/" & strVoyCtrl
%>
