<%@Language="VBScript" %>

<!--#include file="Common.inc"-->
<!--#include file="Vessel.inc"-->

<%
	'	即時搬出システム【港運用】	データ一覧表示

%>

<%
	' セッションのチェック
	CheckLogin "sokuji.asp"

	' 港運コード取得
	sOpe = Trim(Session.Contents("userid"))

	' File System Object の生成
	Set fs=Server.CreateObject("Scripting.FileSystemobject")

	' テンポラリファイル名を作成して、セッション変数に設定
	Dim strFileName
	strFileName = GetNumStr(Session.SessionID, 8) & ".csv"
	Session.Contents("tempfile")=strFileName

	ConnectSvr conn, rsd

	Dim LineNo,strOut
	LineNo = 0

	' 取得したコンテナ情報レコードをテンポラリファイルに書き出し
	' テンポラリファイルのOpen
	strFileName="./temp/" & strFileName
	Set ti=fs.OpenTextFile(Server.MapPath(strFileName),2,True)

	'' DBの読み込み
	sql = "SELECT mShipper.NameAbrev,mShipLine.NameAbrev,mVessel.FullName," & _
				"QuickDel.BLNo,QuickDel.ContNo,QuickDel.RejectFlag,QuickDel.RecSchTime," & _
				"QuickDel.Shipper,QuickDel.ShipLine,QuickDel.VslCode,QuickDel.Forwarder,BL.OpeCode " & _
				"FROM QuickDel,mShipLine,mVessel,mShipper,BL " & _
				"WHERE mShipLine.ShipLine=*QuickDel.ShipLine AND mVessel.VslCode=*QuickDel.VslCode AND " & _
				"mShipper.Shipper=*QuickDel.Shipper AND BL.BLNo=*QuickDel.BLNo"
	rsd.Open sql, conn, 0, 1, 1

	Dim ShipperAbrev(),ShipLineAbrev(),VslFull(),BLNo(),CntrNo(),RejectFlg(),RecSchTime()
	Dim Shipper(),ShipLine(),VslCode(),Forwarder(),OpeCode()
	QdelNo=0
	Do While Not rsd.EOF
		ReDim Preserve ShipperAbrev(QdelNo)
		ReDim Preserve ShipLineAbrev(QdelNo)
		ReDim Preserve VslFull(QdelNo)
		ReDim Preserve BLNo(QdelNo)
		ReDim Preserve CntrNo(QdelNo)
		ReDim Preserve RejectFlg(QdelNo)
		ReDim Preserve RecSchTime(QdelNo)
		ReDim Preserve Shipper(QdelNo)
		ReDim Preserve ShipLine(QdelNo)
		ReDim Preserve VslCode(QdelNo)
		ReDim Preserve Forwarder(QdelNo)
		ReDim Preserve OpeCode(QdelNo)
		ShipperAbrev(QdelNo) = Trim(rsd(0))
		ShipLineAbrev(QdelNo) = Trim(rsd(1))
		VslFull(QdelNo) = Trim(rsd(2))
		BLNo(QdelNo) = Trim(rsd(3))
		CntrNo(QdelNo) = Trim(rsd(4))
		RejectFlg(QdelNo) = Trim(rsd(5))
		RecSchTime(QdelNo) = DispDateTime(rsd(6),0)
		Shipper(QdelNo) = Trim(rsd(7))
		ShipLine(QdelNo) = Trim(rsd(8))
		VslCode(QdelNo) = Trim(rsd(9))
		Forwarder(QdelNo) = Trim(rsd(10))
		OpeCode(QdelNo) = Trim(rsd(11))
		QdelNo=QdelNo+1
	  rsd.MoveNext
	Loop
	rsd.Close

	For i=0 to QdelNo-1
		'' BLが存在しなければ、
		If BLNo(i) = "" Then
			sql = "SELECT BL.OpeCode FROM BL,ImportCont " & _
						"WHERE BL.VslCode=*ImportCont.VslCode AND BL.VoyCtrl=*ImportCont.VoyCtrl AND " & _
						"ImportCont.ContNo='" & CntrNo(i) & "' ORDER BY ImportCont.UpdtTime DESC"
			rsd.Open sql, conn, 0, 1, 1
			Do While Not rsd.EOF
				OpeCode(i) = Trim(rsd(0))
				Exit Do
				rsd.MoveNext
			Loop
			rsd.Close
		End If

		If OpeCode(i)=sOpe Then
			strOut = ShipperAbrev(i) & ","
			strOut = strOut & ShipLineAbrev(i) & ","
			strOut = strOut & VslFull(i) & ","
			strOut = strOut & BLNo(i) & ","
			strOut = strOut & CntrNo(i) & ","
			If RejectFlg(i) = "0" then
				strOut = strOut & "○" & ","
			ElseIf RejectFlg(i) = "1" then
				strOut = strOut & "×" & ","
			Else
				strOut = strOut & "" & ","
			End If
			strOut = strOut & RecSchTime(i) & ","
			strOut = strOut & Shipper(i) & ","
			strOut = strOut & ShipLine(i) & ","
			strOut = strOut & VslCode(i) & ","
			strOut = strOut & Forwarder(i)
			ti.WriteLine strOut
		End If
	Next
	ti.Close

	' 表示ファイルのOpen
	Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

	' テンポラリファイルの読み込み
	Dim strData()
	LineNo=0
	Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)
	Do While Not ti.AtEndOfStream
		strTemp=ti.ReadLine
		ReDim Preserve strData(LineNo)
		strData(LineNo) = strTemp
		LineNo=LineNo+1
	Loop
	ti.Close

	If LineNo>0 Then
		Response.Redirect "sokuji-koun-list.asp"
		Response.End
'	Else
'		Response.Redirect "sokuji-koun-new.asp?kind=1"
'		Response.End
	End If

%>
<html>
<head>
<title>即時搬出申込み情報一覧（港運）</title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
<%
    DispMenuJava
%>
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------ここからログイン入力画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/sokuji2t.gif" width="506" height="73"></td>
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
<%
	strNowTime = Year(Now) & "年" & _
		Right("0" & Month(Now), 2) & "月" & _
		Right("0" & Day(Now), 2) & "日" & _
		Right("0" & Hour(Now), 2) & "時" & _
		Right("0" & Minute(Now), 2) & "分現在の情報"

%>
					&nbsp;&nbsp;<%=strNowTime%>
					</font>
				</td>
			</tr>
		</table>

      <table>
        <tr>
          <td> 

	        <table>
	          <tr>
	            <td><img src="gif/botan.gif" width="17" height="17"></td>
	            <td nowrap><b>（港運用）即時搬出申込み情報一覧</b></td>
	            <td><img src="gif/hr.gif"></td>
	          </tr>
	        </table>
			<BR>
<% 
	If LineNo=0 Then
		Response.Write "<BR>"
		' エラーメッセージの表示
		Dim strErrMsg
		strErrMsg = "表示出来るデータが存在しません。"
		DispInformationMessage strErrMsg 
	End If
%>
			<BR>
			<div align=left>
			<input type=button value="表示データの更新" onclick="window.location.href='sokuji-koun-updtchk.asp'">
			</div>
		  </td>
		</tr>
	  </table>

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
<!-------------ログイン画面終わり--------------------------->
<%
    DispMenuBarBack "http://www.hits-h.com/index.asp"
%>
</body>
</html>

<%
    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")
	' Log作成
    WriteLog fs, "即時搬出システム-即時搬出指示情報（港運用）", ""
%>
