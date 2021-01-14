<%@ LANGUAGE="VBScript" %>
<%
Option Explicit
%>
<!--#include file="Common.inc"-->
<!--#include file="Common2.inc"-->
<!--#include file="SYWB032.inc"-->
<html>

<head>
<title>シャーシ属性登録画面</title>
</head>
<body>

<img border="0" src="image/title01.gif" width="311" height="42">
<br><br>
<center>
<p><img border="0" src="image/title26.gif" width="236" height="34"><p>
</center>

		<font face="ＭＳ ゴシック">
<%
	Dim conn, rsd, sql									'ＤＢ接続
	Dim sUsrID, sGrpID, sUsrName, sGrpName, sOperator	'ユーザ情報
	Dim sYMD, sHH										'指定日付
	Dim sChkChassisID									'シャーシID
	Dim s_chg_GrpID, s_chg_UsrID						'変更先グループ、ユーザ
	Dim sNotDelFlag										'｢搬出コンテナを載せない｣指定（Y：オン）
	Dim sNightFlag										'｢夕積みのみ載せる｣指定（Y：オン）
	Dim iOpeNo, iDualOpeNo								'作業番号、デュアル作業番号
	Dim iOpeOrder										'作業順位
	Dim sWk

	'ＤＢ接続
	Call ConnectSvr(conn, rsd)

	'ユーザ情報の取得
	Call GetUserInfo(conn, rsd, sUsrID, sGrpID, sUsrName, sGrpName, sOperator)
	
	'指定日付取得
	sYMD = TRIM(Request.Form("YMD"))

	'登録チェック
	If Request.Form("sy_zaiko")  <> "" Then
		sChkChassisID = Request.Form("sy_zaiko")	'手入力
	ElseIf Request.Form("SELECT1")  <> "No0" Then
		sChkChassisID = Request.Form("SELECT1")		'在庫選択
	Else
		sChkChassisID = Request.Form("SELECT2")		'非在庫選択
	End If

	sNotDelFlag = ""	'｢搬出コンテナを載せない｣指定（Y：オン）
	sNightFlag = ""		'｢夕積みのみ載せる｣指定（Y：オン）
	If Request.Form("check1") = "on" Then
		sNotDelFlag = "Y"	'｢搬出コンテナを載せない｣指定（Y：オン）
	End If
	If Request.Form("check2") = "on" Then
		sNightFlag = "Y"	'｢夕積みのみ載せる｣指定（Y：オン）
	End If

	'グループ変更時の処理
	if Request.Form("check3") = "on" then		'グループ変更の場合
		'指定シャーシの使用予定があるかチェックする
		If Not ChkAppCha(conn, rsd, sChkChassisID) Then
			%><center><%
			Response.Write sChkChassisID
			Response.Write "　のシャーシは予約されているので他のグループには変更できません。</p>"
			%><A HREF="JavaScript:history.back()">
				<BR>シャーシ属性画面へ戻る</A></CENTER> <%
			Response.Write "</body>"
			Response.Write "</html>"
			Response.end
		End If

		'変更先グループ・ユーザの取得
		s_chg_GrpID = trim(Request.Form("SELECT3"))		'グループコード
		sql = "SELECT UserID,GroupID FROM sMUserGroup" & _
		  " WHERE RTRIM(GroupID) = '" & s_chg_GrpID & "'"
		rsd.Open sql, conn, 0, 1, 1
		if not rsd.eof then
			s_chg_UsrID = rsd("UserID")					'ユーザコード
		End If
		rsd.close	
	End If

	'シャーシ検索
	sql = "SELECT * FROM sChassis" & _
	  " WHERE RTRIM(GroupID) = '" & sGrpID & "'" & _
	  "   AND ChassisId = '" & sChkChassisID & "'"
	rsd.Open sql, conn, 0, 2, 1

	If rsd.EOF Then	
		rsd.close
		%><center><%
		Response.Write sChkChassisID
		Response.Write "　のシャーシは存在しません。</p>"	
		%><A HREF="JavaScript:history.back()">
			<BR>シャーシ属性画面へ戻る</A></CENTER> <%
		Response.Write "</body>"
		Response.Write "</html>"
		Response.end
	End If

	'グループ変更の場合にシャーシの現状をチェック
	If Request.Form("check3") = "on" Then	'グループ変更
		If rsd("ContFlag") = "Y" Then	'コンテナフラグ
			rsd.close
			%><center><%
			Response.Write sChkChassisID
			Response.Write "　のシャーシにはコンテナがあります。</p>"
			%><A HREF="JavaScript:history.back()">
				<BR>シャーシ属性画面へ戻る</A></CENTER> <%
			Response.Write "</body>"
			Response.Write "</html>"
			Response.end
		End If

		If rsd("StackFlag") = "W" Then	'シャトル作業中
			rsd.close
			%><center><%
			Response.Write sChkChassisID
			Response.Write "　のシャーシはシャトル作業中です。</p>"	
			%><A HREF="JavaScript:history.back()">
				<BR>シャーシ属性画面へ戻る</A></CENTER> <%
			Response.Write "</body>"
			Response.Write "</html>"
			Response.end
		End If
	End If

	rsd("UpdtTime") = now()				'更新日
	rsd("UpdtPgCd") = "SYWB0032"		'更新プログラム

	rsd("NotDelFlag") = sNotDelFlag		'｢搬出コンテナを載せない｣指定（Y：オン）
	rsd("NightFlag")  = sNightFlag		'｢夕積みのみ載せる｣指定（Y：オン）
	If Request.Form("check3") = "on" Then	'｢グループ変更｣指定
		rsd("GroupID") = s_chg_GrpID	'グループコード
	End If
	rsd("SendFlag") = "Y"				'送信フラグ
	rsd.update
	rsd.close

	'対象シャーシを使用中の搬入を取得
	Call GetChangeApp(conn, rsd, sChkChassisID, iOpeNo, iDualOpeNo)
	If iOpeNo > 0 Then		'作業あり
		'*** 搬入側を処理 ***
		'申請情報の取得（指定作業番号、更新モード）
		Call GetAppInfoOpeNoUpd(conn, rsd, iOpeNo)
		sHH = Trim(rsd("Term"))

		'交換申請情報の更新
		rsd("UpdtTime")  = now()			'更新日
		rsd("UpdtPgCd")  = "SYWB032"		'更新プログラム

		rsd("NotDelFlag") = sNotDelFlag		'｢搬出コンテナを載せない｣指定（Y：オン）
		rsd("NightFlag")  = sNightFlag		'｢夕積みのみ載せる｣指定（Y：オン）

		'｢搬出コンテナを載せない｣指定あるいは'｢夕積みのみ載せる｣指定の場合
		If sNotDelFlag = "Y" Or _
		   sNightFlag  = "Y" Then
			rsd("DualOpeNo") = 0			'デュアル作業番号
		End If
		rsd("SendFlag")  = "Y"				'送信フラグ
		rsd.update
		rsd.close

		'*** 搬出側を処理 ***
		If iDualOpeNo > 0 And _
		   (sNotDelFlag = "Y" Or _
		    sNightFlag  = "Y") Then
			'新規作業順位の取得（指定日、指定時間帯）
			iOpeOrder = GetNewOpeOrder(conn, rsd, sYMD, sHH, "D")

			'申請情報の取得（指定作業番号、更新モード）
			Call GetAppInfoOpeNoUpd(conn, rsd, iDualOpeNo)
			'申請情報の更新
			rsd("UpdtTime")  = now()			'更新日
			rsd("UpdtPgCd")  = "SYWB033"		'更新プログラム
			rsd("ChassisID") = ""				'シャーシID
			rsd("DualOpeNo") = 0				'デュアル作業番号
			rsd("OpeOrder")  = iOpeOrder		'作業順位
			rsd("SendFlag")  = "Y"				'送信フラグ
			rsd.update
			rsd.close
		End If

		If iDualOpeNo > 0 Then	'デュアルの場合
			sWk = "ＤＵＡＬが崩れます。時間枠本数を越える可能性があります。ダイヤ確定時にご注意ください"
		Else
			sWk = "設定しました。空シャーシの過不足を確認してください。"
		End If
		%><CENTER><%=sWk%>
		  <A HREF=SYWB013.asp?TDATE=<%=sYMD%>>
				<BR>一覧画面へ戻る</A></CENTER>
		</body></html><%
		Response.end
	End If

%>
<CENTER>
<B>更新中</B>
</CENTER>
<FORM NAME="SEND">
	<INPUT TYPE=hidden NAME="YMD" VALUE=<%=sYMD%>>
</FORM>
<SCRIPT LANGUAGE="JavaScript">
	location.replace("SYWB013.asp?TDATE=" + document.SEND.YMD.value);
</SCRIPT>

</body>
</html>
 