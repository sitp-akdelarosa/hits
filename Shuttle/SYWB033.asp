<%@ LANGUAGE="VBScript" %>
<%
Option Explicit
%>
<!--#include file="Common.inc"-->
<!--#include file="Common2.inc"-->
<!--#include file="SYWB033.inc"-->
<html>

<head>
<title>シャーシ予約変更登録画面</title>
</head>
<body>

<img border="0" src="image/title01.gif" width="311" height="42">
<br><br>
<center>
<p><img border="0" src="image/title27.gif" width="236" height="34"><p>
</center>

		<font face="ＭＳ ゴシック">
<CENTER>
<%
	Dim conn, rsd, sql									'ＤＢ接続
	Dim sUsrID, sGrpID, sUsrName, sGrpName, sOperator	'ユーザ情報
	Dim sYMD, iOpeNo									'指定日付、作業番号
	Dim sM_ChassisId, sChkChassisID						'元のシャーシＩＤ、変更（交換）シャーシＩＤ
	Dim sHK_flg											'交換フラグ（K：交換）
	Dim sErr_msg										'エラーメッセージ
 	Dim iChg_OpeNo										'交換する作業番号
	Dim sSize20Flag_O, sMixSizeFlag_O, sGroupID_O		'元シャーシ属性
	Dim sSize20Flag_N, sMixSizeFlag_N, sGroupID_N		'新シャーシ属性
	Dim iDualOpe, sHH									'デュアル作業番号、時間帯
	Dim iOpeOrder										'作業順位
	Dim sBlanks,sBlanke									'空白行(エラー時の画面調整用)


	'エラー時の画面調整用
	sBlanks = "<B><U><font color=#ff0000><br><br><br><br>"
	sBlanke = "<br><br><br><br><br><br></font></U></B>"

	'ＤＢ接続
	Call ConnectSvr(conn, rsd)

	'ユーザ情報の取得
	Call GetUserInfo(conn, rsd, sUsrID, sGrpID, sUsrName, sGrpName, sOperator)

	'指定日付取得
	sYMD = TRIM(Request.QueryString("YMD"))

	'作業番号取得
	iOpeNo = TRIM(Request.QueryString("OPENO"))

	'元のシャーシＩＤ取得
	sM_ChassisId = TRIM(Request.QueryString("M_ChassisId"))

	'変更（交換）シャーシＩＤの取得
	If Len(TRIM(Request.QueryString("sCassis"))) = 5 Then
		sChkChassisID = Left(TRIM(Request.QueryString("sCassis")),4)
	Else
		sChkChassisID = Left(TRIM(Request.QueryString("sCassis")),5)
	End If

	'交換フラグ取得
	sHK_flg = Right(TRIM(Request.QueryString("sCassis")),1)

	'元のシャーシＩＤの属性取得
	If sM_ChassisId <> "" Then
		sql = "SELECT Size20Flag, MixSizeFlag, GroupID FROM sChassis" & _
				" WHERE ChassisId = '" & sM_ChassisId & "'"
		rsd.Open sql, conn, 0, 1, 1
		If	rsd.EOF Then		'レコードがない場合
			rsd.close
			sErr_msg = sBlanks & "入力されたシャーシは存在しません。" & sBlanke
			Response.Write sErr_msg
%><center><input type="button" value="　確認　" onClick="JavaScript:history.back()"><center><%
			Response.Write "</body>"
			Response.Write "</html>"
			Response.end
		End If
		sSize20Flag_O  = Trim(rsd("Size20Flag"))		'２０フィートフラグ
		sMixSizeFlag_O = Trim(rsd("MixSizeFlag"))		'２０／４０兼用シャーシ
		sGroupID_O	   = Trim(rsd("GroupID"))			'グループＩＤ
		rsd.close
	End If

	'変更シャーシＩＤの属性取得
	sql = "SELECT Size20Flag, MixSizeFlag, GroupID FROM sChassis" & _
			" WHERE ChassisId = '" & sChkChassisID & "'"
	rsd.Open sql, conn, 0, 1, 1
	If	rsd.EOF Then		'レコードがない場合
		rsd.close
		sErr_msg = sBlanks & "入力されたシャーシは存在しません。" & sBlanke
		Response.Write sErr_msg
		%><center>
		<input type="button" value="　確認　" onClick="JavaScript:history.back()" id=button1 name=button1>
		<center><%
		Response.Write "</body>"
		Response.Write "</html>"
		Response.end
	End If
	sSize20Flag_N  = Trim(rsd("Size20Flag"))		'２０フィートフラグ
	sMixSizeFlag_N = Trim(rsd("MixSizeFlag"))		'２０／４０兼用シャーシ
	sGroupID_N	   = Trim(rsd("GroupID"))			'グループＩＤ
	rsd.close

	'変更（交換）前後属性チェック処理
	If sM_ChassisId <> "" Then	'元シャーシがある場合
		'シャーシサイズの不適合をチェック
		If sSize20Flag_O <> sSize20Flag_N Or _
		   sMixSizeFlag_O <> sMixSizeFlag_N Then
			sErr_msg = sBlanks & "入力されたシャーシは条件にあいません。１" & sBlanke
			Response.Write sErr_msg
			%><center>
			<input type="button" value="　確認　" onClick="JavaScript:history.back()" id=button2 name=button2>
			<center><%
			Response.Write "</body>"
			Response.Write "</html>"
			Response.end
		End if
		'グループのチェックを行う
		If sGroupID_O <> sGroupID_N then
			sErr_msg = sBlanks & sM_ChassisId & "と" & _
						sChkChassisID & "　のシャーシは変更（交換）できません。" & sBlanke
			Response.Write sErr_msg
			%><center>
			<input type="button" value="　確認　" onClick="JavaScript:history.back()" id=button3 name=button3>
			<center><%
			Response.Write "</body>"
			Response.Write "</html>"
			Response.end
		End if
	End If

	'交換先の作業番号を取得する
	if sHK_flg = "K" then	'交換時
		'同一作業日に予約中の搬出申請を取得
		sql = "SELECT distinct OpeNo FROM sAppliInfo"
		sql = sql & " WHERE RTRIM(sAppliInfo.GroupID) = '" & sGrpID & "'"
		sql = sql & "  AND Status   = '02'"
		sql = sql & "  AND RecDel   = 'D'"
		sql = sql & "  AND DelFlag  = ' '"
		sql = sql & "  AND WorkFlag = ' '"
		sql = sql & "  AND LockFlag = ' '"
		sql = sql & "  AND sAppliInfo.WorkDate = '" & cdate(ChgYMDStr(sYMD)) & "'"
		sql = sql & "  AND RTRIM(ChassisID) = '" & sChkChassisID & "'"
		rsd.Open sql, conn, 0, 1, 1
		If rsd.EOF Then	
			rsd.close
			sErr_msg = sBlanks & "交換相手の予約はシャーシ交換不可です。" & sBlanke
			Response.Write sErr_msg	
			%><center>
			<input type="button" value="　確認　" onClick="JavaScript:history.back()" id=button5 name=button5>
			<center><%
			Response.Write "</body>"
			Response.Write "</html>"
			Response.end
		End If
		iChg_OpeNo = rsd("OpeNo")			'交換する作業番号
		rsd.close
	End If

	'変更元申請情報の取得
	Call GetOApp(conn, rsd, iOpeNo, sYMD, sErr_msg)
	If sErr_msg <> "" Then	'エラーのある場合
		rsd.close
		Response.Write sBlanks & sErr_msg & sBlanke
		%><center><input type="button" value="　確認　" onClick="JavaScript:history.back()" id=button6 name=button6>
		<center><%
		Response.Write "</body>"
		Response.Write "</html>"
		Response.end
	End If

	'シャーシサイズの不適合をチェック(申請情報のサイズを確認する)
	If rsd("ContSize") = "20" Then 
'''		If sSize20Flag_N = "Y" Or rsd("MixSizeFlag") = "Y" Then	'2001/06/02
		If sSize20Flag_N = "Y" Or sMixSizeFlag_N = "Y" Then		'2001/06/02
			sErr_msg = ""
		Else
			sErr_msg = sBlanks & "入力されたシャーシは条件にあいません。" & sBlanke
		End If
	Else
		If sSize20Flag_N = "Y" Then
			sErr_msg = sBlanks & "入力されたシャーシは条件にあいません。" & sBlanke
		Else
			sErr_msg = ""
		End If
	End If

	If sErr_msg <> "" then
		Response.Write sErr_msg
		%><center>
		<input type="button" value="　確認　" onClick="JavaScript:history.back()" id=button4 name=button4>
		<center><%
		Response.Write "</body>"
		Response.Write "</html>"
		Response.end
	End if

	iDualOpe  = rsd("DualOpeNo")		'デュアル作業番号
	sHH       = Trim(rsd("Term"))		'時間帯
	iOpeOrder = rsd("OpeOrder")			'作業順位
	rsd.close

	If iDualOpe > 0 Then	'デュアルの場合はデュアル解除する
		'新規作業順位の取得（指定日、指定時間帯）
		iOpeOrder = GetNewOpeOrder(conn, rsd, sYMD, sHH, "D")
	End If

	'申請情報の取得（指定作業番号、更新モード）
	Call GetAppInfoOpeNoUpd(conn, rsd, iOpeNo)

	'変更元申請情報の更新
	rsd("UpdtTime")  = now()			'更新日
	rsd("UpdtPgCd")  = "SYWB033"		'更新プログラム
	rsd("ChassisID") = sChkChassisID	'シャーシID
	rsd("DualOpeNo") = 0				'デュアル作業番号
	rsd("OpeOrder")  = iOpeOrder		'作業順位
	rsd("SendFlag")  = "Y"				'送信フラグ
	rsd.update
	rsd.close

	If iDualOpe > 0 Then	'デュアルの場合はデュアル解除する
		'申請情報の取得（指定作業番号、更新モード）
		Call GetAppInfoOpeNoUpd(conn, rsd, iDualOpe)
		If Not rsd.EOF Then		'本来レコードは必ずある
			'変更元申請情報の更新
			rsd("UpdtTime")  = now()			'更新日
			rsd("UpdtPgCd")  = "SYWB033"		'更新プログラム
			rsd("DualOpeNo") = 0				'デュアル作業番号
			rsd("SendFlag")  = "Y"				'送信フラグ
			rsd.update
		End If
		rsd.close
	End If

	'交換の時
	If sHK_flg = "K" then	'交換時
		'申請情報の取得（指定作業番号、更新モード）
		Call GetAppInfoOpeNoUpd(conn, rsd, iChg_OpeNo)
		If Not rsd.EOF Then		'本来レコードは必ずある
			iDualOpe  = rsd("DualOpeNo")		'デュアル作業番号
			sHH       = Trim(rsd("Term"))		'時間帯
			iOpeOrder = rsd("OpeOrder")			'作業順位
			rsd.close
			If iDualOpe > 0 Then	'デュアルの場合はデュアル解除する
				'新規作業順位の取得（指定日、指定時間帯）
				iOpeOrder = GetNewOpeOrder(conn, rsd, sYMD, sHH, "D")
			End If

			'申請情報の取得（指定作業番号、更新モード）
			Call GetAppInfoOpeNoUpd(conn, rsd, iChg_OpeNo)

			'交換申請情報の更新
			rsd("UpdtTime")  = now()			'更新日
			rsd("UpdtPgCd")  = "SYWB033"		'更新プログラム
			rsd("ChassisID") = sM_ChassisId		'シャーシID
			rsd("DualOpeNo") = 0				'デュアル作業番号
			rsd("OpeOrder")  = iOpeOrder		'作業順位
			rsd("SendFlag")  = "Y"				'送信フラグ
			rsd.update
			rsd.close
			If iDualOpe > 0 Then	'デュアルの場合はデュアル解除する
				'申請情報の取得（指定作業番号、更新モード）
				Call GetAppInfoOpeNoUpd(conn, rsd, iDualOpe)
				If Not rsd.EOF Then		'本来レコードは必ずある
					'変更元申請情報の更新
					rsd("UpdtTime")  = now()			'更新日
					rsd("UpdtPgCd")  = "SYWB033"		'更新プログラム
					rsd("DualOpeNo") = 0				'デュアル作業番号
					rsd("SendFlag")  = "Y"				'送信フラグ
					rsd.update
				End If
				rsd.close
			End If
		Else
			rsd.close
		End if
	End if
%>
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
