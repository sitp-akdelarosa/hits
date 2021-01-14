<%@ LANGUAGE="VBScript" %>
<%
Option Explicit
%>
<!--#include file="Common.inc"-->
<!--#include file="Common2.inc"-->
<html>

<head>
<title>シャーシ詳細画面</title>
</head>

<body >
<%
	Dim sYMD, sChassisID, sDispChassis1, sDispChassis2  
	Dim conn, rsd, sql
	Dim sUsrID, sGrpID, sUsrName, sGrpName, sOperator
	Dim i, j, sNO, sChk1, sChk2, sChkChassisID
	Dim sSize, sPlateNo, sUserName, sPlace, sZokusei, sGrpNm, sContNo
	Dim sGrp_Mei(100), sWorkDate(100), sWorkTime(100), sRecDel(100), sCont(100), sOpeNo(100)

	'ＤＢ接続
	Call ConnectSvr(conn, rsd)

	'ユーザ情報の取得
	Call GetUserInfo(conn, rsd, sUsrID, sGrpID, sUsrName, sGrpName, sOperator)

	'シャーシＩＤ取得
	sChkChassisID = TRIM(Request.QueryString("sCassis"))

	'シャーシID取得

	'シャーシーIDを選択した場合
	sql = "SELECT sChassis.*,sMGroup.GroupName FROM sChassis,sMGroup" & _
	  " WHERE RTRIM(sChassis.ChassisId) = '" & sChkChassisID & "'" & _
	  "   AND RTRIM(sChassis.GroupID) = RTRIM(sMGroup.GroupID)"
	rsd.Open sql, conn, 0, 1, 1

	If Not rsd.EOF Then
		if rsd("Size20Flag") = "Y" then	
			sSize = "20"
		else
			If rsd("MixSizeFlag") = "Y" then	
				sSize = "20/40兼用"
			Else
				sSize = "40"
			End If
		end if
		sPlateNo = rsd("PlateNo")
		sUserName = rsd("UserName")
		sGrpNm = rsd("GroupName")
		if rsd("StackFlag") <> " " then
			sPlace = "SY"
		else
			sPlace = ""
		end if
			
		if rsd("NightFlag") = "Y" then
			sZokusei = "夕積のみ載せる"
		end if

		if rsd("NotDelFlag") = "Y" then
			sZokusei = "搬出コンテナを載せない"
		end if

	end if
	rsd.Close

%>
<img border="0" src="image/title01.gif" width="311" height="42">
<br><br>
<center>
<p><img border="0" src="image/title28.gif" width="236" height="34"><p>
</center>

<font face="ＭＳ ゴシック">
   
<center>
<%dim sdate
sdate = month(date) & "月" & day(date) & "日" & "　" & hour(time) & "時" & minute(time) & "分現在"
'Response.Write sdate%>
<u><%=sdate%></u><br><br>

<table border="1" width="500"  >
<b><font color=#000080>対象</font></b>　　　
	<tr bgcolor=#ffff99><td>
				シャーシ　　　　　　　<%=sChkChassisID%><br>
				サイズ　　　　　　　　<%=sSize%><br>
				ナンバープレート　　　<%=sPlateNo%><br>
				所有者　　　　　　　　<%=sUserName%><br>
	</td></tr>
</table><br>

<table border="1" width="500">
<b><font color=#000080>現在の状態</font></b>
	<tr bgcolor=#ccffcc><td>
				所属グループ　　　　　<%=sGrpNm%><br>
				場所　　　　　　　　　<%=sPlace%><br>
				属性　　　　　　　　　<%=sZokusei%><br>
<%'申請情報読み込み

	sContNo = ""
	sql = "SELECT ContNo FROM sAppliInfo" & _
	  " WHERE RTRIM(ChassisId) = '" & sChkChassisID & "'" & _
	  "   AND ( RTRIM(Place) = 'SY' or RTRIM(Place) = 'MV' )"
	rsd.Open sql, conn, 0, 1, 1

	If Not rsd.EOF Then
		sContNo = rsd("ContNo")
	end if
	rsd.Close
%>
				搭載コンテナ　　　　　<%=sContNo%>
	</td></tr>
</table><br>

<%'申請情報読み込み
	
	'シャーシーIDを選択した場合
	sql = "SELECT OpeNo,WorkDate,Term,RecDel,ContNo,GroupName,DelFlag FROM sAppliInfo,sMGroup" & _
	  " WHERE RTRIM(sAppliInfo.ChassisId) = '" & sChkChassisID & "'" & _
	  "   AND RTRIM(sAppliInfo.GroupID) = RTRIM(sMGroup.GroupID) " & _
	  "   AND DelFlag = ' '  ORDER BY WorkDate"
	rsd.Open sql, conn, 0, 1, 1

	i = 1
	If Not rsd.EOF Then
				
		Do until rsd.EOF
			
			sGrp_Mei(int(i)) = rsd("GroupName")	'グループ名
			'日にち
			sWorkDate(int(i)) = month(rsd("WorkDate")) & "月" 
			sWorkDate(int(i)) = sWorkDate(int(i)) & day(rsd("WorkDate")) & "日"
			'作業時間
			sWorkTime(int(i)) = trim(rsd("Term"))
			'作業番号
			If len(trim(rsd("OpeNo"))) = 4 Then
				sOpeNo(int(i))    = "0" & trim(rsd("OpeNo"))
			Else
				sOpeNo(int(i))    = trim(rsd("OpeNo"))
			End IF
			'作業種類(VP対応)
			if rsd("RecDel") = "R" then
				sRecDel(int(i)) = "搬入"
			Elseif rsd("RecDel") = "D" then
				sRecDel(int(i)) = "搬出"
			else
				sRecDel(int(i)) = "空バン"
			end if	
			'コンテナ
			if trim(rsd("ContNo")) = "" then
				sCont(int(i)) = "　"
			else
				sCont(int(i)) = trim(rsd("ContNo"))
			end if
				
			i = int(i) + 1
			rsd.movenext
		Loop
		rsd.Close

		for j = 1 to (int(i) - 1)
			if int(j) = 1 then
%>
				<table border="1" width="600"  >   
				<b><font color=#000080>リンク作業</font></b>
					<tr>
						<td bgcolor="#e8ffe8" align=center>グループ</td>
					    <td bgcolor="#e8ffe8" align=center>日にち</td>			
					    <td bgcolor="#e8ffe8" align=center>作業時間</td>			
					    <td bgcolor="#e8ffe8" align=center>予約番号</td>			
					    <td bgcolor="#e8ffe8" align=center>作業種類</td>			
					    <td bgcolor="#e8ffe8" align=center>コンテナ</td>			
					</tr>
<%
			end if%>
			<tr>
				<td align=center><%=sGrp_Mei(int(j))%></td>
			    <td align=center><%=sWorkDate(int(j))%></td>			
			    <td align=center><%=GetTimeSlotStr(conn,rsd,sWorkTime(int(j)))%></td>			
			    <td align=center><%=sOpeNo(int(j))%></td>			
			    <td align=center><%=sRecDel(int(j))%></td>			
			    <td align=center><%=sCont(int(j))%></td>			
			</tr>
<%
		next%>
</table>
<%	else
		rsd.Close
					%>該当の作業なし<%	
	end if%>

</center><br>
<center>
    <form>
    <input type="button" value="　戻る　" onclick="history.back()" >
	</form>
</center>
</body>     
</html>