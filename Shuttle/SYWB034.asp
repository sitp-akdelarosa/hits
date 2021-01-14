<%@ LANGUAGE="VBScript" %>
<%
Option Explicit
%>
<!--#include file="Common.inc"-->
<!--#include file="Common2.inc"-->
<!--#include file="SYWB017.inc"-->
<html>

<head>
<title>利用回数モニタ詳細</title>
</head>
<body>
<%
	Dim sYMD, sChassisID, sDispChassis1, sDispChassis2  
	Dim conn, rsd, sql
	Dim sUsrID, sGrpID, sUsrName, sGrpName, sOperator,sMonthStart
	Dim sNMonth, sBMonth1, sBMonth2, sBMonth3
	Dim sDisp_Date, sDisp_Date1, sDisp_Date2, sDisp_Date3,sDisp_Date4
	Dim sGroupName, sTrgDate, sStartDate, sEndDate
	Dim dCntDate, sWeek, sAmPm, dOldCntDate
	Dim iRDCount, iDelCount, iRecCount, iVPCount, iRVCount,  iUse, iUse_sum

	'ＤＢ接続
	Call ConnectSvr(conn, rsd)

	'ユーザ情報の取得
	Call GetUserInfo(conn, rsd, sUsrID, sGrpID, sUsrName, sGrpName, sOperator)

	'ユーザ名の取得
	sql = "SELECT GroupID,GroupName FROM sMGroup" & _
		  " WHERE RTRIM(GroupID) = '" & sGrpID & "'"
			rsd.Open sql, conn, 0, 1, 1
	if not rsd.EOF then
		sGroupName = rsd("GroupName")
	end if
	rsd.Close

	sGroupName = sGroupName & "　殿"

	'月度開始日の取得
	sMonthStart= GetEnv(conn, rsd, "MonthStart")

	'指定日付取得
	sTrgDate = TRIM(Request.QueryString("TDATE"))

	'過去３ヶ月の年月取得
	Call GetBefore3Month(date(), trim(sMonthStart), sNMonth, sBMonth1, sBMonth2, sBMonth3)
		
	sDisp_Date1 = left(sNMonth,4) & "年" & mid(sNMonth,5) & "月"
	sDisp_Date2 = left(sBMonth1,4) & "年" & mid(sBMonth1,5) & "月"
	sDisp_Date3 = left(sBMonth2,4) & "年" & mid(sBMonth2,5) & "月"
	sDisp_Date4 = left(sBMonth3,4) & "年" & mid(sBMonth3,5) & "月"

	select case	Trim(Request.Form("SELECT1"))
		case sNMonth
			sNMonth = "selected value=" & sNMonth
			sBMonth1 = "value=" & sBMonth1
			sBMonth2 = "value=" & sBMonth2
			sBMonth3 = "value=" & sBMonth3
			sDisp_Date = sDisp_Date1
		case sBMonth1
			sNMonth = "value=" & sNMonth
			sBMonth1 = "selected value=" & sBMonth1
			sBMonth2 = "value=" & sBMonth2
			sBMonth3 = "value=" & sBMonth3
			sDisp_Date = sDisp_Date2
		case sBMonth2
			sNMonth = "value=" & sNMonth
			sBMonth1 = "value=" & sBMonth1
			sBMonth2 = "selected value=" & sBMonth2
			sBMonth3 = "value=" & sBMonth3
			sDisp_Date = sDisp_Date3
		case sBMonth3
			sNMonth = "value=" & sNMonth
			sBMonth1 = "value=" & sBMonth1
			sBMonth2 = "value=" & sBMonth2
			sBMonth3 = "selected value=" & sBMonth3
			sDisp_Date = sDisp_Date4
		case else
			sNMonth = "value=" & sNMonth
			sBMonth1 = "value=" & sBMonth1
			sBMonth2 = "value=" & sBMonth2
			sBMonth3 = "value=" & sBMonth3
			sDisp_Date = ""
	end select
%>
<img border="0" src="image/title01.gif" width="311" height="42">
<br>
<center>
<p><img border="0" src="image/title31.gif" width="236" height="34"><p>
<b><u><font size=3><%=sGroupName %></font></u></b><br>

<FORM ACTION="SYWB034.asp?TDATE=<%=sTrgDate%>" METHOD="post">
<b><font size=3>年月選択（過去３ヶ月）</font></b>
<SELECT NAME="SELECT1">
<OPTION VALUE="No" >
<OPTION <%=sNMonth%>><%=sDisp_Date1%>
<OPTION <%=sBMonth1%>><%=sDisp_Date2%>
<OPTION <%=sBMonth2%>><%=sDisp_Date3%>
<OPTION <%=sBMonth3%>><%=sDisp_Date4%>

</select>
<input type="submit" value="照    会" id=submit4>
</form>
<%
'未入力チェック
if Request.Form("SELECT1") = "No"  then
	Response.Write "<br><p><b>年月を選択してください。</b></p><br>"
	%><form  METHOD="post"  NAME="UPLOAD1" ACTION="SYWB013.asp?TDATE=<%=sTrgDate%>">
		<input type="submit" value="戻    る"id=submit4 name=submit4>
	</form><%
	Response.Write "</body>"
	Response.Write "</html>"
	Response.End
end if 

'該当データチェック

'開始・終了日の取得(入力開始日付より
sStartDate = ""	
sEndDate = ""	
Call GetStartEnd(conn, rsd, sGrpID, Trim(Request.Form("SELECT1")), trim(sMonthStart), sStartDate, sEndDate)

if rsd.EOF then
	rsd.Close
	Response.Write "<br><p><b>該当データがありません。</b></p><br>"
	%><form  METHOD="post"  NAME="UPLOAD1" ACTION="SYWB013.asp?TDATE=<%=sTrgDate%>">
		<input type="submit" value="戻    る"id=submit4 name=submit4>
	</form><%
	Response.Write "</body>"
	Response.Write "</html>"
	Response.End
end if

%>
</center>
<center>
<table border="1" width="800"  >   
	<tr>
		<th width="90" bgcolor="#7fffd4" align=center><%=sDisp_Date%></th>
	    <th bgcolor="#7fffd4" align=center>曜日</th>			
	    <th bgcolor="#7fffd4" align=center>午前<br>午後</th>			
	    <th width="90" bgcolor="#7fffd4" align=center>デュアル<br>(搬入搬出)</th>			
	    <th width="90" bgcolor="#7fffd4" align=center>デュアル<br>(搬入空ﾊﾞﾝ)</th>			
	    <th width="90" bgcolor="#7fffd4" align=center>搬出のみ</th>			
	    <th width="90" bgcolor="#7fffd4" align=center>搬入のみ<br>(含前受け)</th>			
	    <th width="90" bgcolor="#7fffd4" align=center>空バン</th>			
	    <th width="90" bgcolor="#7fffd4" align=center>利用回数</th>			
	</tr>
<%
'計算エリア
	iRDCount	=	0
	iDelCount	=	0
	iRecCount	=	0
	iVPCount	=	0	'VP対応
	iRVCount	=	0	'VP対応
	iUse		=	0
	iUse_sum	=	0
'開始日付セット
	dCntDate    =	sStartDate 
	dOldCntDate =	sStartDate
	sAmPm = "A"
'画面表示
	Do Until dCntDate > sEndDate
		Do Until rsd.EOF
			sWeek = sWeekday(Weekday(cDate(ChgYMDStr(dCntDate))))		'曜日の取得
			If sAmPm = "P" And rsd("RecDelDate") <> dOldCntDate  then
%>				<tr>
				    <td bgcolor=#fff0f5 align=center>午後</td>			
					<td bgcolor=#fff0f5 align=center>0</td>			
					<td bgcolor=#fff0f5 align=center>0</td>			<!--VP対応 -->
					<td bgcolor=#fff0f5 align=center>0</td>			
					<td bgcolor=#fff0f5 align=center>0</td>			
					<td bgcolor=#fff0f5 align=center>0</td>			<!--VP対応 -->
					<td bgcolor=#fff0f5 align=center>0</td>			
				</tr>
<%				sAmPm = "A"
			End IF
      		If rsd("RecDelDate") = dCntDate	then '件数日付と等しい

				iRDCount	=	Int(iRDCount)	+	Int(rsd("RDCount"))
				iDelCount	=	Int(iDelCount)	+	Int(rsd("DelCount"))
				iRecCount	=	Int(iRecCount)	+	Int(rsd("RecCount"))
				iVPCount	=	Int(iVPCount)	+	Int(rsd("VPCount"))	'VP対応
				iRVCount	=	Int(iRVCount)	+	Int(rsd("RVCount"))	'VP対応
				iUse		=	Int(rsd("RDCount")) * 2 + Int(rsd("DelCount")) +  _
				                	Int(rsd("RecCount")) + Int(rsd("VPCount"))     +  _
							Int(rsd("RVCount")) * 2
				iUse_sum	=	Int(iUse_sum)		+	iUse
	'
				if rsd("AmPm") = "A" then
%>						<tr>
					<td bgcolor=#AFEEEE align=center ROWSPAN=2><%=day(ChgYMDStr(dCntDate))%></td>
					<td bgcolor=#AFEEEE align=center ROWSPAN=2><%=sWeek%></td>
					<td bgcolor=#FFFFE0 align=center>午前</td>
					<td bgcolor=#FFFFE0 align=center><%=rsd("RDCount")%></td>
					<td bgcolor=#FFFFE0 align=center><%=rsd("RVCount")%></td>		<!--VP対応 -->
					<td bgcolor=#FFFFE0 align=center><%=rsd("DelCount")%></td>
					<td bgcolor=#FFFFE0 align=center><%=rsd("RecCount")%></td>
					<td bgcolor=#FFFFE0 align=center><%=rsd("VPCount")%></td>		<!--VP対応 -->
					<td bgcolor=#FFFFE0 align=center><%=iUse%></td>
					</tr>
<%					sAmPm = "P"
				else
					If sAmPm = "A" then			'午後のみの場合
%>
						<tr>
						    <td bgcolor=#AFEEEE align=center ROWSPAN=2><%=day(ChgYMDStr(dCntDate))%></td>
						    <td bgcolor=#AFEEEE align=center ROWSPAN=2><%=sWeek%></td>			
						    <td bgcolor=#FFFFE0 align=center>午前</td>
						    <td bgcolor=#FFFFE0 align=center>0</td>
						    <td bgcolor=#FFFFE0 align=center>0</td>		<!--VP対応 -->
						    <td bgcolor=#FFFFE0 align=center>0</td>
						    <td bgcolor=#FFFFE0 align=center>0</td>
						    <td bgcolor=#FFFFE0 align=center>0</td>		<!--VP対応 -->
						    <td bgcolor=#FFFFE0 align=center>0</td>
						</tr>

						<tr>
						    <td bgcolor=#fff0f5 align=center>午後</td>
						    <td bgcolor=#fff0f5 align=center><%=rsd("RDCount")%></td>
						    <td bgcolor=#fff0f5 align=center><%=rsd("RVCount")%></td>	<!--VP対応 -->
						    <td bgcolor=#fff0f5 align=center><%=rsd("DelCount")%></td>
						    <td bgcolor=#fff0f5 align=center><%=rsd("RecCount")%></td>
						    <td bgcolor=#fff0f5 align=center><%=rsd("VPCount")%></td>	<!--VP対応 -->
						    <td bgcolor=#fff0f5 align=center><%=iUse%></td>
						</tr>
<%					Else
%>						<tr>
						    <td bgcolor=#fff0f5 align=center>午後</td>
						    <td bgcolor=#fff0f5 align=center><%=rsd("RDCount")%></td>
						    <td bgcolor=#fff0f5 align=center><%=rsd("RVCount")%></td>	<!--VP対応 -->
						    <td bgcolor=#fff0f5 align=center><%=rsd("DelCount")%></td>
						    <td bgcolor=#fff0f5 align=center><%=rsd("RecCount")%></td>
						    <td bgcolor=#fff0f5 align=center><%=rsd("VPCount")%></td>	<!--VP対応 -->
						    <td bgcolor=#fff0f5 align=center><%=iUse%></td>			
						</tr>
<%				
					End if
					sAmPm = "A"
					dCntDate = GetYMDStr(ChgYMDDate(dCntDate) + 1)	'日付＋１
				End If
				rsd.MoveNext
			else
%>
				<tr>
				    <td bgcolor=#AFEEEE align=center ROWSPAN=2><%=day(ChgYMDStr(dCntDate))%></td>
				    <td bgcolor=#AFEEEE align=center ROWSPAN=2><%=sWeek%></td>
				    <td bgcolor=#FFFFE0 align=center>午前</td>
				    <td bgcolor=#FFFFE0 align=center>0</td>
				    <td bgcolor=#FFFFE0 align=center>0</td>		<!--VP対応 -->
				    <td bgcolor=#FFFFE0 align=center>0</td>
				    <td bgcolor=#FFFFE0 align=center>0</td>
				    <td bgcolor=#FFFFE0 align=center>0</td>		<!--VP対応 -->
				    <td bgcolor=#FFFFE0 align=center>0</td>
				</tr>
				<tr>
				    <td bgcolor=#fff0f5 align=center>午後</td>
				    <td bgcolor=#fff0f5 align=center>0</td>
				    <td bgcolor=#fff0f5 align=center>0</td>		<!--VP対応 -->
				    <td bgcolor=#fff0f5 align=center>0</td>
				    <td bgcolor=#fff0f5 align=center>0</td>
				    <td bgcolor=#fff0f5 align=center>0</td>		<!--VP対応 -->
				    <td bgcolor=#fff0f5 align=center>0</td>			
				</tr>
<%
				sAmPm = "A"
				dCntDate = GetYMDStr(ChgYMDDate(dCntDate) + 1)	'日付＋１
			End if 
			dOldCntDate = dCntDate		'現在のレコードの日付を保存する
		Loop
		rsd.close
		Exit Do
	Loop

'残りがあれば０を書く
	If sAmPm = "P" then '残りの午後データがある場合書く
%>		<tr>
		    <td bgcolor=#fff0f5 align=center>午後</td>
		    <td bgcolor=#fff0f5 align=center>0</td>
		    <td bgcolor=#fff0f5 align=center>0</td>		<!--VP対応 -->
		    <td bgcolor=#fff0f5 align=center>0</td>
		    <td bgcolor=#fff0f5 align=center>0</td>
		    <td bgcolor=#fff0f5 align=center>0</td>		<!--VP対応 -->
		    <td bgcolor=#fff0f5 align=center>0</td>			
		</tr>
<%		dCntDate = GetYMDStr(ChgYMDDate(dCntDate) + 1)	'日付＋１
	End If
	
	Do Until dCntDate > sEndDate
		sWeek = sWeekday(Weekday(cDate(ChgYMDStr(dCntDate))))		'曜日の取得
%>			<tr>
			    <td bgcolor=#AFEEEE align=center ROWSPAN=2><%=day(ChgYMDStr(dCntDate))%></td>
			    <td bgcolor=#AFEEEE align=center ROWSPAN=2><%=sWeek%></td>
			    <td bgcolor=#FFFFE0 align=center>午前</td>
			    <td bgcolor=#FFFFE0 align=center>0</td>
			    <td bgcolor=#FFFFE0 align=center>0</td>		<!--VP対応 -->
			    <td bgcolor=#FFFFE0 align=center>0</td>
			    <td bgcolor=#FFFFE0 align=center>0</td>
			    <td bgcolor=#FFFFE0 align=center>0</td>		<!--VP対応 -->
			    <td bgcolor=#FFFFE0 align=center>0</td>			
			</tr>
			<tr>
			    <td bgcolor=#fff0f5 align=center>午後</td>
			    <td bgcolor=#fff0f5 align=center>0</td>
			    <td bgcolor=#fff0f5 align=center>0</td>		<!--VP対応 -->
			    <td bgcolor=#fff0f5 align=center>0</td>
			    <td bgcolor=#fff0f5 align=center>0</td>
			    <td bgcolor=#fff0f5 align=center>0</td>		<!--VP対応 -->
			    <td bgcolor=#fff0f5 align=center>0</td>			
			</tr>
<%		dCntDate = GetYMDStr(ChgYMDDate(dCntDate) + 1)
	Loop%>
				<tr>
				    <td bgcolor=#b0c4de align=center>合計</td>
				    <td bgcolor=#b0c4de align=center><br><br></td>
				    <td bgcolor=#b0c4de align=center><br><br></td>
				    <td bgcolor=#b0c4de align=center><%=iRDCount%></td>
				    <td bgcolor=#b0c4de align=center><%=iRVCount%></td>		<!--VP対応 -->
				    <td bgcolor=#b0c4de align=center><%=iDelCount%></td>
				    <td bgcolor=#b0c4de align=center><%=iRecCount%></td>
				    <td bgcolor=#b0c4de align=center><%=iVPCount%></td>		<!--VP対応 -->
				    <td bgcolor=#b0c4de align=center><%=iUse_sum%></td>			
				</tr>
		</table>
		</center><br>
	　　　　　　　　　　　　　　　　　（＊）利用回数はデュアルを２回としてカウントします。
		<center>
	    <form  METHOD="post"  NAME="UPLOAD1" ACTION="SYWB013.asp?TDATE=<%=sTrgDate%>">
			<input type="submit" value="戻    る"id=submit4 name=submit4>
		</form>
		</center>

</body>     
</html>     
