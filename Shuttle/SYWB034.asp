<%@ LANGUAGE="VBScript" %>
<%
Option Explicit
%>
<!--#include file="Common.inc"-->
<!--#include file="Common2.inc"-->
<!--#include file="SYWB017.inc"-->
<html>

<head>
<title>���p�񐔃��j�^�ڍ�</title>
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

	'�c�a�ڑ�
	Call ConnectSvr(conn, rsd)

	'���[�U���̎擾
	Call GetUserInfo(conn, rsd, sUsrID, sGrpID, sUsrName, sGrpName, sOperator)

	'���[�U���̎擾
	sql = "SELECT GroupID,GroupName FROM sMGroup" & _
		  " WHERE RTRIM(GroupID) = '" & sGrpID & "'"
			rsd.Open sql, conn, 0, 1, 1
	if not rsd.EOF then
		sGroupName = rsd("GroupName")
	end if
	rsd.Close

	sGroupName = sGroupName & "�@�a"

	'���x�J�n���̎擾
	sMonthStart= GetEnv(conn, rsd, "MonthStart")

	'�w����t�擾
	sTrgDate = TRIM(Request.QueryString("TDATE"))

	'�ߋ��R�����̔N���擾
	Call GetBefore3Month(date(), trim(sMonthStart), sNMonth, sBMonth1, sBMonth2, sBMonth3)
		
	sDisp_Date1 = left(sNMonth,4) & "�N" & mid(sNMonth,5) & "��"
	sDisp_Date2 = left(sBMonth1,4) & "�N" & mid(sBMonth1,5) & "��"
	sDisp_Date3 = left(sBMonth2,4) & "�N" & mid(sBMonth2,5) & "��"
	sDisp_Date4 = left(sBMonth3,4) & "�N" & mid(sBMonth3,5) & "��"

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
<b><font size=3>�N���I���i�ߋ��R�����j</font></b>
<SELECT NAME="SELECT1">
<OPTION VALUE="No" >
<OPTION <%=sNMonth%>><%=sDisp_Date1%>
<OPTION <%=sBMonth1%>><%=sDisp_Date2%>
<OPTION <%=sBMonth2%>><%=sDisp_Date3%>
<OPTION <%=sBMonth3%>><%=sDisp_Date4%>

</select>
<input type="submit" value="��    ��" id=submit4>
</form>
<%
'�����̓`�F�b�N
if Request.Form("SELECT1") = "No"  then
	Response.Write "<br><p><b>�N����I�����Ă��������B</b></p><br>"
	%><form  METHOD="post"  NAME="UPLOAD1" ACTION="SYWB013.asp?TDATE=<%=sTrgDate%>">
		<input type="submit" value="��    ��"id=submit4 name=submit4>
	</form><%
	Response.Write "</body>"
	Response.Write "</html>"
	Response.End
end if 

'�Y���f�[�^�`�F�b�N

'�J�n�E�I�����̎擾(���͊J�n���t���
sStartDate = ""	
sEndDate = ""	
Call GetStartEnd(conn, rsd, sGrpID, Trim(Request.Form("SELECT1")), trim(sMonthStart), sStartDate, sEndDate)

if rsd.EOF then
	rsd.Close
	Response.Write "<br><p><b>�Y���f�[�^������܂���B</b></p><br>"
	%><form  METHOD="post"  NAME="UPLOAD1" ACTION="SYWB013.asp?TDATE=<%=sTrgDate%>">
		<input type="submit" value="��    ��"id=submit4 name=submit4>
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
	    <th bgcolor="#7fffd4" align=center>�j��</th>			
	    <th bgcolor="#7fffd4" align=center>�ߑO<br>�ߌ�</th>			
	    <th width="90" bgcolor="#7fffd4" align=center>�f���A��<br>(�������o)</th>			
	    <th width="90" bgcolor="#7fffd4" align=center>�f���A��<br>(���������)</th>			
	    <th width="90" bgcolor="#7fffd4" align=center>���o�̂�</th>			
	    <th width="90" bgcolor="#7fffd4" align=center>�����̂�<br>(�ܑO��)</th>			
	    <th width="90" bgcolor="#7fffd4" align=center>��o��</th>			
	    <th width="90" bgcolor="#7fffd4" align=center>���p��</th>			
	</tr>
<%
'�v�Z�G���A
	iRDCount	=	0
	iDelCount	=	0
	iRecCount	=	0
	iVPCount	=	0	'VP�Ή�
	iRVCount	=	0	'VP�Ή�
	iUse		=	0
	iUse_sum	=	0
'�J�n���t�Z�b�g
	dCntDate    =	sStartDate 
	dOldCntDate =	sStartDate
	sAmPm = "A"
'��ʕ\��
	Do Until dCntDate > sEndDate
		Do Until rsd.EOF
			sWeek = sWeekday(Weekday(cDate(ChgYMDStr(dCntDate))))		'�j���̎擾
			If sAmPm = "P" And rsd("RecDelDate") <> dOldCntDate  then
%>				<tr>
				    <td bgcolor=#fff0f5 align=center>�ߌ�</td>			
					<td bgcolor=#fff0f5 align=center>0</td>			
					<td bgcolor=#fff0f5 align=center>0</td>			<!--VP�Ή� -->
					<td bgcolor=#fff0f5 align=center>0</td>			
					<td bgcolor=#fff0f5 align=center>0</td>			
					<td bgcolor=#fff0f5 align=center>0</td>			<!--VP�Ή� -->
					<td bgcolor=#fff0f5 align=center>0</td>			
				</tr>
<%				sAmPm = "A"
			End IF
      		If rsd("RecDelDate") = dCntDate	then '�������t�Ɠ�����

				iRDCount	=	Int(iRDCount)	+	Int(rsd("RDCount"))
				iDelCount	=	Int(iDelCount)	+	Int(rsd("DelCount"))
				iRecCount	=	Int(iRecCount)	+	Int(rsd("RecCount"))
				iVPCount	=	Int(iVPCount)	+	Int(rsd("VPCount"))	'VP�Ή�
				iRVCount	=	Int(iRVCount)	+	Int(rsd("RVCount"))	'VP�Ή�
				iUse		=	Int(rsd("RDCount")) * 2 + Int(rsd("DelCount")) +  _
				                	Int(rsd("RecCount")) + Int(rsd("VPCount"))     +  _
							Int(rsd("RVCount")) * 2
				iUse_sum	=	Int(iUse_sum)		+	iUse
	'
				if rsd("AmPm") = "A" then
%>						<tr>
					<td bgcolor=#AFEEEE align=center ROWSPAN=2><%=day(ChgYMDStr(dCntDate))%></td>
					<td bgcolor=#AFEEEE align=center ROWSPAN=2><%=sWeek%></td>
					<td bgcolor=#FFFFE0 align=center>�ߑO</td>
					<td bgcolor=#FFFFE0 align=center><%=rsd("RDCount")%></td>
					<td bgcolor=#FFFFE0 align=center><%=rsd("RVCount")%></td>		<!--VP�Ή� -->
					<td bgcolor=#FFFFE0 align=center><%=rsd("DelCount")%></td>
					<td bgcolor=#FFFFE0 align=center><%=rsd("RecCount")%></td>
					<td bgcolor=#FFFFE0 align=center><%=rsd("VPCount")%></td>		<!--VP�Ή� -->
					<td bgcolor=#FFFFE0 align=center><%=iUse%></td>
					</tr>
<%					sAmPm = "P"
				else
					If sAmPm = "A" then			'�ߌ�݂̂̏ꍇ
%>
						<tr>
						    <td bgcolor=#AFEEEE align=center ROWSPAN=2><%=day(ChgYMDStr(dCntDate))%></td>
						    <td bgcolor=#AFEEEE align=center ROWSPAN=2><%=sWeek%></td>			
						    <td bgcolor=#FFFFE0 align=center>�ߑO</td>
						    <td bgcolor=#FFFFE0 align=center>0</td>
						    <td bgcolor=#FFFFE0 align=center>0</td>		<!--VP�Ή� -->
						    <td bgcolor=#FFFFE0 align=center>0</td>
						    <td bgcolor=#FFFFE0 align=center>0</td>
						    <td bgcolor=#FFFFE0 align=center>0</td>		<!--VP�Ή� -->
						    <td bgcolor=#FFFFE0 align=center>0</td>
						</tr>

						<tr>
						    <td bgcolor=#fff0f5 align=center>�ߌ�</td>
						    <td bgcolor=#fff0f5 align=center><%=rsd("RDCount")%></td>
						    <td bgcolor=#fff0f5 align=center><%=rsd("RVCount")%></td>	<!--VP�Ή� -->
						    <td bgcolor=#fff0f5 align=center><%=rsd("DelCount")%></td>
						    <td bgcolor=#fff0f5 align=center><%=rsd("RecCount")%></td>
						    <td bgcolor=#fff0f5 align=center><%=rsd("VPCount")%></td>	<!--VP�Ή� -->
						    <td bgcolor=#fff0f5 align=center><%=iUse%></td>
						</tr>
<%					Else
%>						<tr>
						    <td bgcolor=#fff0f5 align=center>�ߌ�</td>
						    <td bgcolor=#fff0f5 align=center><%=rsd("RDCount")%></td>
						    <td bgcolor=#fff0f5 align=center><%=rsd("RVCount")%></td>	<!--VP�Ή� -->
						    <td bgcolor=#fff0f5 align=center><%=rsd("DelCount")%></td>
						    <td bgcolor=#fff0f5 align=center><%=rsd("RecCount")%></td>
						    <td bgcolor=#fff0f5 align=center><%=rsd("VPCount")%></td>	<!--VP�Ή� -->
						    <td bgcolor=#fff0f5 align=center><%=iUse%></td>			
						</tr>
<%				
					End if
					sAmPm = "A"
					dCntDate = GetYMDStr(ChgYMDDate(dCntDate) + 1)	'���t�{�P
				End If
				rsd.MoveNext
			else
%>
				<tr>
				    <td bgcolor=#AFEEEE align=center ROWSPAN=2><%=day(ChgYMDStr(dCntDate))%></td>
				    <td bgcolor=#AFEEEE align=center ROWSPAN=2><%=sWeek%></td>
				    <td bgcolor=#FFFFE0 align=center>�ߑO</td>
				    <td bgcolor=#FFFFE0 align=center>0</td>
				    <td bgcolor=#FFFFE0 align=center>0</td>		<!--VP�Ή� -->
				    <td bgcolor=#FFFFE0 align=center>0</td>
				    <td bgcolor=#FFFFE0 align=center>0</td>
				    <td bgcolor=#FFFFE0 align=center>0</td>		<!--VP�Ή� -->
				    <td bgcolor=#FFFFE0 align=center>0</td>
				</tr>
				<tr>
				    <td bgcolor=#fff0f5 align=center>�ߌ�</td>
				    <td bgcolor=#fff0f5 align=center>0</td>
				    <td bgcolor=#fff0f5 align=center>0</td>		<!--VP�Ή� -->
				    <td bgcolor=#fff0f5 align=center>0</td>
				    <td bgcolor=#fff0f5 align=center>0</td>
				    <td bgcolor=#fff0f5 align=center>0</td>		<!--VP�Ή� -->
				    <td bgcolor=#fff0f5 align=center>0</td>			
				</tr>
<%
				sAmPm = "A"
				dCntDate = GetYMDStr(ChgYMDDate(dCntDate) + 1)	'���t�{�P
			End if 
			dOldCntDate = dCntDate		'���݂̃��R�[�h�̓��t��ۑ�����
		Loop
		rsd.close
		Exit Do
	Loop

'�c�肪����΂O������
	If sAmPm = "P" then '�c��̌ߌ�f�[�^������ꍇ����
%>		<tr>
		    <td bgcolor=#fff0f5 align=center>�ߌ�</td>
		    <td bgcolor=#fff0f5 align=center>0</td>
		    <td bgcolor=#fff0f5 align=center>0</td>		<!--VP�Ή� -->
		    <td bgcolor=#fff0f5 align=center>0</td>
		    <td bgcolor=#fff0f5 align=center>0</td>
		    <td bgcolor=#fff0f5 align=center>0</td>		<!--VP�Ή� -->
		    <td bgcolor=#fff0f5 align=center>0</td>			
		</tr>
<%		dCntDate = GetYMDStr(ChgYMDDate(dCntDate) + 1)	'���t�{�P
	End If
	
	Do Until dCntDate > sEndDate
		sWeek = sWeekday(Weekday(cDate(ChgYMDStr(dCntDate))))		'�j���̎擾
%>			<tr>
			    <td bgcolor=#AFEEEE align=center ROWSPAN=2><%=day(ChgYMDStr(dCntDate))%></td>
			    <td bgcolor=#AFEEEE align=center ROWSPAN=2><%=sWeek%></td>
			    <td bgcolor=#FFFFE0 align=center>�ߑO</td>
			    <td bgcolor=#FFFFE0 align=center>0</td>
			    <td bgcolor=#FFFFE0 align=center>0</td>		<!--VP�Ή� -->
			    <td bgcolor=#FFFFE0 align=center>0</td>
			    <td bgcolor=#FFFFE0 align=center>0</td>
			    <td bgcolor=#FFFFE0 align=center>0</td>		<!--VP�Ή� -->
			    <td bgcolor=#FFFFE0 align=center>0</td>			
			</tr>
			<tr>
			    <td bgcolor=#fff0f5 align=center>�ߌ�</td>
			    <td bgcolor=#fff0f5 align=center>0</td>
			    <td bgcolor=#fff0f5 align=center>0</td>		<!--VP�Ή� -->
			    <td bgcolor=#fff0f5 align=center>0</td>
			    <td bgcolor=#fff0f5 align=center>0</td>
			    <td bgcolor=#fff0f5 align=center>0</td>		<!--VP�Ή� -->
			    <td bgcolor=#fff0f5 align=center>0</td>			
			</tr>
<%		dCntDate = GetYMDStr(ChgYMDDate(dCntDate) + 1)
	Loop%>
				<tr>
				    <td bgcolor=#b0c4de align=center>���v</td>
				    <td bgcolor=#b0c4de align=center><br><br></td>
				    <td bgcolor=#b0c4de align=center><br><br></td>
				    <td bgcolor=#b0c4de align=center><%=iRDCount%></td>
				    <td bgcolor=#b0c4de align=center><%=iRVCount%></td>		<!--VP�Ή� -->
				    <td bgcolor=#b0c4de align=center><%=iDelCount%></td>
				    <td bgcolor=#b0c4de align=center><%=iRecCount%></td>
				    <td bgcolor=#b0c4de align=center><%=iVPCount%></td>		<!--VP�Ή� -->
				    <td bgcolor=#b0c4de align=center><%=iUse_sum%></td>			
				</tr>
		</table>
		</center><br>
	�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�i���j���p�񐔂̓f���A�����Q��Ƃ��ăJ�E���g���܂��B
		<center>
	    <form  METHOD="post"  NAME="UPLOAD1" ACTION="SYWB013.asp?TDATE=<%=sTrgDate%>">
			<input type="submit" value="��    ��"id=submit4 name=submit4>
		</form>
		</center>

</body>     
</html>     
