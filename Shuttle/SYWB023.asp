<%@ LANGUAGE="VBScript" %>
<%
Option Explicit
%>
<!--#include file="Common.inc"-->
<!--#include file="Common2.inc"-->
<html>

<head>
<title>�V���[�V�ڍ׉��</title>
</head>

<body >
<%
	Dim sYMD, sChassisID, sDispChassis1, sDispChassis2  
	Dim conn, rsd, sql
	Dim sUsrID, sGrpID, sUsrName, sGrpName, sOperator
	Dim i, j, sNO, sChk1, sChk2, sChkChassisID
	Dim sSize, sPlateNo, sUserName, sPlace, sZokusei, sGrpNm, sContNo
	Dim sGrp_Mei(100), sWorkDate(100), sWorkTime(100), sRecDel(100), sCont(100), sOpeNo(100)

	'�c�a�ڑ�
	Call ConnectSvr(conn, rsd)

	'���[�U���̎擾
	Call GetUserInfo(conn, rsd, sUsrID, sGrpID, sUsrName, sGrpName, sOperator)

	'�V���[�V�h�c�擾
	sChkChassisID = TRIM(Request.QueryString("sCassis"))

	'�V���[�VID�擾

	'�V���[�V�[ID��I�������ꍇ
	sql = "SELECT sChassis.*,sMGroup.GroupName FROM sChassis,sMGroup" & _
	  " WHERE RTRIM(sChassis.ChassisId) = '" & sChkChassisID & "'" & _
	  "   AND RTRIM(sChassis.GroupID) = RTRIM(sMGroup.GroupID)"
	rsd.Open sql, conn, 0, 1, 1

	If Not rsd.EOF Then
		if rsd("Size20Flag") = "Y" then	
			sSize = "20"
		else
			If rsd("MixSizeFlag") = "Y" then	
				sSize = "20/40���p"
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
			sZokusei = "�[�ς̂ݍڂ���"
		end if

		if rsd("NotDelFlag") = "Y" then
			sZokusei = "���o�R���e�i���ڂ��Ȃ�"
		end if

	end if
	rsd.Close

%>
<img border="0" src="image/title01.gif" width="311" height="42">
<br><br>
<center>
<p><img border="0" src="image/title28.gif" width="236" height="34"><p>
</center>

<font face="�l�r �S�V�b�N">
   
<center>
<%dim sdate
sdate = month(date) & "��" & day(date) & "��" & "�@" & hour(time) & "��" & minute(time) & "������"
'Response.Write sdate%>
<u><%=sdate%></u><br><br>

<table border="1" width="500"  >
<b><font color=#000080>�Ώ�</font></b>�@�@�@
	<tr bgcolor=#ffff99><td>
				�V���[�V�@�@�@�@�@�@�@<%=sChkChassisID%><br>
				�T�C�Y�@�@�@�@�@�@�@�@<%=sSize%><br>
				�i���o�[�v���[�g�@�@�@<%=sPlateNo%><br>
				���L�ҁ@�@�@�@�@�@�@�@<%=sUserName%><br>
	</td></tr>
</table><br>

<table border="1" width="500">
<b><font color=#000080>���݂̏��</font></b>
	<tr bgcolor=#ccffcc><td>
				�����O���[�v�@�@�@�@�@<%=sGrpNm%><br>
				�ꏊ�@�@�@�@�@�@�@�@�@<%=sPlace%><br>
				�����@�@�@�@�@�@�@�@�@<%=sZokusei%><br>
<%'�\�����ǂݍ���

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
				���ڃR���e�i�@�@�@�@�@<%=sContNo%>
	</td></tr>
</table><br>

<%'�\�����ǂݍ���
	
	'�V���[�V�[ID��I�������ꍇ
	sql = "SELECT OpeNo,WorkDate,Term,RecDel,ContNo,GroupName,DelFlag FROM sAppliInfo,sMGroup" & _
	  " WHERE RTRIM(sAppliInfo.ChassisId) = '" & sChkChassisID & "'" & _
	  "   AND RTRIM(sAppliInfo.GroupID) = RTRIM(sMGroup.GroupID) " & _
	  "   AND DelFlag = ' '  ORDER BY WorkDate"
	rsd.Open sql, conn, 0, 1, 1

	i = 1
	If Not rsd.EOF Then
				
		Do until rsd.EOF
			
			sGrp_Mei(int(i)) = rsd("GroupName")	'�O���[�v��
			'���ɂ�
			sWorkDate(int(i)) = month(rsd("WorkDate")) & "��" 
			sWorkDate(int(i)) = sWorkDate(int(i)) & day(rsd("WorkDate")) & "��"
			'��Ǝ���
			sWorkTime(int(i)) = trim(rsd("Term"))
			'��Ɣԍ�
			If len(trim(rsd("OpeNo"))) = 4 Then
				sOpeNo(int(i))    = "0" & trim(rsd("OpeNo"))
			Else
				sOpeNo(int(i))    = trim(rsd("OpeNo"))
			End IF
			'��Ǝ��(VP�Ή�)
			if rsd("RecDel") = "R" then
				sRecDel(int(i)) = "����"
			Elseif rsd("RecDel") = "D" then
				sRecDel(int(i)) = "���o"
			else
				sRecDel(int(i)) = "��o��"
			end if	
			'�R���e�i
			if trim(rsd("ContNo")) = "" then
				sCont(int(i)) = "�@"
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
				<b><font color=#000080>�����N���</font></b>
					<tr>
						<td bgcolor="#e8ffe8" align=center>�O���[�v</td>
					    <td bgcolor="#e8ffe8" align=center>���ɂ�</td>			
					    <td bgcolor="#e8ffe8" align=center>��Ǝ���</td>			
					    <td bgcolor="#e8ffe8" align=center>�\��ԍ�</td>			
					    <td bgcolor="#e8ffe8" align=center>��Ǝ��</td>			
					    <td bgcolor="#e8ffe8" align=center>�R���e�i</td>			
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
					%>�Y���̍�ƂȂ�<%	
	end if%>

</center><br>
<center>
    <form>
    <input type="button" value="�@�߂�@" onclick="history.back()" >
	</form>
</center>
</body>     
</html>