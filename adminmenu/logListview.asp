<%@Language="VBScript" %>

<!--#include file="./Common/Common.inc"-->
<%
'  �i�ύX�����j
'   2013-09-26   Y.TAKAKUWA   �X�}�[�g�t�H����ǉ��B
%>
<%
' �W�v���W�b�N
	On Error Resume Next

	Dim sYearF,sMonthF,sDataF
	Dim sYearT,sMonthT,sDataT
	Dim sOldYer,sOldDate
	Dim iDataFlag,iOldFlag,iDateCnt
	Dim conn, rs,connC, rsC,sql
	Dim LogDate(),Count(),MCount(),HDate(),DCount(),TCount()
	'2013-09-30 Y.TAKAKUWA Add-S
	Dim SCount()
	'2013-09-30 Y.TAKAKUWA Add-E
	Dim iTotal,iTemp,LineNo

	iDataFlag=0
	iOldFlag=0
	Set fs=Server.CreateObject("Scripting.FileSystemobject")


	'�p�����[�^�擾
	sYearF=left(Request.QueryString("fDate"),4)
	sMonthF=mid(Request.QueryString("fDate"),6,2)
	sYearT=left(Request.QueryString("tDate"),4)
	sMonthT=mid(Request.QueryString("tDate"),6,2)


	sDataF=sYearF & sMonthF
	sDataT=sYearT & sMonthT


	'3�N�O�̔N�����i�[
	sOldYer=cstr(cint(Year(now))-4) & "12"
	sOldDate=cstr(cint(Year(now))-4) & "/12/01"

	'----------------------------------------
	' �c�a�ڑ�
	'----------------------------------------        
	ConnectSvr conn, rs
	ConnectSvrC connC, rsC
	
	'3�N�ȏ�O������
	if sDataF<=sOldYer then
		iOldFlag=1
		'2013-09-30 Y.TAKAKUWA Upd-S
		'sql =" SELECT LogDate,SUM(DataCount)as DataCount,SUM(MDataCount)as MDataCount  "
		sql =" SELECT LogDate,SUM(DataCount)as DataCount,SUM(MDataCount)as MDataCount, SUM(SDataCount) as SDataCount  "
		'2013-09-30 Y.TAKAKUWA Upd-E
		sql = sql & " FROM ("
		sql = sql & " SELECT "
		sql = sql & " substring(LogDate,1,4) as LogDate "
		sql = sql & " ,CASE WHEN UpdtPgCd='COMMNINC' THEN DataCount ELSE 0 END as DataCount"
		sql = sql & " ,CASE WHEN UpdtPgCd='MCOMNINC' THEN DataCount ELSE 0 END as MDataCount"
		'2013-09-30 Y.TAKAKUWA Add-S
		sql = sql & " ,CASE WHEN UPPER(UpdtPgCd)='BROWSER' THEN DataCount ELSE 0 END as SDataCount"
		'2013-09-30 Y.TAKAKUWA Add-E
		sql = sql & " FROM xLog "
		sql = sql & " WHERE substring(LogDate,1,6)>='" & sDataF & "'"
		sql = sql & " AND substring(LogDate,1,6)<='" & sOldYer & "'"
		sql = sql & " ) MAIN "
		sql = sql & " Group By LogDate"
		sql = sql & " UNION "
	else
		sql=""
	end if

    '2013-09-30 Y.TAKAKUWA Upd-S
	'sql = sql & " SELECT LogDate,SUM(DataCount)as DataCount,SUM(MDataCount)as MDataCount  "
	sql = sql & " SELECT LogDate,SUM(DataCount)as DataCount,SUM(MDataCount)as MDataCount, SUM(SDataCount) as SDataCount  "
	'2013-09-30 Y.TAKAKUWA Upd-E
	sql = sql & " FROM ("
	sql = sql & " SELECT "
	sql = sql & " substring(LogDate,1,6) as LogDate "
	sql = sql & " ,CASE WHEN UpdtPgCd='COMMNINC' THEN DataCount ELSE 0 END as DataCount"
	sql = sql & " ,CASE WHEN UpdtPgCd='MCOMNINC' THEN DataCount ELSE 0 END as MDataCount"
	'2013-09-30 Y.TAKAKUWA Add-S
	sql = sql & " ,CASE WHEN UPPER(UpdtPgCd)='BROWSER' THEN DataCount ELSE 0 END as SDataCount"
	'2013-09-30 Y.TAKAKUWA Add-E
	sql = sql & " FROM xLog "
	'�ߋ�3�N���܂ރf�[�^�̏ꍇ
	if iOldFlag= 1 then
		sql = sql & " WHERE substring(LogDate,1,6)>'" & sOldYer & "'"
	else
		sql = sql & " WHERE substring(LogDate,1,6)>='" & sDataF & "'"
	end if
	sql = sql & " AND substring(LogDate,1,6)<='" & sDataT & "'"
	sql = sql & " ) MAIN "
	sql = sql & " Group By LogDate"
	sql = sql & " Order By LogDate"

	rs.Open sql, conn, 0, 1, 1
		on error resume next
	
	'�f�[�^���݃`�F�b�N
	if rs.eof or err.number<>0 then
		iDataFlag=0
	else
		iDataFlag=1
	end if

	rsC.Open sql, connC, 0, 1, 1
		on error resume next

	if iDataFlag=0 then
		if rsC.eof or err.number<>0 then
			iDataFlag=0
		else
			iDataFlag=1
		end if
	end if

	LineNo=0
	'�f�[�^�����݂���ꍇ
	if iDataFlag=1 then
		'���O�W�v�f�[�^�̎擾
		'Hits�f�[�^�����[�v
		Do While Not rs.EOF
			ReDim Preserve LogDate(LineNo)
			ReDim Preserve Count(LineNo)
			ReDim Preserve MCount(LineNo)
			'2013-09-30 Y.TAKAKUWA Add-S
			ReDim Preserve SCount(LineNo)
			'2013-09-30 Y.TAKAKUWA Add-E
			if len(trim(rs("LogDate")))=4 then
				LogDate(LineNo)=trim(rs("LogDate")) & "�N"
			else
				LogDate(LineNo)=left(trim(rs("LogDate")),4) & "/" & right(trim(rs("LogDate")),2)
			end if
			Count(LineNo)=rs("DataCount")
			MCount(LineNo)=rs("MDataCount")
			'2013-09-30 Y.TAKAKUWA Add-S
			SCount(LineNo)=rs("SDataCount")
			'2013-09-30 Y.TAKAKUWA Add-E
			LineNo=LineNo+1

			rs.MoveNext
		Loop

		'CAM�f�[�^�����[�v
		Do While Not rsC.EOF
			ReDim Preserve LogDate(LineNo)
			ReDim Preserve Count(LineNo)
			ReDim Preserve MCount(LineNo)
			'2013-09-30 Y.TAKAKUWA Add-S
			ReDim Preserve SCount(LineNo)
			'2013-09-30 Y.TAKAKUWA Add-E
			if len(trim(rsC("LogDate")))=4 then
				LogDate(LineNo)=trim(rsC("LogDate")) & "�N"
			else
				LogDate(LineNo)=left(trim(rsC("LogDate")),4) & "/" & right(trim(rsC("LogDate")),2)
			end if
			Count(LineNo)=rsC("DataCount")
			MCount(LineNo)=rsC("MDataCount")
			'2013-09-30 Y.TAKAKUWA Add-S
			SCount(LineNo)=rsC("SDataCount")
			'2013-09-30 Y.TAKAKUWA Add-E
			LineNo=LineNo+1
			rsC.MoveNext
		Loop

		iTotal=0
		If LineNo>0 Then
			'���O�^�C�g���擾
			'�ߋ�3�N���܂ޏꍇ
			if iOldFlag=1 then
				iCount=DateDiff("yyyy", left(sDataF,4) & "/" & mid(sDataF,5,2) & "/01",sOldDate)
				tmpDate=left(sDataF,4) & "/" & mid(sDataF,5,2) & "/01"
				do
					if iCount<iLoop then
						exit do
					end if
				
					ReDim Preserve HDate(iDateCnt)
					HDate(iDateCnt)=left(tmpDate,4) & "�N"
					tmpDate=DateAdd("yyyy", 1, tmpDate)
					iDateCnt=iDateCnt+1
					iLoop=iLoop+1
				Loop
				iCount=DateDiff("m", cstr(cint(left(sOldDate,4))+1) & "/" & "01/01", left(sDataT,4) & "/" & mid(sDataT,5,2) & "/01")
				tmpDate=cstr(cint(left(sOldDate,4))+1) & "/" & "01/01" 
			else
				iCount=DateDiff("m", left(sDataF,4) & "/" & mid(sDataF,5,2) & "/01", left(sDataT,4) & "/" & mid(sDataT,5,2) & "/01")
				tmpDate=left(sDataF,4) & "/" & mid(sDataF,5,2) & "/" & right(sDataF,2)				
			end if
			iLoop=0
			do
				if iCount<iLoop then
					exit do
				end if
				
				ReDim Preserve HDate(iDateCnt)
				HDate(iDateCnt)=left(tmpDate,4) & "/" & mid(tmpDate,6,2)
				tmpDate=DateAdd("m", 1, tmpDate)
				iDateCnt=iDateCnt+1
				iLoop=iLoop+1
			Loop
			'2013-09-30 Y.TAKAKUWA Upd-S
			'ReDim DCount(iDateCnt-1,3)
			Redim DCount(iDateCnt-1,4)
			'2013-09-30 Y.TAKAKUWA Upd-E
			For i=0 to iDateCnt-1

				for j=0 to 4       '2013-09-30 Y.TAKAKUWA Upd 3 -> 4 
					DCount(i,J)=0
				Next

				for j=0 to LineNo-1
					'���t�������ꍇ
					if cstr(HDate(i))=cstr(LogDate(j)) Then
						DCount(i,0)=DCount(i,0)+Count(j)	'PC
						DCount(i,1)=DCount(i,1)+MCount(j)	'�g��
						'2013-09-30 Y.TAKAKUWA Add-S
						DCount(i,2)=DCount(i,2)+SCount(j)   '�X�}�g�t�H��
						'2013-09-30 Y.TAKAKUWA Add-E
						'2013-09-30 Y.TAKAKUWA Upd-S
						'DCount(i,2)=DCount(i,2)+Count(j)+MCount(j)	'���v
						DCount(i,3)=DCount(i,3)+Count(j)+MCount(j)+SCount(j)	'���v
						'2013-09-30 Y.TAKAKUWA Upd-E
					end If 
				Next
				'2013-09-30 Y.TAKAKUWA Upd-S
				'iTotal=iTotal+DCount(i,2)
				'DCount(i,3)=iTotal
				iTotal=iTotal+DCount(i,3)
				DCount(i,4)=iTotal
				'2013-09-30 Y.TAKAKUWA Upd-E
				
			Next
		End If
	End if
    '2013-09-30 Y.TAKAKUWA Upd-S
	'Redim TCount(3)
	Redim TCount(4)
	'2013-09-30 Y.TAKAKUWA Upd-E
	iTemp=0

	for i=0 to iDateCnt-1
		TCount(0)=TCount(0)+DCount(i,0)
		TCount(1)=TCount(1)+DCount(i,1)
		'2013-09-30 Y.TAKAKUWA Upd-S
		'TCount(2)=TCount(2)+DCount(i,2)
		'TCount(3)=iTotal
		TCount(2)=TCount(2)+DCount(i,2)
		TCount(3)=TCount(3)+DCount(i,3)
		TCount(4)=iTotal
		'2013-09-30 Y.TAKAKUWA Upd-E
	Next
	set conn=nothing
	set rs=nothing
	set connC=nothing
	set rsC=nothing

	call Makecsv(sDataF,sDataT,sMode)
'------------------------------
'CSV�t�@�C���쐬
'------------------------------   
function MakeCsv(sDataF,sDataT,sMode)
	dim filenm     '�t�@�C����	
	dim path,ObjFSO, strFileName


	'�f�[�^�����݂���ꍇ
	if iDataFlag=1 then

		strFileName=GetNumStr(Session.SessionID, 8) & ".csv"


		Session.Contents("tempfile")=strFileName

		'�t�@�C���I�u�W�F�N�g�쐬
	    	Set ObjFSO = Server.CreateObject("Scripting.FileSystemObject")

		
		' �t�@�C�����ҏW
		filenm = Server.Mappath("../temp/" & strFileName)


		' �t�@�C���쐬
		Set ObjTS = ObjFSO.OpenTextFile(filenm, 2, True)



		if Err.Number <> 0 then
			response.write Err.description
			response.end
		end if

		For i=0 to iDateCnt-1
		    '2013-09-30 Y.TAKAKUWA Upd-S
			'ObjTS.WriteLine trim(HDate(i)) & "," & DCount(i,0) & "," & DCount(i,1) & "," & DCount(i,2) & "," & DCount(i,3) & ","
			ObjTS.WriteLine trim(HDate(i)) & "," & DCount(i,0) & "," & DCount(i,1) & "," & DCount(i,2) & "," & DCount(i,3) & "," & DCount(i,4) & ","
			'2013-09-30 Y.TAKAKUWA Upd-E
		Next
		'2013-09-30 Y.TAKAKUWA Upd-S
		'ObjTS.WriteLine  "�v," & TCount(0) & "," & TCount(1) & "," & TCount(2) & "," & TCount(3) & ","
		ObjTS.WriteLine  "�v," & TCount(0) & "," & TCount(1) & "," & TCount(2) & "," & TCount(3) & "," & TCount(4) & ","
        '2013-09-30 Y.TAKAKUWA Upd-E
        
		'--- �t�@�C������� ---
		ObjTS.Close   '���O�t�@�C���N���[�Y

	end if
end function
%>

<html>
<head>
	<title>�A�N�Z�X���O�W�v</title>
	<meta http-equiv="Pragma" content="no-cache">
	<meta http-equiv="Content-Type" content="text/html; charset=Sh1ift_JIS">
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="../gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<script language="JavaScript">

</script>
<!-------------����������--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
<tr><td height="20"></td></tr>
<tr>
	<td valign="top">

		<center>
		<table>
		<tr> 
			<td><img src="../gif/botan.gif" width="17" height="17"></td>
			<td nowrap><b>�A�N�Z�X�����݌v�\</b></td>
			<td><img src="../gif/hr.gif" width="400" height="3"></td>
			<INPUT type="hidden" name="Gamen_Mode" size="9" maxlength="1"  readonly tabindex= -1>
		</tr>
		</table>
		<br>
		<table border="0">
		<tr align=left>
		<td align=left>
		<BR>
		</td>
		</tr>
		<tr>
			<td align=left>
			<% If iDataFlag>0 Then %>
				<table border="1" cellpadding="5">
					<tr>
						<th align="center" bgcolor="#6699FF">�敪</th>
						<th align="center" bgcolor="#6699FF" width="100">PC</th>
						<th align="center" bgcolor="#6699FF" width="100">�g�ђ[��</th>
						<!--2013-09-26 Y.TAKAKUWA Add-S-->
						<th align="center" bgcolor="#6699FF" width="100">�X�}�[�g�t�H��</th>
						<!--2013-09-26 Y.TAKAKUWA Add-E-->
						<th align="center" bgcolor="#6699FF" width="100">���v</th>
						<th align="center" bgcolor="#6699FF" width="100">�݌v</th>
					</tr>
					<% For i=0 to iDateCnt-1 %>
					<tr>
						<% If Instr(HDate(i),"�N")<>0 then %>
							<td align="center" bgcolor="#E0FFFF" width="85"><%=HDate(i)%></td>
							<td align="right" bgcolor="#E0FFFF" width="85"><%=FormatNumber(DCount(i,0),0)%> </td>
							<td align="right" bgcolor="#E0FFFF" width="85"><%=FormatNumber(DCount(i,1),0)%> </td>
							<!--2013-09-26 Y.TAKAKUWA Add-S-->
							<!--<td align="right" bgcolor="#E0FFFF" width="85"><%'FormatNumber(DCount(i,2),0)%> </td>
							<td align="right" bgcolor="#E0FFFF" width="85"><%'FormatNumber(DCount(i,3),0)%> </td>-->
							<td align="right" bgcolor="#E0FFFF" width="85"><%=FormatNumber(DCount(i,2),0)%></td>
							<td align="right" bgcolor="#E0FFFF" width="85"><%=FormatNumber(DCount(i,3),0)%> </td>
							<td align="right" bgcolor="#E0FFFF" width="85"><%=FormatNumber(DCount(i,4),0)%> </td>
							<!--2013-09-26 Y.TAKAKUWA Add-E-->
						<% Else %>
							<td align="center" width="85"><%=HDate(i)%></td>
							<td align="right" width="85"><%=FormatNumber(DCount(i,0),0)%> </td>
							<td align="right" width="85"><%=FormatNumber(DCount(i,1),0)%> </td>
							<!--2013-09-26 Y.TAKAKUWA Add-S-->
							<!--<td align="right" width="85"><%=FormatNumber(DCount(i,2),0)%> </td>
							<td align="right" width="85"><%=FormatNumber(DCount(i,3),0)%> </td>-->
							<td align="right" width="85"><%=FormatNumber(DCount(i,2),0)%> </td>
							<td align="right" width="85"><%=FormatNumber(DCount(i,3),0)%> </td>
							<td align="right" width="85"><%=FormatNumber(DCount(i,4),0)%> </td>
							<!--2013-09-26 Y.TAKAKUWA Add-E-->

							
						<% End If %>
					</tr>
					<% Next %>
					<tr>
					<td colspan=1 align="Center">�v</td>
					<td align="right" width="85"><%=FormatNumber(TCount(0),0)%></td>
					<td align="right" width="85"><%=FormatNumber(TCount(1),0)%></td>
					<!--2013-09-26 Y.TAKAKUWA Add-S-->
					<!--<td align="right" width="85"><%=FormatNumber(TCount(2),0)%></td>
					<td align="right" width="85"><%=FormatNumber(TCount(3),0)%></td>-->
					<td align="right" width="85"><%=FormatNumber(TCount(2),0)%></td>
					<td align="right" width="85"><%=FormatNumber(TCount(3),0)%></td>
					<td align="right" width="85"><%=FormatNumber(TCount(4),0)%></td>
					<!--2013-09-26 Y.TAKAKUWA Add-E-->
					
					</tr>
					</table>
			<% Else %>
				<br><div align="center">�f�[�^��1��������܂���B</div><br>
			<% End If %>
			</td>
		</tr>
		<% If LineNo>0 Then %>
		<tr align=Center>
			<td>
			<BR>
			<form action="logListcsvout.asp"><input type="submit" value="CSV�t�@�C���o��">
			</form>
			</td>
		</tr>
		<% End If %>
		</table>
		<a href="javascript:history.back();">�߂�</a>
		<br><br>
		</center>
	</td>
</tr>
</table>
</body>
</html>