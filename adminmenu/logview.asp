<%@Language="VBScript" %>

<!--#include file="./Common/Common.inc"-->
<%
'  �i�ύX�����j
'   2013-09-26   Y.TAKAKUWA   �X�}�[�g�t�H���̃J�E���g��ǉ��B
%>
<%
' �W�v���W�b�N
	On Error Resume Next

	Dim sYear,sMonth,sDay
	Dim sMode
	Dim sYearF,sMonthF,sDayF,sDataF
	Dim sYearT,sMonthT,sDayT,sDataT
	Dim conn, rs,connC, rsC,sql
	Dim iDataFlag
	Dim strLogData()
	Dim iHdRow,iGSum,iTSum,iMTSum,sHdValue,LineNo,MLineNo
    '2013-09-26 Y.TAKAKUWA Add-S
    Dim iSTSum, SLineNo
    '2013-09-26 Y.TAKAKUWA Add-E
	iDataFlag=0

	Set fs=Server.CreateObject("Scripting.FileSystemobject")


	'�p�����[�^�擾
	sYearF=left(Request.QueryString("fDate"),4)
	sMonthF=mid(Request.QueryString("fDate"),6,2)
	sDayF=right(Request.QueryString("fDate"),2)
	sYearT=left(Request.QueryString("tDate"),4)
	sMonthT=mid(Request.QueryString("tDate"),6,2)
	sDayT=right(Request.QueryString("tDate"),2)
	sMode=trim(Request.QueryString("Mode"))


	'�����p���t�쐬
	if sMode="D" then
		sDataF=sYearF & sMonthF & sDayF
		sDataT=sYearT & sMonthT & sDayT
	else
		sDataF=sYearF & sMonthF
		sDataT=sYearT & sMonthT
	end if

	'----------------------------------------
	' �c�a�ڑ�
	'----------------------------------------        
	ConnectSvr conn, rs
	ConnectSvrC connC, rsC

	sql=" SELECT PageNum,WkNum,SUM(DataCount) as DataCount "
	sql=sql & " FROM ("
	sql=sql & " SELECT "
	'���ʂ̏ꍇ
	if sMode="M" then
		sql=sql & " substring(LogDate,1,6) as LogDate "
	else
		sql=sql & " substring(LogDate,1,8) as LogDate "
	end if
	sql=sql & " ,PageNum,WkNum,DataCount "
	sql=sql & " FROM xLog "
	sql=sql & " ) MAIN "
	sql=sql & " WHERE LogDate>='" & sDataF & "'"
	sql=sql & " AND LogDate<='" & sDataT & "'"
	sql=sql & " Group By PageNum,WkNum "
	sql=sql & " ORDER By PageNum,WkNum"

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
	iTSum=0
	MLineNo=0
	iMTSum=0
	'2013-09-26 Y.TAKAKUWA Add-S
	SLineNo=0
	iSTSum=0
	'2013-09-26 Y.TAKAKUWA Add-E
	'�f�[�^�����݂���ꍇ
	if iDataFlag=1 then
		'���O�W�v�f�[�^�̎擾
		'Hits�f�[�^�����[�v
		Do While Not rs.EOF
			ReDim Preserve strLogData(LineNo)
			strLogData(LineNo)=trim(rs("PageNum")) & "," & trim(rs("WkNum")) & "," & trim(rs("DataCount"))
			LineNo=LineNo+1
			rs.MoveNext
		Loop

		'CAM�f�[�^�����[�v
		Do While Not rsC.EOF
			ReDim Preserve strLogData(LineNo)
			strLogData(LineNo)=trim(rsC("PageNum")) & "," & trim(rsC("WkNum")) & "," & trim(rsC("DataCount"))
			LineNo=LineNo+1
			rsC.MoveNext
		Loop

		If LineNo>0 Then
			'' ���O�^�C�g���擾
			Dim iTKind
			Dim PageNum(),WkNum(),PageTitle(),SubTitle()
			Dim strTitleFileName
			strTitleFileName="../logweb.txt"
			Set ti=fs.OpenTextFile(Server.MapPath(strTitleFileName),1,True)
			iTKind=0
			Do While Not ti.AtEndOfStream
				strTemp=ti.ReadLine
				anyTmpTitle=Split(strTemp,",")
				If anyTmpTitle(2)<>"" Then 
					ReDim Preserve PageNum(iTKind)
					ReDim Preserve WkNum(iTKind)
					ReDim Preserve PageTitle(iTKind)
					ReDim Preserve SubTitle(iTKind)
					PageTitle(iTKind) = anyTmpTitle(2)
					PageNum(iTKind) = "<BR>"
					WkNum(iTKind) = "<BR>"
					SubTitle(iTKind) = "<BR>"
					iTKind=iTKind+1
				end if
				ReDim Preserve PageNum(iTKind)
				ReDim Preserve WkNum(iTKind)
				ReDim Preserve PageTitle(iTKind)
				ReDim Preserve SubTitle(iTKind)
				PageNum(iTKind) = anyTmpTitle(0)
				WkNum(iTKind) = anyTmpTitle(1)
				PageTitle(iTKind) = "<BR>"
				'If PageTitle(iTKind)="" Then PageTitle(iTKind)="<BR>"
				SubTitle(iTKind) = anyTmpTitle(3)
				iTKind=iTKind+1
			Loop
			ti.Close

			sHdValue=""
			'' ���O�̏W�v
			ReDim Count(iTKind-1)
			For i=0 to iTKind-1
				'���j���[���ڂ��ς�����ꍇ
				if sHdValue<>PageTitle(i) and PageTitle(i)<>"<BR>" then
					'�擪�s�ȊO
					if i<>0 then
						Count(iHdRow)=iGSum	
					end if
					iHdRow=i
					iGSum=0
					sHdValue=PageTitle(i)
				end if
				Count(i)=0
				For j=0 to LineNo-1
					anyTmp=Split(strLogData(j),",")
					If anyTmp(0)=PageNum(i) and anyTmp(1)=WkNum(i) then
						Count(i)=Count(i)+anyTmp(2)
						iGSum=iGSum+anyTmp(2)
						iTSum=iTSum+anyTmp(2)
					End If
				Next
			Next
			'�ŏI�s�̃f�[�^�𑫂�����
			if iGSum<>0 then
				Count(iHdRow)=iGSum	
			end if

			'�g�їp����
			Dim MPageNum(),MWkNum(),MPageTitle(),MSubTitle()
			strTitleFileName="../logija.txt"
			Set ti=fs.OpenTextFile(Server.MapPath(strTitleFileName),1,True)
			MLineNo=0
			Do While Not ti.AtEndOfStream
				strTemp=ti.ReadLine
				anyTmpTitle=Split(strTemp,",")
				If anyTmpTitle(2)<>"" Then 
					ReDim Preserve MPageNum(MLineNo)
					ReDim Preserve MWkNum(MLineNo)
					ReDim Preserve MPageTitle(MLineNo)
					ReDim Preserve MSubTitle(MLineNo)
					MPageTitle(MLineNo) = anyTmpTitle(2)
					MPageNum(MLineNo) = "<BR>"
					MWkNum(MLineNo) = "<BR>"
					MSubTitle(MLineNo) = "<BR>"
					MLineNo=MLineNo+1
				end if
				ReDim Preserve MPageNum(MLineNo)
				ReDim Preserve MWkNum(MLineNo)
				ReDim Preserve MPageTitle(MLineNo)
				ReDim Preserve MSubTitle(MLineNo)
				MPageNum(MLineNo) = anyTmpTitle(0)
				MWkNum(MLineNo) = anyTmpTitle(1)
				MPageTitle(MLineNo) = "<BR>"
				'If PageTitle(iTKind)="" Then PageTitle(iTKind)="<BR>"
				MSubTitle(MLineNo) = anyTmpTitle(3)
				MLineNo=MLineNo+1
			Loop
			ti.Close
			sHdValue=""
			'' ���O�̏W�v
			ReDim MCount(MLineNo-1)
			For i=0 to MLineNo-1
				'���j���[���ڂ��ς�����ꍇ
				if sHdValue<>MPageTitle(i) and MPageTitle(i)<>"<BR>" then
					'�擪�s�ȊO
					if i<>0 then
						MCount(iHdRow)=iGSum	
					end if
					iHdRow=i
					iGSum=0
					sHdValue=MPageTitle(i)
				end if
				MCount(i)=0
				For j=0 to LineNo-1
					anyTmp=Split(strLogData(j),",")
					If anyTmp(0)=MPageNum(i) and anyTmp(1)=MWkNum(i) then
						MCount(i)=MCount(i)+anyTmp(2)
						iGSum=iGSum+anyTmp(2)
						iMTSum=iMTSum+anyTmp(2)
					End If
				Next
			Next
			'�ŏI�s�̃f�[�^�𑫂�����
			if iGSum<>0 then
				MCount(iHdRow)=iGSum	
			end if
			
			
			'2013-09-26 Y.TAKAKUWA Add-S
			'�X�}�[�g�t�H������
			Dim SPageNum(),SWkNum(),SPageTitle(),SSubTitle()
			strTitleFileName="../logsumafo.txt"
			Set ti=fs.OpenTextFile(Server.MapPath(strTitleFileName),1,True)
			SLineNo=0

			Do While Not ti.AtEndOfStream
				strTemp=ti.ReadLine
				anyTmpTitle=Split(strTemp,",")
				
				If anyTmpTitle(2)<>"" Then 
					ReDim Preserve SPageNum(SLineNo)
					ReDim Preserve SWkNum(SLineNo)
					ReDim Preserve SPageTitle(SLineNo)
					ReDim Preserve SSubTitle(SLineNo)
					SPageTitle(SLineNo) = anyTmpTitle(2)
					SPageNum(SLineNo) = "<BR>"
					SWkNum(SLineNo) = "<BR>"
					SSubTitle(SLineNo) = "<BR>"
					
					SLineNo=SLineNo+1
				end if
				ReDim Preserve SPageNum(SLineNo)
				ReDim Preserve SWkNum(SLineNo)
				ReDim Preserve SPageTitle(SLineNo)
				ReDim Preserve SSubTitle(SLineNo)
				SPageNum(SLineNo) = anyTmpTitle(0)
				SWkNum(SLineNo) = anyTmpTitle(1)
				SPageTitle(SLineNo) = "<BR>"
				SSubTitle(SLineNo) = anyTmpTitle(3)
				SLineNo=SLineNo+1
			Loop
			ti.Close
			sHdValue=""
			'' ���O�̏W�v
			ReDim SCount(SLineNo-1)
			For i=0 to SLineNo-1
				'���j���[���ڂ��ς�����ꍇ
				if sHdValue<>SPageTitle(i) and SPageTitle(i)<>"<BR>" then
					'�擪�s�ȊO
					if i<>0 then
						SCount(iHdRow)=iGSum	
					end if
					iHdRow=i
					iGSum=0
					sHdValue=SPageTitle(i)
				end if
				SCount(i)=0
				For j=0 to LineNo-1
					anyTmp=Split(strLogData(j),",")
					If anyTmp(0)=SPageNum(i) and anyTmp(1)=SWkNum(i) then
						SCount(i)=SCount(i)+anyTmp(2)
						iGSum=iGSum+anyTmp(2)
						iSTSum=iSTSum+anyTmp(2)
					End If
				Next
			Next
			'�ŏI�s�̃f�[�^�𑫂�����
			if iGSum<>0 then
				SCount(iHdRow)=iGSum	
			end if
			'2013-09-26 Y.TAKAKUWA Add-E
		End If
	End if

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


	'----------------------------------------
	' �c�a�ڑ�
	'----------------------------------------        
	ConnectSvr conn, rs
	ConnectSvrC connC, rsC

	sql=" SELECT LogDate,PageNum,WkNum,SUM(DataCount) as DataCount "
	sql=sql & " FROM ("
	sql=sql & " SELECT "
	'���ʂ̏ꍇ
	if sMode="M" then
		sql=sql & " substring(LogDate,1,6) as LogDate "
	else
		sql=sql & " substring(LogDate,1,8) as LogDate "
	end if
	
	sql=sql & " ,PageNum,WkNum,DataCount "
	sql=sql & " FROM xLog "
	sql=sql & " ) MAIN "
	sql=sql & " WHERE LogDate>='" & sDataF & "'"
	sql=sql & " AND LogDate<='" & sDataT & "'"
	sql=sql & " Group By LogDate,PageNum,WkNum "
	sql=sql & " ORDER By PageNum,WkNum,LogDate"

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

		'�o�͏�������������
		ObjTS.WriteLine sDataF & "," & sDataT & "," & sMode

		'Hits�f�[�^�����[�v
		Do While Not rs.EOF
			ObjTS.WriteLine trim(rs("PageNum")) & "," & trim(rs("WkNum")) & "," & trim(rs("LogDate")) & "," & trim(rs("DataCount"))
			rs.MoveNext
		Loop

		'CAM�f�[�^�����[�v
		Do While Not rsC.EOF
			ObjTS.WriteLine trim(rsC("PageNum")) & "," & trim(rsC("WkNum")) & "," & trim(rsC("LogDate")) & "," & trim(rsC("DataCount"))
			rsC.MoveNext
		Loop
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
			<td nowrap><b>���p�����\��(���Ԍ���)</b></td>
			<td><img src="../gif/hr.gif" width="400" height="3"></td>
			<INPUT type="hidden" name="Gamen_Mode" size="9" maxlength="1"  readonly tabindex= -1>
		</tr>
		</table>
		<br>
		<table border="0">
		<tr align=left>
		<td align=left>
		<% If sMode="D" Then %>
			<%=sYearF & "�N" & sMonthF & "��" & sDayF & "��"%>����<%=sYearT & "�N" & sMonthT & "��" & sDayT & "��"%>�܂�
		<% Else %>
			<%=sYearF & "�N" & sMonthF & "��"%>����<%=sYearT & "�N" & sMonthT & "��"%>�܂�
		<% End If %>
		</td>
		</tr>
		<tr>
			<td align=left>
			<% If LineNo>0 Then %>
				���p�\�R����
				<table border="1" cellpadding="5">
					<tr>
						<th align="center" bgcolor="#6699FF">���j���[����</th>
						<th align="center" bgcolor="#6699FF">���</th>
						<th align="center" bgcolor="#6699FF">���No.</th>
						<th align="center" bgcolor="#6699FF" width="100">�A�N�Z�X����</th>
					</tr>
					<% For i=0 to iTKind-1 %>
					<tr>
						<td align="left"><%=PageTitle(i)%></td>
						<td align="left"><%=SubTitle(i)%></td>
						<% If PageNum(i)<>"<BR>" Then %>
						<td align="left"><%=PageNum(i)%>-<%=WkNum(i)%></td>
						<% Else %>
						<td align="left"><%=PageNum(i)%></td>
						<% End If %>
						<td align="right" width="85"><%=FormatNumber(Count(i),0)%> </td>
					</tr>
					<% Next %>
					<tr>
					<td colspan=3 align="Center">���v</td>
					<td align="right" width="85"><%=FormatNumber(iTSum,0)%> </td>
					</tr>
					</table>
				<% If MLineNo>0 Then %>
				<BR>
				���g�ѓd�b��
				<table border="1" cellpadding="5" width="100%">
					<tr>
						<th align="center" bgcolor="#6699FF" >���j���[����</th>
						<th align="center" bgcolor="#6699FF">���</th>
						<th align="center" bgcolor="#6699FF">���No.</th>
						<th align="center" bgcolor="#6699FF" width="100">�A�N�Z�X����</th>
					</tr>
					<% For i=0 to MLineNo-1 %>
					<tr>
						<td align="left"><%=MPageTitle(i)%></td>
						<td align="left"><%=MSubTitle(i)%></td>
						<% If MPageNum(i)<>"<BR>" Then %>
						<td align="left"><%=MPageNum(i)%>-<%=MWkNum(i)%></td>
						<% Else %>
						<td align="left"><%=MPageNum(i)%></td>
						<% End If %>
						<td align="right" width="85"><%=FormatNumber(MCount(i),0)%> </td>
					</tr>
					<% Next %>
					<tr>
					<td colspan=3 align="Center">���v</td>
					<td align="right" width="85"><%=FormatNumber(iMTSum,0)%> </td>
					</tr>
					</table>
					<!--2013-09-26 Y.TAKAKUWA Add-S-->
					<%' If iMTSum<>0 or iTSum<>0 Then %>
					<!--
					<BR>
					<tr align=right><td>
					<table border="1" cellpadding="5" >
					<tr align=right>
					<td colspan=3 align="Center">�����v</td>
					<td align="right" width="100"><%'FormatNumber((iMTSum+iTSum),0)%> </td>
					</tr>
					</table>
					</td></tr>
					-->
					<!--2013-09-26 Y.TAKAKUWA Add-E-->
					<%' End If %>
					
				<% End If %>
				<!--2013-09-26 Y.TAKAKUWA Add-S-->
				<% If SLineNo>0 Then %>
				<BR>
				���X�}�[�g�t�H����
				<table border="1" cellpadding="5" width="100%">
					<tr>
						<th align="center" bgcolor="#6699FF" >���j���[����</th>
						<th align="center" bgcolor="#6699FF">���</th>
						<th align="center" bgcolor="#6699FF">���No.</th>
						<th align="center" bgcolor="#6699FF" width="100">�A�N�Z�X����</th>
					</tr>
					<% For i=0 to SLineNo-1 %>
					<tr>
						<td align="left"><%=SPageTitle(i)%></td>
						<td align="left"><%=SSubTitle(i)%></td>
						<% If SPageNum(i)<>"<BR>" Then %>
						<td align="left"><%=SPageNum(i)%>-<%=SWkNum(i)%></td>
						<% Else %>
						<td align="left"><%=SPageNum(i)%></td>
						<% End If %>
						<td align="right" width="85"><%=FormatNumber(SCount(i),0)%> </td>
					</tr>
					<% Next %>
					<tr>
					<td colspan=3 align="Center">���v</td>
					<td align="right" width="85"><%=FormatNumber(iSTSum,0)%> </td>
					</tr>
					</table>
					<% If iMTSum<>0 or iTSum<>0 or iSTSum<>0 Then %>
					<BR>
					<tr align=right><td>
					<table border="1" cellpadding="5" >
					<tr align=right>
					<td colspan=3 align="Center">�����v</td>
					<td align="right" width="100"><%=FormatNumber((iMTSum+iTSum+iSTSum),0)%> </td>
					</tr>
					</table>
					</td></tr>
					<% End If %>
				<% End If %>
				
				
				<!--2013-09-26 Y.TAKAKUWA Add-E-->
			<% Else %>
				<br><div align="center">�f�[�^��1��������܂���B</div><br>
			<% End If %>
			</td>
		</tr>
		<% If LineNo>0 Then %>
		<tr align=Center>
			<td>
			<BR>
			<form action="logcsvout.asp"><input type="submit" value="CSV�t�@�C���o��">
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