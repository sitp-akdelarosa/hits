<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:sst201.asp				_/
'_/	Function	:�X�e�[�^�X�z�M�˗��V�K�o�^			_/
'_/	Date			:2004/01/15				_/
'_/	Code By		:aspLand HARA			_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTP�R���e���c�^�C�v�ݒ�
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->
<%
	'''�Z�b�V�����̗L�������`�F�b�N
	CheckLoginH

	'''�T�[�o���t�̎擾
	Dim DayTime
	DayTime = Now()

	'''�o�^���܂�����ʂɂāu�ŐV�̏��ɍX�V�v��Submit���ꂽ�ꍇ�̑΍�
	if Session.Contents("InsertSubmitted")="False" then

	'''�f�[�^�擾
	Dim USER, KIND, NUMBER, ErrCode
	USER   = UCase(Session.Contents("userid"))
	KIND = Request.Form("ContORBL")
	NUMBER = Request.Form("ContBLNo")
	ErrCode = 0

	'''�G���[�g���b�v�J�n
	on error resume next
	'''DB�ڑ�
	Dim ObjConn, ObjRS, StrSQL
	ConnDBH ObjConn, ObjRS

	'''�w��R���e�i�ԍ�,�a�k�ԍ��̑��݃`�F�b�N
	'''�b�x���o�ςł��P�O���ȓ��ł���Γo�^�Ƃ���(2004.2.16�d�l�ύX)�B
	'''�a�k�ԍ��w��̏ꍇ�A�b�x���o����Ă��Ȃ����A���o����Ă��Ă��P�O���ȓ��̃R���e�i�̂ݓo�^�ΏۂƂ���(2004.2.16�d�l�ύX)�B
	Dim Num, Num2, ArrayContNo, i
	if KIND = 1 then	'''�R���e�i�ԍ��w��
		StrSQL = "SELECT ContNo, BLNo FROM ImportCont "
		StrSQL = StrSQL & " WHERE ContNo='"& NUMBER &"'"
		StrSQL = StrSQL & " AND UpdtTime = (SELECT max(UpdtTime) FROM ImportCont WHERE ContNo='"& NUMBER &"') "
		StrSQL = StrSQL & " AND (CYDelTime is NULL "
		StrSQL = StrSQL & " OR (CYDelTime is not NULL "
		StrSQL = StrSQL & " AND DATEDIFF(d,CYDelTime,GETDATE()) >= 0 "
		StrSQL = StrSQL & " AND DATEDIFF(d,CYDelTime,GETDATE()) <= 10)) "
	elseif KIND = 2 then	'''�a�k�ԍ��w��
		StrSQL = "SELECT BL.BLNo, IC.ContNo FROM BL, ImportCont IC "
		StrSQL = StrSQL & " WHERE BL.BLNo='"& NUMBER &"'"
		StrSQL = StrSQL & " AND BL.VslCode = IC.VslCode "
		StrSQL = StrSQL & " AND BL.VoyCtrl = IC.VoyCtrl "
		StrSQL = StrSQL & " AND BL.BLNo = IC.BLNo "
		StrSQL = StrSQL & " AND BL.UpdtTime = (SELECT max(BL.UpdtTime) FROM BL WHERE BL.BLNo='"& NUMBER &"') "
		StrSQL = StrSQL & " AND (IC.CYDelTime is NULL "
		StrSQL = StrSQL & " OR (IC.CYDelTime is not NULL "
		StrSQL = StrSQL & " AND DATEDIFF(d,IC.CYDelTime,GETDATE()) >= 0 "
		StrSQL = StrSQL & " AND DATEDIFF(d,IC.CYDelTime,GETDATE()) <= 10)) "
	else
		response.write("KIND error!")
		response.end
	end if

	ObjRS.Open StrSQL, ObjConn, 3, 1
	if err <> 0 then
		'''DB�ؒf
		DisConnDBH ObjConn, ObjRS
		jumpErrorP "1","c102","01","�X�e�[�^�X�z�M�˗��V�K�o�^","101","SQL:<BR>"&strSQL
	end if
	Num = ObjRS.RecordCount

	if KIND=2 then		'''�a�k�ԍ��w��̏ꍇ�A�R�t���Ă���R���e�i�ԍ���ϐ��Ɋi�[
		ReDim ArrayContNo(Num)
		for i=0 to Num-1
			ArrayContNo(i) = ObjRS("ContNo")
			ObjRS.MoveNext
		next
	end if

	if KIND=1 then
		if Num > 0 then
			if Trim(ObjRS("BLNo")) = "" then
				'''�R���e�i�ԍ��̓Z�b�g����Ă��邪�a�k�ԍ����Z�b�g����Ă��Ȃ����R�[�h���w�肳�ꂽ�ꍇ
				Response.Write("<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.0 Transitional//EN'>")
				Response.Write("<HTML>")
				Response.Write("<HEAD>")
				Response.Write("<LINK REL='stylesheet' TYPE='text/css' HREF='./style.css'>")
				Response.Write("<TITLE>�X�e�[�^�X�z�M�˗��V�K�o�^</TITLE>")
				Response.Write("<META content='text/html; charset=Shift_JIS' http-equiv=Content-Type>")
				Response.Write("<SCRIPT Language='JavaScript'>")
				Response.Write("<!--")
				Response.Write("function CloseWin(){")
				Response.Write("try{")
				Response.Write("window.opener.parent.List.location.href='sst100F.asp'")
				Response.Write("}catch(e){}")
				Response.Write("window.close();")
				Response.Write("}")
				Response.Write("// -->")
				Response.Write("</SCRIPT>")
				Response.Write("<META content='MSHTML 5.00.2919.6307' name=GENERATOR></HEAD>")
				Response.Write("<BODY leftMargin='0' topMargin='0' marginheight='0' marginwidth='0'>")
				Response.Write("<TABLE border='0' cellPadding='5' cellSpacing='0' width='100%'>")
				Response.Write("<FORM name='sst201'>")
				Response.Write("<TR><TD>�@</TD></TR>")
				Response.Write("<TR>")
				Response.Write("<TD align='center'>")
				Response.Write("�w�肳�ꂽ�R���e�i�̂a�k�ԍ����A���R���e�i�e�[�u����<BR>�ݒ肳��Ă��܂���B<BR><BR><BR>")
				Response.Write("<INPUT type='button' value='����' onClick='javascript:window.close();'>")
				Response.Write("</TD>")
				Response.Write("</TR>")
				Response.Write("</FORM>")
				Response.Write("</TABLE>")
				Response.Write("</BODY>")
				Response.Write("</HTML>")
				ObjRS.close
				Response.end
			end if
		end if
	end if
	ObjRS.close


	if Num > 0 then  '''�w�肳�ꂽ�R���e�i�ԍ��܂��͂a�k�ԍ����A���R���e�i�e�[�u���ɑ��݂���ꍇ
		'''�w�肳�ꂽ�R���e�i�ԍ��܂��͂a�k�ԍ��𓯂����[�U�����łɎw�肵�Ă��邩�ǂ����̃`�F�b�N
		if KIND = 1 then '''�R���e�i�ԍ��w��
			StrSQL = "SELECT ContNo FROM TargetContainers "
			StrSQL = StrSQL & " WHERE UserCode='"& USER &"' AND Process='R' AND ContNo='" & NUMBER & "'"
			StrSQL = StrSQL & " AND BLNo is NULL"
		elseif KIND = 2 then	'''�a�k�ԍ��w��̏ꍇ
			StrSQL = "SELECT BLNo FROM TargetContainers "
			StrSQL = StrSQL & " WHERE UserCode='"& USER &"' AND Process='R' AND BLNo='" & NUMBER & "'"
		else
			response.write("KIND error!")
			response.end
		end if
		ObjRS.Open StrSQL, ObjConn, 3, 1
		if err <> 0 then
			'''DB�ؒf
			DisConnDBH ObjConn, ObjRS
			jumpErrorP "2","c102","01","�X�e�[�^�X�z�M�˗��V�K�o�^","101","SQL:<BR>"&strSQL
		end if
		Num2 = ObjRS.RecordCount
		ObjRS.close

		if Num2 > 0 then		'''���łɓ������[�U�������R���e�i�ԍ��A�a�k�ԍ���o�^���Ă���
			ErrCode = 1
		else

		'''�f�[�^�o�^
			if KIND = 1 then		''''�R���e�i�ԍ��w��̏ꍇ
				StrSQL = "INSERT INTO TargetContainers (UserCode, UpdtTime, UpdtPgCd, UpdtTmnl, RegisterDate, Process, "
				StrSQL =  StrSQL & "ContNo, BLNo, LatestSentTime, "
				StrSQL =  StrSQL & "FlagETA, FlagTA, FlagInTime, FlagList, FlagDOStatus, FlagDelPermit, "
				StrSQL =  StrSQL & "FlagDemurrageFreeTime, FlagCYDelTime, FlagDetentionFreeTime, FlagReturnTime, "
				StrSQL =  StrSQL & "ETA, TA, InTime, ListDate, DOStatus, PreDelPermitFlag, DelPermitDate, DemurrageFreeTime, "
				StrSQL =  StrSQL & "CYDelTime, DetentionFreeTime, ReturnTime) "
				StrSQL =  StrSQL & "values ('" & USER & "','" & DayTime & "','STATUS01','" & USER & "','" & DayTime & "','R',"
				StrSQL =  StrSQL & "'" & NUMBER & "',Null, Null,"
				StrSQL =  StrSQL & "'0','0','0','0','0','0',"
				StrSQL =  StrSQL & "'0','0','0','0',"
				StrSQL =  StrSQL & "Null,Null,Null,Null,Null,'N',Null,Null,"
				StrSQL =  StrSQL & "Null,Null,Null)"
				ObjConn.Execute(StrSQL)
				if err <> 0 then
					Set ObjRS = Nothing
					jumpErrorPDB ObjConn,"1","c102","01","�X�e�[�^�X�z�M�˗��V�K�o�^","103","SQL:<BR>"&StrSQL
				end if

			elseif KIND = 2 then
				for i=0 to Num-1
					StrSQL = "INSERT INTO TargetContainers (UserCode, UpdtTime, UpdtPgCd, UpdtTmnl, RegisterDate, Process, "
					StrSQL =  StrSQL & "ContNo, BLNo, LatestSentTime, "
					StrSQL =  StrSQL & "FlagETA, FlagTA, FlagInTime, FlagList, FlagDOStatus, FlagDelPermit, "
					StrSQL =  StrSQL & "FlagDemurrageFreeTime, FlagCYDelTime, FlagDetentionFreeTime, FlagReturnTime, "
					StrSQL =  StrSQL & "ETA, TA, InTime, ListDate, DOStatus, PreDelPermitFlag, DelPermitDate, DemurrageFreeTime, "
					StrSQL =  StrSQL & "CYDelTime, DetentionFreeTime, ReturnTime) "
					StrSQL =  StrSQL & "values ('" & USER & "','" & DayTime & "','STATUS01','" & USER & "','" & DayTime & "','R',"
					StrSQL =  StrSQL & "'" & ArrayContNo(i) & "', '" & NUMBER & "', Null,"
					StrSQL =  StrSQL & "'0','0','0','0','0','0',"
					StrSQL =  StrSQL & "'0','0','0','0',"
					StrSQL =  StrSQL & "Null,Null,Null,Null,Null,'N',Null,Null,"
					StrSQL =  StrSQL & "Null,Null,Null)"

					ObjConn.Execute(StrSQL)
					if err <> 0 then
						Set ObjRS = Nothing
						jumpErrorPDB ObjConn,"1","c102","01","�X�e�[�^�X�z�M�˗��V�K�o�^","103","SQL:<BR>"&StrSQL
					end if
				next
			else
				response.write("KIND error!")
				response.end
			end if

			'''���O�o��
			WriteLogH "c102", "�X�e�[�^�X�z�M�˗��V�K�o�^","01",""
			ObjRS.close
		end if

	else		'''�w�肳�ꂽ�R���e�i�ԍ��A�a�k�ԍ������݂��Ȃ�
		ErrCode = 9
	end if

	'''DB�ڑ�����
	DisConnDBH ObjConn, ObjRS
	'''�G���[�g���b�v����
	on error goto 0

	Session.Contents("InsertSubmitted") = "True"

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>�X�e�[�^�X�z�M�˗��V�K�o�^</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT Language="JavaScript">
<!--
function CloseWin(){
	try{
		window.opener.parent.List.location.href="sst100F.asp"
	}catch(e){}
	window.close();
}
// -->
</SCRIPT>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin="0" topMargin="0" marginheight="0" marginwidth="0">
<!-------------�X�e�[�^�X�z�M�˗��V�K�o�^--------------------------->
<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%">
<FORM name="sst201">
	<TR><TD>�@</TD></TR>
<% if ErrCode=0 then %>
	<TR>
		<TD align="center">
			�o�^���܂����B<BR><BR><BR>
			<INPUT type="button" value="����" onClick="CloseWin()">
		</TD>
	</TR>
<% elseif ErrCode=1 then %>
	<TR>
		<TD align="center">
			�w�肳�ꂽ�R���e�i�ԍ��܂��͂a�k�ԍ��͂��łɓo�^����Ă��܂��B<BR><BR><BR>
			<INPUT type="button" value="����" onClick="javascript:window.close();">
		</TD>
	</TR>
<% elseif ErrCode=9 then %>
	<TR>
		<TD align="center">
			�w�肳�ꂽ�R���e�i�ԍ��܂��͂a�k�ԍ��͑��݂��Ȃ����A<BR>
			���o��P�P���ȏ�o�߂��Ă��邽�ߓo�^�ł��܂���B<BR><BR><BR>
			<INPUT type="button" value="����" onClick="javascript:window.close();">
		</TD>
	</TR>
<% end if %>
</FORM>
</TABLE>
<!-------------��ʏI���--------------------------->
</BODY>
</HTML>

<%'''if Session.Contents("InsertSubmitted")="False"��else���� %>
<% else %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>�X�e�[�^�X�z�M�˗��V�K�o�^</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT Language="JavaScript">
<!--
function CloseWin(){
	try{
		window.opener.parent.List.location.href="sst100F.asp"
	}catch(e){}
	window.close();
}
// -->
</SCRIPT>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin="0" topMargin="0" marginheight="0" marginwidth="0">
<!-------------�X�e�[�^�X�z�M�˗��V�K�o�^--------------------------->
<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%">
<FORM name="sst201">
	<TR><TD>�@</TD></TR>
	<TR>
		<TD align="center">
			�o�^�͂��łɊ������Ă��܂��B<BR><BR><BR>
			<INPUT type="button" value="����" onClick="CloseWin()">
		</TD>
	</TR>
</FORM>
</TABLE>
<!-------------��ʏI���--------------------------->
</BODY>
</HTML>
<%'''if Session.Contents("InsertSubmitted")="False"��endif���� %>
<% end if %>
