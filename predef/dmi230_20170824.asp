<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi230.asp				_/
'_/	Function	:���O����o���͊m�F���			_/
'_/	Date		:2003/05/28				_/
'_/	Code By		:SEIKO Electric.Co ��d			_/
'_/	Modify		:C-002	2003/08/06	���l���ǉ�	_/
'_/	Modify		:3th	2003/01/31	3���S�ʉ��C	_/
'_/	Modify		:2017/05/09			�s�����P�O�s�ɕύX	_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTP�R���e���c�^�C�v�ݒ�
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->
<!--#include File="CommonFunc.inc"-->
<%
'�Z�b�V�����̗L�������`�F�b�N
  CheckLoginH

'�f�[�^����
  dim BookNo, COMPcd0, COMPcd1,Mord, ret, ErrerM,i
  dim WkOutFlag,Pcool, OutStyle							'2016/08/25 H.Yoshikawa Add
  dim PickPlace(), Terminal()							'2016/09/07 H.Yoshikawa Add			2017/05/09 H.Yoshikawa Upd(4 �� �Ȃ�)
  dim WarningM											'2016/10/27 H.Yoshikawa Add

  Const RowNum = 10										'2017/05/09 H.Yoshikawa Add
  Redim PickPlace(RowNum-1)								'2017/05/09 H.Yoshikawa Add
  Redim Terminal(RowNum-1)								'2017/05/09 H.Yoshikawa Add

  BookNo = Trim(Request("BookNo"))
  COMPcd0 = Request("COMPcd0")
  COMPcd1 = Request("COMPcd1")
  Mord    = Request("Mord")
  ret = true
  ErrerM = ""
  WarningM = ""											'2016/10/27 H.Yoshikawa Add
'�G���[�g���b�v�J�n
  on error resume next
'DB�ڑ�
  dim ObjConn, ObjRS, StrSQL
  ConnDBH ObjConn, ObjRS

'�u�b�N�̏d���o�^�`�F�b�N
  dim strCodes,dummy1, dummy2
  If Mord=0 OR (Mord=1 AND COMPcd1 <> Request("oldCOMPcd1")) Then
'2006/03/06 mod-s h.matsuda(SQL�����č\�z)
'    checkSPBook ObjConn, ObjRS, BookNo,COMPcd0,COMPcd1,strCodes,dummy1, dummy2, ret
    checkSPBook2 ObjConn, ObjRS, BookNo,COMPcd0,COMPcd1,strCodes,dummy1, dummy2, ret
'2006/03/06 mod-e h.matsuda
    If Not ret Then
      ErrerM="�w�肵���u�b�L���ONo�͎w����u"& Left(strCodes,Len(strCodes)-1) &"�v�Ŋ��ɓo�^����Ă��܂��B"
    End If
  End If
  if err <> 0 then
    DisConnDBH ObjConn, ObjRS	'DB�ؒf
    jampErrerP "2","b303","01","�u�b�L���O�w���e�[�u��","101","SQL�F<BR>"&StrSQL
  end if
  If (ret) Then
   '�w�b�hID�̃`�F�b�N
    dim CMPcd
    'CW-327 Change
    'CMPcd = Array("",COMPcd1,"","","")
    CMPcd = Array("",Trim(COMPcd1),"","","")
    checkHdCd ObjConn, ObjRS, CMPcd, ret
    If (ret) Then
    Else
      ErrerM="�w�肳�ꂽ��ЃR�[�h�͑��݂��܂���B"
    End If
  End If

'�u�b�N�̔��o�����`�F�b�N
  If ret Then
    dim cmpNum
    StrSQL = "SELECT Count(EXC.BookNo) AS numB, Count(Pic.Qty) AS numQ "&_
             "FROM ExportCont AS EXC INNER JOIN Pickup AS Pic ON (EXC.VslCode = Pic.VslCode) "&_
             "AND (EXC.VoyCtrl = Pic.VoyCtrl) AND (EXC.BookNo = Pic.BookNo) "&_
             "WHERE EXC.BookNo='"& BookNo &"' AND EmpDelTime IS NOT NULL"
    ObjRS.Open StrSQL, ObjConn
    if err <> 0 then
      DisConnDBH ObjConn, ObjRS
      jampErrerP "2","b303","01","����o�F���o�����`�F�b�N","101","SQL:<BR>"&strSQL
    end if
    cmpNum=ObjRS("numB")
    If ObjRS("numQ")<>0 Then
      ObjRS.close
      StrSQL = "SELECT Pic.Qty "&_
               "FROM ExportCont AS EXC INNER JOIN Pickup AS Pic ON (EXC.VslCode = Pic.VslCode) "&_
               "AND (EXC.VoyCtrl = Pic.VoyCtrl) AND (EXC.BookNo = Pic.BookNo) "&_
               "WHERE EXC.BookNo='"& BookNo &"' GROUP BY Pic.Qty"
      ObjRS.Open StrSQL, ObjConn
      if err <> 0 then
        DisConnDBH ObjConn, ObjRS
        jampErrerP "2","b303","01","����o�F���o�����`�F�b�N","101","SQL:<BR>"&strSQL
      end if
      If cmpNum = ObjRS("Qty") Then
        WarningM="<����>�w��̃u�b�L���O�ԍ��͔��o���������Ă��܂��B<BR>"
      End If
    End If
    ObjRS.close
  End If
  
  '2016/09/07 H.Yoshikawa Add Start
  If ret Then
 	dim OutNum, OdrNum			'2016/10/26 H.Yoshikawa Add
  	dim SizeChk					'2017/05/10 H.Yoshikawa Add
  	
 	''�{���`�F�b�N
  	For i=0 To RowNum-1			'2017/05/09 H.Yoshikawa Upd(4��RowNum-1)
  		PickPlace(i) = gfTrim(Request("PickPlace" & i))
  		Terminal(i) = gfTrim(Request("Terminal" & i))
		if gfTrim(Request("UpdFlag" & i)) = "1" then
			'2016/10/12 H.Yoshikawa Add Start �i�������������̖̂{�������Z����j
			Dim Sz, Tp, Ht, Qty, j
			Sz = gfTrim(Request("ContSize" & i))
			Tp = gfTrim(Request("ContType" & i))
			Ht = gfTrim(Request("ContHeight" & i))
			Qty = CInt(Request("PickNum" & i))
			for j=0 To RowNum-1						'2017/05/09 H.Yoshikawa Upd(4��RowNum-1)
				if i<>j then
					if gfTrim(Request("DelFlag" & i)) <> "1" then	'2017/05/10 H.Yoshikawa Add
						if gfTrim(Request("ContSize" & j)) = Sz and gfTrim(Request("ContType" & j)) = Tp and gfTrim(Request("ContHeight" & j)) = Ht then
							Qty = Qty + CInt(Request("PickNum" & j))
						end if 
					end if											'2017/05/10 H.Yoshikawa Add
				end if
			next
			'2016/10/12 H.Yoshikawa Add End
			
			'�s�b�N�A�b�v�ꏊ�擾
			StrSQL = "SELECT * FROM Pickup  "
			StrSQL = StrSQL & "WHERE VslCode    = '" & gfSQLEncode(Request("VslCode")) & "'"
			StrSQL = StrSQL & "  AND VoyCtrl    = '" & gfSQLEncode(Request("VoyCtrl")) & "'"
			StrSQL = StrSQL & "  AND BookNo     = '" & gfSQLEncode(BookNo) & "'"
			StrSQL = StrSQL & "  AND ContSize   = '" & gfSQLEncode(Request("ContSize" & i)) & "'"
			StrSQL = StrSQL & "  AND ContType   = '" & gfSQLEncode(Request("ContType" & i)) & "'"
			StrSQL = StrSQL & "  AND ContHeight = '" & gfSQLEncode(Request("ContHeight" & i)) & "'"
			StrSQL = StrSQL & " ORDER BY Qty desc "
		    ObjRS.Open StrSQL, ObjConn
		    if err <> 0 then
		      DisConnDBH ObjConn, ObjRS
		      jampErrerP "2","b303","01","����o�F�{���`�F�b�N","101","SQL:<BR>"&strSQL & "<BR>" & err.description
		    end if
			if ObjRS.eof then
				WarningM=WarningM & "<����>����o�I�[�_�[���o�^����Ă��܂���B�i" & i + 1 & "�s�ځj<BR>"
				PickPlace(i) = ""
				Terminal(i) = ""
			else
				'2016/10/26 H.Yoshikawa Del Start
				'if Qty > CInt(ObjRS("Qty")) then
				'	ret = false
				'	ErrerM=ErrerM & "���͂��ꂽ�{�����A����o�I�[�_�[�{���𒴂��Ă��܂��B�i" & i + 1 & "�s�ځj<BR>"
				'end if
				'2016/10/26 H.Yoshikawa Del Start
				PickPlace(i) = gfTrim(ObjRS("PickPlace"))
				Terminal(i) = gfTrim(ObjRS("Terminal"))
			end if
			ObjRS.close

			'2016/10/26 H.Yoshikawa Add Start
			'�����[�U�o�^�̗\��{�������Z
			StrSQL = "SELECT ISNULL(Sum(Qty1), 0) as NumCont FROM BookingAssign "
			StrSQL = StrSQL & "WHERE VslCode    = '" & gfSQLEncode(Request("VslCode")) & "'"
			StrSQL = StrSQL & "  AND Voyage     = '" & gfSQLEncode(Request("VoyCtrl")) & "'"
			StrSQL = StrSQL & "  AND BookNo     = '" & gfSQLEncode(BookNo) & "'"
			StrSQL = StrSQL & "  AND ContSize1   = '" & gfSQLEncode(Request("ContSize" & i)) & "'"
			StrSQL = StrSQL & "  AND ContType1   = '" & gfSQLEncode(Request("ContType" & i)) & "'"
			StrSQL = StrSQL & "  AND ContHeight1 = '" & gfSQLEncode(Request("ContHeight" & i)) & "'"
			StrSQL = StrSQL & "  AND SenderCode <> '" & gfSQLEncode(COMPcd0) & "'"
			StrSQL = StrSQL & "  AND Process     = 'R'"
			ObjRS.Open StrSQL, ObjConn
			if err <> 0 then
				DisConnDBH ObjConn, ObjRS
				jampErrerP "1","b303","01","����o�F�ʃ��[�U�\��{���擾","101","SQL:<BR>"&strSQL
			end if
			if not ObjRS.eof then
				Qty = Qty + CInt(ObjRS("NumCont"))
			end if
			ObjRS.close

			
			'���ꑮ���̃I�[�_�[�{�����擾
			StrSQL = "SELECT ISNULL(Sum(Qty), 0) as NumQty FROM PickUp "
			StrSQL = StrSQL & "WHERE VslCode    = '" & gfSQLEncode(Request("VslCode")) & "'"
			StrSQL = StrSQL & "  AND VoyCtrl    = '" & gfSQLEncode(Request("VoyCtrl")) & "'"
			StrSQL = StrSQL & "  AND BookNo     = '" & gfSQLEncode(BookNo) & "'"
			StrSQL = StrSQL & "  AND ContSize   = '" & gfSQLEncode(Request("ContSize" & i)) & "'"
			StrSQL = StrSQL & "  AND ContType   = '" & gfSQLEncode(Request("ContType" & i)) & "'"
			StrSQL = StrSQL & "  AND ContHeight = '" & gfSQLEncode(Request("ContHeight" & i)) & "'"
			ObjRS.Open StrSQL, ObjConn
			if err <> 0 then
				DisConnDBH ObjConn, ObjRS
				jampErrerP "1","b303","01","����o�F�I�[�_�[�{���擾","101","SQL:<BR>"&strSQL
			end if
			if not ObjRS.eof then
				OdrNum=CInt(ObjRS("NumQty"))
			end if
			ObjRS.close
			if OdrNum > 0 then
				if Qty > OdrNum then
					ret = false
					ErrerM=ErrerM & "���͂��ꂽ�����̖{�����v���A����o�I�[�_�[�{���𒴂��Ă��܂��B�i" & i + 1 & "�s�ځj<BR>"
				end if
			end if

			'���ꑮ���̔��o�ςݖ{�����擾
			if Qty < CInt(Request("OutNum" & i)) then
				ret = false
				ErrerM=ErrerM & "���͂��ꂽ�����̖{�����v���A���o�ςݖ{����������Ă��܂��B�i" & i + 1 & "�s�ځj<BR>"
			end if
			'2016/10/26 H.Yoshikawa Add End
			
			'2017/05/10 H.Yoshikawa Add Start
			'�T�C�Y�^�^�C�v�^�n�C�g�̑g�������}�X�^�ɑ��݂��邩�`�F�b�N
			SizeChk = 0
			StrSQL = "SELECT Count(*) AS CNT FROM ViewkMSizeTypeHeight "
			StrSQL = StrSQL & "WHERE ContSize   = '" & gfSQLEncode(Request("ContSize" & i)) & "'"
			StrSQL = StrSQL & "  AND ContType   = '" & gfSQLEncode(Request("ContType" & i)) & "'"
			StrSQL = StrSQL & "  AND ContHeight = '" & gfSQLEncode(Request("ContHeight" & i)) & "'"
			ObjRS.Open StrSQL, ObjConn
			if err <> 0 then
				DisConnDBH ObjConn, ObjRS
				jampErrerP "1","b303","01","����o�F�T�C�Y�^�C�v�n�C�g�}�X�^�擾","101","SQL:<BR>"&strSQL
			end if
			if not ObjRS.eof then
				SizeChk=CInt(ObjRS("CNT"))
			end if
			ObjRS.close
			if SizeChk <= 0 then
				ret = false
				ErrerM=ErrerM & "���͂��ꂽ�������T�C�Y�^�C�v�n�C�g�}�X�^�ɓo�^����Ă��܂���B�i" & i + 1 & "�s�ځj<BR>"
			end if
			'2017/05/10 H.Yoshikawa Add End
		end if
		
		'2017/05/10 H.Yoshikawa Add Start
		if gfTrim(Request("DelFlag" & i)) = "1" then
			if CInt(Request("OutNum" & i)) > 0 then
				ret = false
				ErrerM=ErrerM & "���ɔ��o�ς݂̃R���e�i�����邽�߁A�s�폜�ł��܂���B�i" & i + 1 & "�s�ځj<BR>"
			end if
		end if
		'2017/05/10 H.Yoshikawa Add End
	Next
  End If
  '2016/09/07 H.Yoshikawa Add End
  
  '2016/10/27 H.Yoshikawa Add Start
  if ErrerM <> "" then
  	ErrerM = ErrerM & "<BR>�u�߂�v�{�^�����������A�ē��͂��Ă��������B"
  end if
  '2016/10/27 H.Yoshikawa Add End
  
'DB�ڑ�����
  DisConnDBH ObjConn, ObjRS
'�G���[�g���b�v����
  on error goto 0

  dim tmpstr
  If ret Then
    tmpstr=",���͓��e�̐���:0(������)"
  Else
    tmpstr=",���͓��e�̐���:1(���)"
  End If
  If Request("Mord")=0 Then
    WriteLogH "b302", "����o���O������","02",BookNo&"/"&COMPcd1&tmpstr
  Else
    WriteLogH "b302", "����o���O������","13",BookNo&"/"&COMPcd1&tmpstr
  End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>��o���s�b�N�����͊m�F</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--

function setParam(target){
//  window.resizeTo(500,260);
  bgset(target);
}

//�o�^
function GoEntry(printFlag){
  target=document.dmi230F;
  target.SijiF.value=printFlag
  target.action="./dmi240.asp";
  target.submit();
}
//�߂�
function GoBackT(){
  target=document.dmi230F;
  target.action="./dmi220.asp";
  target.submit();
}
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="setParam(document.dmi230F)">
<!-------------����o�����͊m�F���--------------------------->
<FORM name="dmi230F" method="POST">
<TABLE border=0 cellPadding=3 cellSpacing=0 width="100%">
  <TR>
    <TD colspan=2>
      <B>��o���s�b�N�����͊m�F</B></TD></TR>
  <TR>
    <TD><DIV class=bgb>�u�b�L���O�m���D</DIV></TD>
    <TD><INPUT type=text name="BookNoM" value="<%=Request("BookNoM")%>" readOnly size=40>
        <INPUT type=hidden name="BookNo" value="<%=Request("BookNo")%>"></TD></TR>
  <TR>
    <TD><DIV class=bgb>�D��</DIV></TD>
    <TD><INPUT type=text name="shipFact" value="<%=Request("shipFact")%>" readOnly size=40></TD></TR>
  <TR>
    <TD><DIV class=bgb>*�D��</DIV></TD>
    <TD><INPUT type=text name="shipName" value="<%=Request("shipName")%>" readOnly size=40>
    	<INPUT type=hidden name="VslCode" value="<%=Request("VslCode")%>">							<!-- 2016/08/23 H.Yoshikawa Add -->
    </TD></TR>
  <TR>
  	<!-- 2016/08/23 H.Yoshikawa Upd Start -->
    <!-- <TD><DIV class=bgb>�d���n</DIV></TD>
    <TD><INPUT type=text name="delivTo" value="<%=Request("delivTo")%>" readOnly size=40></TD></TR> -->
    <TD><DIV class=bgb>*Voyage</DIV></TD>
    <TD><INPUT type=hidden name="delivTo" value="<%=Request("delivTo")%>">
    	<INPUT type=text name="ExVoyage" value="<%=Request("ExVoyage")%>" readOnly size=12>			<!-- 2016/08/23 H.Yoshikawa Add -->
    	<INPUT type=hidden name="VoyCtrl" value="<%=Request("VoyCtrl")%>">							<!-- 2016/10/17 H.Yoshikawa Upd(text��hidden) -->
    </TD></TR>
  	<!-- 2016/08/23 H.Yoshikawa Upd End -->
  <TR>
    <TD><DIV class=bgb>��ЃR�[�h(���^)</DIV></TD>
    <TD><INPUT type=text name="COMPcd1" value="<%=COMPcd1%>" size=5  readOnly>
        <INPUT type=hidden name="oldCOMPcd1" value="<%=Request("oldCOMPcd1")%>"></TD></TR>
  <TR>
    <TD><DIV class=bgb>�����Ɩ{��</DIV></TD>
    <TD></TD></TR>
  <TR>
    <TD colspan=2>
    <TABLE border=0 cellPadding=0 cellSpacing=0 width=900 align=center>
    <!-- 2016/08/23 H.Yoshikawa Upd Start -->
    <!-- <TR><TD></TD><TD>�T�C�Y</TD><TD>�^�C�v</TD><TD>����</TD><TD>�ގ�</TD><TD>�s�b�N�ꏊ</TD><TD></TD><TD>�{��</TD></TR> -->
    <TR>
    	<TD></TD>
    	<TD>*�T�C�Y</TD>
    	<TD>*�^�C�v</TD>
    	<TD>*����</TD>
    	<TD>�ݒ艷�x</TD>
    	<TD>�v���N�[��</TD>
    	<TD>�x���`���[�V����</TD>
    	<TD>*�s�b�N�\�����(���Ԃ���ڸ�َ��̂ݕK�{)</TD>
    	<TD>�@*�{��</TD>
    	<TD>���o��</TD>
    	<TD>�s�b�N�A�b�v�ꏊ</TD>
    	<TD>�ύX</TD>
    	<TD>�s�폜</TD>									<!-- 2017/05/10 H.Yoshikwawa Add -->
    </TR>
    <!-- 2016/08/23 H.Yoshikawa Upd End -->
<% For i=0 To RowNum-1%>						<!-- 2017/05/09 H.Yoshikawa Upd(4��RowNum-1) -->
      <TR><TD>(<%=i+1%>)</TD>
          <TD><INPUT type=text name="ContSize<%=i%>"   value="<%=Request("ContSize"&i)%>" size=4  readOnly></TD>
          <TD><INPUT type=text name="ContType<%=i%>"   value="<%=Request("ContType"&i)%>" size=4  readOnly></TD>
          <TD><INPUT type=text name="ContHeight<%=i%>" value="<%=Request("ContHeight"&i)%>" size=4  readOnly></TD>
      <!-- 2016/08/23 H.Yoshikawa Upd Start
          <TD><INPUT type=text name="Material<%=i%>"   value="<%=Request("Material"&i)%>"   size=4  readOnly></TD>
          <TD><INPUT type=text name="PickPlace<%=i%>"  value="<%=Request("PickPlace"&i)%>"  size=25 readOnly></TD>
          <TD>�E�E�E</TD>
          <TD><INPUT type=text name="PickNum<%=i%>" value="<%=Request("PickNum"&i)%>" size=4  readOnly></TD></TR> -->
          <TD><INPUT type=text name="SetTemp<%=i%>"  value="<%=Request("SetTemp"&i)%>" size=8 readOnly>��</TD>
          <TD>
          	<%	if gfTrim(Request("Pcool"&i)) = "" then 
          			Pcool = gfTrim(Request("Bef_Pcool"&i))
          	 	else
          	 		Pcool = gfTrim(Request("Pcool"&i))
          	 	end if
          	%>
          	  <select disabled>
				<option value="0"></option>
				<option value="1" <% if Pcool = "1" then %>selected<% end if %> >�L</option>
			  </select>
              <INPUT type=hidden name="Pcool<%=i%>"  value="<%=Pcool%>" size=5 readOnly>
          </TD>
          <TD><INPUT type=text name="Ventilation<%=i%>"  value="<%=Request("Ventilation"&i)%>" size=5 readOnly>%�i�J���j</TD>
          <TD>
              <INPUT type=text name="PickDate<%=i%>"  value="<%=Request("PickDate"&i)%>" size=15 readOnly>
              <INPUT type=text name="PickHour<%=i%>"  value="<%=Request("PickHour"&i)%>" size=4 readOnly>��
              <INPUT type=text name="PickMinute<%=i%>"  value="<%=Request("PickMinute"&i)%>" size=4 readOnly>��
          </TD>
          <TD>�c<INPUT type=text name="PickNum<%=i%>" value="<%=Request("PickNum"&i)%>" size=4 readOnly></TD>
          <% OutStyle=""
             select case Trim(Request("OutFlag"&i))
               case "0"
                 WkOutFlag = "�m�F��"
               case "1"
                 WkOutFlag = "��"
               case "9"
                 WkOutFlag = "�s��"
                 OutStyle="color:red;"
               case else
                 WkOutFlag = ""
             end select
          %>
          <TD style="<%=OutStyle%>"><INPUT type=hidden name="OutFlag<%=i%>"  value="<%=Request("OutFlag"&i)%>" ><%=WkOutFlag %></TD>
          <TD><INPUT type=hidden name="PickPlace<%=i%>"  value="<%=PickPlace(i)%>"><%=gfHTMLEncode(PickPlace(i))%>
              <INPUT type=hidden name="Terminal<%=i%>"  value="<%=Terminal(i)%>">
          </TD>
          <TD><INPUT type=checkbox value="1" disabled <% if Request("UpdFlag"&i) = "1" then%> checked <% end if %>>
              <INPUT type=hidden name="UpdFlag<%=i%>" value="<%=Request("UpdFlag"&i)%>">
          </TD>
		  <% '2017/05/10 H.Yoshikawa Upd Start %>
          <TD><INPUT type=checkbox value="1" disabled <% if Request("DelFlag"&i) = "1" then%> checked <% end if %>>
              <INPUT type=hidden name="DelFlag<%=i%>" value="<%=Request("DelFlag"&i)%>">
          </TD>
		  <% '2017/05/10 H.Yoshikawa Upd End %>
			<% '2016/10/27 H.Yoshikawa Upd Start %>
			<INPUT type=hidden name="OutNum<%=i%>" value="<%=Request("OutNum"&i)%>">  <!-- 2016/10/26 H.Yoshikawa Add -->
			<INPUT type=hidden name="Bef_ContSize<%=i%>"    value="<%=Request("Bef_ContSize"&i)%>">
			<INPUT type=hidden name="Bef_ContType<%=i%>"    value="<%=Request("Bef_ContType"&i)%>">
			<INPUT type=hidden name="Bef_ContHeight<%=i%>"  value="<%=Request("Bef_ContHeight"&i)%>">
			<INPUT type=hidden name="Bef_SetTemp<%=i%>"     value="<%=Request("Bef_SetTemp"&i)%>">
			<INPUT type=hidden name="Bef_Pcool<%=i%>"       value="<%=Request("Bef_Pcool"&i)%>">
			<INPUT type=hidden name="Bef_Ventilation<%=i%>" value="<%=Request("Bef_Ventilation"&i)%>">
			<INPUT type=hidden name="Bef_PickDate<%=i%>"    value="<%=Request("Bef_PickDate"&i)%>">
			<INPUT type=hidden name="Bef_PickHour<%=i%>"    value="<%=Request("Bef_PickHour"&i)%>">
			<INPUT type=hidden name="Bef_PickMinute<%=i%>"  value="<%=Request("Bef_PickMinute"&i)%>">
			<INPUT type=hidden name="Bef_PickNum<%=i%>"     value="<%=Request("Bef_PickNum"&i)%>">
			<INPUT type=hidden name="Bef_OutFlag<%=i%>"   value="<%=Request("Bef_OutFlag"&i)%>">
			<INPUT type=hidden name="Bef_PickPlace<%=i%>"   value="<%=Request("Bef_PickPlace"&i)%>">
			<INPUT type=hidden name="Bef_Terminal<%=i%>"    value="<%=Request("Bef_Terminal"&i)%>">
			<% '2016/10/27 H.Yoshikawa Upd End %>
	  </TR>
      <!-- 2016/08/23 H.Yoshikawa Upd End -->
<% Next %>
    </TABLE>
    </TD></TR>
  <TR>
    <TD><DIV class=bgb>�o���l�ߓ���</DIV></TD>
    <TD><INPUT type=text name="vanMon" value="<%=Request("vanMon")%>" size=3  readOnly>��
        <INPUT type=text name="vanDay" value="<%=Request("vanDay")%>" size=3  readOnly>��
        <INPUT type=text name="vanHou" value="<%=Request("vanHou")%>" size=3  readOnly>��
        <INPUT type=text name="vanMin" value="<%=Request("vanMin")%>" size=3  readOnly>��
        </TD></TR>
  <TR>
    <TD><DIV class=bgb>�o���l�ߏꏊ�P</DIV></TD>
    <TD><INPUT type=text name="vanPlace1" value="<%=Request("vanPlace1")%>" size=73  readOnly></TD></TR>
  <TR>
    <TD><DIV class=bgb>�o���l�ߏꏊ�Q</DIV></TD>
    <TD><INPUT type=text name="vanPlace2" value="<%=Request("vanPlace2")%>" size=73  readOnly></TD></TR>
  <TR>
    <TD><DIV class=bgb>�i��</DIV></TD>
    <TD><INPUT type=text name="goodsName" value="<%=Request("goodsName")%>" size=30  readOnly></TD></TR>
  <TR>
    <TD><DIV class=bgb>������b�x�D�b�x�J�b�g��</DIV></TD>
    <TD><INPUT type=text name="Terminal" value="<%=Request("Terminal")%>" readOnly>
        <INPUT type=text name="CYCut" value="<%=Request("CYCut")%>" readOnly></TD></TR>
  <TR>
    <TD><DIV class=bgb>���l�P</DIV></TD>
    <TD><INPUT type=text name="Comment1" value="<%=Request("Comment1")%>" size=73  readOnly></TD></TR>
  <TR>
    <TD><DIV class=bgb>���l�Q</DIV></TD>
    <TD><INPUT type=text name="Comment2" value="<%=Request("Comment2")%>" size=73  readOnly></TD></TR>
    
  <TR>
<!-- 2009/03/10 R.Shibuta Add-S -->
  	<TD><DIV class=bgy>*�o�^�S����</DIV></TD>
 	<TD><INPUT type=text name="TruckerSubName" readOnly = "readOnly" value="<%=Request("TruckerSubName")%>" maxlength=16></TD></TR>
<!-- 2009/03/10 R.Shibuta Add-E -->
<!-- 2016/08/23 H.Yoshikawa Add Start -->
  <TR>
  	<TD><DIV class=bgy>*�d�b�ԍ�</DIV></TD>
 	<TD><INPUT type=text name="Tel" value="<%=Request("Tel")%>"  readonly></TD></TR>
  <TR>
  	<TD><DIV class=bgy>*���[���A�h���X</DIV></TD>
 	<TD><INPUT type=text name="Mail" value="<%=Request("Mail")%>" readonly size=60>
 		<INPUT type=checkbox value="1" <% if Request("MailFlag") = "1" then %>checked <% end if %> disabled>
 		���o�ۏ�ԕύX���Ƀ��[�����󂯎��
 		<INPUT type=hidden name="MailFlag" value="<%=Request("MailFlag")%>">
 	</TD></TR>
<!-- 2016/08/23 H.Yoshikawa Add End -->
  <TR>
    <TD colspan=2 align=center>
      <INPUT type=hidden name=Mord value="<%=Request("Mord")%>" >
      <INPUT type=hidden name=COMPcd0 value="<%=COMPcd0%>" >
      <INPUT type=hidden name=Res value="<%=Request("Res")%>" >
      <INPUT type=hidden name=SijiF value="" ><P><BR></P>
      <INPUT type=hidden name=shipline value="<%=Request("shipline")%>" ><%'add h.matsuda%>
<%'2016/08/30 H.Yoshikawa Add Start%>
       <INPUT type=hidden name=compFlag value="<%=Request("compFlag")%>" >
<%'2016/08/30 H.Yoshikawa Add End%>
<% IF ret Then %>
	<% if WarningM <> "" then %>
       <P><DIV class=alert><%=WarningM%></DIV></P>
	<% end if %>
       <INPUT type=button value="�m��" onClick="GoEntry('No')">
<% Else %>
       <P><DIV class=alert><%=ErrerM%></DIV></P>
<% End If %>
       <INPUT type=button value="�߂�" onClick="GoBackT()">
<% IF Mord=0  AND ret Then %>
       <P><INPUT type=button value="�m�聕�w�������" onClick="GoEntry('Yes')"></P>
<% End If %>

    </TD></TR>

</TABLE>
</FORM>
<!-------------��ʏI���--------------------------->
</BODY></HTML>
