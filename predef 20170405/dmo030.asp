<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmo030.asp				_/
'_/	Function	:�����o���ꗗ�W�J���			_/
'_/	Date		:2003/07/23				_/
'_/	Code By		:SEIKO Electric.Co ��d			_/
'_/	Modify		:C-001 2003/07/29	CSV�o�͑Ή�	_/
'_/			:C-002 2003/07/29	���l���Ή�	_/
'_/			:3th   2004/01/31	3���Ή��FHTML�C��_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTP�R���e���c�^�C�v�ݒ�
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->
<%
'�Z�b�V�����̗L�������`�F�b�N
  CheckLoginH
  WriteLogH "b109", "�����o���O���W�J","01",""

'�T�[�o�����̎擾
  dim DayTime,day
  getDayTime DayTime
  day = DayTime(0) & "�N" & DayTime(1) & "��" & DayTime(2) & "��" &_
        DayTime(3) & "��" & DayTime(4) & "�����݂̏��"
'INI�t�@�C�����ݒ�l���擾
  dim param(2)
  getIni param

'�f�[�^�擾
  dim No,Num, preDtTbl,DtTbl,Siji,i,j
  Siji  =Array("","�w�肠��","�w��Ȃ�","�ꗗ","�a�k")
  No=Request("targetNo")
  ReDim preDtTbl(1)

  preDtTbl(0)=Split(Request("Datatbl0"), ",", -1, 1)
  preDtTbl(1)=Split(Request("Datatbl"&No), ",", -1, 1)

  Num=1
  ReDim DtTbl(Num)
  DtTbl(0)=preDtTbl(0)
  DtTbl(0)(5)="�R���e�i�ԍ�"
  '�G���[�g���b�v�J�n
    on error resume next
  'DB�ڑ�
    dim ObjConn, ObjRS, StrSQL
    ConnDBH ObjConn, ObjRS
  '�W�J�f�[�^����
    j=1
'3th     If preDtTbl(1)(4)="1" Then		'�w������
    Select Case preDtTbl(1)(4)
      Case "1"			'�w������
        DtTbl(1)=preDtTbl(1)
        DtTbl(1)(11)="�@"
      Case "2" 			'�w��Ȃ�
        '�Ώێ擾
        StrSQL = "SELECT Cnt.ContNo,Cnt.ContSize, INC2.ReturnTime, INC2.CYDelTime, "&_
                 "NULLIF('-',LEFT(DateDiff(day,getdate(),DateAdd(day,"&preDtTbl(1)(21)-param(0)+1&" ,INC2.CYDelTime))*("&preDtTbl(1)(21)&"%6),1)) AS ReturnArrert "&_
                 "From (ImportCont AS INC1 INNER JOIN ImportCont AS INC2 ON "&_
                 "(INC1.VoyCtrl = INC2.VoyCtrl) AND (INC1.VslCode = INC2.VslCode) AND (INC1.BLNo = INC2.BLNo)) "&_
                 "INNER JOIN Container AS Cnt "&_
                 "ON INC2.ContNo=Cnt.ContNo AND INC2.VslCode=Cnt.VslCode AND INC2.VoyCtrl=Cnt.VoyCtrl "&_
                 "WHERE INC1.ContNo='" & preDtTbl(1)(5) & "' AND INC1.BLNo= '"& preDtTbl(1)(11) &"' " &_
                 "ORDER BY INC2.ContNo ASC, INC2.UpdtTime DESC"
        ObjRS.Open StrSQL, ObjConn
        if err <> 0 then
          DisConnDBH ObjConn, ObjRS	'DB�ؒf
          jampErrerP "1","b109","01","�����o�F�W�J���","101","SQL:<BR>"&StrSQL
        end if
        j=2
        Do Until ObjRS.EOF
          If preDtTbl(1)(5) = Trim(ObjRS("ContNo")) Then
            DtTbl(1)=preDtTbl(1)
            DtTbl(1)(12)=ObjRS("ReturnArrert")
            If IsNull(ObjRS("ReturnTime")) Then
              DtTbl(1)(8)="��"
            Else
              DtTbl(1)(8)="��"
            End If
            DtTbl(1)(17)=Trim(ObjRS("ContSize"))
            DtTbl(1)(20)=Trim(ObjRS("CYDelTime"))
          Else
            ReDim Preserve DtTbl(j)
            DtTbl(j)=preDtTbl(1)
            DtTbl(j)(5)=Trim(ObjRS("ContNo"))
            DtTbl(j)(12)=ObjRS("ReturnArrert")
            If IsNull(ObjRS("ReturnTime")) Then
              DtTbl(j)(8)="��"
            Else
              DtTbl(j)(8)="��"
            End If
            DtTbl(j)(17)=Trim(ObjRS("ContSize"))
            DtTbl(j)(20)=Trim(ObjRS("CYDelTime"))
            j=j+1
          End If
          ObjRS.MoveNext
      Loop
      ObjRS.close
      Num=j-1
'3th    ElseIf preDtTbl(1)(4)="3" Then	'�ꗗ
    Case "3" 			'�ꗗ
        '�Ώی����擾
        StrSQL = "SELECT count(ITF.ContNo) AS CNUM FROM "&_
                 "(hITCommonInfo ITC INNER JOIN hITFullOutSelect ITF ON ITC.WkContrlNo = ITF.WkContrlNo) "&_
                 "INNER JOIN ImportCont IPC ON ITF.ContNo =IPC.ContNo AND ITC.BLNo = IPC.BLNo "&_
                 "WHERE ITC.ContNo='"&preDtTbl(1)(5)&"' AND ITC.WkNo='"&preDtTbl(1)(3)&"' AND Process='R' AND ITC.WkType='1'"
'ADD 20030908 This Line:AND Process='R' AND ITC.WkType='1'
'ADD 20030911 This Item:AND ITC.WkNo='"&preDtTbl(1)(3)&"'
        ObjRS.Open StrSQL, ObjConn
        if err <> 0 then
          DisConnDBH ObjConn, ObjRS	'DB�ؒf
          jampErrerP "1","b109","01","�����o�F�W�J���","101","SQL:<BR>"&StrSQL
        end if
        Num = Num + ObjRS("CNUM")-1
        ObjRS.close
        ReDim Preserve DtTbl(Num)
        '�f�[�^�擾
        StrSQL = "SELECT ITF.ContNo, IPC.ReturnTime, IPC.CYDelTime, "&_
                 "NULLIF('-',LEFT(DateDiff(day,getdate(),DateAdd(day,"&preDtTbl(1)(21)-param(0)+1&" ,IPC.CYDelTime))*("&preDtTbl(1)(21)&"%6),1)) AS ReturnArrert "&_
                 "FROM (hITCommonInfo ITC INNER JOIN hITFullOutSelect ITF ON ITC.WkContrlNo = ITF.WkContrlNo) "&_
                 "INNER JOIN ImportCont IPC ON ITF.ContNo =IPC.ContNo AND ITC.BLNo = IPC.BLNo "&_
                 "WHERE ITC.ContNo='"&preDtTbl(1)(5)&"' AND ITC.WkNo='"&preDtTbl(1)(3)&"' AND Process='R' AND ITC.WkType='1'"
'ADD 20030908 This Line:AND Process='R' AND ITC.WkType='1'
'ADD 20030911 This Item:AND ITC.WkNo='"&preDtTbl(1)(3)&"'
        ObjRS.Open StrSQL, ObjConn
        if err <> 0 then
          DisConnDBH ObjConn, ObjRS	'DB�ؒf
          jampErrerP "1","b109","01","�����o�F�W�J���","101","SQL:<BR>"&StrSQL
        end if
        Do Until ObjRS.EOF
          DtTbl(j)=preDtTbl(1)
          DtTbl(j)(5)=Trim(ObjRS("ContNo"))
          DtTbl(j)(12)=ObjRS("ReturnArrert")
          If IsNull(ObjRS("ReturnTime")) Then
            DtTbl(j)(8)="��"
          Else
            DtTbl(j)(8)="��"
          End If
          ObjRS.MoveNext
          j=j+1
        Loop
        ObjRS.close
'3th      ElseIf preDtTbl(1)(4)="2" Or preDtTbl(1)(4)="4" Then	'�w��Ȃ�,BL
      Case "4"			'BL
        '�Ώی����擾
        dim VslCode,VoyCtrl
        '�Ώ�BL�I��
        StrSQL = "SELECT INC.VslCode, INC.VoyCtrl "&_
                 "From ImportCont AS INC  "&_
                 "Where INC.BLNo= '"& preDtTbl(1)(11) &"' ORDER BY INC.UpdtTime DESC"
        ObjRS.Open StrSQL, ObjConn
        if err <> 0 then
          DisConnDBH ObjConn, ObjRS	'DB�ؒf
          jampErrerP "1","b109","01","�����o�F�W�J���","101","SQL:<BR>"&StrSQL
        end if
        VslCode=Trim(ObjRS("VslCode"))
        VoyCtrl=Trim(ObjRS("VoyCtrl"))
        ObjRS.close

        StrSQL = "SELECT count(ContNo) AS CNUM FROM ImportCont WHERE BLNo='"&preDtTbl(1)(11)&"' "&_
                 "AND VoyCtrl =" & VoyCtrl & " AND VslCode= '"& VslCode &"' "
        ObjRS.Open StrSQL, ObjConn
        if err <> 0 then
          DisConnDBH ObjConn, ObjRS	'DB�ؒf
          jampErrerP "1","b109","01","�����o�F�W�J���","101","SQL:<BR>"&StrSQL
        end if
        Num = Num + ObjRS("CNUM")-1
        ObjRS.close
        ReDim Preserve DtTbl(Num)
        '�f�[�^�擾
        StrSQL = "SELECT ContNo, ReturnTime, CYDelTime, "&_
                 "NULLIF('-',LEFT(DateDiff(day,getdate(),DateAdd(day,"&preDtTbl(1)(21)-param(0)+1&" ,CYDelTime))*("&preDtTbl(1)(21)&"%6),1)) AS ReturnArrert "&_
                 "FROM ImportCont WHERE BLNo='"&preDtTbl(1)(11)&"' "&_
                 "AND VoyCtrl =" & VoyCtrl & " AND VslCode= '"& VslCode &"' "
        ObjRS.Open StrSQL, ObjConn
        if err <> 0 then
          DisConnDBH ObjConn, ObjRS	'DB�ؒf
          jampErrerP "1","b109","01","�����o�F�W�J���","101","SQL:<BR>"&StrSQL
        end if
        Do Until ObjRS.EOF
          DtTbl(j)=preDtTbl(1)
          DtTbl(j)(5)=Trim(ObjRS("ContNo"))
          DtTbl(j)(12)=ObjRS("ReturnArrert")
          If IsNull(ObjRS("ReturnTime")) Then
            DtTbl(j)(8)="��"
          Else
            DtTbl(j)(8)="��"
          End If
          ObjRS.MoveNext
          j=j+1
        Loop
        ObjRS.close
      Case Else
          jampErrerP "1","b109","01","�����o�F�W�J���","101","SQL:<BR>"&StrSQL
      End Select
  'DB�ڑ�����
    DisConnDBH ObjConn, ObjRS
  '�G���[�g���b�v����
    on error goto 0

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>�����o���O���W�J</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
window.focus();
// -->
</SCRIPT>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY>
<!-------------�����o���W�J���--------------------------->
<TABLE border="0" cellPadding="0" cellSpacing="0" width="100%">
   <TR>
     <TD align="right" bgColor="#000099" height="25" colspan="3">
       <IMG src="Image/logo_hits_ver2.gif" height="25" width="300"></TD>
   </TR>
   <TR height="48">
       <TD width="506" align=center><FONT size=+1><B>���O������<B></FONT></TD>
       <TD width="20%"><B>�����o���</B></TD>
       <TD nowrap><%=day%></TD></TR>
</TABLE>
<HR>
<CENTER>
<TABLE border="1" cellPadding="3" cellSpacing="0" cols="<%=Num+1%>">
<%If Num<>0 Then%> 
  <% If DtTbl(1)(14)<>"�@" Then %>
  <TR class=bga>
    <TH nowrap><%=DtTbl(0)(1)%></TH><TH nowrap><%=DtTbl(0)(2)%></TH>
    <TH nowrap>�w����<BR>�ւ̉�</TH>
    <TH nowrap><%=DtTbl(0)(3)%></TH><TH nowrap><%=DtTbl(0)(4)%></TH><TH nowrap><%=DtTbl(0)(5)%></TH>
    <TH nowrap><%=DtTbl(0)(15)%></TH><TH nowrap><%=DtTbl(0)(16)%></TH><TH nowrap><%=DtTbl(0)(17)%></TH>
    <TH nowrap><%=DtTbl(0)(18)%></TH><TH nowrap><%=DtTbl(0)(19)%></TH><TH nowrap><%=DtTbl(0)(24)%></TH>
    <!--<TH nowrap><%'=DtTbl(0)(6)%></TH>--><!-- Commented 2003.9.4 -->
    <TH nowrap><%=DtTbl(0)(7)%></TH>
    <TH nowrap><%=DtTbl(0)(8)%></TH><TH nowrap><%=DtTbl(0)(9)%></TH><TH nowrap><%=DtTbl(0)(10)%></TH>
    <TH nowrap><%=DtTbl(0)(22)%></TH><TH nowrap><%=DtTbl(0)(23)%></TH>
  </TR>
    <% For j=1 to Num %>
      <% If DtTbl(j)(12) = "-" Or DtTbl(j)(8) = "��" Then %>
  <TR class=bgw>
      <% Else %>
  <TR class=bgarrt>  
      <% End If%>
    <TD nowrap><%=DtTbl(j)(1)%><BR></TD><TD nowrap><%=DtTbl(j)(2)%></TD>
    <TD nowrap><%=DtTbl(j)(14)%></TD><TD nowrap><%=DtTbl(j)(3)%></TD>
    <TD nowrap><%=Siji(DtTbl(j)(4))%></TD><TD nowrap><%=DtTbl(j)(5)%></TD>
<%'C-001    <TD nowrap>< %=DtTbl(j)(15)% ></TD><TD nowrap>< %=DtTbl(j)(16)% ></TD><TD nowrap>< %=DtTbl(j)(17)% ></TD> %>
    <TD nowrap><%=DtTbl(j)(15)%><BR></TD><TD nowrap><%=DtTbl(j)(16)%><BR></TD><TD nowrap><%=DtTbl(j)(17)%><BR></TD>
    <TD nowrap><%=DtTbl(j)(18)%><BR></TD><TD nowrap><%=DtTbl(j)(19)%><BR></TD><TD nowrap><%=DtTbl(j)(24)%><BR></TD>
    <!--<TD nowrap><%'=DtTbl(j)(6)%></TD>--><!-- Commented 2003.9.4 -->
    <TD nowrap><%=DtTbl(j)(7)%></TD><TD nowrap><%=DtTbl(j)(8)%></TD>
    <TD nowrap><%=DtTbl(j)(9)%><BR></TD><TD nowrap><%=DtTbl(j)(10)%><BR></TD>
    <TD nowrap><%=DtTbl(j)(22)%><BR></TD><TD nowrap><%=DtTbl(j)(23)%><BR></TD>
  </TR>
    <% Next %>
  <% Else %>
  <TR class=bga>
    <TH nowrap><%=DtTbl(0)(1)%></TH><TH nowrap><%=DtTbl(0)(2)%></TH>
    <TH nowrap><%=DtTbl(0)(3)%></TH><TH nowrap><%=DtTbl(0)(4)%></TH><TH nowrap><%=DtTbl(0)(5)%></TH>
    <TH nowrap><%=DtTbl(0)(15)%></TH><TH nowrap><%=DtTbl(0)(16)%></TH><TH nowrap><%=DtTbl(0)(17)%></TH>
    <TH nowrap><%=DtTbl(0)(18)%></TH><TH nowrap><%=DtTbl(0)(19)%></TH><TH nowrap><%=DtTbl(0)(24)%></TH>
    <!--<TH nowrap><%'=DtTbl(0)(6)%></TH>--><!-- Commented 2003.9.4 -->
    <TH nowrap><%=DtTbl(0)(7)%></TH>
    <TH nowrap><%=DtTbl(0)(8)%></TH><TH nowrap><%=DtTbl(0)(9)%></TH><TH nowrap><%=DtTbl(0)(10)%></TH>
    <TH nowrap><%=DtTbl(0)(22)%></TH><TH nowrap><%=DtTbl(0)(23)%></TH>
  </TR>
    <% For j=1 to Num %>
      <% If DtTbl(j)(12) = "-" Or DtTbl(j)(8) = "��" Then %>
  <TR class=bgw>
      <% Else %>
  <TR class=bgarrt>  
      <% End If%>
    <TD nowrap><%=DtTbl(j)(1)%><BR></TD><TD nowrap><%=DtTbl(j)(2)%></TD>
    <TD nowrap><%=DtTbl(j)(3)%></TD><TD nowrap><%=Siji(DtTbl(j)(4))%></TD><TD nowrap><%=DtTbl(j)(5)%></TD>
<%'C-001    <TD nowrap>< %=DtTbl(j)(15)% ></TD><TD nowrap>< %=DtTbl(j)(16)% ></TD><TD nowrap>< %=DtTbl(j)(17)% ></TD> --%>
    <TD nowrap><%=DtTbl(j)(15)%><BR></TD><TD nowrap><%=DtTbl(j)(16)%><BR></TD><TD nowrap><%=DtTbl(j)(17)%><BR></TD>
    <TD nowrap><%=DtTbl(j)(18)%><BR></TD><TD nowrap><%=DtTbl(j)(19)%><BR></TD><TD nowrap><%=DtTbl(j)(24)%><BR></TD>
    <!--<TD nowrap><%'=DtTbl(j)(6)%></TD>--><!-- Commented 2003.9.4 -->
    <TD nowrap><%=DtTbl(j)(7)%></TD><TD nowrap><%=DtTbl(j)(8)%></TD>
    <TD nowrap><%=DtTbl(j)(9)%><BR></TD><TD nowrap><%=DtTbl(j)(10)%><BR></TD>
    <TD nowrap><%=DtTbl(j)(22)%><BR></TD><TD nowrap><%=DtTbl(j)(23)%><BR></TD>
  </TR>
    <% Next %>
  <% End If %>
<% Else %>
  <TR class=bgw><TD nowrap>��ƈČ��͂���܂���</TD></TR>
<% End If %>
</TABLE>
</CENTER>
<!-------------��ʏI���--------------------------->
</BODY></HTML>
