<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmo320.asp				_/
'_/	Function	:���O���������͉��(�\��)		_/
'_/	Date		:2003/05/29				_/
'_/	Code By		:SEIKO Electric.Co ��d			_/
'_/	Modify		:C-002	2003/08/07	���l���ǉ�	_/
'_/	Modify		:3th	2003/01/31	3���ύX	_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTP�R���e���c�^�C�v�ݒ�
	Response.ContentType = "text/html; charset=Shift_JIS"
	Response.AddHeader "Pragma", "no-cache" 
%>
<!--#include File="Common.inc"-->
<!--#include File="../download/download.inc"-->
<!--#include File="../ExcelCreator/include/report.inc"-->
<!--#include file="../ExcelCreator/include/XlsCrt3vbs.inc"-->

<%
'�Z�b�V�����̗L�������`�F�b�N
  CheckLoginH
  WriteLogH "b402", "���������O������","11",""

'�f�[�^���擾
  dim SakuNo,BookNo,CONnum
  dim User,TruckerSubName,TruckerName
  
  dim file1,gerrmsg
  
  SakuNo = Request("SakuNo")
  BookNo = Request("BookNo")
  CONnum = Request("CONnum")

'�f�[�^��DB��茟��
  dim CMPcd,HedId,Hmon,Hday
  dim UpFlag,TruckerFlag,WkCNo,compFlag
  dim Comment1,Comment2,Comment3	'C-002
  
  User   = Session.Contents("userid")
  
'�G���[�g���b�v�J�n
'  on error resume next
'DB�ڑ�
  dim ObjConn, ObjRS, StrSQL
  ConnDBH ObjConn, ObjRS

  StrSQL = "SELECT ITC.WkContrlNo, ITC.RegisterCode, ITC.TruckerSubCode1, ITC.TruckerSubCode2, "&_
           "ITC.TruckerSubCode3, ITC.TruckerSubCode4, ITC.HeadID, ITC.WorkDate, ITC.WorkCompleteDate, "&_
           "ITC.Comment1, ITC.Comment2, ITC.Comment3, "&_
           "ITR.TruckerFlag1, ITR.TruckerFlag2, ITR.TruckerFlag3, ITR.TruckerFlag4, "&_
           "ITC.TruckerSubName1, ITC.TruckerSubName2, ITC.TruckerSubName3, ITC.TruckerSubName4, ITC.TruckerSubName5, "&_
           "T1.Trucked AS Trucked1, T2.Trucked AS Trucked2, T3.Trucked AS Trucked3, T4.Trucked AS Trucked4 "&_
           "FROM hITCommonInfo AS ITC INNER JOIN hITReference AS ITR ON ITC.WkContrlNo = ITR.WkContrlNo "&_
           "LEFT JOIN mTrucker T1 ON (ITC.TruckerSubCode1 = T1.HeadCompanyCode) "&_
           "LEFT JOIN mTrucker T2 ON (ITC.TruckerSubCode2 = T2.HeadCompanyCode) "&_
           "LEFT JOIN mTrucker T3 ON (ITC.TruckerSubCode3 = T3.HeadCompanyCode) "&_
           "LEFT JOIN mTrucker T4 ON (ITC.TruckerSubCode4 = T4.HeadCompanyCode) "&_
           "WHERE ITC.ContNo='"&CONnum&"' AND ITC.WkNo='"& SakuNo &"' AND ITC.WkType='3' AND ITC.Process='R'"

'C-002 ADD This Line :            "ITC.Comment1, ITC.Comment2, ITC.Comment3, "&_
  ObjRS.Open StrSQL, ObjConn
  if err <> 0 then
    DisConnDBH ObjConn, ObjRS	'DB�ؒf
    jampErrerP "1","b402","11","�������F�f�[�^�擾","102","SQL:<BR>"&StrSQL
  end if
  WkCNo = Trim(ObjRS("WkContrlNo"))
  CMPcd    = Array("","","","","")
  CMPcd(0) = Trim(ObjRS("RegisterCode"))
  CMPcd(1) = Trim(ObjRS("TruckerSubCode1"))
  CMPcd(2) = Trim(ObjRS("TruckerSubCode2"))
  CMPcd(3) = Trim(ObjRS("TruckerSubCode3"))
  CMPcd(4) = Trim(ObjRS("TruckerSubCode4"))
'���O�C�����[�U�ɂ���ĉ�ЃR�[�h�\������
      chengeCompCd CMPcd, UpFlag
      '2009/07/24 M.Marquez Add-S
      if UpFlag="" then 
        UpFlag=1
      end if
      '2009/07/24 M.Marquez Add-S
      If UpFlag <> 5 Then
        TruckerFlag= Trim(ObjRS("TruckerFlag"&UpFlag))
      Else
        TruckerFlag = 0
      End If
'���O�C�����[�U�ɂ���ăw�b�hID�\������
    HedId  = Trim(ObjRS("HeadID"))
    IF TruckerFlag = 1 Then 
      HedId  = "*****"
    End If
'2009/08/04 Tanaka Upd-S
'���O�C�����[�U�ɂ���ĒS���Җ��̂𔻒f
'	Select Case Trim(User)
'		Case Trim(ObjRS("RegisterCode"))
'			TruckerSubName = Trim(ObjRS("TruckerSubName1"))
'		Case Trim(ObjRS("Trucked1"))
'			TruckerSubName = Trim(ObjRS("TruckerSubName2"))
'		Case Trim(ObjRS("Trucked2"))
'			TruckerSubName = Trim(ObjRS("TruckerSubName3"))
'		Case Trim(ObjRS("Trucked3"))
'			TruckerSubName = Trim(ObjRS("TruckerSubName4"))
'		Case Trim(ObjRS("Trucked4"))
'			TruckerSubName = Trim(ObjRS("TruckerSubName5"))
'		Case Else
'			TruckerSubName = ""
'	End Select 

	Select Case Trim(User)
		Case Trim(ObjRS("RegisterCode"))
			TruckerSubName = Trim(ObjRS("TruckerSubName1"))
			TruckerName = ObjRS("TruckerSubName1")
		Case Trim(ObjRS("Trucked1"))
			TruckerSubName = Trim(ObjRS("TruckerSubName1"))
			TruckerName = ObjRS("TruckerSubName2")
		Case Trim(ObjRS("Trucked2"))
			TruckerSubName = Trim(ObjRS("TruckerSubName2"))
			TruckerName = ObjRS("TruckerSubName3")
		Case Trim(ObjRS("Trucked3"))
			TruckerSubName = Trim(ObjRS("TruckerSubName3"))
			TruckerName = ObjRS("TruckerSubName4")
		Case Trim(ObjRS("Trucked4"))
			TruckerSubName = Trim(ObjRS("TruckerSubName4"))
			TruckerName = ObjRS("TruckerSubName5")
		Case Else
			TruckerSubName = ""
	End Select 
	
'2009/08/04 Tanaka Upd-E
'�����\���
  dim TmpA
  If ObjRS("WorkDate") = "1900/01/01" Or IsNull(ObjRS("WorkDate")) Then	'���t��Null�ł������ꍇ
    Hmon   = Null
    Hday   = Null
  Else
    TmpA   = Split(ObjRS("WorkDate"), "/", -1, 1)
    Hmon   = TmpA(1)
    Hday   = TmpA(2)
  End If
  compFlag  = isNull(ObjRS("WorkCompleteDate"))
  Comment1=Trim(ObjRS("Comment1"))
  Comment2=Trim(ObjRS("Comment2"))
  Comment3=Trim(ObjRS("Comment3"))
  ObjRS.close

  StrSQL = "SELECT CYV.ShipLine, CYV.VslName, CYV.Voyage, CYV.DPort, CYV.DelivPlace, CYV.ContSize, CYV.ContType, "&_
           "CYV.ContHeight, CYV.Material, CYV.TareWeight, CYV.CustOK, CYV.SealNo, CYV.ContWeight, CYV.ReceiveFrom, "&_
           "CYV.CustClear, CYV.OvHeight, CYV.OvWidthL, CYV.OvWidthR, CYV.OvLengthF, CYV.OvLengthA, CYV.Operator, "&_
           "EXC.RHO, EXC.SetTemp, EXC.Ventilation, EXC.IMDG1, EXC.IMDG2, EXC.IMDG3, EXC.UNNo1, EXC.UNNo2,EXC.UNNo3,"&_
           "BOK.RecTerminal, CASE WHEN mP.FullName IS Null Then Bok.PlaceRec Else mP.FullName END AS PlaceDel, BOK.LPort "&_
           "FROM (CYVanInfo AS CYV LEFT JOIN ExportCont AS EXC ON (CYV.ContNo = EXC.ContNo) AND "&_
           "(CYV.BookNo = EXC.BookNo)) LEFT JOIN Booking AS BOK ON (EXC.VslCode = BOK.VslCode) AND "&_
           "(EXC.VoyCtrl = BOK.VoyCtrl) AND (EXC.BookNo = BOK.BookNo) "&_
           "LEFT JOIN mPort AS mP ON Bok.PlaceRec = mP.PortCode "&_
           "WHERE CYV.BookNo='"& BookNo &"' AND CYV.ContNo='"& CONnum &"' AND CYV.WkNo='"& SakuNo &"' "
'20040227 Change Bok.PlaceDel->CASE WHEN mP.FullName IS Null Then Bok.PlaceRec Else mP.FullName END AS PlaceDel,
'20040227 ADD LEFT JOIN mPort AS mP ON Bok.PlaceRec = mP.PortCode
  ObjRS.Open StrSQL, ObjConn
  if err <> 0 then
    DisConnDBH ObjConn, ObjRS	'DB�ؒf
    jampErrerP "1","b402","11","�������F�f�[�^�擾","102","SQL:<BR>"&StrSQL
  end if
  dim TuSk
  If Trim(ObjRS("CustClear")) = "N" Then
    TuSk="��"
  ElseIf Trim(ObjRS("CustClear")) = "Y" Then
    TuSk="��"
  Else
    TuSk="�@"
  End If
  '2010/02/18 M.Marquez Add-A
  if Request.Form("Gamen_Mode")="R" then 
     wReportName="�����[" 
     wReportID="dmo320" 
     wOutFileName=gfReceiveReport(BookNo,SakuNo,CONnum)
     file1	= server.mappath(gOutFileForder & wOutFileName)
	 if not gfdownloadFile(file1, wOutFileName) then
			wMsg = Replace(gerrmsg,"<br>","\n")
	 end if

  end if
  '2010/02/18 M.Marquez Add-E
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>������������(�\��)</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
function setParam(target){
  window.resizeTo(850,690);
  window.focus();
  len = target.elements.length;
  for (i=0; i<46; i++) target.elements[i].readOnly = true;
  bgset(target);
  if ("<%=wMsg%>"!=""){
        alert("<%=wMsg%>");
  }
  else{
      if ("<%=Request.Form("Gamen_Mode")%>"=="R"){
        if ("<%=wOutFileName%>"!=""){
            //openwinexcel("<%=wMsg%>","<%=wOutFileName%>");
            //fOpenExcel("<%=wIISFilePath%><%=wOutFileName%>");
            //parent.location.replace("<%=wIISFilePath%><%=wOutFileName%>");
        }
        document.dmo320F.Gamen_Mode.value="";
      }
  }
}

//�R���e�i�ڍ׉��
function GoConInfo(){
  target=document.dmo320F;
  target.BookNo.disabled=true;
  BookInfo(target);
  target.BookNo.disabled=false;
}
//�X�V��ʂ�
function GoReEntry(){
  target=document.dmo320F;
  target.action="./dmi320.asp";
  return true;
}
//2010-02-18 M.Marquez Add-S
//���[�o�͉�ʂ�
function GoReport(){
  target=document.dmo320F;
  target.Gamen_Mode.value="R";
  target.submit();
  return true;
}
function openwinexcel(msg,outfile){
    var w=500;
    var h=225;
    var l=0;
    var t=0;
    var target=document.dmo320F;


    if(screen.width){
        l=(screen.width-w)/2;
    }
    if(screen.availWidth){
        l=(screen.availWidth-w)/2;
    }
    if(screen.height){
        t=(screen.height-h)/2;
    }
    if(screen.availHeight){
        t=(screen.availHeight-h)/2;
    }
    var Win = window.open("/ExcelCreator/DownloadScreen.asp?Origin=0&OutFile=" + outfile + "&msg=" + msg, "", "width="+w+",height=" + h +",top="+t+",left="+l+",status=no,resizable=yes,scrollbars=no");
}

function fOpenExcel(lFileName) {
    var Excel, Book; 
    // Create the Excel application object.
    Excel = new ActiveXObject("Excel.Application"); 
    // Make Excel visible.
    Excel.Visible = true; 
    // Open work book.
    Book = Excel.Workbooks.Open(lFileName,false)
}
//2010-02-18 M.Marquez Add-E
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="setParam(document.dmo320F)">
<!-------------������������(�\��)���--------------------------->
<FORM name="dmo320F" method="POST">
<!--2010-02-18 M.Marquez Add-A-->
<INPUT type=hidden name="Gamen_Mode">
<!--2010-02-18 M.Marquez Add-E-->
<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%">
  <TR>
    <TD colspan=2><B>������������(�\�����[�h)</B></TD>
    <TD colspan=2><TABLE border=1 cellPadding=3 cellSpacing=0 align="right">
          <TR bgcolor="#f0f0f0"><TD>��Ɣԍ�</TD><TD><%=SakuNo%></TD></TR>
        </TABLE>
    </TD></TR>
  <TR>
    <TD><DIV class=bgb>�R���e�i�m���D</DIV></TD>
    <TD><INPUT type=text name="CONnum" value="<%=CONnum%>"></TD>
    <TD>
        <DIV class=bgb>�T�C�Y�A�^�C�v�A�����A�ގ��A�e�A�E�F�C�g</DIV></TD>
    <TD><INPUT type=text name="CONsize" value="<%=Trim(ObjRS("ContSize"))%>" size=5>
        <INPUT type=text name="CONtype" value="<%=Trim(ObjRS("ContType"))%>" size=5>
        <INPUT type=text name="CONhite" value="<%=Trim(ObjRS("ContHeight"))%>" size=5>
        <INPUT type=text name="CONsitu" value="<%=Trim(ObjRS("Material"))%>" size=5>
        <INPUT type=text name="CONtear" value="<%=Trim(ObjRS("TareWeight"))%>" size=5>kg
    </TD></TR>
  <TR>
    <TD><DIV class=bgb>�u�b�L���O�m���D</DIV></TD>
    <TD><INPUT type=text name="BookNo" value="<%=BookNo%>"></TD>
    <TD><DIV class=bgb>�ۊ�</DIV></TD>
    <TD><INPUT type=text name="MrSk" value="<%=Trim(ObjRS("CustOK"))%>" size=5></TD></TR>
  <TR>
    <TD><BR><DIV class=bgb>��ЃR�[�h</DIV></TD>
    <TD>�o�^��<BR>
        <INPUT type=text name="CMPcd0" value="<%=CMPcd(0)%>" size=7>
        <INPUT type=text name="CMPcd1" value="<%=CMPcd(1)%>" size=5>
        <INPUT type=text name="CMPcd2" value="<%=CMPcd(2)%>" size=5>
        <INPUT type=text name="CMPcd3" value="<%=CMPcd(3)%>" size=5>
        <INPUT type=text name="CMPcd4" value="<%=CMPcd(4)%>" size=5></TD>
    <TD><DIV class=bgb>�V�[���ԍ�</DIV></TD>
    <TD><INPUT type=text name="SealNo" value="<%=Trim(ObjRS("SealNo"))%>"></TD></TR>
<!-- 2009/10/09 Add-S Fujiyama -->
  <TR>
    <TD Align=right><DIV class=bgb>�w�����S����</DIV></TD>
    <TD>
        <INPUT type=text name="TruckerName@" readonly = "readonly" value="<%=Trim(TruckerName)%>" maxlength=16>
    </TD></TR>
<!-- 2009/10/09 Add-S Fujiyama -->
  <TR>
    <TD><DIV class=bgb>�w�b�h�h�c</DIV></TD>
    <TD><INPUT type=text name="HedId" value="<%=HedId%>"></TD>
    <TD><DIV class=bgb>�O���X�E�F�C�g</DIV></TD>
    <TD><INPUT type=text name="GrosW" value="<%=Trim(ObjRS("ContWeight"))%>" size=5>kg</TD></TR>
  <TR>
    <TD><DIV class=bgb>������</DIV></TD>
    <TD><INPUT type=text name="HTo" value="<%=Trim(ObjRS("RecTerminal"))%>" size=30></TD>
    <TD><DIV class=bgb>������</DIV></TD>
    <TD><INPUT type=text name="HFrom" value="<%=Trim(ObjRS("ReceiveFrom"))%>" size=30></TD></TR>
  <TR>
    <TD><DIV class=bgb>�����\���</DIV></TD>
    <TD><INPUT type=text name="Hmon" value="<%=Hmon%>" size=2>��
        <INPUT type=text name="Hday" value="<%=Hday%>" size=2>��</TD>
    <TD><DIV class=bgb>�ʊ�</DIV></TD>
    <TD><INPUT type=text name="TuSk" value="<%=TuSk%>" size=5></TD></TR>
  <TR>
    <TD><DIV class=bgb>�戵�D��</DIV></TD>
    <TD><INPUT type=text name="ThkSya" value="<%=Trim(ObjRS("ShipLine"))%>" size=27></TD>
    <TD><DIV class=bgb>�q�g�n</DIV></TD>
    <TD><INPUT type=text name="RHO" value="<%=Trim(ObjRS("RHO"))%>" size=5></TD></TR>
  <TR>
    <TD><DIV class=bgb>�{�D��</DIV></TD>
    <TD><INPUT type=text name="ShipN" value="<%=Trim(ObjRS("VslName"))%>"></TD>
    <TD><DIV class=bgb>�ݒ艷�x</DIV></TD>
    <TD><INPUT type=text name="SttiT" value="<%=Trim(ObjRS("SetTemp"))%>" size=5></TD></TR>
  <TR>
    <TD><DIV class=bgb>���q</DIV></TD>
    <TD><INPUT type=text name="NextV" value="<%=Trim(ObjRS("Voyage"))%>"></TD>
    <TD><DIV class=bgb>�u�d�m�s</DIV></TD>
    <TD><INPUT type=text name="VENT" value="<%=Trim(ObjRS("Ventilation"))%>" size=5></TD></TR>
  <TR>
    <TD><DIV class=bgb>�׎�n</DIV></TD>
    <TD><INPUT type=text name="NiukP" value="<%=Trim(ObjRS("PlaceDel"))%>"></TD>
    <TD><DIV class=bgb>�h�l�c�f�P�A�h�l�c�f�Q�A�h�l�c�f�R</DIV></TD>
    <TD><INPUT type=text name="IMDG1" value="<%=Trim(ObjRS("IMDG1"))%>" size=5>
        <INPUT type=text name="IMDG2" value="<%=Trim(ObjRS("IMDG2"))%>" size=5>
        <INPUT type=text name="IMDG3" value="<%=Trim(ObjRS("IMDG3"))%>" size=5></TD></TR>
  <TR>
    <TD><DIV class=bgb>�ύ`</DIV></TD>
    <TD><INPUT type=text name="TumiP" value="<%=Trim(ObjRS("LPort"))%>"></TD>
    <TD><DIV class=bgb>�t�m �m��.�P�A�t�m �m��.�Q�A�t�m �m��.�R</DIV></TD>
    <TD><INPUT type=text name="UNNo1" value="<%=Trim(ObjRS("UNNo1"))%>" size=6>
        <INPUT type=text name="UNNo2" value="<%=Trim(ObjRS("UNNo2"))%>" size=6>
        <INPUT type=text name="UNNo3" value="<%=Trim(ObjRS("UNNo3"))%>" size=6></TD></TR>
  <TR>
    <TD><DIV class=bgb>�g�`</DIV></TD>
    <TD><INPUT type=text name="AgeP" value="<%=Trim(ObjRS("DPort"))%>"></TD>
    <TD><DIV class=bgb>�n/�g�A�n/�v�k�A�n/�v�q�A�n/�k�e�A�n/�k�`</DIV></TD>
    <TD><INPUT type=text name="OH"  value="<%=Trim(ObjRS("OvHeight"))%>"  size=5>
        <INPUT type=text name="OWL" value="<%=Trim(ObjRS("OvWidthL"))%>" size=5>
        <INPUT type=text name="OWR" value="<%=Trim(ObjRS("OvWidthR"))%>" size=5>
        <INPUT type=text name="OLF" value="<%=Trim(ObjRS("OvLengthF"))%>" size=5>
        <INPUT type=text name="OLA" value="<%=Trim(ObjRS("OvLengthA"))%>" size=5>cm</TD></TR>
  <TR>
    <TD><DIV class=bgb>�דn�n</DIV></TD>
    <TD><INPUT type=text name="NiwataP" value="<%=Trim(ObjRS("DelivPlace"))%>"></TD>
    <TD><DIV class=bgb>�I�y���[�^</DIV></TD>
    <TD><INPUT type=text name="Operator" value="<%=Trim(ObjRS("Operator"))%>"></TD></TR>
  <TR>
    <TD><DIV class=bgb>���l�P</DIV></TD>
    <TD colspan=3><INPUT type=text name="Comment1" value="<%=Comment1%>" size=73></TD></TR>
  <TR>
    <TD><DIV class=bgb>���l�Q</DIV></TD>
    <TD colspan=3><INPUT type=text name="Comment2" value="<%=Comment2%>" size=73></TD></TR>
  <TR>
    <TD><DIV class=bgb>���l�R</DIV></TD>
    <TD colspan=3><INPUT type=text name="Comment3" value="<%=Comment3%>" size=73></TD></TR>

<!-- 2009/03/10 R.Shibuta Add-S -->
  <TR>
   <TD><DIV class=bgy>�o�^�S����</DIV></TD>
   <TD><INPUT type=text name="TruckerSubName" readonly = "readonly" maxlength=16></TD>
<!-- 2009/03/10 R.Shibuta Add-E -->
  </TR>

  <TR>
    <TD colspan=4 align=center valign=bottom>
       <INPUT type=hidden name=SakuNo value="<%=SakuNo%>">
       <INPUT type=hidden name="UpFlag"  value="<%=UpFlag%>">
<%'20030909 IF compFlag AND (UCase(Session.Contents("userid"))=CMPcd(0) Or TruckerFlag<>1) Then %>

<%'Chang 20050303 STAT fro Fourth Recon By SEIKO N.Oosige
' IF UCase(Session.Contents("userid"))=CMPcd(0) Or Request("TruckerFlag")<>1 Then 
  IF UCase(Session.Contents("userid"))=CMPcd(0) Or compFlag Then %>
       <INPUT type=hidden name="compFlag" value="<%=compFlag%>">
       <INPUT type=hidden name="WkCNo"     value="<%=WkCNo%>">
       <INPUT type=hidden name="TruckerFlag" value="<%=TruckerFlag%>">
       <INPUT type=hidden name="Mord" value="1">
 <%' 2009/08/04 Tanaka Add-S %>
  <INPUT type=hidden name="TruckerName" value="<%=Trim(TruckerName)%>">
 <%' 2009/08/04 Tanaka Add-E %>
       <INPUT type=submit value="�X�V���[�h" onClick="return GoReEntry()">
<%End IF%>
       <INPUT type=submit value="����" onClick="window.close()">
       <INPUT type=button value="�����[" onClick="GoReport();">
       <P>
       <INPUT type=button value="�R���e�i���" onClick="GoConInfo()"></TD></TR>
</TABLE>
</FORM>
<!-------------��ʏI���--------------------------->
<%

'DB�ڑ�����
  DisConnDBH ObjConn, ObjRS
'�G���[�g���b�v����
  on error goto 0
%>
</BODY></HTML>