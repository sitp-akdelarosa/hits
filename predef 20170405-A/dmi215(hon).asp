<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi215.asp				_/
'_/	Function	:���O����o���͏��擾�@�\	_/
'_/	Date		:2004/01/31				_/
'_/	Code By		:SEIKO Electric.Co ��d			_/
'_/	Modify		:								_/
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

'�f�[�^����
  dim BookNo, COMPcd0,COMPcd1,Mord, ret, ErrerM
  dim shipFact,shipName,delivTo,Terminal,CYCut,Continfo
  dim TFlag,VanTime,VanPlace1,VanPlace2,GoodsName,Comment1,Comment2
  dim i,tmpDate,BookNoM
  Redim Continfo(5)
  For i=0 To 5
    Continfo(i)= Array("","","","","","")
  Next
  VanTime=Array("","","","")
  BookNo = Trim(Request("BookNo"))
  Mord    = Request("Mord")

  'add-s h.matsuda 2006/03/06
  dim ShipLine,VslCodeX,VoyCtrlX,ShoriMode
	ShoriMode = Trim(Request("ShoriMode"))
	ShipLine = Trim(Request("ShipLine"))
	VslCodeX=""
	VoyCtrlX=""
  'add-e h.matsuda 2006/03/06

  ret = true
  ErrerM = ""
'�G���[�g���b�v�J�n
  on error resume next
'DB�ڑ�
  dim ObjConn, ObjRS, StrSQL
  ConnDBH ObjConn, ObjRS
  
  If Mord=0 Then '�V�K�o�^
    COMPcd0 = UCase(Session.Contents("userid"))
    COMPcd1 = ""
  '�u�b�N�̑��݃`�F�b�N
    dim cmpNum
    StrSQL = "SELECT Count(Bok.BookNo) AS numB FROM Booking AS Bok WHERE Bok.BookNo='"& BookNo &"' "
'2006/03/06 add-s h.matsuda
	if ShipLine<>"" and ShoriMode<>"" then
		strSQL=strSQL & " AND BOK.shipline='"& ShipLine &"'"
	End If
'2006/03/06 add-s h.matsuda
    ObjRS.Open StrSQL, ObjConn
    if err <> 0 then
      DisConnDBH ObjConn, ObjRS
      jampErrerP "1","b303","01","����o�F�u�b�L���O�ԍ����݃`�F�b�N","101","SQL:<BR>"&strSQL
    end if
    cmpNum=ObjRS("numB")
    ObjRS.close
    If cmpNum<1 Then
      ret=false
      ErrerM="<P>�w�肳�ꂽ�u�b�L���O�ԍ��u"&BookNo&"�v��<BR>�V�X�e���ɓo�^����Ă��܂���B<BR>"&_
             "���͂̊ԈႢ���Ȃ����ԍ����m�F���Ă��������B</P>"
    End If


'2006/03/06 add-s h.matsuda
    If ShipLine<>"" and ShoriMode<>"" Then
		StrSQL = "select a.vslcode,a.voyctrl,a.bookno,maxtime,shipline "
		StrSQL = StrSQL & "from "
		StrSQL = StrSQL & "(select distinct "
		StrSQL = StrSQL & "BookNo,VslCode ,VoyCtrl,shipline,"
		StrSQL = StrSQL & "max(UpdtTime) maxTime from Booking "
		StrSQL = StrSQL & "group by BookNo ,VslCode ,VoyCtrl,shipline) as a "
		StrSQL = StrSQL & "where "
		StrSQL = StrSQL & "bookno='" & BookNo & "' "
		StrSQL = StrSQL & "and ShipLine='" & ShipLine & "'"
		StrSQL = StrSQL & "order by maxtime desc"
		ObjRS.Open StrSQL, ObjConn
		
		if not objrs.eof then
			VslCodeX=trim(ObjRS("vslcode"))
			VoyCtrlX=trim(ObjRS("voyctrl"))
		end if
		ObjRS.close
		else
    End If
'2006/03/06 add-e h.matsuda


    If ret Then
'2006/03/06 mod-s h.matsuda
'      BookAs ObjConn, ObjRS, BookNo,COMPcd0,ret
	if ShipLine<>"" then
      BookAs2 ObjConn, ObjRS, BookNo,COMPcd0,ret,ShipLine
	else
      BookAs ObjConn, ObjRS, BookNo,COMPcd0,ret
	end if
'2006/03/06 mod-e h.matsuda
      If Not ret Then
        ErrerM="<P>�w�肳�ꂽ�u�b�L���O�ԍ��u"&BookNo&"�v��<BR>"&_
               "�ʂ̓o�^�҂ɂ���Ă��łɓo�^����Ă��邽�߁A<BR>�o�^�ł��܂���B</P>"
      End If
    End If
    If ret Then
    '�u�b�N�̔��o�����`�F�b�N
     StrSQL = "SELECT Count(EXC.BookNo) AS numB, Count(Pic.Qty) AS numQ "&_
              "FROM ExportCont AS EXC INNER JOIN Pickup AS Pic ON (EXC.VslCode = Pic.VslCode) "&_
              "AND (EXC.VoyCtrl = Pic.VoyCtrl) AND (EXC.BookNo = Pic.BookNo) AND (EXC.PickPlace=Pic.PickPlace) "&_
              "AND (EXC.VslCode = Pic.VslCode) "&_
              "WHERE EXC.BookNo='"& BookNo &"' AND EmpDelTime IS NOT NULL"
'2006/03/06 add-s h.matsuda
    If ShipLine<>"" and ShoriMode<>"" Then
		StrSQL = StrSQL & " AND (EXC.VslCode = '" & VslCodeX & "')"
		StrSQL = StrSQL & " AND (EXC.VoyCtrl = '" & VoyCtrlX & "')"
    End If
'2006/03/06 add-e h.matsuda
     ObjRS.Open StrSQL, ObjConn
     if err <> 0 then
       DisConnDBH ObjConn, ObjRS
       jampErrerP "1","b303","01","����o�F���o�����`�F�b�N","101","SQL:<BR>"&strSQL
     end if
     cmpNum=ObjRS("numB")
     If ObjRS("numQ")<>0 Then
       ObjRS.close
       StrSQL = "SELECT Pic.Qty "&_
                "FROM ExportCont AS EXC INNER JOIN Pickup AS Pic ON (EXC.VslCode = Pic.VslCode) "&_
                "AND (EXC.VoyCtrl = Pic.VoyCtrl) AND (EXC.BookNo = Pic.BookNo) AND (EXC.PickPlace=Pic.PickPlace) "&_
                "AND (EXC.VslCode = Pic.VslCode) "&_
                "WHERE EXC.BookNo='"& BookNo &"' GROUP BY Pic.Qty"
'2006/03/06 add-s h.matsuda(SQL�����č\�z)
	    If ShipLine<>"" and ShoriMode<>"" Then
			StrSQL =replace(strsql,"GROUP BY"," AND (EXC.VslCode = '" & VslCodeX & "')"&_
					" AND (EXC.VoyCtrl = '" & VoyCtrlX & "') GROUP BY")
		End If
'2006/03/06 add-e h.matsuda
       ObjRS.Open StrSQL, ObjConn
       if err <> 0 then
         DisConnDBH ObjConn, ObjRS
         jampErrerP "1","b303","01","����o�F���o�����`�F�b�N","102","SQL:<BR>"&strSQL
       end if
       If cmpNum = ObjRS("Qty") Then
         ErrerM="<DIV class=alert><����>�w��̃u�b�L���O�ԍ��͔��o���������Ă��܂��B</DIV>"
       End If
     End If
     ObjRS.close
      '���擾
      StrSQL = "SELECT Bok.RecTerminal, Bok.VslCode, Bok.VoyCtrl, "&_
               "CASE WHEN mV.FullName IS Null Then Bok.VslCode Else mV.FullName END AS shipName, "&_
               "CASE WHEN mS.FullName IS Null Then Bok.ShipLine Else mS.FullName END AS shipfact, "&_
               "CASE WHEN mP.FullName IS Null Then Bok.DPort Else mP.FullName END AS delivTo, "&_
               "Pic.ContSize, Pic.ContType, Pic.ContHeight, Pic.PickPlace, Pic.Material, VSC.CYCut "&_
               "FROM ((((Booking AS Bok LEFT JOIN mVessel AS mV ON Bok.VslCode = mV.VslCode) "&_
               "LEFT JOIN mShipLine AS mS ON Bok.ShipLine = mS.ShipLine) "&_
               "LEFT JOIN mPort AS mP ON Bok.DPort = mP.PortCode) "&_
               "LEFT JOIN Pickup AS Pic ON (Bok.BookNo = Pic.BookNo) AND (Bok.VoyCtrl = Pic.VoyCtrl) AND (Bok.VslCode = Pic.VslCode)) "&_
               "LEFT JOIN VslSchedule AS VSC ON (Bok.VoyCtrl = VSC.VoyCtrl) AND (Bok.VslCode = VSC.VslCode) "&_
               "WHERE Bok.BookNo='"& BookNo &"' ORDER BY Bok.UpdtTime DESC"
'2006/03/06 add-s h.matsuda(SQL�����č\�z)
	    If ShipLine<>"" and ShoriMode<>"" Then
			StrSQL =replace(strsql,"ORDER BY"," AND (Bok.shipline = '" & ShipLine & "')"&_
					" ORDER BY")
		End If
'2006/03/06 add-e h.matsuda

'CW-315 ADD Bok.VslCode, Bok.VoyCtrl,
'20040227C Change Bok.DelivPlace = mP.PortCode -> Bok.DPort = mP.PortCode
'20040227C Change mP.FullName AS delivTo -> CASE WHEN mP.FullName IS Null Then Bok.DPort Else  mP.FullName END AS delivTo, 

      ObjRS.Open StrSQL, ObjConn
      if err <> 0 then
        DisConnDBH ObjConn, ObjRS
        jampErrerP "1","b302","01","����o�F���擾","102","SQL:<BR>"&strSQL
      end if
      shipFact = Trim(ObjRS("shipFact"))		'�D��
      shipName = Trim(ObjRS("shipName"))		'�D��
      delivTo  = Trim(ObjRS("delivTo"))		'�d���n
      Terminal = Trim(ObjRS("RecTerminal"))	'������b�x
      If Not IsNull(ObjRS("CYCut")) Then
        CYCut    = Left(ObjRS("CYCut"),10)			'CY�J�b�g��
        tmpDate  = Split(CYCut, "/", -1, 1)
        CYCut    = tmpDate(0) & "�N" & tmpDate(1) & "��" & tmpDate(2) & "��"
      End if
'CW-315 ADD Start
      Dim VslCode, VoyCtrl
      VslCode =Trim(ObjRS("VslCode"))
      VoyCtrl=Trim(ObjRS("VoyCtrl"))
'CW-315 ADD End
      i=0
      Do Until ObjRS.EOF OR i=5
'CW-315 Change        If shipFact = Trim(ObjRS("shipFact")) AND shipName = Trim(ObjRS("shipName")) Then
        If VslCode = Trim(ObjRS("VslCode")) AND VoyCtrl = Trim(ObjRS("VoyCtrl")) Then
          Continfo(i)(0)= Trim(ObjRS("ContSize"))			'�T�C�Y
          Continfo(i)(1)= Trim(ObjRS("ContType"))			'�^�C�v
          Continfo(i)(2)= Trim(ObjRS("ContHeight"))		'����
          Continfo(i)(3)= Trim(ObjRS("Material"))		'�ގ�
          Continfo(i)(4)= Trim(ObjRS("PickPlace"))	'�s�b�N�ꏊ
          i=i+1
          ObjRS.MoveNext
        Else
          i=5
        End If
      Loop
      ObjRS.close
    End If
    TFlag     =""
    VanPlace1 =""
    VanPlace2 =""
    GoodsName =""
    Comment1  =""
    Comment2  =""
    BookNoM  = BookNo
    dim tmpstr
    If ret Then
      tmpstr=",���͓��e�̐���:0(������)"
    Else
      tmpstr=",���͓��e�̐���:1(���)"
    End If
    WriteLogH "b302", "����o���O������","02",BookNo&tmpstr
  Else		'�X�V

'2006/03/06 add-s h.matsuda
    If ShipLine<>"" and ShoriMode<>"" Then
		StrSQL = "		   select a.vslcode,a.voyctrl,a.bookno,maxtime,shipline	"
		StrSQL = StrSQL & "from													"
		StrSQL = StrSQL & "(select distinct										"
		StrSQL = StrSQL & "BookNo,VslCode ,VoyCtrl,shipline,					"
		StrSQL = StrSQL & "max(UpdtTime) maxTime from Booking					"
		StrSQL = StrSQL & "group by BookNo ,VslCode ,VoyCtrl,shipline) as a		"
		StrSQL = StrSQL & "where												"
		StrSQL = StrSQL & "bookno='" & BookNo & "'								"
		StrSQL = StrSQL & "and ShipLine='" & ShipLine & "'						"
		StrSQL = StrSQL & "order by maxtime desc								"
		ObjRS.Open StrSQL, ObjConn
		
		if not objrs.eof then
			VslCodeX=trim(ObjRS("vslcode"))
			VoyCtrlX=trim(ObjRS("voyctrl"))
		end if
		ObjRS.close
		else
    End If
'2006/03/06 add-e h.matsuda

    dim tmpTimeA,tmpTimeB
    COMPcd0 = Request("COMPcd0")
    COMPcd1 = Request("COMPcd1")

   '���擾
    StrSQL = "SELECT TruckerFlag,ContSize1,ContType1,ContHeight1,ContMaterial1,PickPlace1,Qty1, "&_
             "ContSize2,ContType2,ContHeight2,ContMaterial2,PickPlace2,Qty2, "&_
             "ContSize3,ContType3,ContHeight3,ContMaterial3,PickPlace3,Qty3, "&_
             "ContSize4,ContType4,ContHeight4,ContMaterial4,PickPlace4,Qty4, "&_
             "ContSize5,ContType5,ContHeight5,ContMaterial5,PickPlace5,Qty5, "&_
             "VanTime,VanPlace1,VanPlace2,GoodsName,Comment1,Comment2 "&_
             "FROM BookingAssign "&_
             "WHERE BookNo='"& BookNo &"' AND SenderCode='"& COMPcd0 &"' AND TruckerCode='"& COMPcd1 &"'"
    ObjRS.Open StrSQL, ObjConn
    if err <> 0 then
      DisConnDBH ObjConn, ObjRS
      jampErrerP "1","b302","12","����o�F���擾","102","SQL:<BR>"&strSQL
    end if
    TFlag         = ObjRS("TruckerFlag")
    Continfo(0)(0)= Trim(ObjRS("ContSize1"))
    Continfo(0)(1)= Trim(ObjRS("ContType1"))
    Continfo(0)(2)= Trim(ObjRS("ContHeight1"))
    Continfo(0)(3)= Trim(ObjRS("ContMaterial1"))
    Continfo(0)(4)= Trim(ObjRS("PickPlace1"))
    Continfo(0)(5)= Trim(ObjRS("Qty1"))
    Continfo(1)(0)= Trim(ObjRS("ContSize2"))
    Continfo(1)(1)= Trim(ObjRS("ContType2"))
    Continfo(1)(2)= Trim(ObjRS("ContHeight2"))
    Continfo(1)(3)= Trim(ObjRS("ContMaterial2"))
    Continfo(1)(4)= Trim(ObjRS("PickPlace2"))
    Continfo(1)(5)= Trim(ObjRS("Qty2"))
    Continfo(2)(0)= Trim(ObjRS("ContSize3"))
    Continfo(2)(1)= Trim(ObjRS("ContType3"))
    Continfo(2)(2)= Trim(ObjRS("ContHeight3"))
    Continfo(2)(3)= Trim(ObjRS("ContMaterial3"))
    Continfo(2)(4)= Trim(ObjRS("PickPlace3"))
    Continfo(2)(5)= Trim(ObjRS("Qty3"))
    Continfo(3)(0)= Trim(ObjRS("ContSize4"))
    Continfo(3)(1)= Trim(ObjRS("ContType4"))
    Continfo(3)(2)= Trim(ObjRS("ContHeight4"))
    Continfo(3)(3)= Trim(ObjRS("ContMaterial4"))
    Continfo(3)(4)= Trim(ObjRS("PickPlace4"))
    Continfo(3)(5)= Trim(ObjRS("Qty4"))
    Continfo(4)(0)= Trim(ObjRS("ContSize5"))
    Continfo(4)(1)= Trim(ObjRS("ContType5"))
    Continfo(4)(2)= Trim(ObjRS("ContHeight5"))
    Continfo(4)(3)= Trim(ObjRS("ContMaterial5"))
    Continfo(4)(4)= Trim(ObjRS("PickPlace5"))
    Continfo(4)(5)= Trim(ObjRS("Qty5"))
    VanPlace1 = Trim(ObjRS("VanPlace1"))
    VanPlace2 = Trim(ObjRS("VanPlace2"))
    If Trim(ObjRS("VanTime")) <> "" Then
      tmpTimeA  = Split(ObjRS("VanTime"), " ", -1, 1)
      tmpTimeB  = Split(tmpTimeA(0), "/", -1, 1)
      VanTime(0)= tmpTimeB(1)
      VanTime(1)= tmpTimeB(2)
      If UBound(tmpTimeA)>0 then		'CW-318
        tmpTimeB  = Split(tmpTimeA(1), ":", -1, 1)
        VanTime(2)= tmpTimeB(0)
        VanTime(3)= tmpTimeB(1)
      End If		'CW-318
    End If
    GoodsName = Trim(ObjRS("GoodsName"))
    Comment1  = Trim(ObjRS("Comment1"))
    Comment2  = Trim(ObjRS("Comment2"))
    ObjRS.close
    if err <> 0 then
      DisConnDBH ObjConn, ObjRS
      jampErrerP "1","b303","12","����o�F�f�[�^�擾","200","SQL:<BR>"&UBound(tmpTimeA)
    end if

    '�u�b�N�̑��݃`�F�b�N
    StrSQL = "SELECT Count(Bok.BookNo) AS numB FROM Booking AS Bok WHERE Bok.BookNo='"& BookNo &"' "
    ObjRS.Open StrSQL, ObjConn
    if err <> 0 then
      DisConnDBH ObjConn, ObjRS
      jampErrerP "1","b303","12","����o�F�u�b�L���O�ԍ����݃`�F�b�N","101","SQL:<BR>"&strSQL
    end if
    If ObjRS("numB")>0 Then
      ObjRS.close
      '���擾
      StrSQL = "SELECT Bok.RecTerminal, "&_
               "CASE WHEN mV.FullName IS Null Then Bok.VslCode Else mV.FullName END AS shipName, "&_
               "CASE WHEN mS.FullName IS Null Then Bok.ShipLine Else mS.FullName END AS shipfact, "&_
               "CASE WHEN mP.FullName IS Null Then Bok.DPort Else  mP.FullName END AS delivTo, VSC.CYCut "&_
               "FROM (((Booking AS Bok LEFT JOIN mVessel AS mV ON Bok.VslCode = mV.VslCode) "&_
               "LEFT JOIN mShipLine AS mS ON Bok.ShipLine = mS.ShipLine) "&_
               "LEFT JOIN mPort AS mP ON Bok.DPort = mP.PortCode) "&_
               "LEFT JOIN VslSchedule AS VSC ON (Bok.VoyCtrl = VSC.VoyCtrl) AND (Bok.VslCode = VSC.VslCode) "&_
               "WHERE Bok.BookNo='"& BookNo &"' ORDER BY Bok.UpdtTime DESC"
'2006/03/06 add-s h.matsuda(SQL�����č\�z)
	    If ShipLine<>"" and ShoriMode<>"" Then
			StrSQL =replace(strsql,"ORDER BY"," AND (Bok.shipline = '" & ShipLine & "')"&_
					" ORDER BY")
		End If
'2006/03/06 add-e h.matsuda

'20040227C Change Bok.DelivPlace = mP.PortCode -> Bok.DPort = mP.PortCode
'20040227C Change mP.FullName AS delivTo -> CASE WHEN mP.FullName IS Null Then Bok.DPort Else  mP.FullName END AS delivTo, 
      ObjRS.Open StrSQL, ObjConn
      if err <> 0 then
        DisConnDBH ObjConn, ObjRS
        jampErrerP "1","b302","12","����o�F���擾","102","SQL:<BR>"&strSQL
      end if
      shipFact = Trim(ObjRS("shipFact"))		'�D��
      shipName = Trim(ObjRS("shipName"))		'�D��
      delivTo  = Trim(ObjRS("delivTo"))		'�d���n
      Terminal = Trim(ObjRS("RecTerminal"))	'������b�x
      If Not IsNull(ObjRS("CYCut")) Then
        CYCut    = Left(ObjRS("CYCut"),10)			'CY�J�b�g��
        tmpDate  = Split(CYCut, "/", -1, 1)
        CYCut    = tmpDate(0) & "�N" & tmpDate(1) & "��" & tmpDate(2) & "��"
      Else
        CYCut    = ""
      End If
      BookNoM  = BookNo
    Else
      shipFact = ""
      shipName = ""
      delivTo  = ""
      Terminal = ""
      CYCut    = ""
      BookNoM   = "�u�b�L���O�ԍ����폜����Ă��܂�"
    End If
    ObjRS.close
  End If
'DB�ڑ�����
  DisConnDBH ObjConn, ObjRS
'�G���[�g���b�v����
  on error goto 0

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>����o�����͊m�F</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<% IF ret Then %>
<SCRIPT language=JavaScript>
<!--
//�o�^
function GoNext(){
  target=document.dmi215F;
  <% If Mord=0 Then %>
  target.action="./dmi220.asp";
  <% Else %>
  target.action="./dmo220.asp";
  <% End If%>
  target.submit();
}
// -->
</SCRIPT>
<!-------------DB�����p���--------------------------->
<BODY onLoad="GoNext();">
<FORM name="dmi215F" method="POST">
  <INPUT type=hidden name="BookNo"   value="<%=BookNo%>">
  <INPUT type=hidden name="BookNoM"   value="<%=BookNoM%>">
  <INPUT type=hidden name="COMPcd0"  value="<%=COMPcd0%>">
  <INPUT type=hidden name="COMPcd1"  value="<%=COMPcd1%>">
  <INPUT type=hidden name="oldCOMPcd1" value="<%=COMPcd1%>">
  <INPUT type=hidden name="Mord"     value="<%=Mord%>">
  <INPUT type=hidden name="CompF"    value="<%=Request("CompF")%>" >
  <INPUT type=hidden name="shipFact" value="<%=shipFact%>">
  <INPUT type=hidden name="shipName" value="<%=shipName%>">
  <INPUT type=hidden name="delivTo"  value="<%=delivTo%>">
  <INPUT type=hidden name="Terminal"  value="<%=Terminal%>">
  <INPUT type=hidden name="CYCut"    value="<%=CYCut%>">

  <INPUT type=hidden name="shipline" value="<%=shipline%>">

  <% For i=0 To 4%>
  <INPUT type=hidden name="ContSize<%=i%>"   value="<%=Continfo(i)(0)%>">
  <INPUT type=hidden name="ContType<%=i%>"   value="<%=Continfo(i)(1)%>">
  <INPUT type=hidden name="ContHeight<%=i%>" value="<%=Continfo(i)(2)%>">
  <INPUT type=hidden name="Material<%=i%>"   value="<%=Continfo(i)(3)%>">
  <INPUT type=hidden name="PickPlace<%=i%>"  value="<%=Continfo(i)(4)%>">
  <INPUT type=hidden name="PickNum<%=i%>"    value="<%=Continfo(i)(5)%>">
  <% Next %>
  <INPUT type=hidden name="TFlag"     value="<%=TFlag%>">
  <INPUT type=hidden name="vanMon"    value="<%=VanTime(0)%>">
  <INPUT type=hidden name="vanDay"    value="<%=VanTime(1)%>">
  <INPUT type=hidden name="vanHou"    value="<%=VanTime(2)%>">
  <INPUT type=hidden name="vanMin"    value="<%=VanTime(3)%>">
  <INPUT type=hidden name="vanPlace1" value="<%=VanPlace1%>">
  <INPUT type=hidden name="vanPlace2" value="<%=VanPlace2%>">
  <INPUT type=hidden name="goodsName" value="<%=GoodsName%>">
  <INPUT type=hidden name="Comment1"  value="<%=Comment1%>">
  <INPUT type=hidden name="Comment2"  value="<%=Comment2%>">
  <INPUT type=hidden name="ErrerM"  value="<%=ErrerM%>">
</FORM>
<!-------------��ʏI���--------------------------->
<%Else%>
<!-------------�G���[���--------------------------->
<SCRIPT language=JavaScript>
<!--
window.resizeTo(400,200);
// -->
</SCRIPT>
<BODY>
<CENTER>
<DIV class=alert>
  <%=ErrerM%>
</DIV>
<P><INPUT type=submit value="����" onClick="window.close()"></P>

</CENTER>
<%End If%>

</BODY></HTML>