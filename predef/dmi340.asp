<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits										   _/
'_/	FileName	:dmi320.asp									   _/
'_/	Function	:���O���������͍X�V							   _/
'_/	Date		:2003/05/29									   _/
'_/	Code By		:SEIKO Electric.Co ��d						   _/
'_/	Modify		:C-002	2003/08/08	���l���ǉ�				   _/
'_/	Modify		:3th	2003/01/31	3���ύX					   _/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTP�R���e���c�^�C�v�ݒ�
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->
<!--#include File="../download/download.inc"-->
<!--#include File="../ExcelCreator/include/report.inc"-->
<!--#include file="../ExcelCreator/include/XlsCrt3vbs.inc"-->
<!--#include File="CommonFunc.inc"-->								<!-- 2016/08/10 H.Yoshikawa Add -->

<%
'�Z�b�V�����̗L�������`�F�b�N
  CheckLoginH
  WriteLogH "b402", "���������O������","14",""

'�T�[�o���t�̎擾
  dim DayTime, YY,Yotei
  getDayTime DayTime

'���[�U�f�[�^����
  dim USER,sUN, Utype
  USER   = UCase(Session.Contents("userid"))
  sUN    = Session.Contents("sUN")
  Utype  = Session.Contents("UType")
'�f�[�^�擾
  dim UpFlag,CONnum,SakuNo,BookNo
  dim CMPcd,HedId,HTo,Hmon,Hday,TuSk
  dim FullName,Mord,i
  dim SendUser
  Mord   = Request("Mord")
  UpFlag = Request("UpFlag")
  SakuNo = Request("SakuNo")
  CONnum = Request("CONnum")
  BookNo = Request("BookNo")
  CMPcd  = Array(Request("CMPcd0"),Request("CMPcd1"),Request("CMPcd2"),Request("CMPcd3"),Request("CMPcd4"))
  HedId   = Request("HedId")
  Hmon    = Right("00" & Request("Hmon") ,2)
  Hday    = Right("00" & Request("Hday") ,2)
  '��Ɨ\����̔N�x������
  If DayTime(1) > Hmon Then	'���N
    YY = DayTime(0) +1
  ElseIf DayTime(1) = Hmon AND DayTime(2) > Hday Then	'CW-043
    YY = DayTime(0) +1					'CW-043
  Else
    YY = DayTime(0)
  End If
  If Hmon = "00" Or Hday = "00" Then
    Yotei= "Null"
  Else
    Yotei=  "'"& YY &"/"& Hmon &"/"& Hday &"'"
  End If
  If HedId = "" Then
    HedId   = "Null"
  Else
    HedId = "'"& HedId &"'"
  End If
'2016/08/10 H.Yoshikawa Del Start
''�ʊ�
'  TuSk=Request("TuSk")
'  If TuSk="��" Then
'    TuSk="Y"
'  Else
'    TuSk="N"
'  End If
'2016/08/10 H.Yoshikawa Del End
  FullName= ""
'3th ADD ��������������
  dim OH,OWL,OWR,OLF,OLA
  If Request("OH") <>"" Then OH =Request("OH")  Else OH ="0"
  If Request("OWL")<>"" Then OWL=Request("OWL") Else OWL="0"
  If Request("OWR")<>"" Then OWR=Request("OWR") Else OWR="0"
  If Request("OLF")<>"" Then OLF=Request("OLF") Else OLF="0"
  If Request("OLA")<>"" Then OLA=Request("OLA") Else OLA="0"
'3th ADD ��������������

 dim TruckerSubName
 TruckerSubName = Request("TruckerSubName")
 
'2016/08/10 H.Yoshikawa Add Start
 dim AsDry, SolasChk, AgreeChk, kariflag
 dim LqFlag(5)
 dim Tareweight
 dim Status
 
 if gfTrim(Request("AsDry")) = "1" then
 	AsDry = "1"
 else
 	AsDry = "0"
 end if
 if gfTrim(Request("SolasChk")) = "1" then
 	SolasChk = "1"
 else
 	SolasChk = "0"
 end if
 if gfTrim(Request("AgreeChk")) = "1" then
 	AgreeChk = "1"
 else
 	AgreeChk = "0"
 end if
 kariflag = gfTrim(Request("kariflag"))
 for i = 1 to 5
	if gfTrim(Request("LqFlag" & i)) = "1" then
		LqFlag(i) = "1"
	else
		LqFlag(i) = "0"
	end if
 next
  if gfTrim(Request("CONtear")) = "" then
 	Tareweight = "NULL"
 else
 	Tareweight = Request("CONtear")
 end if
 
 if kariflag = "1" then
 	Status = "0"
 else
 	Status = "7"
 end if

'2016/08/10 H.Yoshikawa Add End
 
'�G���[�g���b�v�J�n
'  on error resume next
'DB�ڑ�
  dim ObjConn, ObjRS, StrSQL,tmpStr 
  ConnDBH ObjConn, ObjRS

  dim ret,ErrerM
  ret = true
'if Request.Form("Gamen_Mode")<>"R" then 				'2016/08/19 H.Yoshikawa Del

'3th ADD START  ��������������
  If Mord = 0 Then	'�V�K�o�^
    dim WkContrlNo,UpdateFlag,RFlag
    RFlag=0
    '�d���o�^�`�F�b�N
    StrSQL = "SELECT Count(ITC.WkContrlNo) AS Num "&_
             "FROM hITCommonInfo AS ITC LEFT JOIN CYVanInfo AS CYV ON (ITC.WkNo = CYV.WkNo) AND (ITC.ContNo = CYV.ContNo) "&_
             "WHERE ITC.ContNo='" & CONnum &"' AND ITC.Process='R' AND ITC.WkType='3' AND CYV.BookNo='"& BookNo &"' "
    ObjRS.Open StrSQL, ObjConn
    if err <> 0 then
      DisConnDBH ObjConn, ObjRS	'DB�ؒf
      jampErrerP "1","b401","03","�������F�d���`�F�b�N","101","SQL:<BR>"&StrSQL
    end if
    If Trim(ObjRS("Num")) <> "0" Then
      ret=false
      ErrerM="���쒆�Ɏw�肵���u�b�L���ONo�A�R���e�i�ԍ����o�^����܂����B<BR>���̂��ߓo�^�����̓L�����Z������܂�</P>"
    End If
    SendUser = CMPcd(1)
    ObjRS.Close
    If ret Then
      'CYVaninfo�e�[�u���ɉߋ��f�[�^���c���Ă��邩�`�F�b�N
      StrSQL = "SELECT Count(CYV.ContNo) AS Num "&_
               "FROM CYVanInfo AS CYV "&_
               "WHERE CYV.ContNo='" & CONnum &"' AND CYV.SenderCode='" & USER &"' AND CYV.BookNo='"& BookNo &"' "
      ObjRS.Open StrSQL, ObjConn
      if err <> 0 then
        DisConnDBH ObjConn, ObjRS	'DB�ؒf
        jampErrerP "1","b401","03","�������FCYVaninfo�e�[�u���`�F�b�N","101","SQL:<BR>"&StrSQL
      end if
      If Trim(ObjRS("Num")) <> "0" Then
        UpdateFlag = true
      Else
        UpdateFlag = false
      End If
      ObjRS.Close
      '�������^�ƎҖ��擾
      If CMPcd(1) <> "" Then
        StrSQL = "SELECT FullName FROM mUsers WHERE mUsers.HeadCompanyCode='" & CMPcd(1) &"'"
        ObjRS.Open StrSQL, ObjConn
        if err <> 0 then
          DisConnDBH ObjConn, ObjRS	'DB�ؒf
          jampErrerPDB ObjConn,"1","b401","03","�������F�f�[�^�o�^","102","SQL:<BR>"&StrSQL
        end if
        FullName = "'" & ObjRS("FullName") & "' "
        ObjRS.close
      Else
        FullName = "Null"
      End If
      '�f�[�^���`
      For i=1 To 4
        If CMPcd(i) = "" Then
          CMPcd(i) = "Null"
        Else
          If CMPcd(i) = USER Then
            RFlag=1
          End If
          CMPcd(i) = "'" & CMPcd(i) & "'"
        End If
      Next
      '��ƊǗ��ԍ��̍̔�
      getWkContrlNo ObjConn, ObjRS, sUN, WkContrlNo
      if err <> 0 then
        Set ObjRS = Nothing
        jampErrerPDB ObjConn,"1","b401","03","�������F�f�[�^�o�^","103","��ƊǗ��ԍ��擾�Ɏ��s<BR>"&StrSQL
      end if
      '��Ɣԍ��̍̔�
      getWkNo ObjConn, ObjRS, USER, SakuNo
      if err <> 0 then
        Set ObjRS = Nothing
        jampErrerPDB ObjConn,"1","b401","03","�������F�f�[�^�o�^","103","��Ɣԍ��擾�Ɏ��s<BR>"&StrSQL
      end if
      'IT���ʃe�[�u���ւ̓o�^
        StrSQL = "Insert Into hITCommonInfo (WkContrlNo,UpdtTime,UpdtPgCd,UpdtTmnl,Status, " &_
                 "Process,WkType,InPutDate,UpdtUserCode,WkNo,ContNo,RegisterType,RegisterName, " &_
                 "RegisterCode,TruckerSubCode1,HeadID,WorkDate,TruckerName,Comment1,Comment2,Comment3,TruckerSubName1) "&_
                 "values ('"& WkContrlNo &"','"& Now() &"','PREDEF01','"& USER &"', "&_
                 "'" & Status & "','R','3','"& Now() &"','"& USER &"','"& SakuNo &"','"& CONnum &"', "&_
                 "'"& Utype &"','"& sUN &"','"& CMPcd(0) &"',"& CMPcd(1) &","& HedId &", "&_
                 Yotei &","& FullName &",'"& gfSQLEncode(Request("Comment1")) &"','"& gfSQLEncode(Request("Comment2")) &"', "&_
                 "'"& gfSQLEncode(Request("Comment3")) &"','" & gfSQLEncode(TruckerSubName) & "'"&  ") "									'2016/09/20 Status�l�ύX
        ObjConn.Execute(StrSQL)
        if err <> 0 then
          Set ObjRS = Nothing
          jampErrerPDB ObjConn,"1","b401","03","�������F�f�[�^�o�^","103","SQL:<BR>"&StrSQL
        end if
      '�Ɖ�e�[�u���o�^
      StrSQL = "Insert Into hITReference (WkContrlNo, UpdtTime, UpdtPgCd,UpdtTmnl," &_
               "TruckerFlag1,TruckerFlag2,TruckerFlag3,TruckerFlag4)" &_
               "values ('"& WkContrlNo &"','"& Now() &"','PREDEF01','"& USER &"'," &_
               "'"&RFlag&"','0','0','0')"
      ObjConn.Execute(StrSQL)
      if err <> 0 then
        Set ObjRS = Nothing
        jampErrerPDB ObjConn,"1","b401","03","�������F�f�[�^�o�^","103","SQL:<BR>"&StrSQL
      end if
      If UpdateFlag Then
        'CYVaninfo�e�[�u���̍X�V
        '2016/08/10 H.Yoshikawa Upd Start
        'StrSQL = "UPDATE CYVanInfo SET ContSize='"&Request("CONsize")&"', ContType='"&Request("CONtype")&"', "&_
        '         "UpdtTime='"& Now() &"',UpdtPgCd='PREDEF01',UpdtTmnl='"& USER &"',"&_
        '         "ContHeight='"&Request("CONhite")&"', Material='"&Request("CONsitu")&"', "&_
        '         "ShipLine='"&Request("ThkSya")&"',VslName='"&Request("ShipN")&"',"&_
        '         "TareWeight="&Request("CONtear")&", CustOK='"&Request("MrSk")&"', "&_
        '         "SealNo='"&Request("SealNo")&"', ContWeight="&Request("GrosW")&", "&_
        '         "ReceiveFrom='"&Request("HFrom")&"', CustClear='"&TuSk&"', "&_
        '         "Voyage='"&Request("NextV")&"',DPort='"&Request("AgeP")&"',"&_
        '         "OvHeight="&OH&", OvWidthL="&OWL&",OvWidthR="&OWR&", OvLengthF="&OLF&", "&_
        '         "OvLengthA="&OLA&",DelivPlace='"&Request("NiwataP")&"',"&_
        '         "Operator='"&Request("Operator")&"', WkNo='"& SakuNo &"' "&_
        '         "WHERE BookNo='"& BookNo &"' AND SenderCode='" & USER &"' AND ContNo='"& CONnum &"'  "
        StrSQL = ""
        StrSQL = StrSQL & "UPDATE CYVanInfo SET "
        StrSQL = StrSQL & "UpdtTime='"& Now() &"',UpdtPgCd='PREDEF01',UpdtTmnl='"& gfSQLEncode(USER) &"',"
        StrSQL = StrSQL & "ShipLine='"&gfSQLEncode(Request("ThkSya"))&"', "
        StrSQL = StrSQL & "VslCode= '"&gfSQLEncode(Request("ShipCode"))&"', "
        StrSQL = StrSQL & "VslName='"&gfSQLEncode(Request("ShipN"))&"',"
        StrSQL = StrSQL & "Voyage='"&gfSQLEncode(Request("NextV"))&"',"
        StrSQL = StrSQL & "DPort='"&gfSQLEncode(Request("AgeP"))&"',"
        StrSQL = StrSQL & "ErrCode='0',"												'2018/04/04 Fujiyama Add
        StrSQL = StrSQL & "DelivPlace='"&gfSQLEncode(Request("NiwataP"))&"',"
        StrSQL = StrSQL & "LPort='"&gfSQLEncode(Request("TumiP"))&"',"					'2016/11/03 H.Yoshikawa Add
        StrSQL = StrSQL & "PlaceRec='"&gfSQLEncode(Request("NiukP"))&"',"				'2016/11/03 H.Yoshikawa Add
        StrSQL = StrSQL & "PRShipper= '"&gfSQLEncode(Request("Shipper"))&"', "
        StrSQL = StrSQL & "PRForwarder= '"&gfSQLEncode(Request("Forwarder"))&"', "
        StrSQL = StrSQL & "PRForwarderTan= '"&gfSQLEncode(Request("FwdStaff"))&"', "
        StrSQL = StrSQL & "PRForwarderTel= '"&gfSQLEncode(Request("FwdTel"))&"', "
        StrSQL = StrSQL & "ContSize='"&gfSQLEncode(Request("CONsize"))&"', "
        StrSQL = StrSQL & "ContType='"&gfSQLEncode(Request("CONtype"))&"', "
        StrSQL = StrSQL & "ContHeight='"&gfSQLEncode(Request("CONhite"))&"', "
        StrSQL = StrSQL & "Material='', "												'2016/10/17 H.Yoshikawa Add
        StrSQL = StrSQL & "TareWeight="&Tareweight&", "
        StrSQL = StrSQL & "CustOK='"&gfSQLEncode(Request("MrSk"))&"', "
        StrSQL = StrSQL & "SealNo='"&gfSQLEncode(Request("SealNo"))&"', "
        StrSQL = StrSQL & "ContWeight="&Request("GrosW")&", "
        StrSQL = StrSQL & "Solas= '"& SolasChk & "', "
        StrSQL = StrSQL & "ReportNo= '"&gfSQLEncode(Request("ReportNo"))&"', "
        StrSQL = StrSQL & "ReceiveFrom='"&gfSQLEncode(Request("HFrom"))&"', "
        if gfTrim(Request("SttiT")) = "" then
        	StrSQL = StrSQL & "SetTemp= '', "
        else
        	StrSQL = StrSQL & "SetTemp= '"&gfSQLEncode(Request("SttiT"))&"C', "
        end if
        StrSQL = StrSQL & "AsDry= '"&AsDry&"', "
        StrSQL = StrSQL & "Ventilation= '"&gfSQLEncode(Request("VENT"))&"', "
        for i = 1 to 5
          StrSQL = StrSQL & "IMDG" & i & "= '"&gfSQLEncode(Request("IMDG"&i))&"', "
          StrSQL = StrSQL & "Label" & i & "= '"&gfSQLEncode(Request("Label"&i))&"', "
          StrSQL = StrSQL & "SubLabel" & i & "= '"&gfSQLEncode(Request("SubLabel"&i))&"', "
          StrSQL = StrSQL & "UNNo" & i & "= '"&gfSQLEncode(Request("UNNo"&i))&"', "
          StrSQL = StrSQL & "LqFlag" & i & "= '"&LqFlag(i)&"', "
        next
        StrSQL = StrSQL & "OvHeight="&OH&", OvWidthL="&OWL&",OvWidthR="&OWR&", OvLengthF="&OLF&", OvLengthA="&OLA&","
        StrSQL = StrSQL & "Operator='"&gfSQLEncode(Request("Operator"))&"', WkNo='"& SakuNo &"', "
        StrSQL = StrSQL & "ContactInfo= '"&gfSQLEncode(Request("TruckerTel"))&"', "
        StrSQL = StrSQL & "Consent= '"&AgreeChk&"' "
        StrSQL = StrSQL & ",kariflag= '"&kariflag&"' "
        StrSQL = StrSQL & ",EntryName='"&gfSQLEncode(Request("EntryName"))&"' "			'2017/04/04 H.Yoshikawa Add
        StrSQL = StrSQL & "WHERE BookNo='"& gfSQLEncode(BookNo) &"' AND SenderCode='" & gfSQLEncode(USER) &"' AND ContNo='"& gfSQLEncode(CONnum) &"' "
        '2016/08/10 H.Yoshikawa Upd End
        ObjConn.Execute(StrSQL)
        if err <> 0 then
           Set ObjRS = Nothing
           jampErrerPDB ObjConn,"1","b401","03","�������F�f�[�^�o�^","104","SQL:<BR>"&StrSQL
        end if
      Else
        'CYVaninfo�e�[�u���ւ̓o�^
        '2016/08/10 H.Yoshikawa Upd Start
        'StrSQL = "Insert Into  CYVanInfo (SenderCode,BookNo,ContNo,UpdtTime,UpdtPgCd,UpdtTmnl, "&_
        '         "ContSize,ContType,ContHeight,ShipLine,VslName,Voyage,DPort,DelivPlace, "&_
        '         "SealNo,ContWeight,CustClear,Material,TareWeight,CustOK,ReceiveFrom, "&_
        '         "OvHeight,OvWidthL,OvWidthR,OvLengthF,OvLengthA,Operator,WkNo) "&_
        '         "values ('" & USER &"','"& BookNo &"','"& CONnum &"','"& Now() &"','PREDEF01','"& USER &"', "&_
        '         "'"&Request("CONsize")&"','"&Request("CONtype")&"','"&Request("CONhite")&"', "&_
        '         "'"&Request("ThkSya")&"','"&Request("ShipN")&"','"&Request("NextV")&"', "&_
        '         "'"&Request("AgeP")&"','"&Request("NiwataP")&"','"&Request("SealNo")&"',"&_
        '         "'"&Request("GrosW")&"','"&TuSk&"','"&Request("CONsitu")&"',"&Request("CONtear")&", " &_
        '         "'"&Request("MrSk")&"','"&Request("HFrom")&"', "&_
        '         OH&", "&OWL&","&OWR&","&OLF&","&OLA&", "&_
        '         "'"&Request("Operator")&"','"& SakuNo &"')"
        StrSQL = ""
        StrSQL = StrSQL & "Insert Into  CYVanInfo (SenderCode,BookNo,ContNo,"
        StrSQL = StrSQL & "UpdtTime,UpdtPgCd,UpdtTmnl, "
        StrSQL = StrSQL & "ContSize,ContType,ContHeight,TareWeight, "
        StrSQL = StrSQL & "Material, "															'2016/10/17 H.Yoshikawa Add
        StrSQL = StrSQL & "ShipLine,VslName,VslCode,Voyage, "
        StrSQL = StrSQL & "DPort,DelivPlace, "
        StrSQL = StrSQL & "LPort,PlaceRec, "													'2016/11/03 H.Yoshikawa Add
        StrSQL = StrSQL & "SealNo,ContWeight,CustOK,ReceiveFrom, "
        StrSQL = StrSQL & "ErrCode, "															'2018/04/04 Fujiyama Add
        StrSQL = StrSQL & "OvHeight,OvWidthL,OvWidthR,OvLengthF,OvLengthA,Operator,WkNo"
        StrSQL = StrSQL & ",PRShipper "
        StrSQL = StrSQL & ",PRForwarder "
        StrSQL = StrSQL & ",PRForwarderTan "
        StrSQL = StrSQL & ",PRForwarderTel "
        StrSQL = StrSQL & ",ReportNo "
        StrSQL = StrSQL & ",SetTemp "
        StrSQL = StrSQL & ",AsDry "
        StrSQL = StrSQL & ",Ventilation "
        StrSQL = StrSQL & ",ContactInfo "
        StrSQL = StrSQL & ",Solas "
        StrSQL = StrSQL & ",Consent "
        for i = 1 to 5
          StrSQL = StrSQL & ",IMDG" & i
          StrSQL = StrSQL & ",Label" & i
          StrSQL = StrSQL & ",SubLabel" & i
          StrSQL = StrSQL & ",UNNo" & i
          StrSQL = StrSQL & ",LqFlag" & i
        next
        StrSQL = StrSQL & ",kariflag "
        StrSQL = StrSQL & ",EntryName "															'2017/04/04 H.Yoshikawa Add
        StrSQL = StrSQL & ")"
        StrSQL = StrSQL & "values ('" & gfSQLEncode(USER) &"','"& gfSQLEncode(BookNo) &"','"& gfSQLEncode(CONnum) &"', "
        StrSQL = StrSQL & "'"& Now() &"','PREDEF01','"& gfSQLEncode(USER) &"', "
        StrSQL = StrSQL & "'"&gfSQLEncode(Request("CONsize"))&"','"&gfSQLEncode(Request("CONtype"))&"','"&gfSQLEncode(Request("CONhite"))&"', "&Tareweight&", "
        StrSQL = StrSQL & "'', "																					'2016/10/17 H.Yoshikawa Add
        StrSQL = StrSQL & "'"&gfSQLEncode(Request("ThkSya"))&"','"&gfSQLEncode(Request("ShipN"))&"', "
        StrSQL = StrSQL & "'"&gfSQLEncode(Request("ShipCode"))&"','"&gfSQLEncode(Request("NextV"))&"', "
        StrSQL = StrSQL & "'"&gfSQLEncode(Request("AgeP"))&"','"&gfSQLEncode(Request("NiwataP"))&"', "
        StrSQL = StrSQL & "'"&gfSQLEncode(Request("TumiP"))&"','"&gfSQLEncode(Request("NiukP"))&"', "				'2016/11/03 H.Yoshikawa Add
        StrSQL = StrSQL & "'"&gfSQLEncode(Request("SealNo"))&"',"&Request("GrosW")&", '"&gfSQLEncode(Request("MrSk"))&"','"&gfSQLEncode(Request("HFrom"))&"', "
        StrSQL = StrSQL & "'0', "																'2018/04/04 Fujiyama Add
        StrSQL = StrSQL & OH&", "&OWL&","&OWR&","&OLF&","&OLA&", "
        StrSQL = StrSQL & " '"&gfSQLEncode(Request("Operator"))&"','"& SakuNo &"' "
        StrSQL = StrSQL & ", '"&gfSQLEncode(Request("Shipper"))&"' "
        StrSQL = StrSQL & ", '"&gfSQLEncode(Request("Forwarder"))&"' "
        StrSQL = StrSQL & ", '"&gfSQLEncode(Request("FwdStaff"))&"' "
        StrSQL = StrSQL & ", '"&gfSQLEncode(Request("FwdTel"))&"' "
        StrSQL = StrSQL & ", '"&gfSQLEncode(Request("ReportNo"))&"' "
        if gfTrim(Request("SttiT")) = "" then
        	StrSQL = StrSQL & ", '' "
        else
        	StrSQL = StrSQL & ", '"&gfSQLEncode(Request("SttiT"))&"C' "
        end if
        StrSQL = StrSQL & ", '"&AsDry&"' "
        StrSQL = StrSQL & ", '"&gfSQLEncode(Request("VENT"))&"' "
        StrSQL = StrSQL & ", '"&gfSQLEncode(Request("TruckerTel"))&"' "
        StrSQL = StrSQL & ", '"&SolasChk&"' "
        StrSQL = StrSQL & ", '"&AgreeChk&"' "
        for i = 1 to 5
          StrSQL = StrSQL & ", '"&gfSQLEncode(Request("IMDG"&i))&"' "
          StrSQL = StrSQL & ", '"&gfSQLEncode(Request("Label"&i))&"' "
          StrSQL = StrSQL & ", '"&gfSQLEncode(Request("SubLabel"&i))&"' "
          StrSQL = StrSQL & ", '"&gfSQLEncode(Request("UNNo"&i))&"' "
          StrSQL = StrSQL & ", '"&LqFlag(i)&"' "
        next
        StrSQL = StrSQL & ", '"&kariflag&"' "
        StrSQL = StrSQL & ", '"&gfSQLEncode(Request("EntryName"))&"' "							'2017/04/04 H.Yoshikawa Add
        StrSQL = StrSQL & ")"
        '2016/08/10 H.Yoshikawa Upd End
'response.write StrSQL
'response.end
        ObjConn.Execute(StrSQL)
        if err <> 0 then
           Set ObjRS = Nothing
           jampErrerPDB ObjConn,"1","b401","03","�������F�f�[�^�o�^","104","SQL:<BR>"&StrSQL
        end if
      End If
    End If
  Else
'3th ADD END  ��������������
'CW-006	ADD START ��������������
   '�����E�X�V�`�F�b�N
    If UpFlag <>5 Then
      StrSQL="SELECT ITC.WorkCompleteDate, ITR.TruckerFlag"& UpFlag &" AS Flag "&_
             "FROM hITCommonInfo AS ITC INNER JOIN hITReference AS ITR ON ITC.WkContrlNo = ITR.WkContrlNo "&_
             "Where WkNo='"& SakuNo &"' AND Process='R' AND WkType='3'"
    Else
      StrSQL="SELECT WorkCompleteDate FROM hITCommonInfo " &_
             "Where WkNo='"& SakuNo &"' AND Process='R' AND WkType='3'"
    End If
    ObjRS.Open StrSQL, ObjConn
    if err <> 0 then
      ObjRS.Close
      Set ObjRS = Nothing
      jampErrerPDB ObjConn,"1","b402","14","�������F�f�[�^�o�^","102","SQL:<BR>"&StrSQL
    end if
    If NOT IsNull(ObjRS("WorkCompleteDate")) Then 
      ret=false
      ErrerM="�w��̍�Ƃ͉�ʑ��쒆�ɍ�Ƃ������������߁A�X�V�̓L�����Z������܂����B"
    End If
'ADD 20050303 STAT fro Fourth Recon By SEIKO N.Oosige
    If Len(Request("partFlg"))=1 Then
      ObjRS.close
      StrSQL = "UPDATE hITCommonInfo SET UpdtTime='"& Now() &"', UpdtPgCd='PREDEF01'," &_
               "UpdtTmnl='"& USER &"', Status='" & Status & "',Process='R',UpdtUserCode='"& USER &"', "&_
               "WorkDate=" & Yotei &_
               " Where WkNo='"& SakuNo &"' AND Process='R' AND WkType='3'"							'2016/09/20 Status�l�ύX
      ObjConn.Execute(StrSQL)
      if err <> 0 then
        Set ObjRS = Nothing
        jampErrerPDB ObjConn,"1","b402","14","�������F�f�[�^�o�^","104","SQL:<BR>"&StrSQL
      end if
      StrSQL = "UPDATE CYVanInfo SET "&_
               "UpdtTime='"& Now() &"',UpdtPgCd='PREDEF01',UpdtTmnl='"& USER &"',"&_
               "SealNo='"&Request("SealNo")&"', ContWeight="&Request("GrosW")&", "&_
               "Solas='"&SolasChk&"', Consent='"&AgreeChk&"' "&_
               "WHERE BookNo='"& BookNo &"' AND ContNo='"& CONnum &"' AND WkNo='"& SakuNo &"' "		'2016/08/10 H.Yoshikawa Upd�iCustClear �폜�j
      ObjConn.Execute(StrSQL)
      if err <> 0 then
        Set ObjRS = Nothing
        jampErrerPDB ObjConn,"1","b402","14","�������F�f�[�^�o�^","104","SQL:<BR>"&StrSQL
      end if
    Else
'ADD 20050303 END
     '�`�F�b�N
      If UpFlag <>5 Then
        If Trim(ObjRS("Flag"))=1 Then 
          ret=false
          ErrerM="�w��̍�Ƃ͉�ʑ��쒆�Ɏw����Ɏ�����ꂽ���߁A�X�V�̓L�����Z������܂����B"
        End If
      End If
      ObjRS.close
      If ret Then
'CW-006	End ADD ��������������
      '�f�[�^�X�V
        If Mord <> 2 Then	'�X�V
          If UpFlag<>5 Then
            If CMPcd(UpFlag)="" Then
              tmpStr=", TruckerSubCode"& UpFlag &"=Null "
            Else
              tmpStr=", TruckerSubCode"& UpFlag &"='"& CMPcd(UpFlag) & "' "
              SendUser = CMPcd(UpFlag)
            End If
          Else
            tmpStr=" "
          End If

          tmpStr = tmpStr & ", TruckerSubName"& UpFlag &"='"& gfSQLEncode(TruckerSubName) & "' "

        '�������^�ƎҖ��擾
          If UpFlag<2 Then
            If CMPcd(1) <> "" Then
              StrSQL = "SELECT FullName FROM mUsers WHERE mUsers.HeadCompanyCode='" & CMPcd(1) &"'"
              ObjRS.Open StrSQL, ObjConn
              if err <> 0 then
                DisConnDBH ObjConn, ObjRS	'DB�ؒf
                jampErrerPDB ObjConn,"1","b402","14","�������F�f�[�^�o�^","102","SQL:<BR>"&StrSQL
              end if
              FullName = ",TruckerName='" & ObjRS("FullName") & "' "
              ObjRS.close
            Else
              FullName = ",TruckerName=Null "
            End If
          End If

          StrSQL = "UPDATE hITCommonInfo SET UpdtTime='"& Now() &"', UpdtPgCd='PREDEF01'," &_
                   "UpdtTmnl='"& USER &"', Status='" & Status & "', Process='R', " &_
                   "UpdtUserCode='"& USER &"', "&_
                   "HeadID=" & HedId &", WorkDate=" & Yotei & tmpstr & FullName &_
                   ", Comment1='"& gfSQLEncode(Request("Comment1")) &"',Comment2='"& gfSQLEncode(Request("Comment2")) &"',Comment3='"& gfSQLEncode(Request("Comment3")) &"' "&_
                   "Where WkNo='"& SakuNo &"' AND Process='R' AND WkType='3'"																'2016/09/20 Status�l�ύX
'C-002 ADD This Line : ", Comment1='"& Request("Comment1") &"',Comment2='"& Request("Comment2") &"',Comment3='"& Request("Comment3") &"' "&_
          ObjConn.Execute(StrSQL)
          if err <> 0 then
            Set ObjRS = Nothing
            jampErrerPDB ObjConn,"1","b402","14","�������F�f�[�^�o�^","104","SQL:<BR>"&StrSQL
          end if
          If UpFlag = 5 Then
            tmpStr = " "
          Else
            If UpFlag = 1 AND CMPcd(1) = UCase(Session.Contents("COMPcd")) Then 
              tmpStr = ", TruckerFlag1=1 "
            Else
              tmpStr = ", TruckerFlag"& UpFlag &"=0 "
            End If
          End If
          UpFlag = UpFlag-1
          If UpFlag = 0 Then
            tmpStr = tmpStr&" "
          Else
            tmpStr = tmpStr&", TruckerFlag"& UpFlag &"=1 "
          End If
       '�Q�ƃt���O�X�V
          StrSQL = "UPDATE hITReference SET UpdtTime='"& Now() &"', UpdtPgCd='PREDEF01'," &_
                   "UpdtTmnl='"& USER &"'"&tmpStr&_
                   "WHERE WkContrlNo IN (select WkContrlNo From hITCommonInfo "&_
                   "WHERE WkNo='"& SakuNo &"' AND Process='R' AND WkType='3')"
          ObjConn.Execute(StrSQL)
            if err <> 0 then
              Set ObjRS = Nothing
              jampErrerPDB ObjConn,"1","b402","14","�������F�f�[�^�o�^","104","SQL:<BR>"&StrSQL
            end if
        '2016/08/10 H.Yoshikawa Upd Start
          'StrSQL = "UPDATE CYVanInfo SET ContSize='"&Request("CONsize")&"', ContType='"&Request("CONtype")&"', "&_
          '         "ContHeight='"&Request("CONhite")&"', Material='"&Request("CONsitu")&"', "&_
          '         "TareWeight="&Request("CONtear")&", CustOK='"&Request("MrSk")&"', "&_
          '         "SealNo='"&Request("SealNo")&"', ContWeight="&Request("GrosW")&", "&_
          '         "ReceiveFrom='"&Request("HFrom")&"', CustClear='"&TuSk&"', "&_
          '         "OvHeight="&OH&", OvWidthL="&OWL&", OvWidthR="&OWR&", OvLengthF="&OLF&", OvLengthA="&OLA&" "&_
          '         "WHERE BookNo='"& BookNo &"' AND ContNo='"& CONnum &"' AND WkNo='"& SakuNo &"' "
          StrSQL = ""
        StrSQL = StrSQL & "UPDATE CYVanInfo SET "
        StrSQL = StrSQL & "UpdtTime='"& Now() &"',UpdtPgCd='PREDEF01',UpdtTmnl='"& gfSQLEncode(USER) &"',"
        StrSQL = StrSQL & "VslCode= '"&gfSQLEncode(Request("ShipCode"))&"', "
        StrSQL = StrSQL & "VslName='"&gfSQLEncode(Request("ShipN"))&"', "
        StrSQL = StrSQL & "Voyage='"&gfSQLEncode(Request("NextV"))&"', "
        StrSQL = StrSQL & "PRShipper= '"&gfSQLEncode(Request("Shipper"))&"', "
        StrSQL = StrSQL & "PRForwarder= '"&gfSQLEncode(Request("Forwarder"))&"', "
        StrSQL = StrSQL & "PRForwarderTan= '"&gfSQLEncode(Request("FwdStaff"))&"', "
        StrSQL = StrSQL & "PRForwarderTel= '"&gfSQLEncode(Request("FwdTel"))&"', "
        StrSQL = StrSQL & "ContSize='"&gfSQLEncode(Request("CONsize"))&"', "
        StrSQL = StrSQL & "ContType='"&gfSQLEncode(Request("CONtype"))&"', "
        StrSQL = StrSQL & "ContHeight='"&gfSQLEncode(Request("CONhite"))&"', "
        StrSQL = StrSQL & "Material='', "																'2016/10/17 H.Yoshikawa Add
        StrSQL = StrSQL & "TareWeight="&Tareweight&", "
        StrSQL = StrSQL & "CustOK='"&gfSQLEncode(Request("MrSk"))&"', "
        StrSQL = StrSQL & "SealNo='"&gfSQLEncode(Request("SealNo"))&"', "
        StrSQL = StrSQL & "ContWeight="&Request("GrosW")&", "
        StrSQL = StrSQL & "Solas= '"&SolasChk&"', "
        StrSQL = StrSQL & "ReportNo= '"&gfSQLEncode(Request("ReportNo"))&"', "
        StrSQL = StrSQL & "ReceiveFrom='"&gfSQLEncode(Request("HFrom"))&"', "
        if gfTrim(Request("SttiT")) = "" then
        	StrSQL = StrSQL & "SetTemp= '', "
        else
        	StrSQL = StrSQL & "SetTemp= '"&gfSQLEncode(Request("SttiT"))&"C', "
        end if
        StrSQL = StrSQL & "AsDry= '"&AsDry&"', "
        StrSQL = StrSQL & "Ventilation= '"&gfSQLEncode(Request("VENT"))&"', "
        for i = 1 to 5
          StrSQL = StrSQL & "IMDG" & i & "= '"&gfSQLEncode(Request("IMDG"&i))&"', "
          StrSQL = StrSQL & "Label" & i & "= '"&gfSQLEncode(Request("Label"&i))&"', "
          StrSQL = StrSQL & "SubLabel" & i & "= '"&gfSQLEncode(Request("SubLabel"&i))&"', "
          StrSQL = StrSQL & "UNNo" & i & "= '"&gfSQLEncode(Request("UNNo"&i))&"', "
          StrSQL = StrSQL & "LqFlag" & i & "= '"&LqFlag(i)&"', "
        next
        StrSQL = StrSQL & "OvHeight="&OH&", OvWidthL="&OWL&",OvWidthR="&OWR&", OvLengthF="&OLF&", OvLengthA="&OLA&","
        StrSQL = StrSQL & "ContactInfo= '"&gfSQLEncode(Request("TruckerTel"))&"', "
        StrSQL = StrSQL & "Consent= '"&AgreeChk&"' "
        StrSQL = StrSQL & ",kariflag= '"&kariflag&"' "												'2016/10/12 H.Yoshikawa Add
        StrSQL = StrSQL & ",DPort='"&gfSQLEncode(Request("AgeP"))&"' "					'2016/11/03 H.Yoshikawa Add
        StrSQL = StrSQL & ",DelivPlace='"&gfSQLEncode(Request("NiwataP"))&"' "			'2016/11/03 H.Yoshikawa Add
        StrSQL = StrSQL & ",LPort='"&gfSQLEncode(Request("TumiP"))&"' "					'2016/11/03 H.Yoshikawa Add
        StrSQL = StrSQL & ",PlaceRec='"&gfSQLEncode(Request("NiukP"))&"' "				'2016/11/03 H.Yoshikawa Add
        StrSQL = StrSQL & ",EntryName='"&gfSQLEncode(Request("EntryName"))&"' "			'2017/04/04 H.Yoshikawa Add
       StrSQL = StrSQL & "WHERE BookNo='"& gfSQLEncode(BookNo) &"' AND ContNo='"& gfSQLEncode(CONnum) &"' AND WkNo='"& gfSQLEncode(SakuNo) &"' "
        '2016/08/10 H.Yoshikawa Upd End
          ObjConn.Execute(StrSQL)
          if err <> 0 then
             Set ObjRS = Nothing
             jampErrerPDB ObjConn,"1","b402","14","�������F�f�[�^�o�^","104","SQL:<BR>"&StrSQL
          end if
        Else		'�ۗ�
          '�w�b�_ID�X�V
            If UpFlag=5 Then
              tmpStr=""
            Else
              tmpStr=", TruckerSubCode"& UpFlag &"=Null"
            End If
            StrSQL = "UPDATE hITCommonInfo SET UpdtTime='"& Now() &"', UpdtPgCd='PREDEF01'," &_
                     "UpdtTmnl='"& USER &"', Status='" & Status & "', Process='R', " &_
                     "UpdtUserCode='"& USER &"'"& tmpStr &", HeadID=Null " &_
                     "Where ContNo='"& CONnum &"' AND WkNo='"& SakuNo &"' AND Process='R' AND WkType='3'"				'2016/09/20 H.Yoshikawa Upd Status�l�ύX
            ObjConn.Execute(StrSQL)
            if err <> 0 then
              Set ObjRS = Nothing
             jampErrerPDB ObjConn,"1","b402","15","�������F�ۗ�","102","SQL:<BR>"&StrSQL
            end if
           '�Q�ƃt���O�X�V
            StrSQL = "UPDATE hITReference SET UpdtTime='"& Now() &"', UpdtPgCd='PREDEF01'," &_
                     "UpdtTmnl='"& USER &"', TruckerFlag"& UpFlag-1 &"=2 "&_
                     "WHERE WkContrlNo IN (select WkContrlNo From hITCommonInfo "&_
                     "WHERE ContNo='"& CONnum &"' AND WkNo='"& SakuNo &"' AND Process='R' AND WkType='3')"
            ObjConn.Execute(StrSQL)
            if err <> 0 then
              Set ObjRS = Nothing
              jampErrerPDB ObjConn,"1","b402","15","�������F�ۗ�","102","SQL:<BR>"&StrSQL
            end if
          End If
      End If		'CW-006
    End If		'ADD 20050303 STAT fro Fourth Recon By SEIKO N.Oosige
  End If
  
'�f�[�^�擾
  if kariflag = "1" then 							'2016/08/19 H.Yoshikawa Add
	Dim Email1, Email2, Email3, Email4, Email5
	Dim UserName,ComInterval,rc

	'''�ʐM�Ԋu�擾
	StrSQL = "SELECT ComInterval FROM mParam WHERE Seq = '1'"

	ObjRS.Open StrSQL, ObjConn
	if err <> 0 then
	'''DB�ؒf
		DisConnDBH ObjConn, ObjRS
		jampErrerPDB ObjConn,"1","b10"&(2+Flag),"16","�������F���[�����M","104","SQL:<BR>"&StrSQL
	end if

	ComInterval = ObjRS("ComInterval")
	ObjRS.Close

	if SendUser <> "" then
	''��Ɣ����z�M���̎擾
		StrSQL = "SELECT T.*, "
		StrSQL = StrSQL & "CASE WHEN U.NameAbrev IS NULL THEN U.FullName ELSE U.NameAbrev END AS USERNAME "
		StrSQL = StrSQL & "FROM mUsers U, "
		StrSQL = StrSQL & "(SELECT T.* FROM TargetOperation T, mUsers U WHERE T.UserCode = U.UserCode "
		StrSQL = StrSQL & "AND U.HeadCompanyCode ='" & SendUser & "') T "
		StrSQL = StrSQL & "WHERE U.UserCode = '" & USER & "'"
		
		ObjRS.Open StrSQL, ObjConn
	    if ObjRS.EOF <> True then
			if err <> 0 then
		'''DB�ؒf
				DisConnDBH ObjConn, ObjRS
				jampErrerPDB ObjConn,"1","b10"&(2+Flag),"16","�������F���[�����M","104","SQL:<BR>"&StrSQL
			end if

			Dim svName, mailTo, mailFrom, attachedFiles, ObjMail
			Dim mailFlag1, mailFlag2, mailFlag3, mailFlag4, mailFlag5
			Dim mailSubject, mailBody,WorkName
			Dim SendTime, UpdateSendTime
			Dim fp, fobj, tfile

	<!-- 2009/03/10 R.Shibuta Add-S -->
		'''SMTP�T�[�o���̐ݒ�
			svName   = "slitdns2.hits-h.com"
	'		svName   = "192.168.17.61"
			attachedFiles = ""
			mailFlag1 = 0
			mailFlag2 = 0
			mailFlag3 = 0
			mailFlag4 = 0
			mailFlag5 = 0
		'''���[�����M���A�h���X�̐ݒ�
			mailFrom = "mrhits@hits-h.com"
			mailTo = ""
			rc = ""
			
			if Trim(ObjRS("Email1")) <> "" AND ObjRS("FlagRecResults1") = "1" then
				mailTo = mailTo & Trim(ObjRS("Email1"))
				mailFlag1 = 1
			else
				mailFlag1 = 0
			end if

			if Trim(ObjRS("Email2")) <> "" AND ObjRS("FlagRecResults2") = "1" then
				if mailFlag1 = 1 then
					mailTo = mailTo & vbtab & Trim(ObjRS("Email2"))
				else
					mailTo = mailTo & Trim(ObjRS("Email2"))
				end if
					mailFlag2 = 1
			else
				mailFlag2 = 0
			end if

			if Trim(ObjRS("Email3")) <> "" AND ObjRS("FlagRecResults3") = "1" then
				if mailFlag1 = 1 or mailFlag2 = 1 then
					mailTo = mailTo & vbtab & Trim(ObjRS("Email3"))
				else
					mailTo = mailTo & Trim(ObjRS("Email3"))
				end if
				mailFlag3 = 1
			else
				mailFlag3 = 0
			end if

			if Trim(ObjRS("Email4")) <> "" AND ObjRS("FlagRecResults4") = "1" then
				if mailFlag1 = 1 or mailFlag2 = 1 or mailFlag3 = 1 then
					mailTo = mailTo & vbtab & Trim(ObjRS("Email4"))
				else
					mailTo = mailTo & Trim(ObjRS("Email4"))
				end if
				mailFlag4 = 1
			else
				mailFlag4 = 0
			end if

			if Trim(ObjRS("Email5")) <> "" AND ObjRS("FlagRecResults5") = "1" then
				if mailFlag1 = 1 or mailFlag2 = 1 or mailFlag3 = 1 or mailFlag4 = 1 then
					mailTo = mailTo & vbtab & Trim(ObjRS("Email5"))
				else
					mailTo = mailTo & Trim(ObjRS("Email5"))
				end if
				mailFlag5 = 1
			else
				mailFlag5 = 0
			end if

			Set ObjMail = Server.CreateObject("BASP21")

			mailSubject = "HiTS ��ƈ˗�"
			mailBody = "���������" & "���� (" & Trim(ObjRS("USERNAME")) & "�l���)" & vbCrLf & vbCrLf
			mailBody = mailBody & "���������" & "���������܂����B" & vbCrLf
			mailBody = mailBody & "�ڂ�����HiTS�̎��O���o�^�̉�ʂ����Q�Ɖ������B"
				
			'���[�����M�������猻�݂̎������ʐM�Ԋu�ȏ�̏ꍇ�̓��[���𑗐M����B

			
			if Trim(mailTo) <> "" Then
				if ObjRS("RecResultsDate") < DateAdd("n",(ComInterval * -1), Now())  OR IsNull(ObjRS("RecResultsDate")) = True then
					rc=ObjMail.Sendmail(svName, mailTo, mailFrom, mailSubject, mailBody, attachedFiles)
					sendTime=Now
				end if

				If rc = "" Then
					'''���[�����M���t�̍X�V���s���B
					StrSQL = "UPDATE TargetOperation SET UpdtTime='" & Now() & "', UpdtPgCd='dmi340',"
					StrSQL = StrSQL & " UpdtTmnl='" & USER & "',"&  "RecResultsDate='" & Now() & "'"
					StrSQL = StrSQL &"WHERE UserCode = '" & Trim(ObjRS("UserCode")) & "'"

					ObjConn.Execute(StrSQL)
					if err <> 0 then
						Set ObjRS = Nothing
						jumpErrorPDB ObjConn,"1","c104","14","�������F���[�����M","104","SQL:<BR>"&StrSQL
					end if
				else
				WriteLogH "b10", "�������ł�", "",""
					fp = Server.MapPath("./mailerror") & "\error.txt"
					set fobj = Server.CreateObject("Scripting.FileSystemObject")
						if rc<>"" then
							if fobj.FileExists(fp) = True then
								set tfile = fobj.OpenTextFile(fp,8)
							else
								set tfile = fobj.CreateTextFile(fp,True,False)
							end if
							tfile.WriteLine sendTime & " " & rc
							tfile.Close
							ErrerM = "���[�����M�Ɏ��s���܂����B<BR>"
							ret = 1
						end if
				end if
			else

			end if
	    end if
<!-- 2009/03/10 R.Shibuta Add-E -->
	end if
  end if
  
  'DB�ڑ�����
  DisConnDBH ObjConn, ObjRS
  'end if													'2016/08/19 H.Yoshikawa Del
  '2010/06/07 M.Marquez Add-A

  if kariflag = "1" then 									'2016/10/12 H.Yoshikawa Add
    dim file1,gerrmsg
    dim file2, quefile										'2016/08/19 H.Yoshikawa Add
	dim fso, outputFile										'2016/08/19 H.Yoshikawa Add

  'if Request.Form("Gamen_Mode")="R" then					'2016/08/19 H.Yoshikawa Del
    wReportName="�����[" 
    wReportID="dmo320" 
    wOutFileName=gfReceiveReport(BookNo,SakuNo,CONnum)
    file1	= server.mappath(gOutFileForder & wOutFileName)
    '2016/09/08 H.Yoshikawa Upd Start
	'if not gfdownloadFile(file1, wOutFileName) then
	'		wMsg = Replace(gerrmsg,"<br>","\n")
	'end if
    'XPS�ŕۑ�
    file2 = server.mappath(gOutFileForder & "PDF/" & gfTrim(BookNo) & "_" & gfTrim(CONnum) & ".xps")
    quefile = server.mappath(gOutFileForder & "que/" & Replace(wOutFileName, "xls", "txt"))
    
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set outputFile = fso.OpenTextFile(quefile, 2, True)
	outputFile.WriteLine file1 & "/" & file2
	outputFile.Close
	set outputFile = Nothing
	set fso = Nothing

    
	''���[�����M�̃X�N���v�g�����s�i�߂��҂��Ȃ��j
	'Set objShell = server.CreateObject("WScript.Shell")
	'objShell.Run server.mappath(gOutFileForder & "XlsPDFChg.bat") & " " & file1 & " " & file2, 0, false
	'ErrerM = gfXlsPDFChg(file1, file2)
    '2016/09/08 H.Yoshikawa Upd End
  end if
  '2010/06/07 M.Marquez Add-E
  
'�G���[�g���b�v����
  on error goto 0

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<%'2010/06/07 M.Marquez Upd-S
  'If Mord =0 AND ret Then
  If ret Then %>
<!-------------���O��������Ɣԍ����s--------------------------->
<TITLE>��Ɣԍ����s</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT language=JavaScript>
   try{
     //window.resizeTo(200,300);
     window.resizeTo(500,300);
     //window.opener.parent.List.location.href="./dmo310F.asp"
   }catch(e){
   }
//2010-02-18 M.Marquez Add-S
//���[�o�͉�ʂ�
function GoReport(){
  target=document.dmi340F;
  target.Gamen_Mode.value="R";
  target.submit();
  return true;
}
</SCRIPT>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY>
<FORM name="dmi340F" method="POST">

  <P align=center><B>��Ɣԍ����s</B></P>
  <BR>
  <P align=center>��Ɣԍ��́u<%=SakuNo%>�v�ł��B</P>
  <BR><P><BR><P><BR><P>
  <!-- '2016/09/08 H.Yoshikawa Add Start -->
  <P align=center>
  <% If gfTrim(ErrerM) <> "" then %>
   <DIV class=alert><%=ErrerM%></DIV><BR>
  <% End If %>
  </P>
  <!-- '2016/09/08 H.Yoshikawa Add Start -->
  <P align=center>
  <INPUT type=hidden name="SakuNo" value=<%=SakuNo%>>
  <INPUT type=hidden name="CONnum" value=<%=CONnum%>>
  <INPUT type=hidden name="BookNo" value=<%=BookNo%>>
  <INPUT type=hidden name="Gamen_Mode">
  <!-- <INPUT type=button value="�����[���" onClick="GoReport();"> 		'2016/08/18 H.Yoshikawa Delete -->
  <INPUT type=button value="����" onClick="window.close()">
  </P>
<%' ELSE '2010/06/07 M.Marquez Del%>
<!--2010/06/07 M.Marquez Del Start -- >
<!-------------���O���������͍X�V--------------------------->
<!--TITLE>���O���������͍X�V</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%">
  <TR><TD align=center-->
  <%' If ret Then%>
   <!--�X�V���܂����B<BR>��ʂ͎����I�ɕ����܂��B
    <SCRIPT language=JavaScript>
     try{
       window.opener.parent.DList.location.href="./dmo310L.asp"
       window.opener.parent.Top.location.href="./dmo310T.asp"
     }catch(e){
     }
     //window.close();
    </SCRIPT-->
  <%' Else %>
   <!--DIV class=alert><%=ErrerM%></DIV><BR>
   <INPUT type=button value="����" onClick="window.close()">
  <%' End If%>
    </TD></TR>
</TABLE-->
<% End If %>

<!-------------��ʏI���--------------------------->
<!--2010/06/07 M.Marquez Del End -->
</FORM>
</BODY></HTML>