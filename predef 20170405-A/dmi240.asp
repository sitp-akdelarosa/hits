<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi240.asp				_/
'_/	Function	:���O����o�o�^�E�X�V			_/
'_/	Date		:2003/05/28				_/
'_/	Code By		:SEIKO Electric.Co ��d			_/
'_/	Modify		:C-002	2003/08/06	���l���ǉ�	_/
'_/	Modify		:3th	2003/01/31	3���S�ʉ��C	_/
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

'�T�[�o���t�̎擾
  dim DayTime
  getDayTime DayTime

'�f�[�^����
  dim BookNo,COMPcd0,COMPcd1,FullName,ret,PFlag
  dim vanMon,vanDay,vanHou,vanMin,VanTime,YY
  dim Qty1,Qty2,Qty3,Qty4,Qty5,i
  dim SendUser
  dim ErrerM													'2016/08/29 H.Yoshikawa Add
  dim oldObjRS													'2016/08/29 H.Yoshikawa Add
  Set oldObjRS = Server.CreateObject("ADODB.Recordset")			'2016/08/29 H.Yoshikawa Add
  
  BookNo = Request("BookNo")
  COMPcd0 = Request("COMPcd0")
  COMPcd1 = Request("COMPcd1")
  SendUser = Request("COMPcd1")
  
  vanMon =Right("00" & Request("vanMon"),2)
  vanDay =Right("00" & Request("vanDay"),2)
  vanHou =Right("00" & Request("vanHou"),2)
  vanMin =Right("00" & Request("vanMin"),2)
  '���̔N�x������
  If DayTime(1) > vanMon Then	'���N
    YY = DayTime(0) +1
  ElseIf DayTime(1) = vanMon AND DayTime(2) > vanDay Then
    YY = DayTime(0) +1
  Else
    YY = DayTime(0)
  End If
  If vanMon = "00" Or vanDay = "00" Then
    VanTime= "Null"
  Else
    VanTime= "'" & YY &"/"& vanMon &"/"& vanDay &" "& vanHou &":"& vanMin &"'"
  End If

'���[�U�f�[�^����
  dim USER
  USER   = UCase(Session.Contents("userid"))


'�G���[�g���b�v�J�n
  on error resume next
'DB�ڑ�
  dim ObjConn, ObjRS, StrSQL
  ConnDBH ObjConn, ObjRS
  dim tFlag

  ret = true
  '�u�b�N�̏d���o�^�`�F�b�N
  dim strCodes,PFlag1, PFlag2
'2006/03/06 mod-s h.matsuda(SQL�����č\�z)
'  checkSPBook ObjConn, ObjRS, BookNo,COMPcd0,COMPcd1,strCodes,PFlag1, PFlag2, ret
  checkSPBook2 ObjConn, ObjRS, BookNo,COMPcd0,COMPcd1,strCodes,PFlag1, PFlag2, ret
'2006/03/06 mod-e h.matsuda

  If Request("Mord") = 0 Then	'�����o�^
    If ret Then
'2006/03/06 mod-s h.matsuda
'      BookAs ObjConn, ObjRS, BookNo,COMPcd0,ret
	if Request("shipline")<>"" then
      BookAs2 ObjConn, ObjRS, BookNo,COMPcd0,ret,Request("shipline")
	else
      BookAs ObjConn, ObjRS, BookNo,COMPcd0,ret
	end if
'2006/03/06 mod-e h.matsuda
    End If
    If ret Then
      If Trim(COMPcd1) <> "" Then
        '�������^�ƎҖ��擾
        StrSQL = "SELECT FullName FROM mUsers WHERE mUsers.HeadCompanyCode='" & COMPcd1 &"'"
        ObjRS.Open StrSQL, ObjConn
        if err <> 0 then
          DisConnDBH ObjConn, ObjRS	'DB�ؒf
          jampErrerP "1","b302","04","����o�F�f�[�^�o�^","102","�������^�ƎҖ��擾�Ɏ��s<BR>"&StrSQL
        end if
        FullName = "'" & gfSQLEncode(ObjRS("FullName")) & "'"
        ObjRS.close
      Else 
        FullName = "Null"
      End If
      If COMPcd1 = UCase(Session.Contents("COMPcd")) Then 
        tFlag=1
      Else
        tFlag=0
      End If
      If PFlag1="0" Then
        StrSQL = "Insert Into SPBookInfo (BookNo, SenderCode, UpdtTime, UpdtPgCd, UpdtTmnl, Status,"&_
                 " Process, InputDate, TruckerCode, TruckerFlag, TruckerName, Comment ) "&_
                 "values ('"& BookNo &"','"& COMPcd0 &"','"& Now() &"','PREDEF01','"& USER &"','0',"&_
                 "'R','"& Now() &"','"& COMPcd1 &"','"& tFlag &"',"& FullName &",'"& Request("Comment") &"')"
        ObjConn.Execute(StrSQL)
        if err <> 0 then
          Set ObjRS = Nothing
          jampErrerPDB ObjConn,"1","b302","04","����o�F�f�[�^�o�^","103","SQL:<BR>"&StrSQL
        end if
      ElseIf PFlag1="2" Then
        StrSQL = "UPDATE SPBookInfo SET SenderCode='"& COMPcd0 &"', UpdtTime='"& Now() &"', UpdtPgCd='PREDEF01', "&_
                 "UpdtTmnl='"& USER &"', Status='0', Process='R', InputDate='"& Now() &"', "&_
                 "TruckerCode='"& COMPcd1 &"', TruckerFlag='"& tFlag &"', TruckerName="& FullName&_
                 ", Comment='"& Request("Comment") &"' "&_
                 "WHERE BookNo='"& BookNo &"' "
'2006/03/06 add-s h.matsuda(SQL�����č\�z)
		StrSQL=StrSQL & "and SenderCode = '" & COMPcd0 & "'"
'2006/03/06 add-e h.matsuda
        ObjConn.Execute(StrSQL)
        if err <> 0 then
          Set ObjRS = Nothing
          jampErrerPDB ObjConn,"1","b302","04","����o�F�f�[�^�o�^","104","SQL:<BR>"&StrSQL
        end if
      Else			'2���ȍ~
        StrSQL = "UPDATE SPBookInfo SET SenderCode='"& COMPcd0 &"', UpdtTime='"& Now() &"', UpdtPgCd='PREDEF01', "&_
                 "UpdtTmnl='"& USER &"', Status='0', Process='R', InputDate='"& Now() &"', "&_
                 "TruckerCode='', TruckerFlag='"& tFlag &"', TruckerName=''"&_
                 ", Comment='"& Request("Comment") &"' "&_
                 "WHERE BookNo='"& BookNo &"' "
'2006/03/06 add-s h.matsuda(SQL�����č\�z)
		StrSQL=StrSQL & "and SenderCode = '" & COMPcd0 & "'"
'2006/03/06 add-e h.matsuda
        ObjConn.Execute(StrSQL)
        if err <> 0 then
          Set ObjRS = Nothing
          jampErrerPDB ObjConn,"1","b302","04","����o�F�f�[�^�o�^","105","SQL:<BR>"&StrSQL'matsuda
        end if
      End If
'2016/08/29 H.Yoshikawa Del Start
'      If Request("PickNum0") = "" Then
'        Qty1="Null"
'      Else
'        Qty1="'" & Request("PickNum0") & "'"
'      End If
'      If Request("PickNum1") = "" Then
'        Qty2="Null"
'      Else
'       Qty2="'" & Request("PickNum1") & "'"
'      End If
'      If Request("PickNum2") = "" Then
'        Qty3="Null"
'      Else
'        Qty3="'" & Request("PickNum2") & "'"
'      End If
'      If Request("PickNum3") = "" Then
'        Qty4="Null"
'      Else
'        Qty4="'" & Request("PickNum3") & "'"
'      End If
'      If Request("PickNum4") = "" Then
'        Qty5="Null"
'      Else
'        Qty5="'" & Request("PickNum4") & "'"
'      End If
'
'      	If PFlag2="0" Then
''2006/03/06 h.matsuda mod-s�@�D�Вǉ�
''        StrSQL = "Insert Into BookingAssign "&_
''                 "(BookNo,SenderCode,TruckerCode,UpdtTime,UpdtPgCd,UpdtTmnl,"&_
''                 "Process,InputDate,TruckerName,TruckerFlag,"&_
''                 "ContSize1,ContType1,ContHeight1,ContMaterial1,PickPlace1,Qty1,"&_
''                 "ContSize2,ContType2,ContHeight2,ContMaterial2,PickPlace2,Qty2,"&_
''                 "ContSize3,ContType3,ContHeight3,ContMaterial3,PickPlace3,Qty3,"&_
''                 "ContSize4,ContType4,ContHeight4,ContMaterial4,PickPlace4,Qty4,"&_
''                 "ContSize5,ContType5,ContHeight5,ContMaterial5,PickPlace5,Qty5,"&_
''                 "VanTime,VanPlace1,VanPlace2,GoodsName,Comment1,Comment2) "&_
''                 "values ('"& BookNo &"','"& COMPcd0 &"','"& COMPcd1 &"','"& Now() &"','PREDEF01','"& USER &"',"&_
''                 "'R','"& Now() &"',"& FullName &",'"& tFlag &"',"&_
''                 "'"& Request("ContSize0") &"','"& Request("ContType0") &"','"& Request("ContHeight0") &"','"& Request("Material0") &"','"& Request("PickPlace0") &"',"& Qty1 &","&_
''                 "'"& Request("ContSize1") &"','"& Request("ContType1") &"','"& Request("ContHeight1") &"','"& Request("Material1") &"','"& Request("PickPlace1") &"',"& Qty2 &","&_
''                 "'"& Request("ContSize2") &"','"& Request("ContType2") &"','"& Request("ContHeight2") &"','"& Request("Material2") &"','"& Request("PickPlace2") &"',"& Qty3 &","&_
''                 "'"& Request("ContSize3") &"','"& Request("ContType3") &"','"& Request("ContHeight3") &"','"& Request("Material3") &"','"& Request("PickPlace3") &"',"& Qty4 &","&_
''                 "'"& Request("ContSize4") &"','"& Request("ContType4") &"','"& Request("ContHeight4") &"','"& Request("Material4") &"','"& Request("PickPlace4") &"',"& Qty5 &","&_
''                 VanTime &",'"& Request("vanPlace1") &"','"& Request("vanPlace2") &"','"& Request("goodsName") &"','"& Request("Comment1") &"','"& Request("Comment2") &"')"
'        StrSQL = "Insert Into BookingAssign "&_
'                 "(BookNo,SenderCode,TruckerCode,UpdtTime,UpdtPgCd,UpdtTmnl,"&_
'                 "Process,InputDate,TruckerName,TruckerFlag,"&_
'                 "ContSize1,ContType1,ContHeight1,ContMaterial1,PickPlace1,Qty1,"&_
'                 "ContSize2,ContType2,ContHeight2,ContMaterial2,PickPlace2,Qty2,"&_
'                 "ContSize3,ContType3,ContHeight3,ContMaterial3,PickPlace3,Qty3,"&_
'                 "ContSize4,ContType4,ContHeight4,ContMaterial4,PickPlace4,Qty4,"&_
'                 "ContSize5,ContType5,ContHeight5,ContMaterial5,PickPlace5,Qty5,"&_
'                 "VanTime,VanPlace1,VanPlace2,GoodsName,Comment1,Comment2,ShipLine,TruckerSubName) "&_
'                 "values ('"& BookNo &"','"& COMPcd0 &"','"& COMPcd1 &"','"& Now() &"','PREDEF01','"& USER &"',"&_
'                 "'R','"& Now() &"',"& FullName &",'"& tFlag &"',"&_
'                 "'"& Request("ContSize0") &"','"& Request("ContType0") &"','"& Request("ContHeight0") &"','"& Request("Material0") &"','"& Request("PickPlace0") &"',"& Qty1 &","&_
'                 "'"& Request("ContSize1") &"','"& Request("ContType1") &"','"& Request("ContHeight1") &"','"& Request("Material1") &"','"& Request("PickPlace1") &"',"& Qty2 &","&_
'                 "'"& Request("ContSize2") &"','"& Request("ContType2") &"','"& Request("ContHeight2") &"','"& Request("Material2") &"','"& Request("PickPlace2") &"',"& Qty3 &","&_
'                 "'"& Request("ContSize3") &"','"& Request("ContType3") &"','"& Request("ContHeight3") &"','"& Request("Material3") &"','"& Request("PickPlace3") &"',"& Qty4 &","&_
'                 "'"& Request("ContSize4") &"','"& Request("ContType4") &"','"& Request("ContHeight4") &"','"& Request("Material4") &"','"& Request("PickPlace4") &"',"& Qty5 &","&_
'                 VanTime &",'"& Request("vanPlace1") &"','"& Request("vanPlace2") &"','"&_
'                 Request("goodsName") &"','"& Request("Comment1") &"','"& Request("Comment2") &"','"& Request("shipline") &"','" & Request("TruckerSubName") & "')"
'                 SendUser = COMPcd1
''2006/03/06 h.matsuda mod-s
'        ObjConn.Execute(StrSQL)
'        if err <> 0 then
'          Set ObjRS = Nothing
'          jampErrerPDB ObjConn,"1","b302","04","����o�F�f�[�^�o�^","103","SQL:<BR>"&StrSQL
'        end if
'      Else
''2006/03/06 h.matsuda mod-s�@�D�Вǉ�
''        StrSQL = "UPDATE  BookingAssign SET  "&_
''                 "UpdtTime='"& Now() &"',UpdtPgCd='PREDEF01',UpdtTmnl='"& USER &"',"&_
''                 "Process='R',InputDate='"& Now() &"',TruckerName="& FullName &",TruckerFlag='"& tFlag &"',"&_
''                 "ContSize1='"& Request("ContSize0") &"',ContType1='"& Request("ContType0") &"',ContHeight1='"& Request("ContHeight0") &"',"&_
''                 "ContMaterial1='"& Request("Material0") &"',PickPlace1='"& Request("PickPlace0") &"',Qty1="& Qty1 &","&_
''                 "ContSize2='"& Request("ContSize1") &"',ContType2='"& Request("ContType1") &"',ContHeight2='"& Request("ContHeight1") &"',"&_
''                 "ContMaterial2='"& Request("Material1") &"',PickPlace2='"& Request("PickPlace1") &"',Qty2="& Qty2 &","&_
''                 "ContSize3='"& Request("ContSize2") &"',ContType3='"& Request("ContType2") &"',ContHeight3='"& Request("ContHeight2") &"',"&_
''                 "ContMaterial3='"& Request("Material2") &"',PickPlace3='"& Request("PickPlace2") &"',Qty3="& Qty3 &","&_
''                 "ContSize4='"& Request("ContSize3") &"',ContType4='"& Request("ContType3") &"',ContHeight4='"& Request("ContHeight3") &"',"&_
''                 "ContMaterial4='"& Request("Material3") &"',PickPlace4='"& Request("PickPlace3") &"',Qty4="& Qty4 &","&_
''                 "ContSize5='"& Request("ContSize4") &"',ContType5='"& Request("ContType4") &"',ContHeight5='"& Request("ContHeight4") &"',"&_
''                 "ContMaterial5='"& Request("Material4") &"',PickPlace5='"& Request("PickPlace4") &"',Qty5="& Qty5 &","&_
''                 "VanTime="& VanTime &",VanPlace1='"& Request("vanPlace1") &"',VanPlace2='"& Request("vanPlace2") &"',"&_
''                 "GoodsName='"& Request("goodsName") &"',Comment1='"& Request("Comment1") &"',Comment2='"& Request("Comment2") &"'"&_
''                 "WHERE BookNo='"& BookNo &"' AND SenderCode='"& COMPcd0 &"' AND TruckerCode='"& COMPcd1 &"'"
'        StrSQL = "UPDATE  BookingAssign SET  "&_
'                 "UpdtTime='"& Now() &"',UpdtPgCd='PREDEF01',UpdtTmnl='"& USER &"',"&_
'                 "Process='R',InputDate='"& Now() &"',TruckerName="& FullName &",TruckerFlag='"& tFlag &"',"&_
'                 "ContSize1='"& Request("ContSize0") &"',ContType1='"& Request("ContType0") &"',ContHeight1='"& Request("ContHeight0") &"',"&_
'                 "ContMaterial1='"& Request("Material0") &"',PickPlace1='"& Request("PickPlace0") &"',Qty1="& Qty1 &","&_
'                 "ContSize2='"& Request("ContSize1") &"',ContType2='"& Request("ContType1") &"',ContHeight2='"& Request("ContHeight1") &"',"&_
'                 "ContMaterial2='"& Request("Material1") &"',PickPlace2='"& Request("PickPlace1") &"',Qty2="& Qty2 &","&_
'                 "ContSize3='"& Request("ContSize2") &"',ContType3='"& Request("ContType2") &"',ContHeight3='"& Request("ContHeight2") &"',"&_
'                 "ContMaterial3='"& Request("Material2") &"',PickPlace3='"& Request("PickPlace2") &"',Qty3="& Qty3 &","&_
'                 "ContSize4='"& Request("ContSize3") &"',ContType4='"& Request("ContType3") &"',ContHeight4='"& Request("ContHeight3") &"',"&_
'                 "ContMaterial4='"& Request("Material3") &"',PickPlace4='"& Request("PickPlace3") &"',Qty4="& Qty4 &","&_
'                 "ContSize5='"& Request("ContSize4") &"',ContType5='"& Request("ContType4") &"',ContHeight5='"& Request("ContHeight4") &"',"&_
'                 "ContMaterial5='"& Request("Material4") &"',PickPlace5='"& Request("PickPlace4") &"',Qty5="& Qty5 &","&_
'                 "VanTime="& VanTime &",VanPlace1='"& Request("vanPlace1") &"',VanPlace2='"& Request("vanPlace2") &"',"&_
'                 "GoodsName='"& Request("goodsName") &"',Comment1='"& Request("Comment1") &"',Comment2='"& Request("Comment2") &"',"&_
'                 "ShipLine='"& Request("ShipLine") &"',"&_
'                 "TruckerSubName='"& Request("TruckerSubName") & "'"&_
'                 "WHERE (BookNo='"& BookNo &"' AND SenderCode='"& COMPcd0 &"' AND TruckerCode='"& COMPcd1 &"') "&_
'                 "and (isnull(shipline,'x')='x' or shipline='" & Request("shipline") & "')"
'                 SendUser = COMPcd1
''2006/03/06 h.matsuda mod-s
'        ObjConn.Execute(StrSQL)
'        if err <> 0 then
'          Set ObjRS = Nothing
'          jampErrerPDB ObjConn,"1","b302","04","����o�F�f�[�^�o�^","103","SQL:<BR>"&StrSQL
'        end if
'      End If
'    End If
'2016/08/29 H.Yoshikawa Del End

      '���̓f�[�^�ۑ�
      ret = InsBookAssign()					'2016/08/29 H.Yoshikawa Add
    End If
  ElseIF Request("Mord") = 1 Then	'�X�V
    ret=true
    Dim oldCOMPcd1
    oldCOMPcd1=Request("oldCOMPcd1")
    If oldCOMPcd1 = "" Then
      oldCOMPcd1= " "
    End If
      If Trim(COMPcd1) <> "" Then
    '�������^�ƎҖ��擾
        StrSQL = "SELECT FullName FROM mUsers WHERE mUsers.HeadCompanyCode='" & COMPcd1 &"'"
        ObjRS.Open StrSQL, ObjConn
        if err <> 0 then
          DisConnDBH ObjConn, ObjRS	'DB�ؒf
          jampErrerP "1","b302","04","����o�F�f�[�^�o�^","102","�������^�ƎҖ��擾�Ɏ��s<BR>"&StrSQL
        end if
        FullName = "'" & gfSQLEncode(ObjRS("FullName")) & "'"
        ObjRS.close
      Else 
        FullName = "Null"
      End If
      If COMPcd1 = UCase(Session.Contents("COMPcd")) Then 
        tFlag=1
      Else
        tFlag=0
      End If
      If PFlag1="1" Then			'1���ڂȂ��
        StrSQL = "UPDATE SPBookInfo SET SenderCode='"& COMPcd0 &"', UpdtTime='"& Now() &"', UpdtPgCd='PREDEF01', "&_
                 "UpdtTmnl='"& USER &"', Status='0', Process='R', "&_
                 "TruckerCode='"& COMPcd1 &"', TruckerFlag='"& tFlag &"', TruckerName="& FullName &" "&_
                 ", Comment='"& Request("Comment") &"' "&_
                 "WHERE BookNo='"& BookNo &"' "
		StrSQL=StrSQL & " and SenderCode = '" & COMPcd0 & "'"					'2016/10/27 H.Yoshikawa Add
'C-002 ADD This Line: ", Comment='"& Request("Comment") &"' "&_
        ObjConn.Execute(StrSQL)
        if err <> 0 then
          Set ObjRS = Nothing
          jampErrerPDB ObjConn,"1","b302","04","����o�F�f�[�^�o�^","104","SQL:<BR>"&StrSQL
        end if

      End If
'2016/08/29 H.Yoshikawa Del Start
'      If Request("PickNum0") = "" Then
'        Qty1="Null"
'      Else
'        Qty1="'" & Request("PickNum0") & "'"
'      End If
'      If Request("PickNum1") = "" Then
'        Qty2="Null"
'      Else
'        Qty2="'" & Request("PickNum1") & "'"
'      End If
'      If Request("PickNum2") = "" Then
'        Qty3="Null"
'      Else
'        Qty3="'" & Request("PickNum2") & "'"
'      End If
'      If Request("PickNum3") = "" Then
'        Qty4="Null"
'      Else
'        Qty4="'" & Request("PickNum3") & "'"
'      End If
'      If Request("PickNum4") = "" Then
'        Qty5="Null"
'      Else
'        Qty5="'" & Request("PickNum4") & "'"
'      End If
'      '�ߋ��f�[�^�̏d���`�F�b�N
'      StrSQL = "SELECT Count(BookNo) AS Num FROM BookingAssign "&_
'               "WHERE BookNo='"& BookNo &"' AND TruckerCode='"& COMPcd1 &"' AND Process='D'"
'      ObjRS.Open StrSQL, ObjConn
'      if err <> 0 then
'        DisConnDBH ObjConn, ObjRS	'DB�ؒf
'        jampErrerP "1","b000","00","�u�b�L���O�w���e�[�u��","101","SQL�F<BR>"&StrSQL
'      end if
'      If ObjRS("Num")>0 Then  '�폜���ꂽ�f�[�^�����݂��Ă���ꍇ
'        StrSQL = "UPDATE BookingAssign SET  "&_
'                 "TruckerCode='"& COMPcd1 &"',UpdtTime='"& Now() &"',UpdtPgCd='PREDEF01',UpdtTmnl='"& USER &"',"&_
'                 "Process='R',TruckerName="& FullName &",TruckerFlag='"& tFlag &"', "&_
'                 "ContSize1='"& Request("ContSize0") &"',ContType1='"& Request("ContType0") &"',ContHeight1='"& Request("ContHeight0") &"',"&_
'                 "ContMaterial1='"& Request("Material0") &"',PickPlace1='"& Request("PickPlace0") &"',Qty1="& Qty1 &", "&_
'                 "ContSize2='"& Request("ContSize1") &"',ContType2='"& Request("ContType1") &"',ContHeight2='"& Request("ContHeight1") &"',"&_
'                 "ContMaterial2='"& Request("Material1") &"',PickPlace2='"& Request("PickPlace1") &"',Qty2="& Qty2 &", "&_
'                 "ContSize3='"& Request("ContSize2") &"',ContType3='"& Request("ContType2") &"',ContHeight3='"& Request("ContHeight2") &"',"&_
'                 "ContMaterial3='"& Request("Material2") &"',PickPlace3='"& Request("PickPlace2") &"',Qty3="& Qty3 &", "&_
'                 "ContSize4='"& Request("ContSize3") &"',ContType4='"& Request("ContType3") &"',ContHeight4='"& Request("ContHeight3") &"',"&_
'                 "ContMaterial4='"& Request("Material3") &"',PickPlace4='"& Request("PickPlace3") &"',Qty4="& Qty4 &", "&_
'                 "ContSize5='"& Request("ContSize4") &"',ContType5='"& Request("ContType4") &"',ContHeight5='"& Request("ContHeight4") &"',"&_
'                 "ContMaterial5='"& Request("Material4") &"',PickPlace5='"& Request("PickPlace4") &"',Qty5="& Qty5 &", "&_
'                 "VanTime="& VanTime &",VanPlace1='"& Request("vanPlace1") &"',VanPlace2='"& Request("vanPlace2") &"', "&_
'                 "GoodsName='"& Request("goodsName") &"',Comment1='"& Request("Comment1") &"',Comment2='"& Request("Comment2") &"', "&_
'                 "TruckerSubName='"& Request("TruckerSubName") & "'"&_
'                 "WHERE BookNo='"& BookNo &"' AND SenderCode='"& COMPcd0 &"' AND TruckerCode='"& COMPcd1 &"'"
'                 SendUser = COMPcd1
'        ObjConn.Execute(StrSQL)
'        if err <> 0 then
'          Set ObjRS = Nothing
'          jampErrerPDB ObjConn,"1","b302","04","����o�F�f�[�^�o�^","104","SQL:<BR>"&StrSQL
'        end if
'          StrSQL = "UPDATE BookingAssign SET  "&_
'                   "UpdtTime='"& Now() &"',UpdtPgCd='PREDEF01',UpdtTmnl='"& USER &"',"&_
'                   "Process='D' WHERE BookNo='"& BookNo &"' AND SenderCode='"& COMPcd0 &"' AND TruckerCode='"& oldCOMPcd1 &"'"
'        ObjConn.Execute(StrSQL)
'        if err <> 0 then
'          Set ObjRS = Nothing
'          jampErrerPDB ObjConn,"1","b302","04","����o�F�f�[�^�o�^","104","SQL:<BR>"&StrSQL
'        end if
'      Else
'        StrSQL = "UPDATE BookingAssign SET  "&_
'                 "TruckerCode='"& COMPcd1 &"',UpdtTime='"& Now() &"',UpdtPgCd='PREDEF01',UpdtTmnl='"& USER &"',"&_
'                 "Process='R',TruckerName="& FullName &",TruckerFlag='"& tFlag &"', "&_
'                 "ContSize1='"& Request("ContSize0") &"',ContType1='"& Request("ContType0") &"',ContHeight1='"& Request("ContHeight0") &"',"&_
'                 "ContMaterial1='"& Request("Material0") &"',PickPlace1='"& Request("PickPlace0") &"',Qty1="& Qty1 &", "&_
'                 "ContSize2='"& Request("ContSize1") &"',ContType2='"& Request("ContType1") &"',ContHeight2='"& Request("ContHeight1") &"',"&_
'                 "ContMaterial2='"& Request("Material1") &"',PickPlace2='"& Request("PickPlace1") &"',Qty2="& Qty2 &", "&_
'                 "ContSize3='"& Request("ContSize2") &"',ContType3='"& Request("ContType2") &"',ContHeight3='"& Request("ContHeight2") &"',"&_
'                 "ContMaterial3='"& Request("Material2") &"',PickPlace3='"& Request("PickPlace2") &"',Qty3="& Qty3 &", "&_
'                 "ContSize4='"& Request("ContSize3") &"',ContType4='"& Request("ContType3") &"',ContHeight4='"& Request("ContHeight3") &"',"&_
'                 "ContMaterial4='"& Request("Material3") &"',PickPlace4='"& Request("PickPlace3") &"',Qty4="& Qty4 &", "&_
'                 "ContSize5='"& Request("ContSize4") &"',ContType5='"& Request("ContType4") &"',ContHeight5='"& Request("ContHeight4") &"',"&_
'                 "ContMaterial5='"& Request("Material4") &"',PickPlace5='"& Request("PickPlace4") &"',Qty5="& Qty5 &", "&_
'                 "VanTime="& VanTime &",VanPlace1='"& Request("vanPlace1") &"',VanPlace2='"& Request("vanPlace2") &"', "&_
'                 "GoodsName='"& Request("goodsName") &"',Comment1='"& Request("Comment1") &"',Comment2='"& Request("Comment2") &"', "&_
'                 "TruckerSubName='"& Request("TruckerSubName") & "'"&_
'                 "WHERE BookNo='"& BookNo &"' AND SenderCode='"& COMPcd0 &"' AND TruckerCode='"& oldCOMPcd1 &"'"
'                 SendUser = oldCOMPcd1
'        ObjConn.Execute(StrSQL)
'        if err <> 0 then
'          Set ObjRS = Nothing
'          jampErrerPDB ObjConn,"1","b302","04","����o�F�f�[�^�o�^","104","SQL:<BR>"&StrSQL
'        end if
'      End If
'      ObjRS.close
'2016/08/29 H.Yoshikawa Del End
'2016/08/29 H.Yoshikawa Add Start
      '�ύX�O�f�[�^�擾
      StrSQL = "SELECT *, CONVERT(CHAR(10), PickDate, 111) AS PickDateStr FROM BookingAssign "
      StrSQL = StrSQL & " WHERE BookNo='"& BookNo &"' AND SenderCode = '" & COMPcd0 & "' AND TruckerCode='"& oldCOMPcd1 &"' ORDER BY Seq "
      oldObjRS.Open StrSQL, ObjConn
      if err <> 0 then
        DisConnDBH ObjConn, oldObjRS	'DB�ؒf
        jampErrerP "1","b000","00","�u�b�L���O�w���e�[�u��","101","SQL�F<BR>"&StrSQL
      end if
      
      '2016/11/29 H.Yoshikawa Del Start
      ''�ߋ��f�[�^�̏d���`�F�b�N
      'StrSQL = "SELECT Count(BookNo) AS Num FROM BookingAssign "&_
      '         "WHERE BookNo='"& BookNo &"' AND TruckerCode='"& COMPcd1 &"' AND Process='D'" & " AND SenderCode = '" & COMPcd0 & "' "
      'ObjRS.Open StrSQL, ObjConn
      'if err <> 0 then
      '  DisConnDBH ObjConn, ObjRS	'DB�ؒf
      '  jampErrerP "1","b000","00","�u�b�L���O�w���e�[�u��","101","SQL�F<BR>"&StrSQL
      'end if
      'If ObjRS("Num")>0 Then  '�폜���ꂽ�f�[�^�����݂��Ă���ꍇ
          '���f�[�^���폜�f�[�^�Ƃ��ĕۑ�
          StrSQL = "UPDATE BookingAssign SET  "&_
                   "UpdtTime='"& Now() &"',UpdtPgCd='PREDEF01',UpdtTmnl='"& USER &"',"&_
                   "Process='D' WHERE BookNo='"& BookNo &"' AND SenderCode='"& COMPcd0 &"' AND TruckerCode='"& oldCOMPcd1 &"'"
          ObjConn.Execute(StrSQL)
          if err <> 0 then
            Set ObjRS = Nothing
            jampErrerPDB ObjConn,"1","b302","04","����o�F�f�[�^�o�^","104","SQL:<BR>"&StrSQL
          end if
      'Else
      '    '���f�[�^���폜
	  ' StrSQL = "DELETE FROM BookingAssign "
	  ' StrSQL = StrSQL & "WHERE BookNo='"& BookNo &"' AND SenderCode='"& COMPcd0 &"' AND TruckerCode='"& oldCOMPcd1 &"'"
	  '  ObjConn.Execute(StrSQL)
	  '	  if err <> 0 then
	  '	    Set ObjRS = Nothing
	  '	    jampErrerPDB ObjConn,"1","b302","04","����o�F�f�[�^�폜","104","SQL:<BR>"&StrSQL
	  '	  end if
      'End If
      'ObjRS.close
      '2016/11/29 H.Yoshikawa Del End
      
      '���̓f�[�^�ۑ�
      ret = InsBookAssign()					'2016/08/29 H.Yoshikawa Add
      oldObjRS.close
'2016/08/29 H.Yoshikawa Add End
  Else		'��
    ret=true
'C-002 ADD Start
    dim tmpstr
    tmpstr = ""
    If Request("Res") = 1 Then
       tmpstr = ", Comment='"& Request("Comment") &"' "
    End If
'C-002 ADD End
'    StrSQL = "UPDATE SPBookInfo SET UpdtTime='"& Now() &"', UpdtPgCd='PREDEF01', "&_
'             "UpdtTmnl='"& USER &"', Status='0', Process='R', TruckerFlag='"&Request("Res")&"' "&_
'             tmpstr & "WHERE BookNo='"& BookNo &"' "
'C-002 ADD This Line: Comment&_
'    ObjConn.Execute(StrSQL)
'    if err <> 0 then
'      Set ObjRS = Nothing
'      jampErrerPDB ObjConn,"1","b302","04","����o�F�ۗ�","102","SQL:<BR>"&StrSQL
'    end if
   StrSQL = "UPDATE BookingAssign SET  "&_
            "UpdtTime='"& Now() &"',UpdtPgCd='PREDEF01',UpdtTmnl='"& USER &"',"&_
            "TruckerFlag='"&Request("Res")&"'"&_
            "WHERE BookNo='"& BookNo &"' AND SenderCode='"& COMPcd0 &"' AND TruckerCode='"& COMPcd1 &"'"
    ObjConn.Execute(StrSQL)
    if err <> 0 then
      Set ObjRS = Nothing
      jampErrerPDB ObjConn,"1","b302","04","����o�F�ۗ�","102","SQL:<BR>"&StrSQL
    end if
    
  End If
  
  SendUser = COMPcd1
  
 'DB�ڑ�����
  DisConnDBH ObjConn, ObjRS
  
 '�f�[�^�擾
	Dim Email1, Email2, Email3, Email4, Email5
	Dim UserName,ComInterval,rc
	Dim ObjRS2,ObjConn2
	ConnDBH ObjConn2, ObjRS2

	'''�ʐM�Ԋu�擾
	StrSQL = "SELECT ComInterval FROM mParam WHERE Seq = '1'"

	ObjRS2.Open StrSQL, ObjConn2
	if err <> 0 then
	'''DB�ؒf
		DisConnDBH ObjConn2, ObjRS2
		jampErrerPDB ObjConn2,"1","b10"&(2+Flag),"16","����o�F���[�����M","104","SQL:<BR>"&StrSQL
	end if

	ComInterval = ObjRS2("ComInterval")
	ObjRS2.Close
	
	if SendUser <> "" then
	''��Ɣ����z�M���̎擾
		StrSQL = "SELECT T.*, "
		StrSQL = StrSQL & "CASE WHEN U.NameAbrev IS NULL THEN U.FullName ELSE U.NameAbrev END AS USERNAME "
		StrSQL = StrSQL & "FROM mUsers U, "
		StrSQL = StrSQL & "(SELECT T.* FROM TargetOperation T, mUsers U WHERE T.UserCode = U.UserCode "
		StrSQL = StrSQL & "AND U.HeadCompanyCode ='" & SendUser & "') T "
		StrSQL = StrSQL & "WHERE U.UserCode = '" & USER & "'"
		
		ObjRS2.Open StrSQL, ObjConn2
		if err <> 0 then
	'''DB�ؒf
			DisConnDBH ObjConn2, ObjRS2
			jampErrerPDB ObjConn2,"1","b10"&(2+Flag),"16","����o�F���[�����M","104","SQL:<BR>"&StrSQL
		end if

		Dim svName, mailTo, mailFrom, attachedFiles, ObjMail
		Dim mailFlag1, mailFlag2, mailFlag3, mailFlag4, mailFlag5
		Dim mailSubject, mailBody,WorkName
		Dim SendTime, UpdateSendTime
		Dim fp, fobj, tfile
		
' 2009/03/10 R.Shibuta Add-S
	'''SMTP�T�[�o���̐ݒ�
		svName   = "slitdns2.hits-h.com"
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

		if Trim(ObjRS2("Email1")) <> "" AND ObjRS2("FlagDelEmp1") = "1" then		
			mailTo = mailTo & Trim(ObjRS2("Email1"))
			mailFlag1 = 1
			
		else
			mailFlag1 = 0
		end if

		if Trim(ObjRS2("Email2")) <> "" AND ObjRS2("FlagDelEmp2") = "1" then
			if mailFlag1 = 1 then
				mailTo = mailTo & vbtab & Trim(ObjRS2("Email2"))
			else
				mailTo = mailTo & Trim(ObjRS2("Email2"))
			end if
				mailFlag2 = 1
		else
			mailFlag2 = 0
		end if

		if Trim(ObjRS2("Email3")) <> "" AND ObjRS2("FlagDelEmp3") = "1" then
			if mailFlag1 = 1 or mailFlag2 = 1 then
				mailTo = mailTo & vbtab & Trim(ObjRS2("Email3"))
			else
				mailTo = mailTo & Trim(ObjRS2("Email3"))
			end if
			mailFlag3 = 1
		else
			mailFlag3 = 0
		end if

		if Trim(ObjRS2("Email4")) <> "" AND ObjRS2("FlagDelEmp4") = "1" then
			if mailFlag1 = 1 or mailFlag2 = 1 or mailFlag3 = 1 then
				mailTo = mailTo & vbtab & Trim(ObjRS2("Email4"))
			else
				mailTo = mailTo & Trim(ObjRS2("Email4"))
			end if
			mailFlag4 = 1
		else
			mailFlag4 = 0
		end if

		if Trim(ObjRS2("Email5")) <> "" AND ObjRS2("FlagDelEmp5") = "1" then
			if mailFlag1 = 1 or mailFlag2 = 1 or mailFlag3 = 1 or mailFlag4 = 1 then
				mailTo = mailTo & vbtab & Trim(ObjRS2("Email5"))
			else
				mailTo = mailTo & Trim(ObjRS2("Email5"))
			end if
			mailFlag5 = 1
		else
			mailFlag5 = 0
		end if

		Set ObjMail = Server.CreateObject("BASP21")

		mailSubject = "HiTS ��ƈ˗�"
		mailBody = "����o���" & "���� (" & Trim(ObjRS2("USERNAME")) & "�l���)" & vbCrLf & vbCrLf
		mailBody = mailBody & "����o���" & "���������܂����B" & vbCrLf
		mailBody = mailBody & "�ڂ�����HiTS�̎��O���o�^�̉�ʂ����Q�Ɖ������B"
			
		'���[�����M�������猻�݂̎������ʐM�Ԋu�ȏ�̏ꍇ�̓��[���𑗐M����B
		
		if Trim(mailTo) <> "" Then

			if ObjRS2("DelEmpDate") < DateAdd("n",(ComInterval * -1), Now()) OR IsNull(ObjRS2("DelEmpDate")) = True then
				rc=ObjMail.Sendmail(svName, mailTo, mailFrom, mailSubject, mailBody, attachedFiles)
				sendTime=Now
			end if

			If rc = "" Then
				'''���[�����M���t�̍X�V���s���B
				StrSQL = "UPDATE TargetOperation SET UpdtTime='" & Now() & "', UpdtPgCd='dmi240',"
				StrSQL = StrSQL & " UpdtTmnl='" & USER & "',"&  "DelEmpDate='" & Now() & "'"
				StrSQL = StrSQL &"WHERE UserCode = '" & Trim(ObjRS2("UserCode")) & "'"

				ObjConn2.Execute(StrSQL)
				if err <> 0 then
					Set ObjRS2 = Nothing
					jumpErrorPDB ObjConn2,"1","c104","14","����o�F���[�����M","104","SQL:<BR>"&StrSQL
				end if
			else
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
					ret = 2
				end if
			end if
		else

		end if
' 2009/03/10 R.Shibuta Add-S
	end if
	'ObjRS.close
	
'DB�ڑ�����
  DisConnDBH ObjConn2, ObjRS2
'�G���[�g���b�v����
  on error goto 0
  
'2016/08/29 H.Yoshikawa Add Start
Function InsBookAssign()
  dim pickTime
  dim OutFlag
  dim MailFlag
  dim ChgFlag
  dim NewFlag
  dim pickHM
  dim i
  dim Seq
  dim Qty1
  dim Operator
  
	InsBookAssign = false

  on error resume next

	SendUser = COMPcd1
	
	'����L�[�ō폜�ς݂̃f�[�^���폜
    StrSQL = "DELETE FROM BookingAssign "
    StrSQL = StrSQL & "WHERE BookNo='"& BookNo &"' AND SenderCode='"& COMPcd0 &"' AND TruckerCode='"& COMPcd1 &"'"
    ObjConn.Execute(StrSQL)
    if err <> 0 then
      Set ObjRS = Nothing
      jampErrerPDB ObjConn,"1","b302","04","����o�F�f�[�^�폜","104","SQL:<BR>"&StrSQL
    end if

	if gfTrim(Request("MailFlag")) = "1" then
		MailFlag = "1"
	else
		MailFlag = "0"
	end if
	
	'�o�^����
    for i = 0 to 4
		if gfTrim(Request("ContSize" & i)) <> "" then
			NewFlag = "1"				'�V�K�o�^
			if Request("Mord") = 1 Then				'�X�V���[�h
				oldObjRS.MoveFirst
				Do Until oldObjRS.eof
					if gfTrim(oldObjRS("Seq")) = "" then
						Seq = 1
					else
						Seq = CInt(gfTrim(oldObjRS("Seq")))
					end if
					if Seq = i + 1 then
						OutFlag = gfTrim(oldObjRS("OutFlag"))
						if gfTrim(oldObjRS("ChgFlag")) = "" then
							ChgFlag = "0000000000000000"
						else
							ChgFlag = gfTrim(oldObjRS("ChgFlag"))
						end if
						
						'�ύX�t���O�ݒ�
						'�s�b�N�\���
						if gfTrim(oldObjRS("PickDateStr")) <> gfTrim(Request("PickDate" & i)) then
							ChgFlag = StrPartChg(ChgFlag, 1, "1")
							OutFlag = "0"
						end if
						'�T�C�Y
						if gfTrim(oldObjRS("ContSize1")) <> gfTrim(Request("ContSize" & i)) then
							ChgFlag = StrPartChg(ChgFlag, 2, "1")
							OutFlag = "0"
						end if
						'�^�C�v
						if gfTrim(oldObjRS("ContType1")) <> gfTrim(Request("ContType" & i)) then
							ChgFlag = StrPartChg(ChgFlag, 3, "1")
							OutFlag = "0"
						end if
						'����
						if gfTrim(oldObjRS("ContHeight1")) <> gfTrim(Request("ContHeight" & i)) then
							ChgFlag = StrPartChg(ChgFlag, 4, "1")
							OutFlag = "0"
						end if
						'�D��
						if gfTrim(oldObjRS("ShipLine")) <> gfTrim(Request("shipline")) then
							ChgFlag = StrPartChg(ChgFlag, 5, "1")
							OutFlag = "0"
						end if
						'�{��
						if gfTrim(oldObjRS("Qty1")) <> gfTrim(Request("PickNum" & i)) then
							ChgFlag = StrPartChg(ChgFlag, 6, "1")
							OutFlag = "0"
						end if
						'�{�D
						if gfTrim(oldObjRS("VslCode")) <> gfTrim(Request("VslCode")) then
							ChgFlag = StrPartChg(ChgFlag, 7, "1")
							OutFlag = "0"
						end if
						'���q
						if gfTrim(oldObjRS("Voyage")) <> gfTrim(Request("VoyCtrl")) then
							ChgFlag = StrPartChg(ChgFlag, 8, "1")
							OutFlag = "0"
						end if
						'�ݒ艷�x
						if Replace(gfTrim(oldObjRS("SetTemp")), "C", "") <> gfTrim(Request("SetTemp" & i)) then
							ChgFlag = StrPartChg(ChgFlag, 9, "1")
							OutFlag = "0"
						end if
						'��ڸ��
						if gfTrim(oldObjRS("Pcool")) <> gfTrim(Request("Pcool" & i)) then
							ChgFlag = StrPartChg(ChgFlag, 10, "1")
							OutFlag = "0"
						end if
						'�x���`���[�V����
						if gfTrim(oldObjRS("Ventilation")) <> gfTrim(Request("Ventilation" & i)) then
							ChgFlag = StrPartChg(ChgFlag, 11, "1")
							OutFlag = "0"
						end if
						'�s�b�N����
						if gfTrim(Request("PickHour" & i)) = "" then
							pickHM = ""
						else
							pickHM = gfTrim(Request("PickHour" & i)) & ":" & gfTrim(Request("PickMinute" & i))
						end if
						if gfTrim(oldObjRS("PickHM")) <> pickHM then
							ChgFlag = StrPartChg(ChgFlag, 12, "1")
							OutFlag = "0"
						end if
						'���l�P
						if gfTrim(oldObjRS("Comment1")) <> gfTrim(Request("Comment1")) then
							ChgFlag = StrPartChg(ChgFlag, 13, "1")
						end if
						'���l�Q
						if gfTrim(oldObjRS("Comment2")) <> gfTrim(Request("Comment2")) then
							ChgFlag = StrPartChg(ChgFlag, 14, "1")
						end if
						'���͎�
						if gfTrim(oldObjRS("TruckerSubName")) <> gfTrim(Request("TruckerSubName")) then
							ChgFlag = StrPartChg(ChgFlag, 15, "1")
						end if
						'�d�b�ԍ�
						if gfTrim(oldObjRS("Tel")) <> gfTrim(Request("Tel")) then
							ChgFlag = StrPartChg(ChgFlag, 16, "1")
						end if

						NewFlag = "0"
						Exit Do
					end if
					oldObjRS.MoveNext
				Loop
			end if
			
			if NewFlag = "1" then
				OutFlag = "0"
				ChgFlag = ""
			end if
			
			Operator = ""

			'�I�y���[�^�擾
			StrSQL = "SELECT Sender From Booking "
			StrSQL = StrSQL & "WHERE VslCode = '" & gfSQLEncode(Request("VslCode")) & "'"
			StrSQL = StrSQL & "  AND VoyCtrl = '" & gfSQLEncode(Request("VoyCtrl")) & "'"
			StrSQL = StrSQL & "  AND BookNo = '" & BookNo & "'"
			
			ObjRS.Open StrSQL, ObjConn
			if err <> 0 then
			'''DB�ؒf
				DisConnDBH ObjConn, ObjRS
				jampErrerP "1","b000","00","�u�b�L���O�e�[�u��","101","SQL�F<BR>"&StrSQL
			end if
			if not ObjRS.eof then
				select case gfTrim(ObjRS("Sender"))
					case "KACCS(�����`�^)"
						Operator = "HKK"
					case "KACCS(��g)"
						Operator = "KAM"
					case "KACCS(�W�F�l�b�N)"
						Operator = "KTC"
					case "KACCS(�O�H�q��)"
						Operator = "MLC"
					case "KACCS(���{�ʉ^)"
						Operator = "NEC"
					case "KACCS(���݉^�A)"
						Operator = "SOG"
				end select
			end if
			ObjRS.Close
		
			if gfTrim(Request("PickNum" & i)) = "" then
				Qty1 = "NULL"
			else
				Qty1 = gfTrim(Request("PickNum" & i))
			end if
			
		    if gfTrim(Request("PickHour" & i)) <> "" then
		    	pickTime = Right("00" & Request("PickHour" & i), 2) & ":" & Right("00" & Request("PickMinute" & i), 2)
		    else
		    	pickTime = ""
		    end if
			
			Seq = i + 1
	        StrSQL = "Insert Into BookingAssign "
	        StrSQL =  StrSQL & "(BookNo "
	        StrSQL =  StrSQL & ",SenderCode "
	        StrSQL =  StrSQL & ",TruckerCode "
	        StrSQL =  StrSQL & ",Seq "
	        StrSQL =  StrSQL & ",UpdtTime "
	        StrSQL =  StrSQL & ",UpdtPgCd "
	        StrSQL =  StrSQL & ",UpdtTmnl "
	        StrSQL =  StrSQL & ",Process "
	        StrSQL =  StrSQL & ",InputDate "
	        StrSQL =  StrSQL & ",TruckerName "
	        StrSQL =  StrSQL & ",TruckerFlag "
	        StrSQL =  StrSQL & ",ContSize1 "
	        StrSQL =  StrSQL & ",ContType1 "
	        StrSQL =  StrSQL & ",ContHeight1 "
	        StrSQL =  StrSQL & ",PickPlace1 "
	        StrSQL =  StrSQL & ",Terminal "
	        StrSQL =  StrSQL & ",Qty1 "
	        StrSQL =  StrSQL & ",VanTime "
	        StrSQL =  StrSQL & ",VanPlace1 "
	        StrSQL =  StrSQL & ",VanPlace2 "
	        StrSQL =  StrSQL & ",GoodsName "
	        StrSQL =  StrSQL & ",Comment1 "
	        StrSQL =  StrSQL & ",Comment2 "
	        StrSQL =  StrSQL & ",ShipLine "
	        StrSQL =  StrSQL & ",TruckerSubName "
	        StrSQL =  StrSQL & ",SetTemp "
	        StrSQL =  StrSQL & ",Pcool "
	        StrSQL =  StrSQL & ",Ventilation "
	        StrSQL =  StrSQL & ",PickDate "
	        StrSQL =  StrSQL & ",PickHM "
	        StrSQL =  StrSQL & ",OutFlag "
	        StrSQL =  StrSQL & ",Tel "
	        StrSQL =  StrSQL & ",Mail "
	        StrSQL =  StrSQL & ",MailFlag "
	        StrSQL =  StrSQL & ",VslCode "
	        StrSQL =  StrSQL & ",Voyage "
	        StrSQL =  StrSQL & ",OPE "
	        StrSQL =  StrSQL & ",ChgFlag "
	        StrSQL =  StrSQL & ") VALUES ( "
	        StrSQL =  StrSQL & " '" & gfSQLEncode(BookNo) & "' "
	        StrSQL =  StrSQL & ",'" & gfSQLEncode(COMPcd0) & "' "
	        StrSQL =  StrSQL & ",'" & gfSQLEncode(COMPcd1) & "' "
	        StrSQL =  StrSQL & ", '" & Seq & "' "
	        StrSQL =  StrSQL & ",'" &  Now() & "' "
	        StrSQL =  StrSQL & ",'PREDEF01' "
	        StrSQL =  StrSQL & ",'" & gfSQLEncode(USER) & "' "
	        StrSQL =  StrSQL & ",'R' "
	        StrSQL =  StrSQL & ",'" & Now() & "' "
	        StrSQL =  StrSQL & ", " & FullName & " "
	        StrSQL =  StrSQL & ",'" & gfSQLEncode(tFlag) & "' "
	        StrSQL =  StrSQL & ",'" & gfSQLEncode(Request("ContSize" & i)) & "' "
	        StrSQL =  StrSQL & ",'" & gfSQLEncode(Request("ContType" & i)) & "' "
	        StrSQL =  StrSQL & ",'" & gfSQLEncode(Request("ContHeight" & i)) & "' "
	        StrSQL =  StrSQL & ",'" & gfSQLEncode(Request("PickPlace" & i)) & "' "
	        StrSQL =  StrSQL & ",'" & gfSQLEncode(Request("Terminal" & i)) & "' "
	        StrSQL =  StrSQL & "," & Qty1 & " "
	        StrSQL =  StrSQL & "," & VanTime & " "
	        StrSQL =  StrSQL & ",'" & gfSQLEncode(Request("vanPlace1")) & "' "
	        StrSQL =  StrSQL & ",'" & gfSQLEncode(Request("vanPlace2")) & "' "
	        StrSQL =  StrSQL & ",'" & gfSQLEncode(Request("goodsName")) & "' "
	        StrSQL =  StrSQL & ",'" & gfSQLEncode(Request("Comment1")) & "' "
	        StrSQL =  StrSQL & ",'" & gfSQLEncode(Request("Comment2")) & "' "
	        StrSQL =  StrSQL & ",'" & gfSQLEncode(Request("shipline")) & "' "
	        StrSQL =  StrSQL & ",'" & gfSQLEncode(Request("TruckerSubName")) & "' "
	        if gfTrim(Request("SetTemp" & i)) = "" then
	        	StrSQL =  StrSQL & ",'' "
	        else
	        	StrSQL =  StrSQL & ",'" & gfSQLEncode(Request("SetTemp" & i)) & "C' "
	        end if
	        StrSQL =  StrSQL & ",'" & gfSQLEncode(Request("Pcool" & i)) & "' "
	        StrSQL =  StrSQL & ",'" & gfSQLEncode(Request("Ventilation" & i)) & "' "
	        StrSQL =  StrSQL & ",'" & gfSQLEncode(Request("PickDate" & i)) & "' "
	        StrSQL =  StrSQL & ",'" & pickTime & "' "
	        StrSQL =  StrSQL & ",'" & OutFlag & "' "
	        StrSQL =  StrSQL & ",'" & gfSQLEncode(Request("Tel")) & "' "
	        StrSQL =  StrSQL & ",'" & gfSQLEncode(Request("Mail")) & "' "
	        StrSQL =  StrSQL & ",'" & MailFlag & "' "
	        StrSQL =  StrSQL & ",'" & gfSQLEncode(Request("VslCode")) & "' "
	        StrSQL =  StrSQL & ",'" & gfSQLEncode(Request("VoyCtrl")) & "' "
	        StrSQL =  StrSQL & ",'" & Operator & "' "
	        StrSQL =  StrSQL & ",'" & ChgFlag & "' "
	        StrSQL =  StrSQL & ") "
	        ObjConn.Execute(StrSQL)
	        if err <> 0 then
	          Set ObjRS = Nothing
	          jampErrerPDB ObjConn,"1","b302","04","����o�F�f�[�^�o�^","103","SQL:<BR>"&StrSQL
	        end if
		end if
	next
	
	InsBookAssign = true

end Function

'������̈ꕔ��ϊ�
'  ������aStr��aPos�Ԗڂ̕�����aSetChr�ɕϊ������������Ԃ�
Function StrPartChg(aStr, aPos, aSetChr)
	dim retStr
	
	retStr = ""
	
	for i = 1 to Len(aStr)
	    if i = aPos then
	    	retStr = retStr & aSetChr
	    else
			retStr = retStr & Mid(aStr, i, 1)
		end if	
	next
	StrPartChg = retStr
end Function

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>��o���s�b�N�o�^�E�X�V</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<SCRIPT language=JavaScript>
  function goNext(){
<% If ret Then %>
    try{
  <% If Request("Mord") = 0 Then %>
      window.opener.parent.List.location.href="./dmo210F.asp"
  <% Else %>
        window.opener.parent.DList.location.href="./dmo210L.asp"
        window.opener.parent.Top.location.href="./dmo210T.asp"
  <% End If %>
    }catch(e){}
  <% If Request("SijiF") = "Yes" Then %>
    document.dmi240F.submit();
  <% Else %>
    window.close();
  <% End If %>
<% End If %>
  }
</SCRIPT>
<BODY onLoad="goNext()">
<!-------------���O����o�o�^�E�X�V--------------------------->
<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%">
  <TR align=center><TD>
<% If ret Then %>
  <% If Request("Mord") = 0 Then %>
   �o�^���܂����B<BR>��ʂ͎����I�ɕ����܂��B
  <% Else %>
   �X�V���܂����B<BR>��ʂ͎����I�ɕ����܂��B
  <% End If %>
      </TD></TR>
<% Elseif ret = 2 then %>
   <P><DIV class=alert><%= ErrerM %></DIV></P>
   <INPUT type=button value="����" onClick="window.close()">
<% Else %>
   <P><DIV class=alert>�w��̃u�b�L���O�ԍ��͑��쒆�ɑ��҂ɂ���ēo�^����܂����B</DIV></P>
   <INPUT type=button value="����" onClick="window.close()">
<% End If %>
</TABLE>
<% If Request("SijiF") = "Yes" Then %>
<FORM name="dmi240F" action="./dmo291.asp" method="POST">
<INPUT type=hidden name="BookNo" value="<%=Request("BookNo")%>">
<INPUT type=hidden name="COMPcd0"  value="<%=COMPcd0%>">
<INPUT type=hidden name="COMPcd1"  value="<%=COMPcd1%>">
<INPUT type=hidden name="shipFact" value="<%=Request("shipFact")%>">
<INPUT type=hidden name="shipName" value="<%=Request("shipName")%>">
<INPUT type=hidden name="delivTo"  value="<%=Request("delivTo")%>">
  <% For i=0 To 4%>
  <INPUT type=hidden name="ContSize<%=i%>"   value="<%=Request("ContSize"&i)%>">
  <INPUT type=hidden name="ContType<%=i%>"   value="<%=Request("ContType"&i)%>">
  <INPUT type=hidden name="ContHeight<%=i%>" value="<%=Request("ContHeight"&i)%>">
  <INPUT type=hidden name="Material<%=i%>"   value="<%=Request("Material"&i)%>">
  <INPUT type=hidden name="PickPlace<%=i%>"  value="<%=Request("PickPlace"&i)%>">
  <INPUT type=hidden name="PickNum<%=i%>"    value="<%=Request("PickNum"&i)%>">
  <% Next %>
<INPUT type=hidden name="vanMon"    value="<%=Request("vanMon")%>">
<INPUT type=hidden name="vanDay"    value="<%=Request("vanDay")%>">
<INPUT type=hidden name="vanHou"    value="<%=Request("vanHou")%>">
<INPUT type=hidden name="vanMin"    value="<%=Request("vanMin")%>">
<INPUT type=hidden name="vanPlace1" value="<%=Request("VanPlace1")%>">
<INPUT type=hidden name="vanPlace2" value="<%=Request("VanPlace2")%>">
<INPUT type=hidden name="goodsName" value="<%=Request("GoodsName")%>">
<INPUT type=hidden name="Terminal"  value="<%=Request("Terminal")%>">
<INPUT type=hidden name="CYCut"    value="<%=Request("CYCut")%>">
<INPUT type=hidden name="Comment1"  value="<%=Request("Comment1")%>">
<INPUT type=hidden name="Comment2"  value="<%=Request("Comment2")%>">
</FORM>
<% End If %>
<!-------------��ʏI���--------------------------->
</BODY></HTML>
