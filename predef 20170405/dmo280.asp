<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmo280.asp				_/
'_/	Function	:��������ꗗCSV�t�@�C���_�E�����[�h	_/
'_/	Date		:2003/08/06				_/
'_/	Code By		:SEIKO Electric.Co ��d			_/
'_/	Modify		:3th   2004/01/31	3���Ή�	_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTP�R���e���c�^�C�v�ݒ�
'	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="./Common.inc"-->
<%
'�Z�b�V�����̗L�������`�F�b�N
  CheckLoginH
  WriteLogH "b308", "��������O���CSV�t�@�C���_�E�����[�h","01",""

'�f�[�^�擾
  dim Num,preDtTbl,i,j,ResA
  dim abspage
  Get_Data Num,preDtTbl
'Chenge 20030908  ResA=Array("�@","Yes","No","��")
  ResA=Array("","Yes","No","��")
'3th ADD Start ��������������
'�\���f�[�^�z��̏���
  dim DtTbl
  ReDim DtTbl(Num)
'���[�U�f�[�^�擾
  dim USER, COMPcd
  USER   = UCase(Session.Contents("userid"))
  COMPcd = Session.Contents("COMPcd")
'���������擾
  dim StrSQL,strWhere
  strWhere= Request("strWhere")
  abspage= Request("absPage")
'�G���[�g���b�v�J�n
 ' on error resume next
'DB�ڑ�
  dim ObjConn, ObjRS, ErrerM
  ConnDBH ObjConn, ObjRS


'�ǉ����̎擾���W�J�f�[�^�̎擾
  Dim Flag,tmp
  If Num>0 Then
'ADD 20050228 Fro survive ViewBookAssing ViewTable By SEIKO N.Oosige
'DEL 20100420 Start
 '   StrSQL = "IF (EXISTS( select * from ViewBookAssing ) ) BEGIN DROP VIEW ViewBookAssing END "
'    ObjConn.Execute(StrSQL)
'    if err <> 0 then
'      err.Clear
'    end if
'      
'    StrSQL = "BEGIN TRAN TRAN1 "
'    ObjConn.Execute(StrSQL)
'    if err <> 0 then
'      Set ObjRS = Nothing
'      jampErrerP "2","b301","01","����o�F�ꗗ�f�[�^�擾","101","SQL:<BR>"&strSQL
'    end if
''ADD 20050228 End
'    StrSQL = "CREATE VIEW ViewBookAssing AS SELECT Max(InputDate) AS MAXDATE,BookNo "&_
'             "FROM BookingAssign GROUP BY BookNo,Process "&_
'             "HAVING Process='R'"
'    ObjConn.Execute(StrSQL)
'    if err <> 0 then
'      DisConnDBH ObjConn, ObjRS	'DB�ؒf
'      jampErrerP "1","b208","01","����o�FCSV�t�@�C���_�E�����[�hmakeview","101","SQL:<BR>"&StrSQL
'    end if
'DEL 20100420 End
    
   StrSQL = "SELECT SPB.BookNo,SPB.SenderCode,SPB.TruckerCode, "&_
             "SPB.ContSize1,SPB.ContType1,SPB.ContHeight1,SPB.ContMaterial1,SPB.PickPlace1,SPB.Qty1, "&_
             "SPB.ContSize2,SPB.ContType2,SPB.ContHeight2,SPB.ContMaterial2,SPB.PickPlace2,SPB.Qty2, "&_
             "SPB.ContSize3,SPB.ContType3,SPB.ContHeight3,SPB.ContMaterial3,SPB.PickPlace3,SPB.Qty3, "&_
             "SPB.ContSize4,SPB.ContType4,SPB.ContHeight4,SPB.ContMaterial4,SPB.PickPlace4,SPB.Qty4, "&_
             "SPB.ContSize5,SPB.ContType5,SPB.ContHeight5,SPB.ContMaterial5,SPB.PickPlace5,SPB.Qty5, "&_
             "SPB.VanTime,SPB.VanPlace1,SPB.VanPlace2,SPB.GoodsName, "&_
             "BOK.RecTerminal, VSL.CYCut, mP.FullName "&_
             "FROM BookingAssign AS SPB "&_
             "LEFT JOIN ViewBookAssing AS VBA ON SPB.BookNO=VBA.BookNo "&_
             "LEFT JOIN Booking AS BOK ON SPB.BookNo = BOK.BookNo "&_
             "LEFT JOIN VslSchedule AS VSL ON (BOK.VoyCtrl = VSL.VoyCtrl) AND (BOK.VslCode = VSL.VslCode) "&_
             "LEFT JOIN mPort AS mP ON BOK.DelivPlace = mP.PortCode "&_
             "WHERE (SPB.SenderCode='"& USER &"' OR SPB.TruckerCode='"& COMPcd &"') AND SPB.Process='R' "&_			 
			 strWhere &_
             "ORDER BY VBA.MAXDATE DESC,SPB.InputDate DESC, SPB.BookNo ASC"
             '
'CW-320             "FROM ((BookingAssign AS SPB LEFT JOIN Booking AS BOK ON SPB.BookNo = BOK.BookNo) "&_
'CW-320             "LEFT JOIN mPort AS mP ON BOK.DelivPlace = mP.PortCode) "&_
'CW-320             "LEFT JOIN VslSchedule AS VSL ON (BOK.VoyCtrl = VSL.VoyCtrl) AND (BOK.VslCode = VSL.VslCode) "&_
'CW-320             "WHERE (SPB.SenderCode='"& USER &"' OR SPB.TruckerCode='"& COMPcd &"') AND SPB.Process='R' "&_
'CW-320              strWhere &_
'CW-320             "ORDER BY DATEPART(yy,SPB.InputDate) DESC,DATEPART(m,SPB.InputDate) DESC,DATEPART(d,SPB.InputDate) DESC, SPB.BookNo ASC, Bok.UpdtTime DESC"  
  ObjRS.PageSize = 200
  ObjRS.CacheSize = 200
  ObjRS.CursorLocation = 3	
  ObjRS.Open StrSQL, ObjConn  
  ObjRS.AbsolutePage = abspage

  if err <> 0 then
    DisConnDBH ObjConn, ObjRS	'DB�ؒf
    jampErrerP "1","b208","01","����o�FCSV�t�@�C���_�E�����[�h","101","SQL:<BR>"&StrSQL
  end if
  i=1
  j=0
  Flag=0
  '2009/10/09 Upd-S Tanaka
'  Do Until ObjRS.EOF
''Response.Write "i="&i&":j="&j&":"&preDtTbl(i)(2)&"="&Trim(ObjRS("BookNo"))&":"&preDtTbl(i)(1)&"="&Trim(ObjRS("SenderCode"))&":"&Trim(preDtTbl(i)(12))&"="&Trim(ObjRS("TruckerCode"))&"<P>"
''2009/02/25 Del-S G.Ariola
'    'If preDtTbl(i)(2)=Trim(ObjRS("BookNo")) AND preDtTbl(i)(1)=Trim(ObjRS("SenderCode")) AND Trim(preDtTbl(i)(12))=Trim(ObjRS("TruckerCode")) then
''2009/02/25 Del-E G.Ariola	
'      tmp=preDtTbl(i)
'      ReDim Preserve tmp(51)
'      tmp(14)= Trim(ObjRS("FullName"))
'      tmp(15)= Trim(ObjRS("ContSize1"))
'      tmp(16)= Trim(ObjRS("ContType1"))
'      tmp(17)= Trim(ObjRS("ContHeight1"))
'      tmp(18)= Trim(ObjRS("ContMaterial1"))
'      tmp(19)= Trim(ObjRS("PickPlace1"))
'      tmp(20)= Trim(ObjRS("Qty1"))
'      tmp(21)= Trim(ObjRS("ContSize2"))
'      tmp(22)= Trim(ObjRS("ContType2"))
'      tmp(23)= Trim(ObjRS("ContHeight2"))
'      tmp(24)= Trim(ObjRS("ContMaterial2"))
'      tmp(25)= Trim(ObjRS("PickPlace2"))
'      tmp(26)= Trim(ObjRS("Qty2"))
'      tmp(27)= Trim(ObjRS("ContSize3"))
'      tmp(28)= Trim(ObjRS("ContType3"))
'      tmp(29)= Trim(ObjRS("ContHeight3"))
'      tmp(30)= Trim(ObjRS("ContMaterial3"))
'      tmp(31)= Trim(ObjRS("PickPlace3"))
'      tmp(32)= Trim(ObjRS("Qty3"))
'      tmp(33)= Trim(ObjRS("ContSize4"))
'      tmp(34)= Trim(ObjRS("ContType4"))
'      tmp(35)= Trim(ObjRS("ContHeight4"))
'      tmp(36)= Trim(ObjRS("ContMaterial4"))
'      tmp(37)= Trim(ObjRS("PickPlace4"))
'      tmp(38)= Trim(ObjRS("Qty4"))
'      tmp(39)= Trim(ObjRS("ContSize5"))
'      tmp(40)= Trim(ObjRS("ContType5"))
'      tmp(41)= Trim(ObjRS("ContHeight5"))
'      tmp(42)= Trim(ObjRS("ContMaterial5"))
'      tmp(43)= Trim(ObjRS("PickPlace5"))
'      tmp(44)= Trim(ObjRS("Qty5"))
'      tmp(45)= Trim(ObjRS("VanTime"))
'      tmp(46)= Trim(ObjRS("VanPlace1"))
'      tmp(47)= Trim(ObjRS("VanPlace2"))
'      tmp(48)= Trim(ObjRS("GoodsName"))
'      tmp(49)= Trim(ObjRS("RecTerminal"))
'      tmp(50)= Trim(ObjRS("CYCut"))
'      DtTbl(j)=tmp
'      j=j+1
'      i=i+1
'	  
'      ObjRS.MoveNext
''2009/02/25 Del-S G.Ariola	  
'    'Else
'    '  ObjRS.MoveNext
'    'End If
''2009/02/25 Del-E G.Ariola	
'  Loop  
  For i=1 to ubound(preDtTbl)
    ObjRS.MoveFirst
    Do Until ObjRS.EOF	  
      If preDtTbl(i)(2)=Trim(ObjRS("BookNo")) AND preDtTbl(i)(1)=Trim(ObjRS("SenderCode")) AND Trim(preDtTbl(i)(12))=Trim(ObjRS("TruckerCode")) then
        tmp=preDtTbl(i)
        ReDim Preserve tmp(51)
        tmp(14)= Trim(ObjRS("FullName"))
        tmp(15)= Trim(ObjRS("ContSize1"))
        tmp(16)= Trim(ObjRS("ContType1"))
        tmp(17)= Trim(ObjRS("ContHeight1"))
        tmp(18)= Trim(ObjRS("ContMaterial1"))
        tmp(19)= Trim(ObjRS("PickPlace1"))
        tmp(20)= Trim(ObjRS("Qty1"))
        tmp(21)= Trim(ObjRS("ContSize2"))
        tmp(22)= Trim(ObjRS("ContType2"))
        tmp(23)= Trim(ObjRS("ContHeight2"))
        tmp(24)= Trim(ObjRS("ContMaterial2"))
        tmp(25)= Trim(ObjRS("PickPlace2"))
        tmp(26)= Trim(ObjRS("Qty2"))
        tmp(27)= Trim(ObjRS("ContSize3"))
        tmp(28)= Trim(ObjRS("ContType3"))
        tmp(29)= Trim(ObjRS("ContHeight3"))
        tmp(30)= Trim(ObjRS("ContMaterial3"))
        tmp(31)= Trim(ObjRS("PickPlace3"))
        tmp(32)= Trim(ObjRS("Qty3"))
        tmp(33)= Trim(ObjRS("ContSize4"))
        tmp(34)= Trim(ObjRS("ContType4"))
        tmp(35)= Trim(ObjRS("ContHeight4"))
        tmp(36)= Trim(ObjRS("ContMaterial4"))
        tmp(37)= Trim(ObjRS("PickPlace4"))
        tmp(38)= Trim(ObjRS("Qty4"))
        tmp(39)= Trim(ObjRS("ContSize5"))
        tmp(40)= Trim(ObjRS("ContType5"))
        tmp(41)= Trim(ObjRS("ContHeight5"))
        tmp(42)= Trim(ObjRS("ContMaterial5"))
        tmp(43)= Trim(ObjRS("PickPlace5"))
        tmp(44)= Trim(ObjRS("Qty5"))
        tmp(45)= Trim(ObjRS("VanTime"))
        tmp(46)= Trim(ObjRS("VanPlace1"))
        tmp(47)= Trim(ObjRS("VanPlace2"))
        tmp(48)= Trim(ObjRS("GoodsName"))
        tmp(49)= Trim(ObjRS("RecTerminal"))
        tmp(50)= Trim(ObjRS("CYCut"))
        DtTbl(j)=tmp
        j=j+1

        Exit Do

      End If
      ObjRS.MoveNext
    Loop

  Next
'2009/10/09 Upd-E Tanaka
  ObjRS.close
'Change 20050228 Fro survive ViewBookAssing ViewTable By SEIKO N.Oosige
'   StrSQL = "DROP VIEW ViewBookAssing"
'DEL 20100420 Start
'  StrSQL = "COMMIT TRAN TRAN1 "
'Change 20050228 End
'  ObjConn.Execute(StrSQL)
'  if err <> 0 then
'    Set ObjRS = Nothing
'    jampErrerP "1","b208","01","����o�FCSV�t�@�C���_�E�����[�hDoropview","101","SQL:<BR>"&StrSQL
'  end if
'DEL 20100420 END
  End If
'DB�ڑ�����
  DisConnDBH ObjConn, ObjRS
'�G���[�g���b�v����
  on error goto 0

'3th ADD END ��������������

' �t�@�C���̃_�E�����[�h
  Response.ContentType="application/octet-stream"
  Response.AddHeader "Content-Disposition","attachment; filename=output.csv"

'3th Chage    Response.Write "���͓�,�w����,�w�����ւ̉�,�u�b�L���O�ԍ�,�D��,�D��,�s�b�N�ϖ{��,�w����,�w�����,���l"
'3th Chage    Response.Write Chr(13) & Chr(10)
'3th Chage    For j=1 To Num
'3th Chage      Response.Write Trim(DtTbl(j)(0))&","&Trim(DtTbl(j)(10))&","&ResA(DtTbl(j)(6))&","&Trim(DtTbl(j)(2))&","
'3th Chage      Response.Write Trim(DtTbl(j)(8))&","&Trim(DtTbl(j)(9))&","&Trim(DtTbl(j)(3))&","
'3th Chage      Response.Write Trim(DtTbl(j)(4))&","&Trim(DtTbl(j)(5))&","&Trim(DtTbl(j)(11))
'3th Chage      Response.Write Chr(13) & Chr(10)
'3th Chage    Next
    Response.Write "���͓�,�w����,�w�����ւ̉�,�u�b�L���O�ԍ�,�D��,�D��,�d���n,"
    Response.Write "�T�C�Y�P,�^�C�v�P,�����P,�ގ��P,�s�b�N�ꏊ�P,�{���P,"
    Response.Write "�T�C�Y�Q,�^�C�v�Q,�����Q,�ގ��Q,�s�b�N�ꏊ�Q,�{���Q,"
    Response.Write "�T�C�Y�R,�^�C�v�R,�����R,�ގ��R,�s�b�N�ꏊ�R,�{���R,"
    Response.Write "�T�C�Y�S,�^�C�v�S,�����S,�ގ��S,�s�b�N�ꏊ�S,�{���S,"
    Response.Write "�T�C�Y�T,�^�C�v�T,�����T,�ގ��T,�s�b�N�ꏊ�T,�{���T,"
    Response.Write "�o���l�ߓ���,�o���l�ߏꏊ�P,�o���l�ߏꏊ�Q,�i��,������b�x,�b�x�J�b�g��,"
    Response.Write "�s�b�N�ϖ{��,�w����,�w�����,���l�P,���l�Q"
    Response.Write Chr(13) & Chr(10)

    For j=0 To Num-1
      Response.Write Trim(DtTbl(j)(0))&","&Trim(DtTbl(j)(10))&","&ResA(DtTbl(j)(6))&","
      Response.Write Trim(DtTbl(j)(2))&","&Trim(DtTbl(j)(8))&","&Trim(DtTbl(j)(9))&","&Trim(DtTbl(j)(14))&","
      Response.Write DtTbl(j)(15)&","&DtTbl(j)(16)&","&DtTbl(j)(17)&","&DtTbl(j)(18)&","&DtTbl(j)(19)&","&DtTbl(j)(20)&","
      Response.Write DtTbl(j)(21)&","&DtTbl(j)(22)&","&DtTbl(j)(23)&","&DtTbl(j)(24)&","&DtTbl(j)(25)&","&DtTbl(j)(26)&","
      Response.Write DtTbl(j)(27)&","&DtTbl(j)(28)&","&DtTbl(j)(29)&","&DtTbl(j)(30)&","&DtTbl(j)(31)&","&DtTbl(j)(32)&","
      Response.Write DtTbl(j)(33)&","&DtTbl(j)(34)&","&DtTbl(j)(35)&","&DtTbl(j)(36)&","&DtTbl(j)(37)&","&DtTbl(j)(38)&","
      Response.Write DtTbl(j)(39)&","&DtTbl(j)(40)&","&DtTbl(j)(41)&","&DtTbl(j)(42)&","&DtTbl(j)(43)&","&DtTbl(j)(44)&","
      Response.Write DtTbl(j)(45)&","&DtTbl(j)(46)&","&DtTbl(j)(47)&","&DtTbl(j)(48)&","&DtTbl(j)(49)&","&DtTbl(j)(50)&","
      Response.Write Trim(DtTbl(j)(3))&","&Trim(DtTbl(j)(4))&","&Trim(DtTbl(j)(5))&","&Trim(DtTbl(j)(11))&","&Trim(DtTbl(j)(13))
      Response.Write Chr(13) & Chr(10)
    Next
  Response.End

%>
