<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmo080.asp				_/
'_/	Function	:�����o���ꗗCSV�t�@�C���_�E�����[�h	_/
'_/	Date		:2003/07/29				_/
'_/	Code By		:SEIKO Electric.Co ��d			_/
'_/	Modify		:3th   2004/01/31	3���Ή�	_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<!--#include File="./Common.inc"-->
<%
'�Z�b�V�����̗L�������`�F�b�N
  CheckLoginH
  WriteLogH "b108", "�����o���O���CSV�t�@�C���_�E�����[�h","01",""

'�f�[�^�擾
  dim preNum, preDtTbl,Num,DtTbl,Siji,i,j
  dim StrSQL,strWhere
  strWhere= Request("strWhere")
'���[�U�f�[�^�擾
  dim USER, COMPcd
  USER   = UCase(Session.Contents("userid"))
  COMPcd = Session.Contents("COMPcd")
'INI�t�@�C�����ݒ�l���擾
  dim param(2)
  getIni param

  Get_Data Num,preDtTbl
  preNum=Num
  Siji  =Array("","�w�肠��","�w��Ȃ�","�ꗗ","�a�k")

  ReDim DtTbl(Num+1)
'3th del  DtTbl(0)=preDtTbl(0)
  '�G���[�g���b�v�J�n
    on error resume next
  'DB�ڑ�
    dim ObjConn, ObjRS, ErrerM
    ConnDBH ObjConn, ObjRS

'3th Chage Steart
'�ǉ����̎擾���W�J�f�[�^�̎擾
    dim tmpNasiConNo,tmpItiConNo,tmpBLNo
    tmpNasiConNo =Array("","","")
    tmpItiConNo  =Array("","")
    tmpBLNo      =Array("","")

    Dim Flag,tmp
    StrSQL = "SELECT ITC.FullOutType,ITC.WkNo,ITC.WorkDate,ITC.HeadID,ITC.GoodsName,ITC.DeliverTo2,ITC.DeliverDate, "&_
             "INC.ReturnPlace, Cnt.ContType, Cnt.ContHeight, Cnt.ContWeight "&_
             "FROM hITCommonInfo AS ITC "&_
			 "LEFT JOIN ImportCont AS INC ON ITC.ContNo =INC.ContNo "&_
             "LEFT JOIN Container AS Cnt ON INC.ContNo =Cnt.ContNo AND INC.VslCode =Cnt.VslCode AND INC.VoyCtrl =Cnt.VoyCtrl "&_
             "WHERE WkType='1' AND (ITC.RegisterCode='"& USER &"' " &_
             "OR ITC.TruckerSubCode1='"& COMPcd &"' OR ITC.TruckerSubCode2='"& COMPcd &"'" &_
             "OR ITC.TruckerSubCode3='"& COMPcd &"' OR ITC.TruckerSubCode4='"& COMPcd &"') AND Process='R' "&_
             strWhere & "ORDER BY isnull(ITC.WorkDate,DATEADD(Year,100,getdate())),ITC.InputDate ASC"
    
    ObjRS.Open StrSQL, ObjConn
    if err <> 0 then
      DisConnDBH ObjConn, ObjRS	'DB�ؒf
      jampErrerP "1","b108","01","�����o�FCSV�t�@�C���_�E�����[�h","101","SQL:<BR>"&StrSQL
    end if
    i=1
    j=0
    Flag=0
    '2009/10/09 Upd-S Tanaka
'    Do Until ObjRS.EOF
'      If i>Num+0 Then		'CW-309 add
'        Exit Do				'CW-309 add
'      End If				'CW-309 add
''change 20050530 Fro CSV�o�͂ɏo�͂���Ȃ����R�[�h�����݂���s�
''      If preDtTbl(i)(3)=Trim(ObjRS("WkNo")) then
'      If preDtTbl(i-1)(3)<>Trim(ObjRS("WkNo")) then
'        tmp=preDtTbl(i)
'        ReDim Preserve tmp(32)
'        tmp(1) =ObjRS("WorkDate")				'���o�\�����
'        If  preDtTbl(i)(10) = "Yes" Then
'          tmp(25)="*****"
'        Else
'          tmp(25)=Trim(ObjRS("HeadID"))		'�w�b�h�h�c
'        End If
'        tmp(29)=Trim(ObjRS("GoodsName"))		'�i��
'        tmp(30)=Trim(ObjRS("DeliverTo2"))		'�[����Q
'        tmp(31)=ObjRS("DeliverDate")			'�[������
'        tmp(4) =ObjRS("FullOutType")		'�w���^�C�v
'        If tmp(4) <>"1" Then
'          If tmp(4) ="2" Then
'            '�w��Ȃ��̏ꍇ�ABL�ԍ��ƃR���e�i�ԍ���ێ�
'            tmpNasiConNo(0) = tmpNasiConNo(0) &","& j
'            tmpNasiConNo(1) = tmpNasiConNo(1) &",'"& preDtTbl(i)(5) & "'"
'            tmpNasiConNo(2) = tmpNasiConNo(2) &",'"& preDtTbl(i)(11) & "'"
'          ElseIf tmp(4) ="3" Then
'            '�ꗗ�̏ꍇ�A��ƊǗ��ԍ���ێ�
'            tmpItiConNo(0) = tmpItiConNo(0) &","& j
'            tmpItiConNo(1) = tmpItiConNo(1) &",'"& preDtTbl(i)(3) & "'"
'          Else
'            'BL�̏ꍇ�ABL�ԍ���ێ�
'            tmpBLNo(0) = tmpBLNo(0) &","& j
'            tmpBLNo(1) = tmpBLNo(1) &",'"& preDtTbl(i)(11) & "'"
'          End If
'        Else
'          tmp(26)=ObjRS("ContType")				'�^�C�v
'          tmp(27)=ObjRS("ContHeight")			'����
'          tmp(28)=ObjRS("ContWeight")*100		'�O���X
'          tmp(32)=Trim(ObjRS("ReturnPlace"))	'��o���ԋp��
'        End If
'        DtTbl(j)=tmp
'        j=j+1
'        i=i+1
'        Flag=0
'        ObjRS.MoveNext
'      Else
''change 20050530 Fro CSV�o�͂ɏo�͂���Ȃ����R�[�h�����݂���s�
''        If Flag=0 Then
''          Flag=1
''          i=i+1
''        Else
''          ObjRS.MoveNext
''        End If
'         ObjRS.MoveNext
'      End If
'    Loop
'    ObjRS.close

    IF ObjRS.EOF=False Then
        For i=1 To Ubound(preDtTbl)
            ObjRS.MoveFirst
            Do Until ObjRS.EOF
			    'Y.TAKAKUWA Upd-S 2015-01-26
				'Y.TAKAKUWA Upd-S 2015-01-30
			    'IF preDtTbl(i)(3)=Trim(ObjRS("WkNo")) then
				IF preDtTbl(i)(3)=Trim(ObjRS("WkNo")) then
				'Y.TAKAKUWA Upd-E 2015-01-30
				'Y.TAKAKUWA Del-S 2015-01-27
                'IF preDtTbl(i-1)(3)<>Trim(ObjRS("WkNo")) then
				'Y.TAKAKUWA Del-E 2015-01-27
				'Y.TAKAKUWA Upd-E 2015-01-26
                    tmp=preDtTbl(i)
                    ReDim Preserve tmp(32)
                    tmp(1) =ObjRS("WorkDate")				'���o�\�����
                    If  preDtTbl(i)(10) = "Yes" Then
                        tmp(25)="*****"
                    Else
                        tmp(25)=Trim(ObjRS("HeadID"))		'�w�b�h�h�c
                    End If
                    tmp(29)=Trim(ObjRS("GoodsName"))		'�i��
                    tmp(30)=Trim(ObjRS("DeliverTo2"))		'�[����Q
                    tmp(31)=ObjRS("DeliverDate")			'�[������
                    tmp(4) =ObjRS("FullOutType")		'�w���^�C�v
                    If tmp(4) <>"1" Then
                        If tmp(4) ="2" Then
                            '�w��Ȃ��̏ꍇ�ABL�ԍ��ƃR���e�i�ԍ���ێ�
                            tmpNasiConNo(0) = tmpNasiConNo(0) &","& j
                            tmpNasiConNo(1) = tmpNasiConNo(1) &",'"& preDtTbl(i)(5) & "'"
                            tmpNasiConNo(2) = tmpNasiConNo(2) &",'"& preDtTbl(i)(11) & "'"
                        ElseIf tmp(4) ="3" Then
                            '�ꗗ�̏ꍇ�A��ƊǗ��ԍ���ێ�
                            tmpItiConNo(0) = tmpItiConNo(0) &","& j
                            tmpItiConNo(1) = tmpItiConNo(1) &",'"& preDtTbl(i)(3) & "'"
                        Else
                            'BL�̏ꍇ�ABL�ԍ���ێ�
                            tmpBLNo(0) = tmpBLNo(0) &","& j
                            tmpBLNo(1) = tmpBLNo(1) &",'"& preDtTbl(i)(11) & "'"
                        End If
                    Else
                        tmp(26)=ObjRS("ContType")				'�^�C�v
                        tmp(27)=ObjRS("ContHeight")			'����
                        tmp(28)=ObjRS("ContWeight")*100		'�O���X
                        tmp(32)=Trim(ObjRS("ReturnPlace"))	'��o���ԋp��
                    End If
                    DtTbl(j)=tmp
                    j=j+1
                    Flag=0
                    
                    Exit Do
				'Y.TAKAKUWA Del-S 2015-01-27
				'Y.TAKAKUWA Upd-S 2015-01-30
                'Else
                'End If
				Else
                End If
				'Y.TAKAKUWA Upd-E 2015-01-30
				'Y.TAKAKUWA Del-E 2015-01-27
                ObjRS.MoveNext
            Loop
        Next
    End If
    ObjRS.close
    '2009/10/09 Upd-E Tanaka
	
	Num=j-1
	
'�R�t���f�[�^�̎擾
    dim tmpNum,tmpNasiConNoA,tmpItiConNoA,tmpBLNoA
    ReDim tmpNasiConNoA(2)
    ReDim tmpItiConNoA(2)
    ReDim tmpBLNoA(2)
    '�w��Ȃ��̕R�t�����擾
    If tmpNasiConNo(0) <> "" Then

'2009/10/02 Upd-S Fujiyama
'      StrSQL="SELECT COUNT(*) AS Num "&_
'             "FROM ImportCont AS INC1 INNER JOIN ImportCont AS INC2 ON (INC1.VslCode = INC2.VslCode) AND (INC1.VoyCtrl = INC2.VoyCtrl) AND (INC1.BLNo = INC2.BLNo) "&_
'             "WHERE INC1.ContNo IN("&Mid(tmpNasiConNo(1),2)&") AND INC1.BLNo IN("&Mid(tmpNasiConNo(2),2)&") "
      StrSQL="SELECT COUNT(*) AS Num "&_
             "FROM ImportCont AS INC1 "&_
             "WHERE INC1.ContNo IN("&Mid(tmpNasiConNo(1),2)&") AND INC1.BLNo IN("&Mid(tmpNasiConNo(2),2)&") "
'2009/10/02 Upd-E Fujiyama
      ObjRS.Open StrSQL, ObjConn
      if err <> 0 then
        DisConnDBH ObjConn, ObjRS	'DB�ؒf
        jampErrerP "1","b108","01","�����o�FCSV�t�@�C���_�E�����[�h(�w��Ȃ��ǉ�)","102","SQL:<BR>"&StrSQL
      end if
      tmpNum=ObjRS("Num")+2
      ObjRS.close
      ReDim Preserve tmpNasiConNoA(tmpNum)
      tmpNasiConNoA(0)=Split(tmpNasiConNo(0), ",", -1, 1)
      tmpNasiConNoA(1)=Split(tmpNasiConNo(2), ",", -1, 1)
'2009/10/02 Upd-S Fujiyama
'      StrSQL="SELECT INC2.ContNo, INC2.BLNo, INC2.ReturnTime, INC2.CYDelTime, INC2.ReturnPlace, "&_
'             "Cnt.ContSize, Cnt.ContType, Cnt.ContHeight, Cnt.ContWeight "&_
'             "FROM ImportCont AS INC1 INNER JOIN ImportCont AS INC2 ON INC1.VslCode=INC2.VslCode AND INC1.VoyCtrl=INC2.VoyCtrl AND INC1.BLNo=INC2.BLNo "&_
'             "LEFT JOIN Container AS Cnt ON INC2.ContNo =Cnt.ContNo AND INC2.VslCode =Cnt.VslCode AND INC2.VoyCtrl =Cnt.VoyCtrl "&_
'             "WHERE INC1.ContNo IN("&Mid(tmpNasiConNo(1),2)&") AND INC1.BLNo IN("&Mid(tmpNasiConNo(2),2)&") "&_
'             "ORDER BY INC1.BLNo ASC"
      StrSQL="SELECT INC1.ContNo, INC1.BLNo, INC1.ReturnTime, INC1.CYDelTime, INC1.ReturnPlace, "&_
             "Cnt.ContSize, Cnt.ContType, Cnt.ContHeight, Cnt.ContWeight "&_
             "FROM ImportCont AS INC1 "&_
             "LEFT JOIN Container AS Cnt ON INC1.ContNo =Cnt.ContNo AND INC1.VslCode =Cnt.VslCode AND INC1.VoyCtrl =Cnt.VoyCtrl "&_
             "WHERE INC1.ContNo IN("&Mid(tmpNasiConNo(1),2)&") AND INC1.BLNo IN("&Mid(tmpNasiConNo(2),2)&") "&_
             "ORDER BY INC1.BLNo ASC"
'2009/10/02 Upd-E Fujiyama
      
      ObjRS.Open StrSQL, ObjConn
      if err <> 0 then
        DisConnDBH ObjConn, ObjRS	'DB�ؒf
        jampErrerP "1","b108","01","�����o�FCSV�t�@�C���_�E�����[�h(�w��Ȃ��ǉ�)","102","SQL:<BR>"&StrSQL
      end if
      i=2
      j=0
      Do Until ObjRS.EOF
        If tmpNasiConNoA(i-1)(1)<>Trim(ObjRS("BLNo")) Then
          For j = 1 to UBound(tmpNasiConNoA(0))
            If tmpNasiConNoA(1)(j)="'"&Trim(ObjRS("BLNo"))&"'" Then
              Exit For
            End If
          Next
        End If
        tmpNasiConNoA(i)=Array(tmpNasiConNoA(0)(j),"","","","","","","","")
        tmpNasiConNoA(i)(1)=Trim(ObjRS("BLNo"))
        tmpNasiConNoA(i)(2)=Trim(ObjRS("ContNo"))			'�R���e�i�ԍ�
        If IsNull(ObjRS("ReturnTime")) Then				'�ԋp
          tmpNasiConNoA(i)(3)="��"
        Else
          tmpNasiConNoA(i)(3)="��"
        End If
        tmpNasiConNoA(i)(4)=ObjRS("ContSize") 			'�T�C�Y
        tmpNasiConNoA(i)(5)=ObjRS("ContType")			'�^�C�v
        tmpNasiConNoA(i)(6)=ObjRS("ContHeight")		'����
        tmpNasiConNoA(i)(7)=ObjRS("ContWeight")*100	'�O���X
        tmpNasiConNoA(i)(8)=Trim(ObjRS("ReturnPlace"))	'��o���ԋp��
        i=i+1
        ObjRS.MoveNext
      Loop
      ObjRS.close
    End If

    '�ꗗ�̕R�t�����擾
    If tmpItiConNo(0) <> "" Then
      StrSQL="SELECT COUNT(*) AS Num "&_
             "FROM (hITCommonInfo ITC INNER JOIN hITFullOutSelect ITF ON ITC.WkContrlNo = ITF.WkContrlNo) "&_
             "WHERE ITC.WkNo IN("&Mid(tmpItiConNo(1),2)&") AND ITC.Process='R' "
      ObjRS.Open StrSQL, ObjConn
      if err <> 0 then
        DisConnDBH ObjConn, ObjRS	'DB�ؒf
        jampErrerP "1","b108","01","�����o�FCSV�t�@�C���_�E�����[�h(�ꗗ�ǉ�)","102","SQL:<BR>"&StrSQL
      end if
      tmpNum=ObjRS("Num")+2
      ObjRS.close
      ReDim Preserve tmpItiConNoA(tmpNum)
      tmpItiConNoA(0) =Split(tmpItiConNo(0), ",", -1, 1)
      tmpItiConNoA(1) =Split(tmpItiConNo(1), ",", -1, 1)
      StrSQL="SELECT ITC.WkNo,ITF.ContNo, INC.ReturnTime, INC.CYDelTime, INC.ReturnPlace, "&_
             "Cnt.ContSize, Cnt.ContType, Cnt.ContHeight, Cnt.ContWeight "&_
             "FROM hITCommonInfo ITC INNER JOIN hITFullOutSelect ITF ON ITC.WkContrlNo = ITF.WkContrlNo "&_
             "LEFT JOIN ImportCont AS INC ON ITF.ContNo =INC.ContNo "&_
             "LEFT JOIN Container AS Cnt ON INC.ContNo =Cnt.ContNo AND INC.VslCode =Cnt.VslCode AND INC.VoyCtrl =Cnt.VoyCtrl "&_
             "WHERE ITC.WkNo IN("&Mid(tmpItiConNo(1),2)&") AND ITC.Process='R'"&_
             "ORDER BY ITF.WkContrlNo ASC"
      ObjRS.Open StrSQL, ObjConn
      if err <> 0 then
        DisConnDBH ObjConn, ObjRS	'DB�ؒf
        jampErrerP "1","b108","01","�����o�FCSV�t�@�C���_�E�����[�h(�ꗗ�ǉ�)","102","SQL:<BR>"&StrSQL
      end if
      i=2
      Do Until ObjRS.EOF
        If tmpItiConNoA(i-1)(1)<>Trim(ObjRS("WkNo")) Then
          For j = 1 to UBound(tmpItiConNoA(0))
            If tmpItiConNoA(1)(j)="'"&Trim(ObjRS("WkNo"))&"'" Then
              Exit For
            End If
          Next
        End If
        tmpItiConNoA(i)=Array(tmpItiConNoA(0)(j),"","","","","","","","")
        tmpItiConNoA(i)(1)=Trim(ObjRS("WkNo"))
        tmpItiConNoA(i)(2)=Trim(ObjRS("ContNo"))			'�R���e�i�ԍ�
        If IsNull(ObjRS("ReturnTime")) Then				'�ԋp
          tmpItiConNoA(i)(3)="��"
        Else
          tmpItiConNoA(i)(3)="��"
        End If
        tmpItiConNoA(i)(4)=ObjRS("ContSize") 			'�T�C�Y
        tmpItiConNoA(i)(5)=ObjRS("ContType")			'�^�C�v
        tmpItiConNoA(i)(6)=ObjRS("ContHeight")		'����
        tmpItiConNoA(i)(7)=ObjRS("ContWeight")*100	'�O���X
        tmpItiConNoA(i)(8)=Trim(ObjRS("ReturnPlace"))	'��o���ԋp��
        i=i+1
        ObjRS.MoveNext
      Loop
      ObjRS.close
    End If

    'BL�̕R�t�����擾
    If tmpBLNo(0) <> "" Then
      tmpBLNoA(0)=Split(tmpBLNo(0), ",", -1, 1)
      tmpBLNoA(1)=Split(tmpBLNo(1), ",", -1, 1)
      dim VslCode,VoyCtrl
      tmpNum=2
      i=2
      For j=1 To UBound(tmpBLNoA(0))
        '�Ώ�BL�I��
        StrSQL = "SELECT INC.VslCode, INC.VoyCtrl "&_
                 "From ImportCont AS INC  "&_
                 "Where INC.BLNo= "& tmpBLNoA(1)(j) &" ORDER BY INC.UpdtTime DESC"
        ObjRS.Open StrSQL, ObjConn
        if err <> 0 then
          DisConnDBH ObjConn, ObjRS	'DB�ؒf
          jampErrerP "1","b108","01","�����o�FCSV�t�@�C���_�E�����[�h(BL�ǉ�)","102","SQL:<BR>"&StrSQL
        end if
        VslCode=Trim(ObjRS("VslCode"))
        VoyCtrl=Trim(ObjRS("VoyCtrl"))
        ObjRS.close
        '�Ώی����擾
        StrSQL = "SELECT count(ContNo) AS Num FROM ImportCont WHERE BLNo="&tmpBLNoA(1)(j)&" "&_
                 "AND VoyCtrl ='" & VoyCtrl & "' AND VslCode= '"& VslCode &"' "
        ObjRS.Open StrSQL, ObjConn
        if err <> 0 then
          DisConnDBH ObjConn, ObjRS	'DB�ؒf
          jampErrerP "1","b108","01","�����o�FCSV�t�@�C���_�E�����[�h(BL�ǉ�)","102","SQL:<BR>"&StrSQL
        end if
        tmpNum=tmpNum+ObjRS("Num")
        ObjRS.close
        ReDim Preserve tmpBLNoA(tmpNum)
        '�Ώۃf�[�^�擾
        StrSQL="SELECT INC.ContNo, INC.BLNo, INC.ReturnTime, INC.CYDelTime, INC.ReturnPlace, "&_
               "Cnt.ContSize, Cnt.ContType, Cnt.ContHeight, Cnt.ContWeight "&_
               "FROM ImportCont AS INC "&_
               "LEFT JOIN Container AS Cnt ON INC.ContNo=Cnt.ContNo AND INC.VslCode=Cnt.VslCode AND INC.VoyCtrl=Cnt.VoyCtrl "&_
               "WHERE INC.BLNo="&tmpBLNoA(1)(j)&" AND INC.VoyCtrl=" & VoyCtrl & " AND INC.VslCode='"& VslCode &"' "&_
               "ORDER BY INC.ContNo ASC"
        ObjRS.Open StrSQL, ObjConn
        if err <> 0 then
          DisConnDBH ObjConn, ObjRS	'DB�ؒf
          jampErrerP "1","b108","01","�����o�FCSV�t�@�C���_�E�����[�h(BL�ǉ�)","102","SQL:<BR>"&StrSQL
        end if
        Do Until ObjRS.EOF
          tmpBLNoA(i)=Array(tmpBLNoA(0)(j),"","","","","","","","")
          tmpBLNoA(i)(1)=Trim(ObjRS("BLNo"))
          tmpBLNoA(i)(2)=Trim(ObjRS("ContNo"))			'�R���e�i�ԍ�
          If IsNull(ObjRS("ReturnTime")) Then				'�ԋp
            tmpBLNoA(i)(3)="��"
          Else
            tmpBLNoA(i)(3)="��"
          End If
          tmpBLNoA(i)(4)=ObjRS("ContSize") 			'�T�C�Y
          tmpBLNoA(i)(5)=ObjRS("ContType")			'�^�C�v
          tmpBLNoA(i)(6)=ObjRS("ContHeight")		'����
          tmpBLNoA(i)(7)=ObjRS("ContWeight")*100	'�O���X
          tmpBLNoA(i)(8)=Trim(ObjRS("ReturnPlace"))	'��o���ԋp��
          i=i+1
          ObjRS.MoveNext
        Loop
        ObjRS.close
      Next
    End If
'3th del  '�W�J�f�[�^����
'3th del    j=1
'3th del    For i=1 To preNum
'3th del      If preDtTbl(i)(4)="1" Then		'�w������
'3th del        DtTbl(j)=preDtTbl(i)
'3th del        DtTbl(j)(11)="�@"
'3th del        j=j+1
'3th del      ElseIf preDtTbl(i)(4)="3" Then	'�ꗗ
'3th del        '�Ώی����擾
'3th del        StrSQL = "SELECT count(ITF.ContNo) AS CNUM FROM "&_
'3th del                 "(hITCommonInfo ITC INNER JOIN hITFullOutSelect ITF ON ITC.WkContrlNo = ITF.WkContrlNo) "&_
'3th del                 "INNER JOIN ImportCont IPC ON ITF.ContNo =IPC.ContNo AND ITC.BLNo = IPC.BLNo "&_
'3th del                 "WHERE ITC.ContNo='"&preDtTbl(i)(5)&"'"
'3th del        ObjRS.Open StrSQL, ObjConn
'3th del        if err <> 0 then
'3th del          DisConnDBH ObjConn, ObjRS	'DB�ؒf
'3th del          jampErrerP "1","b108","01","�����o�FCSV�t�@�C���_�E�����[�h","101","SQL:<BR>"&StrSQL
'3th del        end if
'3th del        Num = Num + ObjRS("CNUM")-1
'3th del        ObjRS.close
'3th del        ReDim Preserve DtTbl(Num)
'3th del        '�f�[�^�擾
'3th del        StrSQL = "SELECT ITF.ContNo, IPC.ReturnTime, IPC.CYDelTime, "&_
'3th del                 "NULLIF('-',LEFT(DateDiff(day,getdate(),DateAdd(day,"&preDtTbl(i)(21)-param(0)+1&" ,IPC.CYDelTime))*("&preDtTbl(i)(21)&"%6),1)) AS ReturnArrert "&_
'3th del                 "FROM (hITCommonInfo ITC INNER JOIN hITFullOutSelect ITF ON ITC.WkContrlNo = ITF.WkContrlNo) "&_
'3th del                 "INNER JOIN ImportCont IPC ON ITF.ContNo =IPC.ContNo AND ITC.BLNo = IPC.BLNo "&_
'3th del                 "WHERE ITC.ContNo='"&preDtTbl(i)(5)&"'"
'3th del        ObjRS.Open StrSQL, ObjConn
'3th del        if err <> 0 then
'3th del          DisConnDBH ObjConn, ObjRS	'DB�ؒf
'3th del          jampErrerP "1","b108","01","�����o�FCSV�t�@�C���_�E�����[�h","101","SQL:<BR>"&StrSQL
'3th del        end if
'3th del        Do Until ObjRS.EOF
'3th del          DtTbl(j)=preDtTbl(i)
'3th del          DtTbl(j)(5)=Trim(ObjRS("ContNo"))
'3th del          DtTbl(j)(12)=ObjRS("ReturnArrert")
'3th del          If IsNull(ObjRS("ReturnTime")) Then
'3th del            DtTbl(j)(8)="��"
'3th del          Else
'3th del            DtTbl(j)(8)="��"
'3th del          End If
'3th del          ObjRS.MoveNext
'3th del          j=j+1
'3th del        Loop
'3th del        ObjRS.close
'3th del      ElseIf preDtTbl(i)(4)="2" Or preDtTbl(i)(4)="4" Then	'�ꗗ'�w��Ȃ�,BL
'3th del        '�Ώی����擾
'3th del        StrSQL = "SELECT count(ContNo) AS CNUM FROM ImportCont WHERE BLNo='"&preDtTbl(i)(11)&"'"
'3th del        ObjRS.Open StrSQL, ObjConn
'3th del        if err <> 0 then
'3th del          DisConnDBH ObjConn, ObjRS	'DB�ؒf
'3th del          jampErrerP "1","b108","01","�����o�FCSV�t�@�C���_�E�����[�h","101","SQL:<BR>"&StrSQL
'3th del        end if
'3th del        Num = Num + ObjRS("CNUM")-1
'3th del        ObjRS.close
'3th del        ReDim Preserve DtTbl(Num)
'3th del        '�f�[�^�擾
'3th del        StrSQL = "SELECT ContNo, ReturnTime, CYDelTime, "&_
'3th del                 "NULLIF('-',LEFT(DateDiff(day,getdate(),DateAdd(day,"&preDtTbl(i)(21)-param(0)+1&" ,CYDelTime))*("&preDtTbl(i)(21)&"%6),1)) AS ReturnArrert "&_
'3th del                 "FROM ImportCont WHERE BLNo='"&preDtTbl(i)(11)&"'"
'3th del        ObjRS.Open StrSQL, ObjConn
'3th del        if err <> 0 then
'3th del          DisConnDBH ObjConn, ObjRS	'DB�ؒf
'3th del          jampErrerP "1","b108","01","�����o�FCSV�t�@�C���_�E�����[�h","101","SQL:<BR>"&StrSQL
'3th del        end if
'3th del        Do Until ObjRS.EOF
'3th del          DtTbl(j)=preDtTbl(i)
'3th del          DtTbl(j)(5)=Trim(ObjRS("ContNo"))
'3th del          DtTbl(j)(12)=ObjRS("ReturnArrert")
'3th del          If IsNull(ObjRS("ReturnTime")) Then
'3th del            DtTbl(j)(8)="��"
'3th del          Else
'3th del            DtTbl(j)(8)="��"
'3th del          End If
'3th del          ObjRS.MoveNext
'3th del          j=j+1
'3th del        Loop
'3th del        ObjRS.close
'3th del      Else
'3th del          jampErrerP "1","b108","01","�����o�FCSV�t�@�C���_�E�����[�h","101","SQL:<BR>"&StrSQL
'3th del      End If
'3th del    Next
'3th change End

  'DB�ڑ�����
    DisConnDBH ObjConn, ObjRS
  '�G���[�g���b�v����
    on error goto 0
'  FOR i = 0 TO UBOUND(tmp)

' �t�@�C���̃_�E�����[�h
  Response.ContentType="application/octet-stream"
  Response.AddHeader "Content-Disposition","attachment; filename=output.csv"

'3th change    Response.Write "���o�\���,�w����,�w�����ւ̉�,��Ɣԍ�,�w����,�R���e�i�ԍ�,�a�k�ԍ�,"
'3th change    Response.Write "�D��,�D��,�T�C�Y,�b�x,�t���[�^�C��,���o��,��������,�ԋp�\��,�ԋp,�w����,�w�����,"
'3th change    Response.Write "���l�P,���l�Q,���l�R"
'3th change    Response.Write Chr(13) & Chr(10)
'3th change    For j=1 To Num
'3th change      Response.Write Trim(DtTbl(j)(1))&","&Trim(DtTbl(j)(2))&","&Trim(DtTbl(j)(14))&","&Trim(DtTbl(j)(3))&","
'3th change      Response.Write Siji(DtTbl(j)(4))&","&Trim(DtTbl(j)(5))&","&Trim(DtTbl(j)(11))&","&Trim(DtTbl(j)(15))&","
'3th change      Response.Write Trim(DtTbl(j)(16))&","&Trim(DtTbl(j)(17))&","&Trim(DtTbl(j)(18))&","&Trim(DtTbl(j)(19))&","
'3th change      Response.Write Trim(DtTbl(j)(0))&","&Trim(DtTbl(j)(6))&","&Trim(DtTbl(j)(7))&","
'3th change      Response.Write Trim(DtTbl(j)(8))&","&Trim(DtTbl(j)(9))&","&Trim(DtTbl(j)(10))&","
'3th change      Response.Write Trim(DtTbl(j)(22))&","&Trim(DtTbl(j)(23))&","&Trim(DtTbl(j)(24))
'3th change      Response.Write Chr(13) & Chr(10)
'3th change    Next
    Response.Write "���o�\���,�w����,�w�����ւ̉�,��Ɣԍ�,�w�b�h�h�c,�w����,�R���e�i�ԍ�,�a�k�ԍ�,"
    Response.Write "�D��,�D��,�T�C�Y,�^�C�v,����,�O���X,�i��,�b�x,�t���[�^�C��,�[����P,�[����Q,�[������,��������,"
    Response.Write "��R���ԋp��,�ԋp�\��,�ԋp,�w����,�w�����,���l�P,���l�Q"
    Response.Write Chr(13) & Chr(10)
    For j=0 To Num
      Flag=0
      Select Case DtTbl(j)(4)
        Case "1"
          Response.Write Trim(DtTbl(j)(1))&","&Trim(DtTbl(j)(2))&","&Trim(DtTbl(j)(14))&","&Trim(DtTbl(j)(3))&","
          Response.Write Trim(DtTbl(j)(25))&","&Siji(DtTbl(j)(4))&","&Trim(DtTbl(j)(5))&","&Trim(DtTbl(j)(11))&","
          Response.Write Trim(DtTbl(j)(15))&","&Trim(DtTbl(j)(16))&","&Trim(DtTbl(j)(17))&","&Trim(DtTbl(j)(26))&","
          Response.Write Trim(DtTbl(j)(27))&","&Trim(DtTbl(j)(28))&","&Trim(DtTbl(j)(29))&","&Trim(DtTbl(j)(18))&","
          Response.Write Trim(DtTbl(j)(19))&","&Trim(DtTbl(j)(24))&","&Trim(DtTbl(j)(30))&","&Trim(DtTbl(j)(31))&","
          Response.Write Trim(DtTbl(j)(6))&","&Trim(DtTbl(j)(32))&","&Trim(DtTbl(j)(7))&","&Trim(DtTbl(j)(8))&","
          Response.Write Trim(DtTbl(j)(9))&","&Trim(DtTbl(j)(10))&","&Trim(DtTbl(j)(22))&","&Trim(DtTbl(j)(23))
          Response.Write Chr(13) & Chr(10)
        Case "2"
          For i=2 To UBound(tmpNasiConNoA)-1
            If tmpNasiConNoA(i)(0)=Trim(j) Then
              Response.Write Trim(DtTbl(j)(1))&","&Trim(DtTbl(j)(2))&","&Trim(DtTbl(j)(14))&","&Trim(DtTbl(j)(3))&","
              Response.Write Trim(DtTbl(j)(25))&","&Siji(DtTbl(j)(4))&","&tmpNasiConNoA(i)(2)&","&Trim(DtTbl(j)(11))&","
              Response.Write Trim(DtTbl(j)(15))&","&Trim(DtTbl(j)(16))&","&tmpNasiConNoA(i)(4)&","&tmpNasiConNoA(i)(5)&","
              Response.Write tmpNasiConNoA(i)(6)&","&tmpNasiConNoA(i)(7)&","&Trim(DtTbl(j)(29))&","&Trim(DtTbl(j)(18))&","
              Response.Write Trim(DtTbl(j)(19))&","&Trim(DtTbl(j)(24))&","&Trim(DtTbl(j)(30))&","&Trim(DtTbl(j)(31))&","
              Response.Write Trim(DtTbl(j)(6))&","&tmpNasiConNoA(i)(8)&","&Trim(DtTbl(j)(7))&","&tmpNasiConNoA(i)(3)&","
              Response.Write Trim(DtTbl(j)(9))&","&Trim(DtTbl(j)(10))&","&Trim(DtTbl(j)(22))&","&Trim(DtTbl(j)(23))
              Response.Write Chr(13) & Chr(10)
              Flag=1
            ElseIf Flag=1 Then
              Exit For
            End If
          Next
        Case "3"
          For i=2 To UBound(tmpItiConNoA)-1
            If tmpItiConNoA(i)(0)=Trim(j) Then
              Response.Write Trim(DtTbl(j)(1))&","&Trim(DtTbl(j)(2))&","&Trim(DtTbl(j)(14))&","&Trim(DtTbl(j)(3))&","
              Response.Write Trim(DtTbl(j)(25))&","&Siji(DtTbl(j)(4))&","&tmpItiConNoA(i)(2)&","&Trim(DtTbl(j)(11))&","
              Response.Write Trim(DtTbl(j)(15))&","&Trim(DtTbl(j)(16))&","&tmpItiConNoA(i)(4)&","&tmpItiConNoA(i)(5)&","
              Response.Write tmpItiConNoA(i)(6)&","&tmpItiConNoA(i)(7)&","&Trim(DtTbl(j)(29))&","&Trim(DtTbl(j)(18))&","
              Response.Write Trim(DtTbl(j)(19))&","&Trim(DtTbl(j)(24))&","&Trim(DtTbl(j)(30))&","&Trim(DtTbl(j)(31))&","
              Response.Write Trim(DtTbl(j)(6))&","&tmpItiConNoA(i)(8)&","&Trim(DtTbl(j)(7))&","&tmpItiConNoA(i)(3)&","
              Response.Write Trim(DtTbl(j)(9))&","&Trim(DtTbl(j)(10))&","&Trim(DtTbl(j)(22))&","&Trim(DtTbl(j)(23))
              Response.Write Chr(13) & Chr(10)
              Flag=1
            ElseIf Flag=1 Then
              Exit For
            End If
          Next
        Case "4"
          For i=2 To UBound(tmpBLNoA)-1
            If tmpBLNoA(i)(0)=Trim(j) Then
              Response.Write Trim(DtTbl(j)(1))&","&Trim(DtTbl(j)(2))&","&Trim(DtTbl(j)(14))&","&Trim(DtTbl(j)(3))&","
              Response.Write Trim(DtTbl(j)(25))&","&Siji(DtTbl(j)(4))&","&tmpBLNoA(i)(2)&","&Trim(DtTbl(j)(11))&","
              Response.Write Trim(DtTbl(j)(15))&","&Trim(DtTbl(j)(16))&","&tmpBLNoA(i)(4)&","&tmpBLNoA(i)(5)&","
              Response.Write tmpBLNoA(i)(6)&","&tmpBLNoA(i)(7)&","&Trim(DtTbl(j)(29))&","&Trim(DtTbl(j)(18))&","
              Response.Write Trim(DtTbl(j)(19))&","&Trim(DtTbl(j)(24))&","&Trim(DtTbl(j)(30))&","&Trim(DtTbl(j)(31))&","
              Response.Write Trim(DtTbl(j)(6))&","&tmpBLNoA(i)(8)&","&Trim(DtTbl(j)(7))&","&tmpBLNoA(i)(3)&","
              Response.Write Trim(DtTbl(j)(9))&","&Trim(DtTbl(j)(10))&","&Trim(DtTbl(j)(22))&","&Trim(DtTbl(j)(23))
              Response.Write Chr(13) & Chr(10)
              Flag=1
            ElseIf Flag=1 Then
              Exit For
            End If
          Next
      End Select
    Next
    Response.End
%>
