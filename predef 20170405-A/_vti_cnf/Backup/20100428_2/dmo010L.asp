<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmo010L.asp				_/
'_/	Function	:�����o���ꗗ��ʃ��X�g�o��		_/
'_/	Date		:2003/05/27				_/
'_/	Code By		:SEIKO Electric.Co ��d			_/
'_/	Modify		:C-001 2003/07/29	CSV�o�͑Ή�	_/
'_/			:C-002 2003/07/29	���l���Ή�	_/
'_/			:C-003 2003/08/22	��Ɣԍ��ł̌���_/
'_/			:C-004 2003/08/22	�\�������`	_/
'_/			:3th   2004/01/31	3���Ή�	_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTP�R���e���c�^�C�v�ݒ�
	Response.ContentType = "text/html; charset=Shift_JIS"
				
		Const CONST_ASC = "<BR><IMG border=0 src=Image/ascending.gif>"
		Const CONST_DESC = "<BR><IMG border=0 src=Image/descending.gif>"
%>
<!--#include File="Common.inc"-->
<%
'�Z�b�V�����̗L�������`�F�b�N
  CheckLoginH

'���[�U�f�[�^����
  dim USER, COMPcd
  USER   = UCase(Session.Contents("userid"))
  COMPcd = Session.Contents("COMPcd")
'INI�t�@�C�����ݒ�l���擾
  dim param(2),calcDate1
  getIni param
  calcDate1 = DateAdd("d", "-"&param(1), Date)


'�f�[�^�擾
  dim Num, DtTbl,i,j,SortFlag,SortKye,InfoFlag,Siji
  Siji  =Array("","�w�肠��","�w��Ȃ�","�ꗗ","�a�k")

  If Request("SortFlag") = "" Then
    SortFlag = 0
  Else
    SortFlag = Request("SortFlag")
  End If

  If Request("InfoFlag") = "" Then
    InfoFlag = 0
  Else
    InfoFlag = Request("InfoFlag")
  End If

  '�\�[�g�P�[�X
  dim strWrer,ErrerM

'2009/02/25 Add-S G.Ariola   
    dim strOrder
  dim FieldName
   ReDim FieldName(19)
  
'  FieldName=Array("ISNULL(ITC.WorkDate,DATEADD(Year,100,getdate()))","Code1","Name1","ITC.WkNo","ITC.FullOutType","BLContNo","mV.ShipLine"," ShipName","CON.ContSize","CY","INC.FreeTime","ITC.DeliverTo1","ISNULL(ITC.WorkCompleteDate,DATEADD(Year,100,getdate()))","ITC.ReturnDateStr","ReturnValue","Code2","Name2","Flag1","ITC.Comment1","ITC.Comment2")
  FieldName=Array("WorkDate","Code1","WkNo","FullOutType","BLContNo","ShipLine","ShipName","ContSize","CY","FreeTime","DeliverTo1","WorkCompleteDate","ReturnDateStr","ReturnValue","Code2","Flag1","Comment1","Comment2","Name1")
  
  strOrder = getSort(Session("Key1"),Session("KeySort1"),"")
  strOrder = getSort(Session("Key2"),Session("KeySort2"),strOrder)
  strOrder = getSort(Session("Key3"),Session("KeySort3"),strOrder)
'2009/02/25 Add-E G.Ariola

  Select Case SortFlag
  
'20030910 �����\������o�\������ɕ\��(�����ȍ~�̂�)�ɕύX
      Case "0" '�����\��:���o�\������ɕ\��(�����ȍ~�̂�)
          WriteLogH "b101", "�����o���O���ꗗ", "01", ""
		  '2010/04/23 G.Ariola Upd-S
          'strWrer = "AND (DateDiff(day,ITC.WorkCompleteDate,'"&calcDate1&"')<=0 Or ITC.WorkCompleteDate IS Null) "&_
		  strWrer = "AND ('"&calcDate1&"' <= ITC.WorkCompleteDate Or ITC.WorkCompleteDate IS Null) "&_
		  
                    "AND (DateDiff(day,ITC.WorkDate,'"&Date&"')<=0 Or ITC.WorkDate IS Null) "
		  '2010/04/23 G.Ariola Upd-E
''3th          getData DtTbl,strWrer
          '2009/11/02 Add-S Tanaka
          'strOrder=" ORDER BY isnull(T.WorkDate,DATEADD(Year,100,getdate())),T.InputDate ASC "
          '2009/11/02 Add-E Tanaka
          GetData DtTbl, strWrer, 1

''3th         j=1
''3th         DtTbl(0)(14) = 0
''3th          For i=1 To Num
''3th            If DtTbl(i)(8)  <> "��" Then
''3th              DtTbl(j)=DtTbl(i)
''3th              DtTbl(0)(14) = DtTbl(0)(14) + DtTbl(j)(13)
''3th              j=j+1
''3th            End If
''3th          Next
''3th          Num=j-1
      Case "12" '�������\��:���o�\������ɕ\��(�����������\��)
          WriteLogH "b101", "�����o���O���ꗗ", "01", ""
		  '2010/04/23 G.Ariola Upd-S
          'strWrer = "AND (DateDiff(day,ITC.WorkCompleteDate,'"&calcDate1&"')<=0 Or ITC.WorkCompleteDate IS Null) "
		  strWrer = "AND ('"&calcDate1&"' <= ITC.WorkCompleteDate Or ITC.WorkCompleteDate IS Null) "
		  '2010/04/23 G.Ariola Upd-E
'3th          getData DtTbl,strWrer
          GetData DtTbl, strWrer, 1
'3th          j=1
'3th          DtTbl(0)(14) = 0
'3th          For i=1 To Num
'3th            If DtTbl(i)(8)  <> "��" Then
'3th              DtTbl(j)=DtTbl(i)
'3th              DtTbl(0)(14) = DtTbl(0)(14) + DtTbl(j)(13)
'3th              j=j+1
'3th            End If
'3th          Next
'3th          Num=j-1
'Del 030722      Case "1" '�ԋp��v����R���e�i��
      Case "2" '���Ɖ�
          WriteLogH "b101", "�����o���O���ꗗ", "03", ""
		  '2010/04/23 G.Ariola Upd-S
          'strWrer = "AND (DateDiff(day,ITC.WorkCompleteDate,'"&calcDate1&"')<=0 Or ITC.WorkCompleteDate IS Null) "
		  strWrer = "AND ('"&calcDate1&"' <= ITC.WorkCompleteDate Or ITC.WorkCompleteDate IS Null) "
		  '2010/04/23 G.Ariola Upd-E
'3th          getData DtTbl,strWrer
          GetData DtTbl, strWrer, 2
'3th          j=1
'3th          DtTbl(0)(14) = 0
'3th          For i=1 To Num
'3th            If DtTbl(i)(10) = "��" Then
'3th              DtTbl(j)=DtTbl(i)
'3th              DtTbl(0)(14) = DtTbl(0)(14) + DtTbl(j)(13)
'3th              j=j+1
'3th            End If
'3th          Next
'3th          Num=j-1
      Case "7" '�ۗ�
          WriteLogH "b101", "�����o���O���ꗗ", "07", ""
		  '2010/04/23 G.Ariola Upd-S
          'strWrer = "AND (DateDiff(day,ITC.WorkCompleteDate,'"&calcDate1&"')<=0 Or ITC.WorkCompleteDate IS Null) "
		  strWrer = "AND ('"&calcDate1&"' <= ITC.WorkCompleteDate Or ITC.WorkCompleteDate IS Null) "
		  '2010/04/23 G.Ariola Upd-E
'3th          getData DtTbl,strWrer
          GetData DtTbl, strWrer, 3
'3th          j=1
'3th          DtTbl(0)(14) = 0
'3th          For i=1 To Num
'3th'CW-031            If DtTbl(i)(10) = "no" Then
'3th            If DtTbl(i)(10) = "No" Then
'3th              DtTbl(j)=DtTbl(i)
'3th              DtTbl(0)(14) = DtTbl(0)(14) + DtTbl(j)(13)
'3th              j=j+1
'3th            End If
'3th          Next
'3th          Num=j-1
      Case "3" '�R���e�i�ԍ��Ō���
          SortKye=Request("SortKye")
          WriteLogH "b101", "�����o���O���ꗗ", "11",SortKye
'3th chage          Get_Data Num,DtTbl
          strWrer = "AND ITC.ContNo LIKE '%" & SortKye & "'"
          getData DtTbl,strWrer,0

'3th          j=1
'3th          DtTbl(0)(14) = 0
'3th          For i=1 To Num
'3th            If Right(DtTbl(i)(5),Len(SortKye))= SortKye AND DtTbl(i)(4) <> 4 Then
'3th              DtTbl(j)=DtTbl(i)
'3th              DtTbl(0)(14) = DtTbl(0)(14) + DtTbl(j)(13)
'3th              j=j+1
'3th            End If
'3th          Next
'3th          Num=j-1
'2010/04/24 Upd-S Tanaka BL�ԍ��ł̌������S���ƂȂ��Ă���̂ŕύX(���\�[�X�ɖ߂�)
'      Case "4" '�R���e�i�ԍ��Ō���
'          SortKye=Request("SortKye")
'          WriteLogH "b101", "�����o���O���ꗗ", "11",SortKye
''3th chage          Get_Data Num,DtTbl
'          strWrer = "AND ITC.ContNo LIKE '%" & SortKye & "'"
'          getData DtTbl,strWrer,0

      Case "4" 'BL�ԍ��Ō���
          SortKye=Request("SortKye")
          WriteLogH "b101", "�����o���O���ꗗ", "11",SortKye
'3th chage          Get_Data Num,DtTbl
          strWrer = "AND ITC.BLNo LIKE '%" & SortKye & "' AND ITC.FullOutType='4'"
          getData DtTbl,strWrer,0


'3th          j=1
'3th          DtTbl(0)(14) = 0
'3th          For i=1 To Num
'3th            If Right(DtTbl(i)(5),Len(SortKye))= SortKye AND DtTbl(i)(4) <> 4 Then
'3th              DtTbl(j)=DtTbl(i)
'3th              DtTbl(0)(14) = DtTbl(0)(14) + DtTbl(j)(13)
'3th              j=j+1
'3th            End If
'3th          Next
'3th          Num=j-1
      Case "11" '��Ɣԍ��Ō���
          SortKye=Request("SortKye")
          WriteLogH "b101", "�����o���O���ꗗ", "11",SortKye
'3th          Get_Data Num,DtTbl
          strWrer = "AND ITC.WkNo LIKE '%" & SortKye & "'"
          getData DtTbl,strWrer,0
'3th          j=1
'3th          DtTbl(0)(14) = 0
'3th          For i=1 To Num
'3th            If Right(DtTbl(i)(3),Len(SortKye))= SortKye Then
'3th              DtTbl(j)=DtTbl(i)
'3th              DtTbl(0)(14) = DtTbl(0)(14) + DtTbl(j)(13)
'3th              j=j+1
'3th            End If
'3th          Next
'3th          Num=j-1
'ADD End C-003
      Case "5" '�S���\��
          WriteLogH "b101", "�����o���O���ꗗ", "04",""
          strWrer = " "
'3th          getData DtTbl,strWrer
          getData DtTbl,strWrer,0
'3th          DtTbl(0)(14) = 0
'3th          For i=1 To Num
'3th              DtTbl(0)(14) = DtTbl(0)(14) + DtTbl(i)(13)
'3th          Next
      Case "6" '���o�������������ׂĕ\��
          WriteLogH "b101", "�����o���O���ꗗ", "06",""
          strWrer = "AND ITC.WorkCompleteDate IS Null "
'3th          getData DtTbl,strWrer
          getData DtTbl,strWrer,1
'3th          j=1
'3th          DtTbl(0)(14) = 0
'3th          For i=1 To Num
'3th            If DtTbl(i)(8)  <> "��" Then
'3th              DtTbl(j)=DtTbl(i)
'3th              DtTbl(0)(14) = DtTbl(0)(14) + DtTbl(j)(13)
'3th              j=j+1
'3th            End If
'3th          Next
'3th          Num=j-1
      Case "8" '�Ɖ�
          WriteLogH "b107", "�����o���O���ꗗ", "01",""
          Get_Data Num,DtTbl
        '�G���[�g���b�v�J�n
          on error resume next
        'DB�ڑ�
          dim ObjConn, ObjRS, StrSQL
          ConnDBH ObjConn, ObjRS
          For i=1 To Num
'CW-002            If DtTbl(i)(13) <> 0 Then
            If DtTbl(i)(13) <> 0 AND DtTbl(i)(6)="" AND DtTbl(i)(14)="��" Then
              StrSQL = "UPDATE hITReference SET UpdtTime='"& Now() &"', UpdtPgCd='PREDEF01'," &_
                       "UpdtTmnl='"& USER &"', TruckerFlag"& DtTbl(i)(13) &"=1 "&_
                       "WHERE WkContrlNo IN (select WkContrlNo From hITCommonInfo "&_
                       "WHERE WkNo='"& DtTbl(i)(3) &"' AND WkType='1' AND Process='R' )"
              ObjConn.Execute(StrSQL)
              if err <> 0 then
                Set ObjRS = Nothing
                jampErrerPDB ObjConn,"2","b107","01","�����o:�Љ�Ϗ���","104","SQL:<BR>"&strSQL
              end if
            End If
          Next
        'DB�ڑ�����
          DisConnDBH ObjConn, ObjRS
        '�G���[�g���b�v����
          on error goto 0
          Response.Redirect "./dmo010L.asp"
      Case else '�S���\��
          WriteLogH "b101", "�����o���O���ꗗ", "04",""
          strWrer = " "
          getData DtTbl,strWrer,0
  End Select

'�f�[�^�擾�֐�
'3th chage Function getData(DtTbl,strWhere)
'2009/02/25 Add-S G.Ariola
Function getSort(Key,SortKey,str)
getSort = str
	if Key <> "" then
	
		if str = "" then
			'getSort = " ORDER BY (Case When LTRIM(ISNULL(" & FieldName(Key) & "),'')='' Then 1 Else 0 End), " & FieldName(Key) & " " & SortKey
			if (FieldName(Key) = "WorkDate" OR FieldName(Key) = "FreeTime" OR FieldName(Key) = "WorkCompleteDate") AND SortKey = "ASC" then 
			getSort = " ORDER BY isnull(" & FieldName(Key) & ",DATEADD(Year,100,getdate())) " & SortKey	
			else
			getSort = " ORDER BY " & FieldName(Key) & " " & SortKey	
			end if
			
		else
			'getSort = str & " , (Case When LTRIM(ISNULL(" & FieldName(Key) & "),'')='' Then 1 Else 0 End), " & FieldName(Key) & " " & SortKey
			if (FieldName(Key) = "WorkDate" OR FieldName(Key) = "FreeTime" OR FieldName(Key) = "WorkCompleteDate") AND SortKey = "ASC"  then 
			getSort = str & " , isnull(" & FieldName(Key) & ",DATEADD(Year,100,getdate())) " & SortKey	
			else
			getSort = str & " , " & FieldName(Key) & " " & SortKey	
			end if
			
		end if	
	end if	
end function

Function getImage(SortKey)
getImage = ""
		if SortKey = "ASC" then
			getImage = CONST_ASC	
		else
			getImage = CONST_DESC
		end if	
end function
'2009/02/25 Add-E G.Ariola
Function getData(DtTbl,strWhere,DelType)

  ReDim DtTbl(1)
'C-002  DtTbl(0)=Array("���o��","���o<BR>�\���","�w����","���<BR>�ԍ�","�w����","�R���e�i�ԍ�<BR>�^�a�k�ԍ�","��������","�ԋp�\��","�ԋp","�w����","�w����<BR>��","BL�ԍ�","�A���[���t���O","�Ɖ��","�w�����։�","�D��","�D��","�T�C�Y","�b�x","�t���[<BR>�^�C��","���o������","�ԋp�l")
'3th  DtTbl(0)=Array("���o��","���o<BR>�\���","�w����","���<BR>�ԍ�","�w����","�R���e�i�ԍ�<BR>�^�a�k�ԍ�","��������","�ԋp�\��","�ԋp","�w����","�w����<BR>��","BL�ԍ�","�A���[���t���O","�Ɖ��","�w�����։�","�D��","�D��","�T�C�Y","�b�x","�t���[<BR>�^�C��","���o������","�ԋp�l","���l�P","���l�Q","���l�R")
'Chang 20050303 STAT fro 4th Recon By SEIKO N.Oosige
'  DtTbl(0)=Array("���o��","���o<BR>�\���","�w����","���<BR>�ԍ�","�w����","�R���e�i�ԍ�<BR>�^�a�k�ԍ�","��������","�ԋp�\��","�ԋp","�w����","�w����<BR>��","BL�ԍ�","�A���[���t���O","�Ɖ��","�w�����։�","�D��","�D��","�T�C�Y","�b�x","�t���[<BR>�^�C��","���o������","�ԋp�l","���l�P","���l�Q","�[����P")
  DtTbl(0)=Array("���o��","���o<BR>�\���","�w����","���<BR>�ԍ�","�w����","�R���e�i�ԍ�<BR>�^�a�k�ԍ�","��������","�ԋp�\��","�ԋp","�w����","�w����<BR>��","BL�ԍ�","�A���[���t���O","�Ɖ��","�w�����։�","�D��","�D��","SZ","�b�x","�t���[<BR>�^�C��","���o������","�ԋp�l","���l�P","���l�Q","�[����P","�R�[�h","�S��")
'Chang 20050303 END
'2009/02/25 Add-S G.Ariola
dim ctr
for ctr = 1 to 3
Session(CSTR("Key" & ctr))
if Session(CSTR("Key" & ctr)) <> "" then
	Select Case Session(CSTR("Key" & ctr))
		Case "0" '�����\���
			DtTbl(0)(1) = DtTbl(0)(1) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "1" '�w���� �| �R�[�h
			DtTbl(0)(25) = DtTbl(0)(25) & getImage(Session(CSTR("KeySort" & ctr)))		
		Case "2" '��Ɣԍ�
			DtTbl(0)(3) = DtTbl(0)(3) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "3" '�w����
			DtTbl(0)(4) = DtTbl(0)(4) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "4" '�R���e�i�ԍ�/BL�ԍ�
			DtTbl(0)(5) = DtTbl(0)(5) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "5" '�D��
			DtTbl(0)(15) = DtTbl(0)(15) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "6" '�D��
			DtTbl(0)(16) = DtTbl(0)(16) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "7" 'SZ
			DtTbl(0)(17) = DtTbl(0)(17) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "8" '�b�x
			DtTbl(0)(18) = DtTbl(0)(18) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "9" '�t���[�^�C��
			DtTbl(0)(19) = DtTbl(0)(19) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "10" '�[����P
			DtTbl(0)(24) = DtTbl(0)(24) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "11" '��������
			DtTbl(0)(6) = DtTbl(0)(6) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "12" '�ԋp�\��
			DtTbl(0)(7) = DtTbl(0)(7) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "13" '�ԋp
			DtTbl(0)(8) = DtTbl(0)(8) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "14" '�w���� �| �R�[�h
			DtTbl(0)(9) = DtTbl(0)(9) & getImage(Session(CSTR("KeySort" & ctr)))
'		Case "16" '�w���� �| �S��
'			DtTbl(0)(28) = DtTbl(0)(28) & getImage(Session(CSTR("KeySort" & ctr)))		
		Case "15" '�w�����
			DtTbl(0)(10) = DtTbl(0)(10) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "16" '���l�P
			DtTbl(0)(22) = DtTbl(0)(22) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "17" '���l�Q
			DtTbl(0)(23) = DtTbl(0)(23) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "18" '�w���� �| �S��
			DtTbl(0)(26) = DtTbl(0)(26) & getImage(Session(CSTR("KeySort" & ctr)))
	  End Select
end if	  
next
'2009/02/25 Add-E G.Ariola
'3th Add Start
  Dim DelStr,DelTarget
  DelStr=Array("","��","��","No")
  DelTarget=Array(0,8,10,10)
  DtTbl(0)(14) = 0
'3th Add End
  '�G���[�g���b�v�J�n
    on error resume next
  'DB�ڑ�
    dim ObjConn, ObjRS, StrSQL
    ConnDBH ObjConn, ObjRS
  '�Ώی����擾
    StrSQL = "SELECT count(WkContrlNo) AS CNUM FROM hITCommonInfo ITC "&_
             "WHERE WkType='1' AND (RegisterCode='"& USER &"' "&_
             "OR TruckerSubCode1='"& COMPcd &"' OR TruckerSubCode2='"&_
              COMPcd &"' OR TruckerSubCode3='"& COMPcd &"' OR TruckerSubCode4='"& COMPcd &"') AND Process='R' " &_
              strWhere
    ObjRS.Open StrSQL, ObjConn
    if err <> 0 then
      DisConnDBH ObjConn, ObjRS	'DB�ؒf
      jampErrerP "2","b101","00","�����o:�ꗗ�\��(�Ώی����擾)","101","SQL:<BR>"&strSQL
      Exit Function
    end if
    Num = ObjRS("CNUM")
    ObjRS.close
    ReDim Preserve DtTbl(Num)

  '�f�[�^�擾 '2009/11/02 Tanaka �����\���\�[�g�p��,ITC.InputDate ��ǉ�
    StrSQL ="SELECT T.* FROM (SELECT ITC.DeliverTo,ITC.BLNo,ITC.WorkDate, "&_
			"(CASE WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN ITC.TruckerSubCode3 "&_
			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubCode2 "&_
			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubCode1 "&_
			"WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN CASE WHEN mU.UserType='5' THEN mU.HeadCompanyCode ELSE ITC.RegisterCode END "&_
			"ELSE CASE WHEN mU.UserType='5' THEN mU.HeadCompanyCode ELSE ITC.RegisterCode END "&_
			"END) as Code1, "&_
			"(CASE WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN ITC.TruckerSubName4 "&_
			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubName3 "&_
			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubName2 "&_
			"WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITC.TruckerSubName1 "&_
			"ELSE ITC.TruckerSubName1 "&_
			"END) as Name1, "&_
			"(CASE "&_
			"(CASE WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN ITR.TruckerFlag4 "&_
			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITR.TruckerFlag3 "&_
			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITR.TruckerFlag2 "&_
			"WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITR.TruckerFlag1 "&_
			"ELSE Null END) "&_
			"WHEN 0 THEN '��' "&_
			"WHEN 1 THEN 'Yes' "&_
			"WHEN 2 THEN 'No' "&_
			"ELSE ' ' END) as Flag2, "&_
			"ITC.WkNo, "&_
			"ITC.FullOutType, "&_
			"ITC.BLNo as BLContNo, A.ShipLine, A.FullName as ShipName, ''  as ContSize, "&_
			"SUBSTRING(A.RecTerminal,1,2) as CY, A.FreeTime, ITC.DeliverTo1, "&_
			"ITC.WorkCompleteDate, "&_
			"ITC.ReturnDateStr, "&_
			"(CASE WHEN ITC.FullOutType = '1' THEN (CASE WHEN INC.ReturnTime IS NULL THEN '��' ELSE '��' END) "&_
			"ELSE ' ' END) as ReturnValue, "&_
			"(CASE WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN '' "&_
			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubCode4 "&_
			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubCode3 "&_
			"WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITC.TruckerSubCode2 "&_
			"ELSE ITC.TruckerSubCode1 "&_
			"END) as Code2, "&_
			"(CASE WHEN "&_
			"(CASE WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITC.TruckerSubCode2 "&_
			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubCode3 "&_
			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubCode4 "&_
			"WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN NULL "&_
			"ELSE ITC.TruckerSubCode1 "&_
			"END) IS NULL THEN ' ' ELSE "&_
			"(CASE (CASE WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITR.TruckerFlag2 "&_
			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITR.TruckerFlag3 "&_
			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITR.TruckerFlag4 "&_
			"WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN NULL "&_
			"ELSE ITR.TruckerFlag1 "&_
			"END) "&_
			"WHEN '0' THEN '��' "&_
			"WHEN '1' THEN 'Yes' "&_
			"ELSE 'No' END) "&_
			"END) as Flag1, "&_
			"ITC.Comment1, ITC.Comment2, "&_
			"ITC.ReturnDateVal, ITC.UpdtUserCode, "&_
			"ITR.TruckerFlag1,ITR.TruckerFlag2, ITR.TruckerFlag3, ITR.TruckerFlag4, "&_
			"mU.HeadCompanyCode, mU.UserType, "&_
			"A.CYDelTime, INC.ReturnTime,ITC.InputDate "&_
			"FROM hITCommonInfo ITC "&_
			"LEFT JOIN hITReference ITR ON ITC.WkContrlNo = ITR.WkContrlNo "&_
			"INNER JOIN mUsers AS mU ON ITC.RegisterCode = mU.UserCode "&_
			"LEFT JOIN ImportCont AS INC ON (ITC.ContNo=INC.ContNo) "&_
			"LEFT JOIN mVessel AS mV On (INC.VslCode=mV.VslCode) "&_
			"LEFT JOIN (SELECT Distinct BL.BLNo, INC.FreeTime, MIN(INC.CYDelTime) AS CYDelTime,mV.ShipLine, mV.FullName, BL.RecTerminal "&_
			"FROM ImportCont AS INC "&_
			"LEFT JOIN mVessel AS mV ON INC.VslCode = mV.VslCode "&_
			"LEFT JOIN BL ON INC.VslCode=BL.VslCode AND INC.VoyCtrl=BL.VoyCtrl AND INC.BLNo=BL.BLNo "&_
			"GROUP BY BL.BLNo,INC.FreeTime,mV.ShipLine, mV.FullName, BL.RecTerminal) A ON  A.BLNo=ITC.BLNo "&_
			"WHERE ITC.Process='R' AND WkType='1' AND ITC.FullOutType='4' AND (ITC.RegisterCode='"& USER &"' "&_
			"OR ITC.TruckerSubCode1='"& COMPcd &"' "&_
			"OR ITC.TruckerSubCode2='"& COMPcd &"' "&_
			"OR ITC.TruckerSubCode3='"& COMPcd &"' "&_
			"OR ITC.TruckerSubCode4='"& COMPcd &"')  "& strWhere &" "&_
			"UNION ALL "&_
			"SELECT ITC.DeliverTo,ITC.BLNo,ITC.WorkDate, "&_
			"(CASE WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN ITC.TruckerSubCode3 "&_
			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubCode2 "&_
			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubCode1 "&_
			"WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN CASE WHEN mU.UserType='5' THEN mU.HeadCompanyCode ELSE ITC.RegisterCode END "&_
			"ELSE CASE WHEN mU.UserType='5' THEN mU.HeadCompanyCode ELSE ITC.RegisterCode END "&_
			"END) as Code1, "&_
			"(CASE WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN ITC.TruckerSubName4 "&_
			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubName3 "&_
			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubName2 "&_
			"WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITC.TruckerSubName1 "&_
			"ELSE ITC.TruckerSubName1 "&_
			"END) as Name1, "&_
			"(CASE "&_
			"(CASE WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN ITR.TruckerFlag4 "&_
			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITR.TruckerFlag3 "&_
			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITR.TruckerFlag2 "&_
			"WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITR.TruckerFlag1 "&_
			"ELSE Null END) "&_
			"WHEN 0 THEN '��' "&_
			"WHEN 1 THEN 'Yes' "&_
			"WHEN 2 THEN 'No' "&_
			"ELSE ' ' END) as Flag2, "&_
			"ITC.WkNo, "&_
			"ITC.FullOutType, "&_
			"ITC.ContNo as BLContNo, "&_
			"mV.ShipLine, "&_
			"mV.FullName as ShipName, "&_
			"CON.ContSize, "&_
			"SUBSTRING(BL.RecTerminal,1,2) as CY, "&_
			"INC.FreeTime, "&_
			"ITC.DeliverTo1, "&_
			"ITC.WorkCompleteDate, "&_
			"ITC.ReturnDateStr, "&_
			"(CASE WHEN ITC.FullOutType = '1' THEN (CASE WHEN INC.ReturnTime IS NULL THEN '��' "&_
			" ELSE '��' END) "&_
			"ELSE ' ' END) as ReturnValue, "&_
			"(CASE WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN '' "&_
			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubCode4 "&_
			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubCode3 "&_
			"WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITC.TruckerSubCode2 "&_
			"ELSE ITC.TruckerSubCode1 "&_
			"END) as Code2, "&_
			"(CASE WHEN "&_
			"(CASE WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITC.TruckerSubCode2 "&_
			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubCode3 "&_
			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubCode4 "&_
			"WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN NULL "&_
			"ELSE ITC.TruckerSubCode1 "&_
			"END) IS NULL THEN ' ' "&_
			"ELSE "&_
			"(CASE (CASE WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITR.TruckerFlag2 "&_
			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITR.TruckerFlag3 "&_
			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITR.TruckerFlag4 "&_
			"WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN NULL "&_
			"ELSE ITR.TruckerFlag1 "&_
			"END) "&_
			"WHEN '0' THEN '��' WHEN '1' THEN 'Yes' ELSE 'No' END) END) as Flag1, "&_
			"ITC.Comment1, ITC.Comment2, "&_
			"ITC.ReturnDateVal, ITC.UpdtUserCode, "&_
			"ITR.TruckerFlag1,ITR.TruckerFlag2, ITR.TruckerFlag3, ITR.TruckerFlag4, "&_
			"mU.HeadCompanyCode, mU.UserType, "&_
			"INC.CYDelTime, INC.ReturnTime,ITC.InputDate "&_
			"FROM hITCommonInfo ITC "&_
			"LEFT JOIN hITReference ITR ON ITC.WkContrlNo = ITR.WkContrlNo "&_
			"INNER JOIN mUsers AS mU ON ITC.RegisterCode = mU.UserCode "&_
			"LEFT JOIN ImportCont AS INC ON ITC.ContNo=INC.ContNo "&_
			"LEFT JOIN mVessel AS mV On INC.VslCode=mV.VslCode "&_
			"LEFT JOIN BL ON INC.VslCode=BL.VslCode AND INC.VoyCtrl=BL.VoyCtrl AND INC.BLNo=BL.BLNo "&_
			"LEFT JOIN Container AS CON ON INC.VslCode = CON.VslCode AND INC.VoyCtrl=CON.VoyCtrl AND INC.ContNo=CON.ContNo "&_
			"WHERE ITC.Process='R' AND WkType='1' AND ITC.FullOutType<>'4' AND (ITC.RegisterCode='"& USER &"' "&_
			"OR ITC.TruckerSubCode1='"& COMPcd &"' OR ITC.TruckerSubCode2='"& COMPcd &"' OR ITC.TruckerSubCode3='"& COMPcd &"' OR ITC.TruckerSubCode4='"& COMPcd &"') "& strWhere &") AS T "&_
             strOrder
			 '"ORDER BY isnull(ITC.WorkDate,DATEADD(Year,100,getdate())),ITC.InputDate ASC"
'C-002 ADD This Line : "ITC.Comment1, ITC.Comment2, ITC.Comment3, "&_
'C-004 ADD This Line : "ORDER BY isnull(ITC.WorkDate,DATEADD(Year,100,getdate())),ITC.InputDate ASC"
'3th chage This Line : "ITC.Comment1, ITC.Comment2, ITC.Comment3, "&_
'3th add INC.VoyCtrl, INC.VslCode,
'response.Write(StrSQL)
'response.end
    ObjRS.Open StrSQL, ObjConn
    if err <> 0 then
      DisConnDBH ObjConn, ObjRS	'DB�ؒf
      jampErrerP "2","b101","00","�����o:�ꗗ�\��(�f�[�^�擾)","102","SQL:<BR>"&strSQL
      Exit Function
    end if
    dim tmpBLNo(1),tmptime
    tmpBLNo(0) = ""
    tmpBLNo(1) = ""
    i=1
    Do Until ObjRS.EOF
     If DtTbl(i-1)(3)<>Trim(ObjRS("WkNo")) Then
'C-002      DtTbl(i)=Array("","","","","","","","","","","","","","","","","","","","","","")
      'DtTbl(i)=Array("","","","","","","","","","","","","","","","","","","","","","","","","")
	  DtTbl(i)=Array("","","","","","","","","","","","","","","","","","","","","","","","","","","")
      DtTbl(i)(0)=Trim(ObjRS("DeliverTo"))
      DtTbl(i)(1)=Mid(ObjRS("WorkDate"),3,8)
      DtTbl(i)(3)=Trim(ObjRS("WkNo"))
      DtTbl(i)(4)=Trim(ObjRS("FullOutType"))
      DtTbl(i)(5)=Trim(ObjRS("BLContNo"))
      DtTbl(i)(6)=Trim(Mid(ObjRS("WorkCompleteDate"),3,14))
      If Trim(Mid(DtTbl(i)(6),10))<>"" Then
        tmptime=Split(Mid(DtTbl(i)(6),10),":",3,1)
        DtTbl(i)(6)=Left(DtTbl(i)(6),9)&Right(0&tmptime(0),2)&":"&tmptime(1)
      End If
      DtTbl(i)(7)=Trim(ObjRS("ReturnDateStr"))
'chenge 030722      DtTbl(i)(12)=ObjRS("ReturnArrert")
      DtTbl(i)(8)=ObjRS("ReturnValue")
      DtTbl(i)(21)=Trim(ObjRS("ReturnDateVal"))
'      If DtTbl(i)(4) = 4 Then
'        DtTbl(i)(5)=Trim(ObjRS("BLNo"))
'        DtTbl(i)(11)=Trim(ObjRS("BLNo"))
'      ElseIf DtTbl(i)(4) = 2 Then
'        DtTbl(i)(5)=Trim(ObjRS("ConTNo"))
        DtTbl(i)(11)=Trim(ObjRS("BLNo"))
'      Else
'        DtTbl(i)(5)=Trim(ObjRS("ConTNo"))
'      End If
	  
'2009/02/25 Add-S G.Ariola		
		DtTbl(i)(2) = Trim(ObjRS("Code1"))
		DtTbl(i)(25) = Trim(ObjRS("Name1"))
		DtTbl(i)(9) = Trim(ObjRS("Code2"))
		DtTbl(i)(26) = Trim(ObjRS("Name2"))

		DtTbl(i)(10) = ObjRS("Flag1")
		DtTbl(i)(14) = ObjRS("Flag2")

If Trim(ObjRS("TruckerSubCode4")) = COMPcd Then
DtTbl(i)(13) = 4
ElseIf Trim(ObjRS("TruckerSubCode3")) = COMPcd Then
DtTbl(i)(13) = 3
ElseIf Trim(ObjRS("TruckerSubCode2")) = COMPcd Then
DtTbl(i)(13) = 2
ElseIf Trim(ObjRS("TruckerSubCode1")) = COMPcd Then
DtTbl(i)(13) = 1
Else
DtTbl(i)(13) = 0
end if
'2009/02/25 Add-E G.Ariola	  
'2009/02/25 Del-S G.Ariola	  
'   '�w����Ɖ�ς݃t���O
'      If Trim(ObjRS("TruckerSubCode4")) = COMPcd Then
'        DtTbl(i)(2) = Trim(ObjRS("TruckerSubCode3"))
'        DtTbl(i)(9) = Null
'        DtTbl(i)(13) = 4
'        DtTbl(i)(14) = ObjRS("TruckerFlag4")
'      ElseIf Trim(ObjRS("TruckerSubCode3")) = COMPcd Then
'        DtTbl(i)(2) = Trim(ObjRS("TruckerSubCode2"))
'        DtTbl(i)(9) = Trim(ObjRS("TruckerSubCode4"))
'        DtTbl(i)(13) = 3
'        DtTbl(i)(14) = ObjRS("TruckerFlag3")
'        DtTbl(i)(10) = ObjRS("TruckerFlag4")
'      ElseIf Trim(ObjRS("TruckerSubCode2")) = COMPcd Then
'        DtTbl(i)(2) = Trim(ObjRS("TruckerSubCode1"))
'        DtTbl(i)(9) = Trim(ObjRS("TruckerSubCode3"))
'        DtTbl(i)(13) = 2
'        DtTbl(i)(14) = ObjRS("TruckerFlag2")
'        DtTbl(i)(10) = ObjRS("TruckerFlag3")
'      ElseIf Trim(ObjRS("TruckerSubCode1")) = COMPcd Then
'        If ObjRS("UserType") = "5" Then         'CW-051
'          DtTbl(i)(2) = Trim(ObjRS("HeadCompanyCode"))  'CW-051
'        Else                        'CW-051
'          DtTbl(i)(2) = Trim(ObjRS("RegisterCode"))
'        End If                      'CW-051
'        DtTbl(i)(9) = Trim(ObjRS("TruckerSubCode2"))
'        DtTbl(i)(13) = 1
'        DtTbl(i)(14) = ObjRS("TruckerFlag1")
'        DtTbl(i)(10) = ObjRS("TruckerFlag2")
'      Else
'        If ObjRS("UserType") = "5" Then         'CW-051
'          DtTbl(i)(2) = Trim(ObjRS("HeadCompanyCode"))  'CW-051
'        Else                        'CW-051
'          DtTbl(i)(2) = Trim(ObjRS("RegisterCode"))
'        End If                      'CW-051
'        DtTbl(i)(9) = Trim(ObjRS("TruckerSubCode1"))
'        DtTbl(i)(13) = 0
'        DtTbl(i)(14) = Null
'        DtTbl(i)(10) = ObjRS("TruckerFlag1")
'      End If
'2009/02/25 Del-E G.Ariola	  
	  
'      If IsNull(DtTbl(i)(9)) Then
'        DtTbl(i)(10) ="�@"
'      ElseIf DtTbl(i)(10) = 0 Then
'        DtTbl(i)(10) ="��"
'      ElseIf DtTbl(i)(10) = 1 Then
'        DtTbl(i)(10) ="Yes"
'      Else
'        DtTbl(i)(10) ="No"
'      End If
'      If DtTbl(i)(14)=0 Then
'        DtTbl(i)(14) ="��"
'      ElseIf DtTbl(i)(14) = 1 Then
'        DtTbl(i)(14) ="Yes"
'      ElseIf DtTbl(i)(14) = 2 Then
'        DtTbl(i)(14) ="No"
'      Else
'        DtTbl(i)(14) ="�@"
'      End If

'2009/02/25 Del-S G.Ariola	  
'      If DtTbl(i)(4) = 1  Then
'        If IsNull(ObjRS("ReturnTime")) Then
'          DtTbl(i)(8)="��"
'        Else
'          DtTbl(i)(8)="��"
'        End If
'      Else
'        DtTbl(i)(8)="�@"
'      End If
'2009/02/25 Del-E G.Ariola	  	  
	  
      DtTbl(i)(12)="-"
'      If DtTbl(i)(4) = 4  Then
'        tmpBLNo(0) = tmpBLNo(0) &","& i  
'        tmpBLNo(1) = tmpBLNo(1) &",'"& DtTbl(i)(5) & "'"
'        DtTbl(i)(17)=""
'        DtTbl(i)(15)=""
'        DtTbl(i)(16)=""
'        DtTbl(i)(18)=""
'        DtTbl(i)(19)=""
'        DtTbl(i)(20)=""
'      Else
        DtTbl(i)(17)=Trim(ObjRS("ContSize"))
        DtTbl(i)(15)=Trim(ObjRS("ShipLine"))
'C-001        DtTbl(i)(16)=Left(ObjRS("FullName"),12)
        DtTbl(i)(16)=Trim(ObjRS("ShipName"))
        DtTbl(i)(18)=Trim(ObjRS("CY"))
        DtTbl(i)(19)=Mid(ObjRS("FreeTime"),3,8)
        DtTbl(i)(20)=Trim(ObjRS("CYDelTime"))
'      End If
'C-002 ADD START
      DtTbl(i)(22)= Trim(ObjRS("Comment1"))
      DtTbl(i)(23)= Trim(ObjRS("Comment2"))
'3th Chage Start
'3th'      DtTbl(i)(24)= Trim(ObjRS("Comment3"))
      DtTbl(i)(24)= Trim(ObjRS("DeliverTo1"))
'3th Change End
'C-002 ADD END

'3th Add Start
      If DelType=0 OR (DelType=1 AND DtTbl(i)(DelTarget(DelType)) <> DelStr(DelType)) OR ((DelType=2 OR DelType=3) AND DtTbl(i)(DelTarget(DelType)) = DelStr(DelType)) Then
        DtTbl(0)(14) = DtTbl(0)(14) + DtTbl(i)(13)
        i=i+1
      Else
        Num=Num-1
      End If
'      If (DelType=1 AND DtTbl(i)(DelTarget(DelType)) = DelStr(DelType)) OR ((DelType=2 OR DelType=3) AND DtTbl(i)(DelTarget(DelType)) <> DelStr(DelType)) Then
'        Num=Num-1
'      Else
'        DtTbl(0)(14) = DtTbl(0)(14) + DtTbl(i)(13)
'        i=i+1
'      End If
'      i=i+1
'3th Add End
     End If
      ObjRS.MoveNext
    Loop
    ObjRS.close

    If i - 1 < Num Then
      ErrerM = "<DIV class=alert>�o�^�f�[�^�̂���"& Num-i+1 &"���ɂ��Ċ֘A�f�[�^�擾���s�̂���"&_
               "�\������Ă��܂���B<BR>�V�X�e���Ǘ��҂ɖ₢���킹�Ă��������B</DIV><P>"
      Num = i - 1
'        DisConnDBH ObjConn, ObjRS
'        jampErrerP "2","b101","00","�����o:�ꗗ�\��","106",Num-i+1&"���̃f�["
    ElseIf i > Num Then     'CW-325 ADD
      Num = i - 1           'CW-325 ADD
    End If
    '2010/04/23 M.Marquez Add-S
    if Ubound(DtTbl) < Num Then 
        Num=Ubound(DtTbl)
    end if
    '2010/04/23 M.Marquez Add-E
'    If Err <> 0 Then
'      DisConnDBH ObjConn, ObjRS 'DB�ؒf
'      jampErrerP "2","b101","00","�����o:�ꗗ�\��(�f�[�^�ҏW)","200",i&"�Ԗڂ̃f�[�^�ҏW�G���["
'      Exit Function
'    End If
'
''ADD 20030729 BL�w��̏ꍇ�̒ǉ��f�[�^�擾 Start
'    If tmpBLNo(1) <> "" Then
'      StrSQL = "SELECT INC.BLNo, INC.FreeTime,INC.CYDelTime, mV.ShipLine, mV.FullName, BL.RecTerminal "&_
'               "FROM ImportCont AS INC LEFT JOIN mVessel AS mV ON INC.VslCode = mV.VslCode "&_
'               "LEFT JOIN BL ON (INC.BLNo=BL.BLNo) AND (INC.VoyCtrl=BL.VoyCtrl) AND (INC.VslCode=BL.VslCode) "&_
'               "WHERE INC.BLNo IN("& Mid(tmpBLNo(1),2) &") ORDER BY INC.BLNo,INC.UpdtTime DESC"
''3th add INC.VoyCtrl, INC.VslCode,
'      ObjRS.Open strSQL, ObjConn
'      If Err <> 0 Then
'        DisConnDBH ObjConn, ObjRS
'        jampErrerP "2", "b101", "00", "�����o:�ꗗ�\��(�ǉ����ڎ擾)", "101", "SQL:<BR>" & strSQL
'      End If
'      Dim tmpBLNoA(1)
'      tmpBLNoA(0) = Split(tmpBLNo(0), ",", -1, 1)
'      tmpBLNoA(1) = Split(tmpBLNo(1), ",", -1, 1)
'      tmpBLNo(1) = ""
'      Do Until ObjRS.EOF
'        If tmpBLNo(1) <> Trim(ObjRS("BLNo")) Then
'          For i = 1 To UBound(tmpBLNoA(0))
'            If tmpBLNoA(1)(i) = "'" & Trim(ObjRS("BLNo")) & "'" Then
'              DtTbl(tmpBLNoA(0)(i))(15) = Trim(ObjRS("ShipLine"))
''C-001              DtTbl(tmpBLNoA(0)(i))(16)=Left(ObjRS("FullName"),12)
'              DtTbl(tmpBLNoA(0)(i))(16) = Trim(ObjRS("FullName"))
'              DtTbl(tmpBLNoA(0)(i))(18) = Trim(ObjRS("RecTerminal"))
'              DtTbl(tmpBLNoA(0)(i))(19) = Mid(ObjRS("FreeTime"), 3, 8)
'              DtTbl(tmpBLNoA(0)(i))(20) = Trim(ObjRS("CYDelTime"))
'              tmpBLNo(1) = Trim(ObjRS("BLNo"))
'            End If
'         Next
'       End If
'       ObjRS.MoveNext
'      Loop
'      ObjRS.Close
'      If Err <> 0 Then
'        DisConnDBH ObjConn, ObjRS   'DB�ؒf
'        jampErrerP "2","b101","00","�����o:�ꗗ�\��(�ǉ����ڃf�[�^�ҏW)","200",i&"�Ԗڂ̃f�[�^�ҏW�G���["
'        Exit Function
'      End If
'   End If

'ADD 20030729 BL�w��̏ꍇ�̒ǉ��f�[�^�擾 End
  'DB�ڑ�����
    DisConnDBH ObjConn, ObjRS
  '�G���[�g���b�v����
    on error goto 0
End Function
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>�����o���ꗗ</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
//�f�[�^�������ꍇ�̕\������
function vew(){
<%If Num<>0 Then%>
 //var obj1=document.getElementById("HDIV");
 //var obj2=document.getElementById("VDIV");
 var obj3=document.getElementById("BDIV");

/* if((document.body.offsetWidth-10) < 1361){
    obj1.style.width=document.body.offsetWidth-10;
	obj1.style.overflow="auto";
 }
  else
 {
 	 obj1.style.width=document.body.offsetWidth-10;
	 obj1.style.overflowX="auto";
 }
 obj2.style.height=document.body.offsetHeight-80;*/
 if((document.body.offsetWidth-120)  < 880){
    obj3.style.width=document.body.offsetWidth;
    obj3.style.overflowX="auto";
 }
 else{
 	 <% If DtTbl(0)(14)<>0 Then %>
	 obj3.style.width=document.body.offsetWidth-120;
	 <% Else %>
  	 obj3.style.width=document.body.offsetWidth-79;
	 <% End If %>
  	 obj3.style.overflowX="auto";
 }
 obj3.style.height=document.body.offsetHeight;
 obj3.style.overflowY="auto";
<% End If %>
}
//�X�V
function GoRenew(No){
  Fname=document.dmo010F;
  Fname.targetNo.value=No;
  Fname.action="./dmi015.asp";
  newWin = window.open("", "ReEntry", "left=10,top=10,status=yes,resizable=yes,scrollbars=yes");
  Fname.target="ReEntry";
  Fname.elements[No].disabled=false;
  Fname.submit();
  Fname.elements[No].disabled=true;
  Fname.target="_self";
}
//�R���e�i�ڍ�
function GoConinf(wkcNo,flag,conNo){
  Fname=document.dmo010F;
  Fname.SakuNo.value=wkcNo;
  Fname.flag.value=flag;
  Fname.CONnum.value=conNo;
//  Fname.InfoFlag.value="9";
//  Fname.action="./dmi015.asp";
  Fname.action="./dmo900.asp";
  newWin = window.open("", "ConInfo", "left=30,top=10,status=yes,scrollbars=yes,resizable=yes,menubar=yes");
  Fname.target="ConInfo";
  Fname.submit();
  Fname.target="_self";
  Fname.InfoFlag.value="0";
}
//����
function SerchC(SortFlag,Kye){
  Fname=document.dmo010F;
  Fname.SortFlag.value=SortFlag;
  Fname.SortKye.value=Kye;
  Fname.target="_self";
  Fname.action="./dmo010L.asp";
  Fname.submit();
}
//�Ɖ��
function GoSyokaizumi(){
  target=document.dmo010F;
  if(target.DataNum.value>0){
    flag = confirm('���񓚂̉񓚂��uYes�v�ɂ��܂����H');
    if(flag==true){
      len=target.elements.length;
      for(i=0;i<len;i++){
        target.elements[i].disabled=false;
      }
      target.SortFlag.value=8;
      target.target="_self";
      target.action="./dmo010L.asp";
      target.submit();
    }
  }
}
//�W�J
function GoDevelop(No){
  Fname=document.dmo010F;
  Fname.targetNo.value=No;
  Fname.action="./dmo030.asp";
  newWin = window.open("", "ConInfo", "left=50,top=10,status=yes,scrollbars=yes,resizable=yes,menubar=yes");
  Fname.target="ConInfo";
  Fname.elements[0].disabled=false;
  Fname.elements[No].disabled=false;
  Fname.submit();
  Fname.elements[0].disabled=true;
  Fname.elements[No].disabled=true;
  Fname.target="_self";
}
//CSV		ADD C-001
function GoCSV(){
  target=document.dmo010F;
  target.target="Bottom";
  len=target.elements.length;
  for(i=0;i<len;i++){
    target.elements[i].disabled=false;
  }
  target.action="./dmo080.asp";
  target.submit();
}
function showContent(){
    var target=null;
    while (target==null) {
	    target=parent.window.frames(0);
	}
    var target1 = target.window.document.getElementById("loading");
    target1.style.display='none';
    //show content
    document.getElementById("content").style.display='block';
}
// -->

</SCRIPT>
<style type="text/css">
INPUT.chrReadOnly
{
    BORDER-BOTTOM: 0px solid;
    BORDER-LEFT: 0px solid;
    BORDER-RIGHT: 0px solid;
    BORDER-TOP: 0px solid;
	PADDING-BOTTOM: 0px solid;
    PADDING-LEFT: 0px solid;
    PADDING-RIGHT: 0px solid;
    PADDING-TOP: 0px solid;
	margin-bottom:0px solid;
	margin-left:0px solid;
	margin-right:0px solid;
	margin-top:0px solid;
    FONT-FAMILY: '�l�r �o�S�V�b�N';
    FONT-SIZE: 10pt;
    TEXT-ALIGN: left
}
DIV.BDIV
{
    position: relative;
    border-width: 0px 0px 1px 0px;
}
thead tr 
{
    position: relative;
    top: expression(this.offsetParent.scrollTop);
}
th.hlist 
{
    position: relative;
}
table {
    border-width: 0px 1px 1px 0px;
}
th {
    border-width: 1px 1px 1px 1px;
    padding: 4px;
    background-color: #ffcc33;
}
</style>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin="0" topMargin="0" marginheight="0" marginwidth="0" onLoad="vew();" onResize="vew()">
<!--setTimeout('showContent()', 500); -->
<!-------------�����o���ꗗ���List--------------------------->


<!--<div id="content" style="display:none;"> -->
<%=ErrerM%>
<Form name="dmo010F" method="POST">
<!--<DIV id="HDIV" style=" overflow-x:scroll;">
<DIV style=" width:1361; height:41;"><!--2009/02/25 G.Ariola --> 
<div id="BDIV">
<TABLE id="testt" border="1" cellPadding="2" cellSpacing="0">
<%If Num>0 Then%> 
  <% If DtTbl(0)(14)<>0 Then %>
  <thead>
  <tr>
    <TH class="hlist" nowrap><%=DtTbl(0)(1)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(2)%><BR><%=DtTbl(0)(25)%></TH>
    <!--<TH class="hlist" nowrap>�w����<BR>�։�</TH> -->
    <TH class="hlist" nowrap><%=DtTbl(0)(3)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(4)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(5)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(15)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(16)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(17)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(18)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(19)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(24)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(6)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(7)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(8)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(9)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(10)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(22)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(23)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(26)%>
    <INPUT type=hidden name='Datatbl0' disabled value='<%=DtTbl(0)(0)%>,<%=DtTbl(0)(1)%>,<%=DtTbl(0)(2)%>,<%=DtTbl(0)(3)%>,<%=DtTbl(0)(4)%>,<%=DtTbl(0)(5)%>,<%=DtTbl(0)(6)%>,<%=DtTbl(0)(7)%>,<%=DtTbl(0)(8)%>,<%=DtTbl(0)(9)%>,<%=DtTbl(0)(10)%>,<%=DtTbl(0)(11)%>,<%=DtTbl(0)(12)%>,<%=DtTbl(0)(13)%>,<%=DtTbl(0)(14)%>,<%=DtTbl(0)(15)%>,<%=DtTbl(0)(16)%>,<%=DtTbl(0)(17)%>,<%=DtTbl(0)(18)%>,<%=DtTbl(0)(19)%>,<%=DtTbl(0)(20)%>,<%=DtTbl(0)(21)%>,<%=DtTbl(0)(22)%>,<%=DtTbl(0)(23)%>,<%=DtTbl(0)(24)%>,<%=Left(DtTbl(j)(25),8)%>'>
    </TH>
  </TR>
  <!--2009/02/25 Add-S G.Ariola test1 -->
  	<!--<TR>
	<TH width="100"><'%=DtTbl(0)(25)%></TH>
	<TH width="100"><'%=DtTbl(0)(26)%></TH>
	<!--<TH width="60"><%'=DtTbl(0)(27)%></TH>
	<TH width="60"><%'=DtTbl(0)(28)%></TH> -->
	<!--</TH></TR> -->
	<!--2009/02/25 Add-E G.Ariola -->
<!--</TABLE> --> <!--2009/02/25 G.Ariola -->
<!--</DIV> --><!--2009/02/25 G.Ariola -->
</THEAD>
<%'If Num>10 Then%><!--<DIV id="VDIV" style=" width:1337; height:242; overflow-y:auto;"> --><!--2009/02/25 G.Ariola -->
  <%'else%><!--<DIV id="VDIV" style=" width:1321; height:242; overflow-y:auto;"> --><!--2009/02/25 G.Ariola -->
 <%'end if%> <!--2009/02/25 G.Ariola -->
<!--<TABLE border="1" cellPadding="2" cellSpacing="0" cols="<%=Num+20%>"> --><!--2009/02/25 G.Ariola -->  
	<tbody>
    <% For j=1 to Num%>
  <TR class=bgw>
    <TD nowrap><%=DtTbl(j)(1)%><BR></TD><TD nowrap><%=DtTbl(j)(2)%><BR></TD>
    <!--<TD nowrap><%=DtTbl(j)(14)%><BR></TD>  -->
    <TD nowrap><A HREF="JavaScript:GoRenew('<%=j%>');"><%=DtTbl(j)(3)%></A><BR></TD>
    <TD nowrap><A HREF="JavaScript:GoDevelop('<%=j%>');"><%=Siji(DtTbl(j)(4))%></A><BR></TD>
    <TD nowrap><A HREF="JavaScript:GoConinf('<%=DtTbl(j)(3)%>',<%=DtTbl(j)(4)%>,'<%=DtTbl(j)(5)%>')"><%=DtTbl(j)(5)%></A><BR>
    </TD>
<%'C-001    <TD nowrap><%=DtTbl(j)(15)% >�@</TD><TD nowrap><%=DtTbl(j)(16)% >�@</TD><TD nowrap><%=DtTbl(j)(17)% >�@</TD> -->%>
    <TD nowrap><%=DtTbl(j)(15)%><BR></TD><TD nowrap><%=Left(DtTbl(j)(16),12)%><BR></TD><TD nowrap><%=DtTbl(j)(17)%><BR></TD>
    <TD nowrap><%=Left(DtTbl(j)(18),2)%><BR></TD><TD nowrap><%=DtTbl(j)(19)%><BR></TD><TD nowrap><%=Left(DtTbl(j)(24),10)%><BR></TD>
    <TD nowrap><%=DtTbl(j)(6)%><BR></TD><TD nowrap><%=DtTbl(j)(7)%><BR></TD><TD nowrap><%=DtTbl(j)(8)%><BR></TD>
    <TD nowrap><%=DtTbl(j)(9)%><BR></TD><!--<TD width="60"><%'=Left(DtTbl(j)(25),4)%><BR></TD> --><TD nowrap><%=DtTbl(j)(10)%><BR></TD>
    <TD nowrap><%=Left(DtTbl(j)(22),10)%><BR></TD><TD nowrap><%=Left(DtTbl(j)(23),10)%><BR></TD><TD nowrap><%=Left(DtTbl(j)(25),8)%><BR>
    <INPUT type=hidden name='Datatbl<%=j%>' disabled value='<%=DtTbl(j)(0)%>,<%=DtTbl(j)(1)%>,<%=DtTbl(j)(2)%>,<%=DtTbl(j)(3)%>,<%=DtTbl(j)(4)%>,<%=DtTbl(j)(5)%>,<%=DtTbl(j)(6)%>,<%=DtTbl(j)(7)%>,<%=DtTbl(j)(8)%>,<%=DtTbl(j)(9)%>,<%=DtTbl(j)(10)%>,<%=DtTbl(j)(11)%>,<%=DtTbl(j)(12)%>,<%=DtTbl(j)(13)%>,<%=DtTbl(j)(14)%>,<%=DtTbl(j)(15)%>,<%=DtTbl(j)(16)%>,<%=DtTbl(j)(17)%>,<%=DtTbl(j)(18)%>,<%=DtTbl(j)(19)%>,<%=DtTbl(j)(20)%>,<%=DtTbl(j)(21)%>,<%=DtTbl(j)(22)%>,<%=DtTbl(j)(23)%>,<%=DtTbl(j)(24)%>,<%=Left(DtTbl(j)(25),8)%>'>
    </TD>
  </TR>
    <% Next %>
<!--</TABLE> -->	<!--2009/02/25 G.Ariola -->
</tbody>
<%'If Num>10 Then%><!--</DIV> --><%'end if%> 	<!--2009/02/25 G.Ariola -->	
  <% Else %>
  <thead>  
     <tr >
  <INPUT type=hidden name='Datatbl0' disabled value='<%=DtTbl(0)(0)%>,<%=DtTbl(0)(1)%>,<%=DtTbl(0)(2)%>,<%=DtTbl(0)(3)%>,<%=DtTbl(0)(4)%>,<%=DtTbl(0)(5)%>,<%=DtTbl(0)(6)%>,<%=DtTbl(0)(7)%>,<%=DtTbl(0)(8)%>,<%=DtTbl(0)(9)%>,<%=DtTbl(0)(10)%>,<%=DtTbl(0)(11)%>,<%=DtTbl(0)(12)%>,<%=DtTbl(0)(13)%>,<%=DtTbl(0)(14)%>,<%=DtTbl(0)(15)%>,<%=DtTbl(0)(16)%>,<%=DtTbl(0)(17)%>,<%=DtTbl(0)(18)%>,<%=DtTbl(0)(19)%>,<%=DtTbl(0)(20)%>,<%=DtTbl(0)(21)%>,<%=DtTbl(0)(22)%>,<%=DtTbl(0)(23)%>,<%=DtTbl(0)(24)%>,<%=Left(DtTbl(j)(25),8)%>'>
    <TH class="hlist" nowrap><%=DtTbl(0)(1)%></TH><TH nowrap><%=DtTbl(0)(2)%><BR><%=DtTbl(0)(25)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(3)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(4)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(5)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(15)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(16)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(17)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(18)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(19)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(24)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(6)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(7)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(8)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(9)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(10)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(22)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(23)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(26)%>    
    </TH>
  </TR>
  <!--2009/02/25 Add-S G.Ariola -->
  <!--<TR>
	<TH width="100"><'%=DtTbl(0)(25)%></TH>
	<TH width="100"><'%=DtTbl(0)(26)%></TH> -->
	<!--<TH width="60"><%'=DtTbl(0)(27)%></TH>
	<TH width="60"><%'=DtTbl(0)(28)%></TH> -->
	<!--</TH></TR> -->
</thead>  
  <!--</TABLE>  --><!--2009/02/25 G.Ariola -->
 <!-- </DIV> --><!--2009/02/25 G.Ariola -->
 <%'If Num>10 Then%><!--<DIV id="VDIV" style=" width:1337; height:242; overflow-y:auto;"> --><!--2009/02/25 G.Ariola -->
  <%'else%><!--<DIV id="VDIV" style=" width:1321; height:242; overflow-y:auto;"> --><!--2009/02/25 G.Ariola -->
 <%'end if%> <!--2009/02/25 G.Ariola -->
  <!--<TABLE border="1" cellPadding="2" cellSpacing="0" cols="<%=Num+20%>">   -->
  <tbody>
    <% For j=1 to Num %>
  <TR class=bgw><INPUT type=hidden name='Datatbl<%=j%>' disabled value='<%=DtTbl(j)(0)%>,<%=DtTbl(j)(1)%>,<%=DtTbl(j)(2)%>,<%=DtTbl(j)(3)%>,<%=DtTbl(j)(4)%>,<%=DtTbl(j)(5)%>,<%=DtTbl(j)(6)%>,<%=DtTbl(j)(7)%>,<%=DtTbl(j)(8)%>,<%=DtTbl(j)(9)%>,<%=DtTbl(j)(10)%>,<%=DtTbl(j)(11)%>,<%=DtTbl(j)(12)%>,<%=DtTbl(j)(13)%>,<%=DtTbl(j)(14)%>,<%=DtTbl(j)(15)%>,<%=DtTbl(j)(16)%>,<%=DtTbl(j)(17)%>,<%=DtTbl(j)(18)%>,<%=DtTbl(j)(19)%>,<%=DtTbl(j)(20)%>,<%=DtTbl(j)(21)%>,<%=DtTbl(j)(22)%>,<%=DtTbl(j)(23)%>,<%=DtTbl(j)(24)%>,<%=Left(DtTbl(j)(25),8)%>'>
    <TD nowrap><%=DtTbl(j)(1)%><BR></TD><TD nowrap><%=DtTbl(j)(2)%><BR></TD>
    <TD nowrap><A HREF="JavaScript:GoRenew('<%=j%>');"><%=DtTbl(j)(3)%></A><BR></TD>
    <TD nowrap><A HREF="JavaScript:GoDevelop('<%=j%>');"><%=Siji(DtTbl(j)(4))%></A><BR></TD>
    <TD nowrap><A HREF="JavaScript:GoConinf('<%=DtTbl(j)(3)%>',<%=DtTbl(j)(4)%>,'<%=DtTbl(j)(5)%>')"><%=DtTbl(j)(5)%></A><BR>
    </TD>
<%'C-001    <TD nowrap><%=DtTbl(j)(15)% ></TD><TD nowrap><%=DtTbl(j)(16)% ></TD><TD nowrap><%=DtTbl(j)(17)% ></TD> %>
    <TD nowrap><%=DtTbl(j)(15)%><BR></TD><TD nowrap><%=Left(DtTbl(j)(16),12)%><BR></TD><TD nowrap><%=DtTbl(j)(17)%><BR></TD>
    <TD nowrap><%=Left(DtTbl(j)(18),2)%><BR></TD><TD nowrap><%=DtTbl(j)(19)%><BR></TD><TD nowrap><%=Left(DtTbl(j)(24),10)%><BR></TD>
    <TD nowrap><%=DtTbl(j)(6)%><BR></TD><TD nowrap><%=DtTbl(j)(7)%><BR></TD><TD nowrap><%=DtTbl(j)(8)%><BR></TD>
    <TD nowrap><%=DtTbl(j)(9)%><BR></TD><!--<TD width="60"><%'=Left(DtTbl(j)(26),4)%><BR></TD> --><TD nowrap><%=DtTbl(j)(10)%><BR></TD>
    <TD nowrap><%=Left(DtTbl(j)(22),10)%><BR></TD><TD nowrap><%=Left(DtTbl(j)(23),10)%><BR></TD><TD nowrap><%=Left(DtTbl(j)(25),8)%><BR></TD>
  </TR>
    <% Next %>
  <!--</TABLE> -->	<!--2009/02/25 G.Ariola -->
  </tbody>  
 <%'If Num>10 Then%><!--</DIV> --><%'end if%> 	<!--2009/02/25 G.Ariola -->
  <% End If %>
  <!--</DIV> -->
<% Else %>
  <TR class=bgw><TD nowrap>��ƈČ��͂���܂���</TD></TR>
<% End If %>
</TABLE>
</div>
<%'3th del Set_Data Num,DtTbl %>
  <INPUT type=hidden name=DataNum value="<%=Num%>">
  <INPUT type=hidden name=SortFlag value="<%=SortFlag%>" >
  <INPUT type=hidden name=SortKye value="<%=SortKye%>" >
  <INPUT type=hidden name=InfoFlag value="" >
  <INPUT type=hidden name=SakuNo value="" >
  <INPUT type=hidden name=flag value="" >
  <INPUT type=hidden name=targetNo value="" >
  <INPUT type=hidden name=CONnum value="" >
  <INPUT type=hidden name=strWhere value="<%=strWrer%>" disabled>
</Form>
<!--</div> -->
<!-------------��ʏI���--------------------------->
</BODY></HTML>
