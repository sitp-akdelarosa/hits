<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/     SystemName      :Hits                                   _/
'_/     FileName        :dmo310L.asp                            _/
'_/     Function        :�����o���ꗗ��ʃ��X�g�o��           _/
'_/     Date            :2003/05/29                             _/
'_/     Code By         :SEIKO Electric.Co ��d                 _/
'_/     Modify          :C-001 2003/08/07       CSV�o�͑Ή�     _/
'_/                     :C-002 2003/08/07       ���l���Ή�      _/
'_/                     :C-003 2003/08/22       ��Ɣԍ��ł̌���_/
'_/                     :C-004 2003/08/22       �\�������`      _/
'_/						:3th   2004/01/31	3���Ή�	_/
'_/						:		2006/03/06	Booking�d���Ή�	_/
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
<!--#include File="CommonFunc.inc"-->
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
  dim Num, DtTbl,i,j,SortFlag,SortKye
  dim Num2
  dim ObjConn, ObjRS
  dim RecCtr
  dim abspage 
  dim pagecnt
  const gcPage = 10
  
  If Request("SortFlag") = "" Then
    SortFlag = 0
  Else
    SortFlag = Request("SortFlag")
  End If

  '�\�[�g�P�[�X
  dim strWrer,ErrerM

  ErrerM =""
  
  dim strOrder
  dim FieldName
  ReDim FieldName(22)
  
  Dim Kari()								'2016/10/12 H.Yoshikawa Add
  Dim bgclr									'2016/08/17 H.Yoshikawa Add

  FieldName=Array("WorkDate","Code1","WkNo","BookNo","ContNo","ShipLine","VslName","ContSize","ContHeight","TareWeight","ReceiveFrom","RecTerminal","CYCut","WorkComplete1","Code2","Flag1","Comment1", "Comment2","Comment3","Name1")
  
  strOrder = getSort(Session("Key1"),Session("KeySort1"),"")
  strOrder = getSort(Session("Key2"),Session("KeySort2"),strOrder)
  strOrder = getSort(Session("Key3"),Session("KeySort3"),strOrder)

  SortKye=gfTrim(Request("SortKye"))
  if SortKye = "" then
  	ErrerM ="�����������w�肵�Ă�������"
  else
	  Select Case SortFlag
	      Case "0" '�����\�������\������ɕ\��
	          WriteLogH "b401", "���������O���ꗗ","01",""
	          strWrer = "AND ('" & calcDate1 & "' <= ITC.WorkCompleteDate Or ITC.WorkCompleteDate IS Null) "
	          getData DtTbl,strWrer,0
	      Case "1" '�w���悪���Ɖ�̃R���e�i�ꗗ
	          WriteLogH "b401", "���������O���ꗗ","03",""
	          strWrer = "AND ('" & calcDate1 & "' <= ITC.WorkCompleteDate Or ITC.WorkCompleteDate IS Null) "
	          getData DtTbl,strWrer,1
	      Case "7" '�ۗ�
	          WriteLogH "b401", "���������O���ꗗ","07",""
	          strWrer = "AND ('" & calcDate1 & "' <= ITC.WorkCompleteDate Or ITC.WorkCompleteDate IS Null) "
	          getData DtTbl,strWrer,2
	      Case "2" '�����������������ׂĕ\��
	        WriteLogH "b401", "���������O���ꗗ","02",""
	        strWrer="AND ITC.WorkCompleteDate IS Null "
	        getData DtTbl,strWrer,0
	      Case "3" '�S���\��
	          WriteLogH "b401", "���������O���ꗗ","04",""
	          strWrer = " "
	        getData DtTbl,strWrer,0
	      Case "4" '�u�b�L���O�ԍ��Ō���
	          SortKye=Request("SortKye")
	          WriteLogH "b401", "���������O���ꗗ","11",SortKye
	          strWrer = "AND CYV.BookNo LIKE '%" & SortKye & "'"
	          getData DtTbl,strWrer,0
	      Case "5" '�R���e�i�ԍ��Ō���
	          SortKye=Request("SortKye")
	          WriteLogH "b401", "���������O���ꗗ","11",SortKye
	          strWrer = "AND ITC.ContNo LIKE '%" & SortKye & "'"
	          getData DtTbl,strWrer,0
	      Case "11" '��Ɣԍ��Ō���
	          SortKye=Request("SortKye")
	          WriteLogH "b401", "���������O���ꗗ","11",SortKye
	          strWrer = "AND ITC.WkNo LIKE '%" & SortKye & "'"
	          getData DtTbl,strWrer,0
	      Case "8" '�Ɖ��
	          WriteLogH "b407", "���������O���Ɖ�","01",SortKye
	          Get_Data Num2,DtTbl
	        '�G���[�g���b�v�J�n
	          on error resume next
	        'DB�ڑ�
	          dim StrSQL
	          ConnDBH ObjConn, ObjRS
	          For i=1 To Num2
	            If DtTbl(i)(9) <> 0 AND DtTbl(i)(6)="" AND DtTbl(i)(8)="�@" AND DtTbl(i)(10)="��" Then
	              StrSQL = "UPDATE hITReference SET UpdtTime='"& Now() &"', UpdtPgCd='PREDEF01'," &_
	                       "UpdtTmnl='"& USER &"', TruckerFlag"& DtTbl(i)(9) &"=1 "&_
	                       "WHERE WkContrlNo IN (select WkContrlNo From hITCommonInfo "&_
	                       "WHERE WkNo='"& DtTbl(i)(3) &"' AND WkType='3' AND Process='R' )"
	              ObjConn.Execute(StrSQL)
	              if err <> 0 then
	                Set ObjRS = Nothing
	                jampErrerPDBH ObjConn,"2","b407","01","�������F�Ɖ�","104","SQL:<BR>"&strSQL
	              end if
	              DtTbl(i)(10)="Yes"
	            End If
	          Next
	        'DB�ڑ�����
	          DisConnDBH ObjConn, ObjRS
	        '�G���[�g���b�v����
	          on error goto 0
	  End Select
  end if

'�f�[�^�擾�֐�
'2009/02/25 Add-S G.Ariola
Function getSort(Key,SortKey,str)
        
		if str = "" AND Key<>"" then
		    str = " ORDER BY "
		elseif str <> "" AND Key<>"" Then 
		    str = str & ","
		elseif str = "" AND Key = "" then
		    str =" ORDER BY ISNULL(WorkDate,DATEADD(Year,100,getdate())),InputDate ASC"		
		end if
		if Key <> "" then 
		    if (FieldName(CInt(Key)) = "WorkDate" OR FieldName(CInt(Key)) = "CYCut" OR FieldName(CInt(Key)) = "WorkComplete1") AND SortKey = "ASC" then 
		        str = str & " ISNULL(" & FieldName(CInt(Key)) & ",DATEADD(Year,100,getdate())) " & SortKey	
		    else
		        str = str & FieldName(Key) & " " & SortKey	
		    end if			
        end if
       getSort = str  
end function

Function getImage(SortKey)
getImage = ""
		if SortKey = "ASC" then
			getImage = CONST_ASC	
		else
			getImage = CONST_DESC
		end if	
end function
Function getData(DtTbl,strWrer,DelType)
  ReDim DtTbl(1)
  DtTbl(0)=Array("���͓�","����<BR>�\���","�w����<BR>�R�[�h","���<BR>�ԍ�","�u�b�L���O�ԍ�","�R���e�i�ԍ�","��������","�w����","�w����<BR>��","�Ɖ��Frag","�w�����։�","�D��","�D��","SZ","H","������","CY","CY<BR>�J�b�g��","���l�P","���l�Q","���l�R","���l�S","TW","�R�[�h","�w����<BR>�S��","���")
dim ctr
for ctr = 1 to 3
Session(CSTR("Key" & ctr))
if Session(CSTR("Key" & ctr)) <> "" then
	Select Case Session(CSTR("Key" & ctr))
		Case "0" '�����\���
			DtTbl(0)(1) = DtTbl(0)(1) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "1" '�w���� �| �R�[�h
			DtTbl(0)(2) = DtTbl(0)(2) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "2" '��Ɣԍ�
			DtTbl(0)(3) = DtTbl(0)(3) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "3" '�u�b�L���O�ԍ�
			DtTbl(0)(4) = DtTbl(0)(4) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "4" '�R���e�i�ԍ�
			DtTbl(0)(5) = DtTbl(0)(5) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "5" '�D��
			DtTbl(0)(11) = DtTbl(0)(11) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "6" '�D��
			DtTbl(0)(12) = DtTbl(0)(12) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "7" 'SZ
			DtTbl(0)(13) = DtTbl(0)(13) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "8" 'H
			DtTbl(0)(14) = DtTbl(0)(14) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "9" 'TW
			DtTbl(0)(22) = DtTbl(0)(22) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "10" '������
			DtTbl(0)(15) = DtTbl(0)(15) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "11" 'CY
			DtTbl(0)(16) = DtTbl(0)(16) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "12" 'CY�J�b�g��
			DtTbl(0)(17) = DtTbl(0)(17) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "13" '��������
			DtTbl(0)(6) = DtTbl(0)(6) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "14" '�w���� �| �R�[�h
			DtTbl(0)(7) = DtTbl(0)(7) & getImage(Session(CSTR("KeySort" & ctr)))
'		Case "16" '�w���� �| �S��
'			DtTbl(0)(26) = DtTbl(0)(26) & getImage(Session(CSTR("KeySort" & ctr)))		
		Case "15" '�w�����
			DtTbl(0)(8) = DtTbl(0)(8) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "16" '���l�P
			DtTbl(0)(18) = DtTbl(0)(18) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "17" '���l�Q
			DtTbl(0)(19) = DtTbl(0)(19) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "18" '���l�R
			DtTbl(0)(20) = DtTbl(0)(20) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "19" '�w���� �| �S��
			DtTbl(0)(24) = DtTbl(0)(24) & getImage(Session(CSTR("KeySort" & ctr)))
	  End Select
end if	  
next
'2009/02/25 Add-E G.Ariola
  DtTbl(0)(10)=0
'3th Add Start
  Dim DelStr,DelTarget
  DelStr=Array("","��","No")
  DelTarget=Array(0,8,8)
'3th Add End
  '�G���[�g���b�v�J�n
    on error resume next
  'DB�ڑ�
    dim StrSQL
    ConnDBH ObjConn, ObjRS
    
    '2016/11/18 H.Yoshikawa Add Start
    dim UserType
    StrSQL = "SELECT UserType FROM mUsers WHERE UserCode = '" & gfSQLEncode(USER) & "'"
    ObjRS.Open StrSQL, ObjConn
    if err <> 0 then
      DisConnDBH ObjConn, ObjRS 'DB�ؒf
      jampErrerP "2","b401","01","�������F���[�U�f�[�^�擾","101","SQL:<BR>"&StrSQL
      Exit Function
    end if
    if not ObjRS.EOF then
    	UserType = ObjRS("UserType")
    end if
    ObjRS.close    

	if UserType <> "4" then
		ErrerM ="�����������w�肵�Ă�������"
		exit function
	end if
  '�Ώی����擾
	dim OpeName
	select case gfTrim(User)
	case "HKK"
		OpeName = "�����`�^"
	case "KAM"
		OpeName = "��g"
	case "KTC"
		OpeName = "�W�F�l�b�N"
	case "MLC"
		OpeName = "�O�H�q��"
	case "NEC"
		OpeName = "���{�ʉ^"
	case "SOG"
		OpeName = "���݉^�A"
	case else
		OpeName = ""
	end select
	
    StrSQL = "SELECT count(WkContrlNo) AS CNUM FROM hITCommonInfo AS ITC "&_
             "INNER JOIN CYVanInfo AS CYV ON ITC.WkNo = CYV.WkNo AND ITC.ContNo=CYV.ContNo "&_
             "LEFT JOIN ExportCont AS EPC ON CYV.BookNo = EPC.BookNo AND CYV.ContNo = EPC.ContNo "&_
             "LEFT JOIN VslSchedule AS VSLS ON EPC.VoyCtrl = VSLS.VoyCtrl AND EPC.VslCode = VSLS.VslCode "&_
             "LEFT JOIN Booking AS BOK ON EPC.VslCode = BOK.VslCode AND EPC.VoyCtrl = BOK.VoyCtrl AND EPC.BookNo = BOK.BookNo "&_
             "WHERE WkType='3' AND BOK.Sender like '%" & OpeName & "%' AND Process='R' " &_
              strWrer 
    ObjRS.Open StrSQL, ObjConn
    if err <> 0 then
      DisConnDBH ObjConn, ObjRS 'DB�ؒf
      jampErrerP "2","b401","01","�������F�f�[�^�擾","101","SQL:<BR>"&StrSQL
      Exit Function
    end if
    Num = ObjRS("CNUM")
    ObjRS.close    
  '�f�[�^�擾
    StrSQL = "SELECT T.* FROM "&_
             "(SELECT ITC.InputDate, ITC.WorkDate, ITC.WkNo, ITC.WorkCompleteDate, ITC.ContNo, ITC.RegisterCode, "&_
             "ITC.TruckerSubCode1, ITC.TruckerSubCode2, ITC.TruckerSubCode3, ITC.TruckerSubCode4, ITC.UpdtUserCode,"&_
             "ITC.TruckerSubName1, ITC.TruckerSubName2, ITC.TruckerSubName3, ITC.TruckerSubName4, " &_
             "ITC.Comment1, ITC.Comment2, ITC.Comment3, ITC.WkContrlNo, "&_
             "ITR.TruckerFlag1, ITR.TruckerFlag2, ITR.TruckerFlag3, ITR.TruckerFlag4, "&_
             "CYV.BookNo, CYV.ShipLine, CYV.VslName, CYV.ContSize, CYV.ContHeight, CYV.ReceiveFrom, "&_
             "CASE ISNULL(CYV.TareWeight,0) "&_
			 "   WHEN 0 THEN '-' "&_
			 "   ELSE CYV.TareWeight "&_ 
		     "END  TareWeight,"&_
             "BOK.RecTerminal, VSLS.CYCut,mU.HeadCompanyCode, mU.UserType "&_
             ",(CASE WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN ITC.TruckerSubCode3 "&_
			 "       WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubCode2 "&_
			 "       WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubCode1 "&_
			 "       WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN CASE WHEN mU.UserType='5' THEN mU.HeadCompanyCode ELSE ITC.RegisterCode END "&_
			 "       ELSE CASE WHEN mU.UserType='5' THEN mU.HeadCompanyCode ELSE ITC.RegisterCode END "&_
			 "  END) as Code1, "&_
			 "ITC.WorkCompleteDate AS WorkComplete1," &_
			 "(CASE WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN '' "&_
			 "      WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubCode4 "&_
			 "      WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubCode3 "&_
			 "      WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITC.TruckerSubCode2 "&_
			 "      ELSE ITC.TruckerSubCode1 "&_
			 " END) as Code2, "&_
             "RTRIM(CASE WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN ITC.TruckerSubName4 "&_
			 "			 WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubName3 "&_
			 "			 WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubName2 "&_
			 "			 WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITC.TruckerSubName1 "&_
			 "			 ELSE ITC.TruckerSubName1 "&_
			 "	    END) as Name1, "&_
		     "CASE   "&_
			 "	   (CASE WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITR.TruckerFlag1 "&_
			 "		     WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITR.TruckerFlag2 "&_
			 "		     WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITR.TruckerFlag3 "&_
			 "		     WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN ITR.TruckerFlag4 "&_
			 "		     ELSE Null END) "&_
			 "		WHEN 0 THEN '��' "&_
			 "		WHEN 1 THEN 'Yes' "&_
			 "		WHEN 2 THEN 'No' "&_
			 "		ELSE ' ' END as Flag2, "&_
			 "CASE ISNULL(CASE WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN '' "&_
			 "	         WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITR.TruckerFlag4 "&_
			 "	         WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITR.TruckerFlag3 "&_
			 "	         WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITR.TruckerFlag2 "&_
			 "	         ELSE ITR.TruckerFlag1 "&_
			 "      END,'') "&_
			 "      WHEN 0 THEN '��' "&_
			 "      WHEN 1 THEN 'Yes' "&_
			 "      WHEN '' THEN '' "&_
			 "      ELSE 'No' "&_
			 "END Flag1, "&_
			 "CASE  WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN '4' "&_
			 "      WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN '3' "&_
			 "      WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN '2' "&_
			 "      WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN '1' "&_
			 "      ELSE '0'  "&_
			 "END NINE  "&_		
			 ", CYV.kariflag "&_		
             "FROM hITCommonInfo AS ITC INNER JOIN CYVanInfo AS CYV ON ITC.WkNo = CYV.WkNo "&_
             "AND ITC.ContNo=CYV.ContNo "&_
             "INNER JOIN hITReference AS ITR ON ITC.WkContrlNo = ITR.WkContrlNo "&_
             "INNER JOIN mUsers AS mU ON ITC.RegisterCode = mU.UserCode "&_
             "LEFT JOIN ExportCont AS EPC ON CYV.BookNo = EPC.BookNo AND CYV.ContNo = EPC.ContNo "&_
             "LEFT JOIN VslSchedule AS VSLS ON EPC.VoyCtrl = VSLS.VoyCtrl AND EPC.VslCode = VSLS.VslCode "&_
             "LEFT JOIN Booking AS BOK ON EPC.VslCode = BOK.VslCode AND EPC.VoyCtrl = BOK.VoyCtrl AND EPC.BookNo = BOK.BookNo "
             "WHERE WkType='3' AND BOK.Sender like '%" & OpeName & "%' AND Process='R' " &_
             strWrer & ") T " &_
             strOrder
             
	ObjRS.PageSize = 200
	ObjRS.CacheSize = 200
	ObjRS.CursorLocation = 3	
    ObjRS.Open StrSQL, ObjConn
	Num2 = ObjRS.recordcount
	ReDim Preserve DtTbl(Num2)
	ReDim Preserve Kari(Num2)					'2016/10/12 H.Yoshikawa Add
	
	if CInt(Num2) > ObjRS.PageSize then		
		If CInt(Request("pagenum")) = 0 Then
			ObjRS.AbsolutePage = 1
		Else
			If CInt(Request("pagenum")) <= ObjRS.PageCount Then
				ObjRS.AbsolutePage = CInt(Request("pagenum"))				
			Else
				ObjRS.AbsolutePage = 1				
			End If			
		End If		
		abspage = ObjRS.AbsolutePage
		pagecnt = ObjRS.PageCount
	else
		abspage = 1
		pagecnt = 1
	End If	
	
    if err <> 0 then
      DisConnDBH ObjConn, ObjRS 'DB�ؒf
      jampErrerP "2","b401","01","�������F�f�[�^�擾","102","SQL:<BR>"&StrSQL
      Exit Function
    end if
    dim tmptime
    i=1
	RecCtr = 0	
    Do Until ObjRS.EOF
	  if RecCtr <= ObjRS.PageSize - 1 then	
      If DtTbl(i-1)(3)<>Trim(ObjRS("WkNo")) Then
        DtTbl(i)=Array("","","","","","","","","","","","","","","","","","","","","","","","","","")
        DtTbl(i)(0)=Mid(ObjRS("InPutDate"),3,8)
        DtTbl(i)(1)=Mid(ObjRS("WorkDate"),3,8)
        DtTbl(i)(3)=Trim(ObjRS("WkNo"))
        DtTbl(i)(4)=Trim(ObjRS("BookNo"))
        DtTbl(i)(5)=Trim(ObjRS("ContNo"))
        DtTbl(i)(6)=Trim(Mid(ObjRS("WorkCompleteDate"),3,14))
        If Trim(Mid(DtTbl(i)(6),10))<>"" Then
          tmptime=Split(Mid(DtTbl(i)(6),10),":",3,1)
          DtTbl(i)(6)=Left(DtTbl(i)(6),9)&Right(0&tmptime(0),2)&":"&tmptime(1)
        End If
        DtTbl(i)(11)=Trim(ObjRS("ShipLine"))
        DtTbl(i)(12)=Trim(ObjRS("VslName"))
        DtTbl(i)(13)=Trim(ObjRS("ContSize"))
        DtTbl(i)(14)=Trim(ObjRS("ContHeight"))
        DtTbl(i)(15)=Trim(ObjRS("ReceiveFrom"))
        DtTbl(i)(16)=Trim(ObjRS("RecTerminal"))
        DtTbl(i)(17)=Trim(Mid(ObjRS("CYCut"),3,8))
        DtTbl(i)(18)=Trim(ObjRS("Comment1"))      'C-002
        DtTbl(i)(19)=Trim(ObjRS("Comment2"))      'C-002
        DtTbl(i)(20)=Trim(ObjRS("Comment3"))      'C-002
        DtTbl(i)(21)=Trim(ObjRS("WkContrlNo"))    '3th
        DtTbl(i)(22)=Trim(ObjRS("TareWeight"))
        DtTbl(i)(2)=Trim(ObjRS("Code1"))
        DtTbl(i)(7)=Trim(ObjRS("Code2"))
        DtTbl(i)(9)=Trim(ObjRS("NINE"))
        DtTbl(i)(8)=Trim(ObjRS("Flag1"))
        DtTbl(i)(10)=Trim(ObjRS("Flag2"))
        DtTbl(i)(23)=""
        DtTbl(i)(24)=Trim(ObjRS("Name1"))
		if gfTrim(ObjRS("kariflag")) = "1" then
			DtTbl(i)(25) = ""
		else
			DtTbl(i)(25) = "��"
		end if
		Kari(i) = gfTrim(ObjRS("kariflag"))
      End If
        If DelType=0 OR DtTbl(i)(DelTarget(DelType)) = DelStr(DelType) Then
          DtTbl(0)(10) = DtTbl(0)(10) + DtTbl(i)(9)
          i=i+1
        Else
          Num2=Num2-1
        End If
'        DtTbl(0)(10) = DtTbl(0)(10) + DtTbl(i)(9)
'       i=i+1
'3th Add End
	
	  RecCtr = RecCtr + 1 
	  End If	  
      ObjRS.MoveNext
    Loop
    ObjRS.close	
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
<%If Num2<>0 Then%>
    
	//Y.TAKAKUWA Add-S 2015-03-06
	var IEVersion = getInternetExplorerVersion();
	//Y.TAKAKUWA Add-E 2015-03-06
    var obj3=document.getElementById("BDIV");
	//Y.TAKAKUWA Upd-S 2015-03-10
	if(IEVersion < 10)
	{
		if((document.body.offsetWidth-10)  < 1118){
			
			obj3.style.width=document.body.offsetWidth;
			obj3.style.overflowX="auto";
		}
		else{
		    
			obj3.style.width=document.body.offsetWidth-(document.body.offsetWidth-1115);
			obj3.style.overflowX="auto";
		}
		obj3.style.height=document.body.offsetHeight-70;
		obj3.style.overflowY="scroll";
	}
	else
	{
	    var initialHeight = document.documentElement.clientHeight;
		if((document.body.offsetWidth) >= 1250){
			obj3.style.width=document.body.offsetWidth-(document.body.offsetWidth-1120);
			obj3.style.overflowX="auto";
		}
		else
		{
			if((document.body.offsetWidth-10)  < 1100){
				obj3.style.width=document.body.offsetWidth;
				obj3.style.overflowX="auto";
			}
			else{
				obj3.style.width=document.body.offsetWidth-(document.body.offsetWidth-1120);
				obj3.style.overflowX="auto";
			}
		}
		obj3.style.height=initialHeight-70;
		obj3.style.overflowY="scroll";
	}
	//Y.TAKAKUWA Upd-E 2015-03-10
	//Y.TAKAKUWA Add-S 2015-03-06
	var obj3header=document.getElementById("BDIVHEADER");
	if(IEVersion < 10)
	{
	    obj3header.style.width=obj3.clientWidth;//.body.offsetWidth-17;
		obj3header.style.height = 35;
	}
	else
	{
		if((document.body.offsetWidth) >= 1250){
			obj3header.style.width=obj3.clientWidth;//document.body.offsetWidth;
			obj3header.style.height = 35;
		}
		else
		{
			obj3header.style.width=obj3.clientWidth;//document.body.offsetWidth-17;
			obj3header.style.height = 35;
		}
	}
	//Y.TAKAKUWA Add-S 2015-03-06
<% End If %>
}
//�X�V
function GoRenew(sakuNo,bookNo,conNo){
  Fname=document.dmo310F;
  Fname.SakuNo.value=sakuNo;
  Fname.BookNo.value=bookNo;
  Fname.CONnum.value=conNo;
  Fname.action="./dmo320.asp";
  newWin = window.open("", "ReEntry", "status=yes,width=500,height=500,left=10,top=10,resizable=yes,scrollbars=yes");
  Fname.target="ReEntry";
  Fname.submit();
}
//�u�b�L���O���
//function GoBookI(bookNo,sShipLine){//2006/03/06 mod
function GoBookI(bookNo,sShipLine){//2006/03/06 mod
  Fname=document.dmo310F;
  Fname.BookNo.value=bookNo;
  Fname.CONnum.value="";        //CW-021 ADD
  Fname.ShipLine.value=sShipLine;// 2006/03/06 add
  BookInfo(Fname);
}
//�R���e�i�ڍ�
function GoConinf(conNo){
  Fname=document.dmo310F;
  Fname.CONnum.value=conNo;
  Fname.BookNo.value="";        //CW-021 ADD
  BookInfo(Fname);
}
//����
function SerchC(SortFlag,Kye){
  Fname=document.dmo310F;
  Fname.SortFlag.value=SortFlag;
  Fname.SortKye.value=Kye;
  Fname.target="_self";
  Fname.action="./dmo310L.asp";
  Fname.submit();
}
//�Ɖ��
function GoSyokaizumi(){
  target=document.dmo310F;
  if(target.DataNum.value>0){
    flag = confirm('���񓚂̉񓚂��uYes�v�ɂ��܂����H');
    if(flag==true){
      target.SortFlag.value=8;
      len=target.elements.length;
      for(i=0;i<len;i++){
        target.elements[i].disabled=false;
      }
      target.target="_self";
      target.action="./dmo310L.asp";
      target.submit();
    }
  }
}
//CSV           ADD C-001
function GoCSV(){
  //2013/05/09 Add-S Tanaka �f�[�^�Ȃ��͏������Ȃ��悤�ɏC��
  if (document.getElementById("DataNum").value != 0){
  //2013/05/09 Add-E Tanaka
    target=document.dmo310F;
    len=target.elements.length;
    for(i=0;i<len;i++){
      target.elements[i].disabled=false;
    }
    target.target="Bottom";
    //2013/05/09 Upd-S Tanaka 200���ȏ�Ή��ŕ\���������p�����[�^�œn���B
    //target.action="./dmo380.asp";
    target.action="./dmo380.asp?RCnt=" + "<%=RecCtr%>";
    //2013/05/09 Upd-E Tanaka
    target.submit();
    //2013/05/09 Add-S Tanaka
    //�_�E�����[�h��Ƀy�[�W�J�ڂ���ƃ_�E�����[�h��ʂ��J���̂Ō��ɖ߂�
    target.target="_self";
    target.action="./dmo310L.asp";
    //2013/05/09 Add-S Tanaka
  }
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
//Y.TAKAKUWA Add-S 2015-01-22
function showPage(pageNo)
{
   var url = window.location.pathname;
   var filename = url.substring(url.lastIndexOf('/')+1);
   filename = "./" + filename;
   target=document.dmo310F;
   len=target.elements.length;
   for(i=0;i<len;i++){
     target.elements[i].disabled=false;
   }
   document.forms[0].pagenum.value=pageNo;
   target.target="_self";
   target.action="./dmo310L.asp";
   target.submit();  
   return false;
}
//Y.TAKAKUWA Add-E 2015-01-22
//Y.TAKAKUWA Add-S 2015-03-09
function getInternetExplorerVersion()
{
  var rv = -1;
  if (navigator.appName == 'Microsoft Internet Explorer')
  {
    var ua = navigator.userAgent;
    var re  = new RegExp("MSIE ([0-9]{1,}[\.0-9]{0,})");
    if (re.exec(ua) != null)
      rv = parseFloat( RegExp.$1 );
  }
  else if (navigator.appName == 'Netscape')
  {
    var ua = navigator.userAgent;
    var re  = new RegExp("Trident/.*rv:([0-9]{1,}[\.0-9]{0,})");
    if (re.exec(ua) != null)
      rv = parseFloat( RegExp.$1 );
  }
  
  return rv;
}
function cloneTable(tblSource, tblDestination)
{
    <%If Num2<>0 Then%>
	var source = document.getElementById(tblSource);
	var destination = document.getElementById(tblDestination);
	var copy = source.cloneNode(true);
	copy.setAttribute('id', tblDestination);
	//Y.TAKAKUWA Add-S 2015-04-06
	//Change the name of cloned elements
	var rowCount = copy.rows.length;
	for(var i=0; i<rowCount; i++) {
		var row = copy.rows[i];
		element_i = row.getElementsByTagName ('input')[0];
		element_i.removeAttribute('name');
	}
	//Y.TAKAKUWA Add-E 2015-04-06
	destination.parentNode.replaceChild(copy, destination);
	source.style.marginTop = "-35px";
	<%End If%>
}
function onScrollDiv(Scrollablediv,Scrolleddiv) {
    document.getElementById(Scrolleddiv).scrollLeft = Scrollablediv.scrollLeft;
}
//Y.TAKAKUWA Add-E 2015-03-09
// -->
</SCRIPT>
<!--2009/10/02 Add-S Fujiyama-->
<%
'-----------------------------------------
'���b�Z�[�W�{�b�N�X�\���֐�
'mes:�\�����b�Z�[�W(�J���}�ŉ��s���܂��B)
'-----------------------------------------
Public Function ShowMessage(mes)
	dim strMsgWk
	dim strMessage
	dim intRcnt

	strMsgWk=Split(mes, ",")

	For intRcnt=0 To ubound(strMsgWk)
		strMessage= strMessage & strMsgWk(intRcnt) & vbcrlf
	Next

'���b�Z�[�W�{�b�N�X�\��
    ShowMessage = MsgBox(strMessage, vbYesNo + vbQuestion) = vbYes
End Function
%>

<!--2009/10/02 Add-E Fujiyama-->
<style>
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
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="vew();" onResize="vew();">
<!--setTimeout('showContent()', 500); -->
<!-------------�����o���ꗗ���List--------------------------->
<!--<div id="content" style="display:none;"> -->
<%=ErrerM%>
<Form name="dmo310F" method="POST">
<TABLE border="0" cellPadding="2" cellSpacing="0" width="100%">
  <tr>
	<td align="right">
	<%	if Num2 > 0 then
			call gfPutPageSort2(Num2,abspage,pagecnt,"pagenum",SortFlag)
		end if		
		DisConnDBH ObjConn, ObjRS
	%>
	</td>
	<td width="50"></td>	
  </tr>
</TABLE>
<!--Y.TAKAKUWA Add-S 2015-03-05-->	
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
	<td>
		<DIV ID="BDIVHEADER" style="overflow:hidden;">
			<table border="1" cellpadding="0" cellspacing="0" width="100%" Id="maintable1">			
			</table>
		</DIV>
	</td>
</tr>
<tr>
<td>
<!--Y.TAKAKUWA Add-E 2015-03-05-->
<div id="BDIV" onscroll="onScrollDiv(this,'BDIVHEADER')">
<!--Y.TAKAKUWA Upd-S 2015-03-10-->
<!--<TABLE border="1" cellPadding="3" cellSpacing="0" cols="<%=Num+1%>">-->
<TABLE border="1" cellPadding="3" cellSpacing="0" cols="<%=Num+1%>" id="maintable">
<!--Y.TAKAKUWA Upd-E 2015-03-10-->
<% If Num2>0 Then%>
<%   '�G���[�g���b�v�J�n
    on error resume next  %>
  <% If DtTbl(0)(10) = 0 Then %>
  <thead>
  <TR>
    <TH class="hlist" nowrap><%=DtTbl(0)(1)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(2)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(3)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(25)%></TH><!--2016/10/11 Add by Yoshikawa --><TH class="hlist" nowrap><%=DtTbl(0)(4)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(5)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(11)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(12)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(13)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(14)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(22)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(15)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(16)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(17)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(6)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(7)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(8)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(18)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(19)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(20)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(24)%></TH>
    <INPUT type=hidden name='Datatbl0' disabled value='<%=DtTbl(0)(0)%>,<%=DtTbl(0)(1)%>,<%=DtTbl(0)(2)%>,<%=DtTbl(0)(3)%>,<%=DtTbl(0)(4)%>,<%=DtTbl(0)(5)%>,<%=DtTbl(0)(6)%>,<%=DtTbl(0)(7)%>,<%=DtTbl(0)(8)%>,<%=DtTbl(0)(9)%>,<%=DtTbl(0)(10)%>,<%=DtTbl(0)(11)%>,<%=DtTbl(0)(12)%>,<%=DtTbl(0)(13)%>,<%=DtTbl(0)(14)%>,<%=DtTbl(0)(15)%>,<%=DtTbl(0)(16)%>,<%=DtTbl(0)(17)%>,<%=DtTbl(0)(18)%>,<%=DtTbl(0)(19)%>,<%=DtTbl(0)(20)%>,<%=DtTbl(0)(21)%>,<%=DtTbl(0)(22)%>,<%=DtTbl(0)(23)%>,<%=DtTbl(0)(24)%>'>
  </TR>
  </thead>
  <tbody>
    <% For j=1 to RecCtr %>
    <% '2016/08/17 H.Yoshikawa Add Start
         if Kari(j) = "1" then
           bgclr = "bgw"
         else
           bgclr = "bgarrt"
         end if
       '2016/08/17 H.Yoshikawa Add Start
    %>
  <TR class=<%=bgclr%>>
    <TD nowrap><%=DtTbl(j)(1)%><BR></TD><TD nowrap><%=DtTbl(j)(2)%><BR></TD>
    <TD nowrap><A HREF="JavaScript:GoRenew('<%=DtTbl(j)(3)%>','<%=DtTbl(j)(4)%>','<%=DtTbl(j)(5)%>');"><%=DtTbl(j)(3)%></A></TD>
    <TD nowrap><%=DtTbl(j)(25)%><BR></TD><!--2016/10/11 Add by Yoshikawa -->
<%'Mod-s 2006/03/06 h.matsuda%>
<!--    <TD nowrap><A HREF="JavaScript:GoBookI('<%=DtTbl(j)(4)%>');"><%=DtTbl(j)(4)%></A></TD>-->
    <TD nowrap><A HREF="JavaScript:GoBookI('<%=DtTbl(j)(4)%>','<%=DtTbl(j)(11)%>');"><%=DtTbl(j)(4)%></A></TD>
<%'Mod-e 2006/03/06 h.matsuda%>
    <TD nowrap><A HREF="JavaScript:GoConinf('<%=DtTbl(j)(5)%>');"><%=DtTbl(j)(5)%></A></TD>
<!-- C-001    <TD nowrap><%=DtTbl(j)(11)%></TD><TD nowrap><%=DtTbl(j)(12)%></TD><TD nowrap><%=DtTbl(j)(13)%></TD> -->
    <TD nowrap><%=DtTbl(j)(11)%><BR></TD><TD nowrap><%=Left(DtTbl(j)(12),12)%><BR></TD><TD nowrap><%=DtTbl(j)(13)%><BR></TD>
    <TD nowrap><%=DtTbl(j)(14)%><BR></TD><TD nowrap><%=DtTbl(j)(22)%></TD><TD nowrap><%=Left(DtTbl(j)(15),20)%><BR></TD>
    <TD nowrap><%=Left(DtTbl(j)(16),2)%><BR></TD><TD nowrap><%=DtTbl(j)(17)%><BR></TD>
    <TD nowrap><%=DtTbl(j)(6)%><BR></TD><TD nowrap><%=DtTbl(j)(7)%><BR></TD><TD nowrap><%=DtTbl(j)(8)%><BR></TD>
    <TD nowrap><%=Left(DtTbl(j)(18),10)%><BR></TD><TD nowrap><%=Left(DtTbl(j)(19),10)%><BR></TD><TD nowrap><%=Left(DtTbl(j)(20),10)%><BR>
    <TD nowrap><%=DtTbl(j)(24)%><BR></TD>
    <INPUT type=hidden name='Datatbl<%=j%>' disabled value='<%=DtTbl(j)(0)%>,<%=DtTbl(j)(1)%>,<%=DtTbl(j)(2)%>,<%=DtTbl(j)(3)%>,<%=DtTbl(j)(4)%>,<%=DtTbl(j)(5)%>,<%=DtTbl(j)(6)%>,<%=DtTbl(j)(7)%>,<%=DtTbl(j)(8)%>,<%=DtTbl(j)(9)%>,<%=DtTbl(j)(10)%>,<%=DtTbl(j)(11)%>,<%=DtTbl(j)(12)%>,<%=DtTbl(j)(13)%>,<%=DtTbl(j)(14)%>,<%=DtTbl(j)(15)%>,<%=DtTbl(j)(16)%>,<%=DtTbl(j)(17)%>,<%=DtTbl(j)(18)%>,<%=DtTbl(j)(19)%>,<%=DtTbl(j)(20)%>,<%=DtTbl(j)(21)%>,<%=DtTbl(j)(22)%>,<%=DtTbl(j)(23)%>,<%=DtTbl(j)(24)%>'>
  </TD></TR>
    <% Next %>
  </tbody>
  <% Else %>
  <thead>
  <TR class=bga>
    <TH class="hlist" nowrap><%=DtTbl(0)(1)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(2)%></TH><TH class="hlist" nowrap>�w����<BR>�։�</TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(3)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(25)%></TH><!--2016/10/11 Add by Yoshikawa --><TH class="hlist" nowrap><%=DtTbl(0)(4)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(5)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(11)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(12)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(13)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(14)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(22)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(15)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(16)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(17)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(6)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(7)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(8)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(18)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(19)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(20)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(24)%></TH>
    <INPUT type=hidden name='Datatbl0' disabled value='<%=DtTbl(0)(0)%>,<%=DtTbl(0)(1)%>,<%=DtTbl(0)(2)%>,<%=DtTbl(0)(3)%>,<%=DtTbl(0)(4)%>,<%=DtTbl(0)(5)%>,<%=DtTbl(0)(6)%>,<%=DtTbl(0)(7)%>,<%=DtTbl(0)(8)%>,<%=DtTbl(0)(9)%>,<%=DtTbl(0)(10)%>,<%=DtTbl(0)(11)%>,<%=DtTbl(0)(12)%>,<%=DtTbl(0)(13)%>,<%=DtTbl(0)(14)%>,<%=DtTbl(0)(15)%>,<%=DtTbl(0)(16)%>,<%=DtTbl(0)(17)%>,<%=DtTbl(0)(18)%>,<%=DtTbl(0)(19)%>,<%=DtTbl(0)(20)%>,<%=DtTbl(0)(21)%>,<%=DtTbl(0)(22)%>,<%=DtTbl(0)(23)%>,<%=DtTbl(0)(24)%>'>
  </TR>
  </thead>
  <tbody>
    <% For j=1 to RecCtr %>
    <% '2016/08/17 H.Yoshikawa Add Start
         if Kari(j) = "1" then
           bgclr = "bgw"
         else
           bgclr = "bgarrt"
         end if
       '2016/08/17 H.Yoshikawa Add Start
    %>
  <TR class=<%=bgclr%>>
    <TD nowrap><%=DtTbl(j)(1)%><BR></TD><TD nowrap><%=DtTbl(j)(2)%><BR></TD>
    <TD nowrap><%=DtTbl(j)(10)%><BR></TD>
    <TD nowrap><A HREF="JavaScript:GoRenew('<%=DtTbl(j)(3)%>','<%=DtTbl(j)(4)%>','<%=DtTbl(j)(5)%>');"><%=DtTbl(j)(3)%></A></TD>
    <TD nowrap><%=DtTbl(j)(25)%><BR></TD><!--2016/10/11 Add by Yoshikawa -->
<%'Mod-s 2006/03/06 h.matsuda%>
<!--    <TD nowrap><A HREF="JavaScript:GoBookI('<%=DtTbl(j)(4)%>');"><%=DtTbl(j)(4)%></A></TD>-->
    <TD nowrap><A HREF="JavaScript:GoBookI('<%=DtTbl(j)(4)%>','<%=DtTbl(j)(11)%>');"><%=DtTbl(j)(4)%></A></TD>
<%'Mod-e 2006/03/06 h.matsuda%>
    <TD nowrap><A HREF="JavaScript:GoConinf('<%=DtTbl(j)(5)%>');"><%=DtTbl(j)(5)%></A></TD>
<!-- C-001    <TD nowrap><%=DtTbl(j)(11)%></TD><TD nowrap><%=DtTbl(j)(12)%></TD><TD nowrap><%=DtTbl(j)(13)%></TD> -->
    <TD nowrap><%=DtTbl(j)(11)%><BR></TD><TD nowrap><%=Left(DtTbl(j)(12),12)%><BR></TD><TD nowrap><%=DtTbl(j)(13)%><BR></TD>
    <TD nowrap><%=DtTbl(j)(14)%><BR></TD><TD nowrap><%=DtTbl(j)(22)%></TD><TD nowrap><%=Left(DtTbl(j)(15),20)%><BR></TD>
    <TD nowrap><%=Left(DtTbl(j)(16),2)%><BR></TD><TD nowrap><%=DtTbl(j)(17)%><BR></TD>
    <TD nowrap><%=DtTbl(j)(6)%><BR></TD><TD nowrap><%=DtTbl(j)(7)%><BR></TD><TD nowrap><%=DtTbl(j)(8)%><BR></TD>
    <TD nowrap><%=Left(DtTbl(j)(18),10)%><BR></TD><TD nowrap><%=Left(DtTbl(j)(19),10)%><BR></TD><TD nowrap><%=Left(DtTbl(j)(20),10)%><BR>
    <!-- Y.TAKAKUWA Upd-S 2015-02-20 -->
	<!--<TD nowrap><%=DtTbl(j)(24) & abspage & strWrer%><BR></TD>-->
	<TD nowrap><%=DtTbl(j)(24)%><BR></TD>
    <!-- Y.TAKAKUWA Upd-E 2015-02-20 -->
	<INPUT type=hidden name='Datatbl<%=j%>' disabled value='<%=DtTbl(j)(0)%>,<%=DtTbl(j)(1)%>,<%=DtTbl(j)(2)%>,<%=DtTbl(j)(3)%>,<%=DtTbl(j)(4)%>,<%=DtTbl(j)(5)%>,<%=DtTbl(j)(6)%>,<%=DtTbl(j)(7)%>,<%=DtTbl(j)(8)%>,<%=DtTbl(j)(9)%>,<%=DtTbl(j)(10)%>,<%=DtTbl(j)(11)%>,<%=DtTbl(j)(12)%>,<%=DtTbl(j)(13)%>,<%=DtTbl(j)(14)%>,<%=DtTbl(j)(15)%>,<%=DtTbl(j)(16)%>,<%=DtTbl(j)(17)%>,<%=DtTbl(j)(18)%>,<%=DtTbl(j)(19)%>,<%=DtTbl(j)(20)%>,<%=DtTbl(j)(21)%>,<%=DtTbl(j)(22)%>,<%=DtTbl(j)(23)%>,<%=DtTbl(j)(24)%>'>
  </TD></TR>
    <% Next %>
  <% End If %>
</tbody>
<% Else %>
  <TR class=bgw><TD nowrap>��ƈČ��͂���܂���</TD></TR>
<% End If %>
</TABLE>
</div>
</td>
</tr>
<% If Num2>0 Then%>
<tr><td>�@�� �Ԃ��\���͉��o�^��\���܂��B</td></tr>
<% End If %>
</table>
<!--Y.TAKAKUWA Add-S 2015-03-06-->
<SCRIPT Language="JavaScript">
    cloneTable("maintable", "maintable1")
</SCRIPT>
 <!--Y.TAKAKUWA Add-E 2015-03-06-->	
<%'3th del Set_Data Num,DtTbl %>
  <INPUT type=hidden name=DataNum ID="DataNum" value="<%=Num%>">
  <INPUT type=hidden name=SakuNo value="" >
  <INPUT type=hidden name=BookNo value="" >
  <INPUT type=hidden name=CONnum value="" >
  <INPUT type=hidden name="SortFlag" value="<%=SortFlag%>">
  <INPUT type=hidden name=SortKye value="<%=SortKye%>" >
  <INPUT type=hidden name=strWhere value="<%=strWrer%>" disabled>
  <INPUT type=hidden name=absPage value="<%=abspage%>" disabled>
  <INPUT type=hidden name=pagenum value="" >
<%'Mod-s 2006/03/06 h.matsuda%>
	  <INPUT type=hidden name="ShoriMode" value="EMoutInf">
	  <INPUT type=hidden name="ShipLine" value="">
<%'Mod-e 2006/03/06 h.matsuda%>
</Form>
<!--</div> -->
<!-------------��ʏI���--------------------------->
</BODY></HTML>
