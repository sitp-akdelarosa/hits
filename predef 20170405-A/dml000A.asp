<%
'**********************************************
'  �y�v���O�����h�c�z�@: dml000A
'  �y�v���O�������́z�@: ���b�N�I��
'
'  �i�ύX�����j
'   2013-02-14   Y.TAKAKUWA   �쐬
'   2013-07-24   Y.TAKAKUWA   Filter SQL by IC and KA.
'   2013-09-25   Y.TAKAKUWA   ���M���O�̋@�\��ǉ�
'**********************************************
Option Explicit
Response.Expires = 0
'HTTP�R���e���c�^�C�v�ݒ�
Response.ContentType = "text/html; charset=Shift_JIS"
Response.AddHeader "Pragma", "no-cache" 
%>
<%'**********************************************
  '���ʂ̑O�񏈗�
  '���ʊ֐�  (Commonfunc.inc)
%>
<!--#include File="Common.inc"-->
<!--#include File="../download/download.inc"-->
<!--#include file="../ExcelCreator/include/XlsCrt3vbs.inc"-->
<!--#include File="../ExcelCreator/include/report.inc"-->
<%
	'**********************************************

	'�Z�b�V�����̗L�������`�F�b�N
	CheckLoginH
	
	'���[�U�f�[�^����
	dim USER, COMPcd  			
	dim v_GamenMode
	dim v_DataCnt2
		
	'����o���O���̃e�[�u��	
	dim HeaderTbl2
	dim Num2
	dim v_SortKey2	
	dim strOrder2
	dim FieldName2	
	dim chkAns2No()
	dim ObjRS2,ObjConn2
	dim TruckerSubCode
	dim blnSorted2
	
	dim wk
	dim i,x
	dim v_ItemName
	dim v_loop
	dim v_work
	dim file1,gerrmsg 
	dim Arr_SCACCode()
	dim v_SCACCode
	dim v_InOutFlag
	dim strSort
	dim abspage, pagecnt,reccnt	
	
	dim Arr_WrkCtrlNo()
	dim Arr_LOID()
	dim Arr_WkNo()
	dim Arr_ContNo()
	dim Arr_DriverID()
	dim Arr_DriverName()
	dim Arr_LoStatus()
	dim Arr_ImpExp()
	dim Arr_SeqNumber()
	dim Arr_Check()
	dim Arr_DriverInfoText()
	dim Arr_DriverInfoVal()
	dim Arr_RecTerminalCode()
	
	dim Arr_DriverIDInfo()
	dim Arr_DriverHeadIDInfo()
	dim Arr_DriverNameInfo()
	dim Arr_DriverCompanyInfo()
	
	dim Arr_ErrID()
	dim Arr_Err()
	dim Arr_WorkDate()
	
	dim Arr_GroupID
	dim LimitErrMsg
	LimitErrMsg = ""
	
	dim v_GroupInfo
	dim v_DriverInfo
	dim v_driverInfoChkFlg
	
	v_InOutFlag=""

    Redim Arr_DriverInfoText(0)
	Redim Arr_DriverInfoVal(0)
	
	Redim Arr_DriverIDInfo(0)
	Redim Arr_DriverHeadIDInfo(0)
	Redim Arr_DriverNameInfo(0)
	Redim Arr_DriverCompanyInfo(0)
	
	dim v_GroupDrivers
	    
	'INI�t�@�C�����ݒ�l���擾
	dim param(2),calcDate,calcDate1
	getIni param
  	calcDate = DateAdd("d", "-"&param(2), Date)
	calcDate1 = DateAdd("d", "-"&param(1), Date)
		
	ReDim FieldName2(14)	
	FieldName2=Array("WorkDate","Code1","Flag2","WkNo","BLContNo","ShipLine","ShipName","ContSize","ReceiveFrom","CY","DelPermitDate","FreeTime","CYCut","Code2","Flag1")    
    
    Redim Arr_GroupID(8)
    Arr_GroupID=Array("G0","G1","G2","G3","G4","G5","G6","G7","G8","G9")
    
	Const CONST_ASC = "<BR><IMG border=0 src=Image/ascending.gif>"
	Const CONST_DESC = "<BR><IMG border=0 src=Image/descending.gif>"
	const gcPage = 10
	
	blnSorted2 = False
	USER   = UCase(Session.Contents("userid"))
	COMPcd = Session.Contents("COMPcd")  	
	
	'----------------------------------------
    ' �ĕ`��O�̍��ڎ擾
   	'----------------------------------------			
	call LfGetRequestItem
		
	call LfgetSCAC("",Arr_SCACCode)
	
	'Set Parameters Start
	if v_GamenMode = "S" OR  v_GamenMode = "I" OR  v_GamenMode = "R" OR  v_GamenMode = "GI" Then
	  if trim(v_InOutFlag)="" and Trim(Request.form("InOutSrhFlag")) = "" Then
		v_InOutFlag=Request("InOutF")
	  End If
	else
	  if Session("InOutFlag") <> "" then
	    if Trim(Session("InOutFlag")) = "1" then
	      v_InOutFlag = "2"
	    elseif Trim(Session("InOutFlag")) = "2" then
	      v_InOutFlag = "1"
	    else
	      v_InOutFlag = Session("InOutFlag")
	    end if
	  Else
	    v_InOutFlag = ""
	  end if
    end if

	strSort = request.cookies("SortTbl2")	
	
	if strSort <> "" then			
		if Mid(strSort,1,2) = "XX" then
			Session("TB2Key1") = ""
		else			
			Session("TB2Key1") = Mid(strSort,1,2)
		end if
		
		if Mid(strSort,3,1) = "0" then
			Session("TB2KeySort1") = "ASC"
		else
			Session("TB2KeySort1") = "DESC"
		end if	
		
		if Mid(strSort,4,2) = "XX" then
			Session("TB2Key2") = ""
		else
			Session("TB2Key2") = Mid(strSort,4,2)
		end if
		
		if Mid(strSort,6,1) = "0" then
			Session("TB2KeySort2") = "ASC"
		else
			Session("TB2KeySort2") = "DESC"
		end if			
		
		if Mid(strSort,7,2) = "XX" then
			Session("TB2Key3") = ""
		else
			Session("TB2Key3") = Mid(strSort,7,2)
		end if
				
		if Mid(strSort,9,1) = "0" then
			Session("TB2KeySort3") = "ASC"
		else
			Session("TB2KeySort3") = "DESC"
		end if
	end if
	
	'cookies�ɒl�����݂���ꍇ
	if strSort <> "" then
		strOrder2 = getSort2(Session("TB2Key1"),Session("TB2KeySort1"),"")
		strOrder2 = getSort2(Session("TB2Key2"),Session("TB2KeySort2"),strOrder2)
		strOrder2 = getSort2(Session("TB2Key3"),Session("TB2KeySort3"),strOrder2)
	Else
		strOrder2=" ORDER BY WorkDate_Sort ,InputDate  "
	End If
	'Set Parameters End		
	
	
	'�o�^
	if v_GamenMode = "I" then		
		call LfUpdLOInfo()
		'Response.redirect "./dml000A.asp?pagenum=" & CInt(Request("pagenum")) & "&pagenum2=" & CInt(Request("pagenum2"))
	end if
	
	'����
	if v_GamenMode = "R" then		
		call LfReleaseLOInfo()
	end if
	
	'�I�����ēo�^
	if v_GamenMode = "GI" and (v_driverInfoChkFlg = "1" or v_driverInfoChkFlg = "2") then		
		call LfUpdGrpLOInfo()
	end if
	
	'Placed here to update the datagrid
	Call getDataTbl2()
	Call getDriverInfo()
	Call getGroupName()
	
	if v_GamenMode = "S" Then
	  Call SetValue(Num2)
	end if
	
Function LfGetRequestItem()

	v_GamenMode = Request.form("Gamen_Mode")
	v_DataCnt2 = Request.form("DataCnt2")
	'v_DataCnt2 = Num2
	v_DriverInfo = Request.Form("driverInfo")
	v_GroupInfo = Request.Form("groupInfo")
	
	if Request.Form("groupInfoChk2") <> "" and v_GroupInfo <> "" then
	  v_driverInfoChkFlg = Request.Form("groupInfoChk2")
	elseIf Request.Form("groupInfoChk1") <> "" and v_DriverInfo <> "" then
	  v_driverInfoChkFlg = Request.Form("groupInfoChk1")
	end if
	
    v_InOutFlag = Request.form("cmbINOut")
	ReDimension(v_DataCnt2)
	For i = 1 to (v_DataCnt2) - 1 
	    Arr_WrkCtrlNo(i) = Trim(Request.Form("WorkControlNo" & i))
	    Arr_LOID(i) = UCase(Trim(Request.form("InputID" & i)))
	    Arr_WkNo(i) =  Trim(Request.form("WkNo" & i))
	    Arr_ContNo(i) = TRIM(Request.form("ContNo" & i))
	    Arr_DriverID(i) = TRIM(Request.form("LODriverID" & i))
	    Arr_SeqNumber(i) = TRIM(Request.form("SeqNumber" & i))
	    Arr_ImpExp(i) = TRIM(Request.form("ImpExp" & i))
	    Arr_Check(i) = Trim(Request.form("chkInOut" & i))
	    Arr_LoStatus(i) = Trim(Request.form("LoStatus" & i))
	    Arr_WorkDate(i) = Trim(Request.Form("WorkDate" & i))
	    Arr_RecTerminalCode(i) = Trim(Request.Form("RecTerminalCode" & i))
	    'Response.Write "TerminalCode:" & Arr_RecTerminalCode(i) & "<BR/>"
	Next

End Function

Function SetValue(index)
  	ReDimension(index)
	For i = 1 to (index) - 1 
	    Arr_WrkCtrlNo(i) = Trim(Request.Form("WorkControlNo" & i))
	    Arr_LOID(i) = UCase(Trim(Request.form("InputID" & i)))
	    Arr_WkNo(i) =  Trim(Request.form("WkNo" & i))
	    Arr_ContNo(i) = TRIM(Request.form("ContNo" & i))
	    Arr_DriverID(i) = TRIM(Request.form("LODriverID" & i))
	    Arr_SeqNumber(i) = TRIM(Request.form("SeqNumber" & i))
	    Arr_ImpExp(i) = TRIM(Request.form("ImpExp" & i))
	    Arr_Check(i) = Trim(Request.form("chkInOut" & i))
	    Arr_LoStatus(i) = Trim(Request.form("LoStatus" & i))
	    Arr_WorkDate(i) = Trim(Request.Form("WorkDate" & i))
	    Arr_RecTerminalCode(i) = Trim(Request.Form("RecTerminalCode" & i))
	Next
End function

Function ReDimension(index)
   Redim Arr_WrkCtrlNo(index)
   Redim Arr_LOID(index)
   Redim Arr_WkNo(index)
   Redim Arr_ContNo(index)
   Redim Arr_ErrID(index)
   Redim Arr_Err(index)
   Redim Arr_Check(index)
   Redim Arr_DriverID(index)
   Redim Arr_LoStatus(index)
   Redim Arr_ImpExp(index)
   Redim Arr_SeqNumber(index)
   Redim Arr_WorkDate(index)
   Redim Arr_RecTerminalCode(index)
   'response.Write "Data:" & index
End Function

Function getGroupName()
  dim StrSQL
  dim ObjConnLO, ObjRSLO
  dim cnt
  
  ConnDBH ObjConnLO, ObjRSLO	
  WriteLogH "", "", "", ""
  
  StrSQL = " SELECT TGroupID.*,ISNULL(LomGroup.LoDriverCompany,'') AS DriverCompany FROM "
  StrSQL = StrSQL & " ( "
  StrSQL = StrSQL & "   SELECT 'G0' AS GroupID "
  StrSQL = StrSQL & "   UNION "
  StrSQL = StrSQL & "   SELECT 'G1' AS GroupID "
  StrSQL = StrSQL & "   UNION "
  StrSQL = StrSQL & "   SELECT 'G2' AS GroupID "
  StrSQL = StrSQL & "   UNION "
  StrSQL = StrSQL & "   SELECT 'G3' AS GroupID "
  StrSQL = StrSQL & "   UNION "
  StrSQL = StrSQL & "   SELECT 'G4' AS GroupID "
  StrSQL = StrSQL & "   UNION "
  StrSQL = StrSQL & "   SELECT 'G5' AS GroupID "
  StrSQL = StrSQL & "   UNION "
  StrSQL = StrSQL & "   SELECT 'G6' AS GroupID "
  StrSQL = StrSQL & "   UNION "
  StrSQL = StrSQL & "   SELECT 'G7' AS GroupID "
  StrSQL = StrSQL & "   UNION "
  StrSQL = StrSQL & "   SELECT 'G8' AS GroupID "
  StrSQL = StrSQL & "   UNION "
  StrSQL = StrSQL & "   SELECT 'G9' AS GroupID "
  StrSQL = StrSQL & " ) TGroupID "
  StrSQL = StrSQL & " LEFT JOIN LomGroup ON TGroupID.GroupID= LomGroup.LoGroupID "
  StrSQL = StrSQL & " AND HiTSUserID='" & USER & "'"
  cnt=0

  ObjRSLO.Open StrSQL, ObjConnLO    
  If ObjRSLO.recordcount > 0 then
    While NOT ObjRSLO.EOF
      if Trim(Arr_GroupID(cnt)) = Trim(ObjRSLO("GroupID")) AND Trim(ObjRSLO("DriverCompany")) <> "" then
         Arr_GroupID(cnt) = Arr_GroupID(cnt) & "&nbsp;&nbsp;" & Trim(ObjRSLO("DriverCompany"))
      end if
      cnt = cnt + 1
      ObjRSLO.MoveNext
    Wend
  end if
               
  ObjRSLO.Close  
  DisConnDBH ObjConnLO, ObjRSLO	'DB�ؒf
		
End Function

Function getDataTbl2()
	dim StrSQL		
  	dim x
	dim ctr
	dim tmptime
	dim strInOutWhere

	On Error Resume Next

	If Trim(v_InOutFlag)="2" Then
		strInOutWhere =" WHERE Type='FULLOUT' "
		
	ElseIf Trim(v_InOutFlag)="1" Then
		strInOutWhere =" WHERE Type='FULLIN' "
	Else
		strInOutWhere =" "
	End If

	ConnDBH ObjConn2, ObjRS2	
	
	'WriteLogH "", "���b�N�I�����O���ꗗ", "", ""
	'2013-09-25 Y.TAKAKUWA Add-S 
	WriteLogH "b501", "�R���e�i���b�N�Ώۈꗗ�\��", "01",""
	'2013-09-25 Y.TAKAKUWA Add-S
	
	ReDim HeaderTbl2(16)	
	
	HeaderTbl2(0) = "�I��"	
	HeaderTbl2(1) = "ID����"	
	HeaderTbl2(2) = "�h���C�o����<BR/>/�O���[�vID"
	HeaderTbl2(3) = "���b�N��<BR/>�w�b�hID"
	HeaderTbl2(4) = "���o��<BR/>�\���"		
	HeaderTbl2(5) = "���<BR>�ԍ�"
	HeaderTbl2(6) = "�R���e�i�ԍ�<BR/>/BL�ԍ�"
	HeaderTbl2(7) = "�D��"
	HeaderTbl2(8) = "�D��"
	HeaderTbl2(9) = "SZ"
	HeaderTbl2(10) = "������"
	HeaderTbl2(11) = "CY"
	HeaderTbl2(12) = "���o����"
	HeaderTbl2(13) = "�t���[<BR>�^�C��"
	HeaderTbl2(14) = "CY�J�b�g��"							
	HeaderTbl2(15) = "�w����"
	HeaderTbl2(16) = "�w����<BR>��"

	for ctr = 1 to 3	
		Session(CSTR("TB2Key" & ctr))	
		if Session(CSTR("TB2Key" & ctr)) <> "" then
			Select Case Session(CSTR("TB2Key" & ctr))
			
			    'Y.TAKAKUWA Upd-S 2013-02-18
				Case "00" '���o���\���
					'HeaderTbl2(1) = HeaderTbl2(1) & getImage(Session(CSTR("TB2KeySort" & ctr)))
					'blnSorted2 = True
				Case "01" '�w����
					'HeaderTbl2(2) = HeaderTbl2(2) & getImage(Session(CSTR("TB2KeySort" & ctr)))
				Case "02" '�w�����։�
					'HeaderTbl2(3) = HeaderTbl2(3) & getImage(Session(CSTR("TB2KeySort" & ctr)))
					'blnSorted2 = True
				'Y.TAKAKUWA Upd-E 2013-02-18		
				Case "03" '��Ɣԍ�
					HeaderTbl2(5) = HeaderTbl2(5) & getImage(Session(CSTR("TB2KeySort" & ctr)))
					blnSorted2 = True
				Case "04" '�R���e�i�ԍ�/BL�ԍ�
					HeaderTbl2(6) = HeaderTbl2(6) & getImage(Session(CSTR("TB2KeySort" & ctr)))
					blnSorted2 = True
				Case "05" '�D��
					HeaderTbl2(7) = HeaderTbl2(7) & getImage(Session(CSTR("TB2KeySort" & ctr)))
				Case "06" '�D��
					HeaderTbl2(8) = HeaderTbl2(8) & getImage(Session(CSTR("TB2KeySort" & ctr)))
				Case "07" 'SZ
					HeaderTbl2(9) = HeaderTbl2(9) & getImage(Session(CSTR("TB2KeySort" & ctr)))
				Case "08" '������
					HeaderTbl2(10) = HeaderTbl2(10) & getImage(Session(CSTR("TB2KeySort" & ctr)))
				Case "09" 'CY
					HeaderTbl2(11) = HeaderTbl2(11) & getImage(Session(CSTR("TB2KeySort" & ctr)))
				Case "10" '���o��
					HeaderTbl2(12) = HeaderTbl2(12) & getImage(Session(CSTR("TB2KeySort" & ctr)))
				Case "11" '�t���[�^�C��
					HeaderTbl2(13) = HeaderTbl2(13) & getImage(Session(CSTR("TB2KeySort" & ctr)))
					blnSorted2 = True
				Case "12" 'CY�J�b�g��
					HeaderTbl2(14) = HeaderTbl2(14) & getImage(Session(CSTR("TB2KeySort" & ctr)))			
				Case "13" '�w����
					HeaderTbl2(15) = HeaderTbl2(15) & getImage(Session(CSTR("TB2KeySort" & ctr)))			
				Case "14" '�w����<BR>��
					HeaderTbl2(16) = HeaderTbl2(16) & getImage(Session(CSTR("TB2KeySort" & ctr)))			
			End Select
		end if	  
	next					
	
	StrSQL ="SELECT  * FROM ( " 
	StrSQL = StrSQL & "SELECT T.* FROM ( " &_
	    "SELECT DISTINCT 'FULLOUT' As Type,ITC.DeliverTo,ITC.BLNo, "&_
		"ISNULL(CONVERT(varchar(10),ITC.WorkDate,111),'') as WorkDate,"&_
		"CASE WHEN ITC.WorkDate Is NULL THEN '9999/12/31' WHEN ITC.WorkDate ='' THEN '9999/12/31' ELSE ITC.WorkDate END as WorkDate_Sort, "&_
			"(CASE WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN ITC.TruckerSubCode3 "&_
			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubCode2 "&_
			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubCode1 "&_
			"WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN (CASE WHEN mU.UserType = 5 THEN mU.HeadCompanyCode ELSE ITC.RegisterCode END) "&_
			"ELSE (CASE WHEN mU.UserType = 5 THEN mU.HeadCompanyCode ELSE ITC.RegisterCode END) "&_
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
			"ISNULL(CONVERT(varchar(10),INC.FreeTime,111),'') as FreeTime, "&_			
			"CASE WHEN INC.FreeTime Is NULL THEN '9999/12/31' WHEN INC.FreeTime ='' THEN '9999/12/31' ELSE INC.FreeTime END as FreeTime_Sort,"&_			
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
			"INC.CYDelTime, INC.ReturnTime,ITC.InputDate, "&_
			"'' as InputDate1,'' as WorkDate1,'' as BookNo,'' as ContNo,'' as VslName,'' as ContHeight,'' as TareWeight, "&_
			"'' as ReceiveFrom,ISNULL(CONVERT(varchar(10),INC.DelPermitDate,111),'-') as DelPermitDate,"&_
			"CASE WHEN INC.DelPermitDate Is NULL THEN '9999/12/31' WHEN INC.DelPermitDate ='' THEN '9999/12/31' ELSE INC.DelPermitDate END as DelPermitDate_Sort, "&_
			"'' as CYCut,'9999/12/31' as CYCut_Sort,'' as CYCut1,'' as WorkComplete,'' as WorkComplete1, "&_
			"ITC.TruckerSubCode1,ITC.TruckerSubCode2,ITC.TruckerSubCode3,ITC.TruckerSubCode4,ITC.WkContrlNo, "&_
			"'' as Nine , '' as Comment3,'' as RegisterCode "&_
			
			" ,LOInfo.InputID As LOID "&_
			" ,LOInfo.LoStatus As LoStatus "&_
			" ,LOInfo.SeqNumber AS SeqNumber "&_
			
			" ,CASE WHEN LEN(LOInfo.InputID)= 2 THEN "&_
            "  LOInfo.InputID "&_
            "  ELSE "&_
            "  LODriver.LoDriverName "&_
            "  END AS LoDriverName "&_
            
            
            "  ,CASE WHEN LEN(LOInfo.InputID)= 2 THEN "&_ 
            "    CASE WHEN LOInfo.LoStatus <> 2 THEN "&_
            "      CASE WHEN LoOwnGroup.LoGroupID IS NULL THEN '0' "&_
            "      ELSE '2' END "&_
            "    ELSE '2' END "&_
            "  ELSE '1' END AS  "&_
            "  GroupFlag "&_
            
            
            
			" ,LOInfo.LoHeadID AS LoHeadID "&_
			" ,BL.RecTerminalCode AS RecTerminalCode "&_
			" ,'O' AS ImpExp "&_
				
			"FROM hITCommonInfo ITC "&_
		"LEFT JOIN hITReference ITR ON ITC.WkContrlNo = ITR.WkContrlNo "&_
			"INNER JOIN mUsers AS mU ON ITC.RegisterCode = mU.UserCode "&_
			"LEFT JOIN ImportCont AS INC ON ITC.ContNo=INC.ContNo "&_
			"LEFT JOIN mVessel AS mV On INC.VslCode=mV.VslCode "&_
			"LEFT JOIN BL ON INC.VslCode=BL.VslCode AND INC.VoyCtrl=BL.VoyCtrl AND INC.BLNo=BL.BLNo "&_
			"LEFT JOIN Container AS CON ON INC.VslCode = CON.VslCode AND INC.VoyCtrl=CON.VoyCtrl AND INC.ContNo=CON.ContNo "&_
			
	        "LEFT JOIN ("&_
                        "SELECT LoLockOnInfo.WkNo, "&_
                                 "LoLockOnInfo.ContNo, "&_
                                 "LoLockOnInfo.SeqNumber, "&_
                                 "LoLockOnInfo.LoStatus, "&_
                                 "LoLockOnInfo.KaccsSend, "&_
                                 "LoLockOnInfo.WorkDate, "&_
                                 "LoLockOnInfo.HiTSUserID, "&_
                                 "CASE WHEN ISNULL(LoLockOnInfo.LoDriverID,'') <> '' THEN LoLockOnInfo.LoDriverID "&_ 
                                      "ELSE LoLockOnInfo.InputID "&_
                                 "END AS InputID, "&_
                                 "LoLockOnInfo.LoDriverID, "&_
                                 "LoLockOnInfo.LoHeadID "&_      
                        "FROM LoLockOnInfo "&_ 
                        
	        ") AS LOInfo ON ITC.WkNo = LOInfo.WkNo AND ITC.ContNo = LOInfo.ContNo AND LOInfo.LoStatus<>'c' AND LOInfo.LoStatus<>'d' "&_
			"LEFT JOIN LomDriver AS LODriver ON LOInfo.InputID = "&_
			"CASE WHEN(LEN(LOInfo.InputID) = 2) THEN LODriver.LoDriverHeadID "&_
			"     WHEN(LEN(LOInfo.InputID)>=6 AND LEN(LOInfo.InputID)<=12) THEN LODriver.LoDriverID "&_
			"     ELSE LODriver.LoDriverID "&_
			"END "&_
			"LEFT JOIN LomGroup ON LOInfo.InputID = LomGroup.LoGroupID AND LomGroup.HiTSUserID='" & USER & "'" &_
			
			"OUTER APPLY "&_
            "( "&_
            "SELECT TOP 1 * FROM "&_
            "LoGroupeDriver  "&_
            "WHERE LoGroupeDriver.LoGroupID = LOInfo.InputID AND LoGroupeDriver.HiTSUserID='" & USER & "' "&_
            ") LoGroupeDriver "&_
            
            "        OUTER APPLY "&_
            "        ( "&_
            "        SELECT TOP 1 LoGroupeDriver.* FROM LomGroup "&_
            "        INNER JOIN LoGroupeDriver ON  LomGroup.LoGroupID=LoGroupeDriver.LoGroupID AND LomGroup.HiTSUserID = LoGroupeDriver.HiTSUserID "&_
            "        WHERE LomGroup.HiTSUserID='" & USER & "' AND LomGroup.LoGroupID=LOInfo.InputID "&_
            "        ) LoOwnGroup "&_
			
			
			"WHERE ITC.Process='R' AND WkType='1' AND ITC.FullOutType<>'4' AND ITC.FullOutType='1' AND (ITC.RegisterCode='"& USER &"' "&_
			"OR ITC.TruckerSubCode1='"& COMPcd &"' OR ITC.TruckerSubCode2='"& COMPcd &"' OR ITC.TruckerSubCode3='"& COMPcd &"' OR ITC.TruckerSubCode4='"& COMPcd &"')" &_
			"AND ( ITC.WorkCompleteDate IS Null) "&_
			" AND (BL.RecterminalCode = 'KA' OR BL.RecterminalCode = 'IC') " &_
			") AS T " 	
        StrSQL = StrSQL &  "UNION ALL "		
		StrSQL = StrSQL & 	"SELECT T.* FROM (SELECT DISTINCT 'FULLIN' As Type,'' as DeliverTo, '' as BLNo, ISNULL(CONVERT(varchar(10),ITC.WorkDate,111),'') as WorkDate," &_
			"CASE WHEN ITC.WorkDate Is NULL THEN '9999/12/31' WHEN ITC.WorkDate ='' THEN '9999/12/31' ELSE ITC.WorkDate END as WorkDate_Sort, "&_
            "        (CASE " &_
    	    "            WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN (CASE WHEN mU.UserType = 5 THEN mU.HeadCompanyCode ELSE ITC.RegisterCode END) " &_
            "            WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubCode1 " &_
            "            WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubCode2 " &_
            "            WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN ITC.TruckerSubCode3 " &_
            "            ELSE (CASE WHEN mU.UserType = 5 THEN mU.HeadCompanyCode ELSE ITC.RegisterCode END) " &_
		    "        END) as Code1, " &_
		    "		 (CASE WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN ITC.TruckerSubName4 "&_
			"			WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubName3 "&_
			"			WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubName2 "&_
			"			WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITC.TruckerSubName1 "&_
			"			ELSE ITC.TruckerSubName1 "&_
			"		 END) as Name1, "&_
		    "		(CASE "&_
			"		(CASE WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITR.TruckerFlag1 "&_
			"		WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITR.TruckerFlag2 "&_
			"		WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITR.TruckerFlag3 "&_
			"		WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN ITR.TruckerFlag4 "&_
			"		ELSE Null END) "&_
			"		WHEN 0 THEN '��' "&_
			"		WHEN 1 THEN 'Yes' "&_
			"		WHEN 2 THEN 'No' "&_
			"		ELSE ' ' END) as Flag2, "&_			
		    "       ITC.WkNo, '' as FullOutType,ITC.ContNo as BLContNo, CYV.ShipLine,CYV.VslName as ShipName, " &_
			"		CYV.ContSize,BOK.RecTerminal as CY, '' as FreeTime,'9999/12/31' as FreeTime2, '' as DeliverTo1, " &_			
		    "       '' as WorkCompleteDate,'' as ReturnDateStr,'' as ReturnValue, " &_
		    "        (CASE " &_
			"            WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN NULL " &_
			"            WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubCode4 " &_
			"            WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubCode3 " &_
		    "            WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITC.TruckerSubCode2 " &_
		    "            ELSE ITC.TruckerSubCode1 " &_
		    "        END) as Code2, " &_
		    "        (CASE WHEN " &_
		    "              (CASE WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITC.TruckerSubCode2 " &_
		    "                    WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubCode3 " &_
		    "                    WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubCode4 " &_
		    "                    WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN NULL " &_
		    "                    ELSE ITC.TruckerSubCode1 " &_
		    "               END) IS NULL THEN ' '" &_
		    "              ELSE " &_
		    "                    (CASE (CASE WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITR.TruckerFlag2 " &_
		    "                                WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITR.TruckerFlag3 " &_
		    "                                WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITR.TruckerFlag4 " &_
		    "                                WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN NULL " &_
		    "                                ELSE ITR.TruckerFlag1 " &_
		    "                     END) " &_
		    "             WHEN '0' THEN '��' " &_
		    "             WHEN '1' THEN 'Yes' " &_
		    "             ELSE 'No' END) " &_
		    "          END) as Flag1, " &_
			"        SUBSTRING(ITC.Comment1,1,10) as Comment1, SUBSTRING(ITC.Comment2,1,10) as Comment2, " &_
			" 		'' as ReturnDateVal, ITC.UpdtUserCode,ITR.TruckerFlag1, ITR.TruckerFlag2, ITR.TruckerFlag3, " &_
			"		ITR.TruckerFlag4, mU.HeadCompanyCode, mU.UserType,'' as CYDelTime,'' as ReturnTime," &_
			"		CONVERT(varchar,ITC.InputDate,11) as InputDate,ITC.InputDate as InputDate1, " &_ 
			"		ITC.WorkDate as WorkDate1,CYV.BookNo,ITC.ContNo,CYV.VslName as VslName, " &_
			" 		CYV.ContHeight,CASE ISNULL(CYV.TareWeight,0) WHEN 0 THEN '-' ELSE CYV.TareWeight END TareWeight, " &_
			"		SUBSTRING(CYV.ReceiveFrom,1,20) as ReceiveFrom, '-' as DelPermitDate,'9999/12/31' as DelPermitDate_Sort," &_
			"		ISNULL(CONVERT(varchar(10),VSLS.CYCut,111),'') as CYCut,"&_
			"		CASE WHEN VSLS.CYCut Is NULL THEN '9999/12/31' WHEN VSLS.CYCut ='' THEN '9999/12/31' ELSE VSLS.CYCut END as CYCut_Sort, "&_
			" 		VSLS.CYCut as CYCut1, " &_			
			"		CONVERT(varchar,ITC.WorkCompleteDate,11) + ' ' + Substring(CONVERT(varchar,ITC.WorkCompleteDate,8),1,5) as WorkComplete, " &_
			"		ITC.WorkCompleteDate as WorkComplete1, " &_
			"		ITC.TruckerSubCode1, ITC.TruckerSubCode2, ITC.TruckerSubCode3, ITC.TruckerSubCode4,  ITC.WkContrlNo, " & _
			"          (CASE "&_
			"                WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN '1' "&_
			"                WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN '2' "&_
			"                WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN '3' "&_
			"                WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN '4' "&_
			"                ELSE '0' END) as Nine, "&_		
			"        SUBSTRING(ITC.Comment3,1,10) as Comment3, " &_
		    "        ITC.RegisterCode " &_ 
		    
		    "        ,LOInfo.InputID As LOID "&_
		    "        ,LOInfo.LoStatus As LoStatus "&_
		    "        ,LOInfo.SeqNumber AS SeqNumber "&_
		    
		    "        ,CASE WHEN LEN(LOInfo.InputID)= 2 THEN "&_
            "         LOInfo.InputID "&_
            "         ELSE "&_
            "         LODriver.LoDriverName "&_
            "         END AS LoDriverName "&_
		    
            "  ,CASE WHEN LEN(LOInfo.InputID)= 2 THEN "&_ 
            "    CASE WHEN LOInfo.LoStatus <> 2 THEN "&_
            "      CASE WHEN LoOwnGroup.LoGroupID IS NULL THEN '0' "&_
            "      ELSE '2' END "&_
            "    ELSE '2' END "&_
            "  ELSE '1' END AS  "&_
            "  GroupFlag "&_
		    
		    
		    
		    "        ,LOInfo.LoHeadID AS LoHeadID "&_
		    "        ,BOK.RecTerminalCode AS RecTerminalCode "&_
		    "        ,'I' AS ImpExp "&_ 
		    
    	    "        FROM hITCommonInfo AS ITC " &_
		    "        INNER JOIN CYVanInfo AS CYV ON ITC.ContNo=CYV.ContNo AND ITC.WkNo = CYV.WkNo " &_
		    "        INNER JOIN hITReference AS ITR ON ITC.WkContrlNo = ITR.WkContrlNo " &_
		    "        INNER JOIN mUsers AS mU ON ITC.RegisterCode = mU.UserCode " &_
		    "        LEFT JOIN ExportCont AS EPC ON CYV.ContNo = EPC.ContNo AND CYV.BookNo = EPC.BookNo and CYV.VslCode = EPC.VslCode " &_
		    "        LEFT JOIN VslSchedule AS VSLS ON EPC.VslCode = VSLS.VslCode AND EPC.VoyCtrl = VSLS.VoyCtrl " &_
		    "        LEFT JOIN Booking AS BOK ON EPC.VslCode = BOK.VslCode AND EPC.VoyCtrl = BOK.VoyCtrl AND EPC.BookNo = BOK.BookNo " &_
		    
		    "        LEFT JOIN ("&_
		                         "SELECT LoLockOnInfo.WkNo, "&_
                                        "LoLockOnInfo.ContNo, "&_
                                        "LoLockOnInfo.SeqNumber, "&_
                                        "LoLockOnInfo.LoStatus, "&_
                                        "LoLockOnInfo.KaccsSend, "&_
                                        "LoLockOnInfo.WorkDate, "&_
                                        "LoLockOnInfo.HiTSUserID, "&_
                                        "CASE WHEN ISNULL(LoLockOnInfo.LoDriverID,'') <> '' THEN LoLockOnInfo.LoDriverID "&_ 
                                             "ELSE LoLockOnInfo.InputID "&_
                                        "END AS InputID, "&_
                                        "LoLockOnInfo.LoDriverID, "&_
                                        "LoLockOnInfo.LoHeadID "&_      
                                 "FROM LoLockOnInfo "&_ 
		    ") AS LOInfo ON ITC.WkNo = LOInfo.WkNo AND ITC.ContNo = LOInfo.ContNo AND LOInfo.LoStatus<>'c' AND LOInfo.LoStatus<>'d' "&_
		    
		    
		    
		    
		    "        LEFT JOIN LomDriver AS LODriver ON LOInfo.InputID = "&_
			"        CASE WHEN(LEN(LOInfo.InputID) = 2) THEN LODriver.LoDriverHeadID "&_
			"             WHEN(LEN(LOInfo.InputID)>=6 AND LEN(LOInfo.InputID)<=12) THEN LODriver.LoDriverID "&_
			"             ELSE LODriver.LoDriverID "&_
			"        END "&_
			"        LEFT JOIN LomGroup ON LOInfo.InputID = LomGroup.LoGroupID AND LomGroup.HiTSUserID='" & USER & "'" &_
			
			"        OUTER APPLY "&_
            "        ( "&_
            "        SELECT TOP 1 * FROM "&_
            "        LoGroupeDriver  "&_
            "        WHERE LoGroupeDriver.LoGroupID = LOInfo.InputID AND LoGroupeDriver.HiTSUserID='" & USER & "' "&_
            "        ) LoGroupeDriver "&_
			
			"        OUTER APPLY "&_
            "        ( "&_
            "        SELECT TOP 1 LoGroupeDriver.* FROM LomGroup "&_
            "        INNER JOIN LoGroupeDriver ON  LomGroup.LoGroupID=LoGroupeDriver.LoGroupID AND LomGroup.HiTSUserID = LoGroupeDriver.HiTSUserID "&_
            "        WHERE LomGroup.HiTSUserID='" & USER & "' AND LomGroup.LoGroupID=LOInfo.InputID" &_
            "        ) LoOwnGroup "&_
			
			
		    "        WHERE ITC.Process='R' AND ITC.WkType='3' AND (ITC.RegisterCode='"& USER &"' " &_
		    "        OR ITC.TruckerSubCode1='"& COMPcd &"' OR ITC.TruckerSubCode2='"& COMPcd &"' " &_
		    "        OR ITC.TruckerSubCode3='"& COMPcd &"' OR ITC.TruckerSubCode4='"& COMPcd &"')" &_
		    "        AND (ITC.WorkCompleteDate IS Null) "&_
		    "        AND (BOK.RecterminalCode = 'KA' OR BOK.RecterminalCode = 'IC') "&_
		    "        ) AS T "
			StrSQL = StrSQL & " ) AS A " &_ 
			
    		strInOutWhere & strOrder2
    'Y.TAKAKUWA 2013-02-19 Add-S
    'Response.Write strSQL
    'Response.end
    'Y.TAKAKUWA 2013-02-19 Add-E
    
	ObjRS2.PageSize = 100
	ObjRS2.CacheSize = 100
	ObjRS2.CursorLocation = 3
	ObjRS2.Open StrSQL, ObjConn2
	
	Num2 = ObjRS2.recordcount	

	if Num2 > 100 then
		If CInt(Request("pagenum2")) = 0 Then
			ObjRS2.AbsolutePage = 1
		Else
			If CInt(Request("pagenum2")) <= ObjRS2.PageCount Then
				ObjRS2.AbsolutePage = CInt(Request("pagenum2"))
			Else
				ObjRS2.AbsolutePage = 1
			End If
		End If		 
	end if
	
	if err <> 0 then
		DisConnDBH ObjConn2, ObjRS2	'DB�ؒf
		jampErrerP "2","b301","01","���b�N�I�����O���","102","SQL:<BR>" & StrSQL & err.description & Err.number
		Exit Function
	end if			
	'�G���[�g���b�v����
    on error goto 0	
    
End Function


Function LfUpdLOInfo()
  dim StrSQL
  dim ObjConnLO, ObjRSLO
  dim ErrFlg
  dim iSeq
  dim Arr_GroupDrivers
  dim icnt
  
  ConnDBH ObjConnLO, ObjRSLO	
  WriteLogH "", "", "", ""
	
  ErrFlg = false
	
  For i = 1 to v_DataCnt2-1
    If Trim(Arr_LOID(i)) <> "" Then
        
      If Len(Trim(Arr_LOID(i))) = 2 then
        If CheckGroupID(Trim(Arr_LOID(i))) = true then
          If CheckGroup(UCase(Arr_LOID(i))) = "" then
            Arr_ErrID(i) = UCase(Arr_LOID(i))
            Arr_Err(i) = "�o�^��"
            ErrFlg = false
          else
            ErrFlg = false
          End If
        else
          If CheckDriver(Arr_LOID(i)) = "" then
            If Trim(LimitErrMsg) = "" Then
              Arr_ErrID(i) = "ID�s��"
            Else
              Arr_ErrID(i) = LimitErrMsg
            End If
            Arr_Err(i) = "�o�^�Ȃ�"
            ErrFlg = true
          else
            ErrFlg = false
          End If 
        end if
      else
        If CheckDriver(Arr_LOID(i)) = "" then
          If Trim(LimitErrMsg) = "" Then
            Arr_ErrID(i) = "ID�s��"
          Else
            Arr_ErrID(i) = LimitErrMsg
          End If
          Arr_Err(i) = "�o�^�Ȃ�"
          ErrFlg = true
        else
          ErrFlg = false
        End If 
      end if  
      
      if ErrFlg = false then
        
        If Trim(Arr_DriverID(i)) = "" then
          
          'QUERY VALUES FOR INSERTION
          StrSQL = "SELECT * FROM LoLockOnInfo WHERE WkNo ='" & Arr_WkNo(i)  & "'"&_
                                               " AND ContNo='" & Arr_ContNo(i) & "'"                
          ObjRSLO.Open StrSQL, ObjConnLO
      
          If ObjRSLO.recordcount > 0 then
            While NOT ObjRSLO.EOF
              iSeq = CInt(Trim(ObjRSLO("SeqNumber")))
              ObjRSLO.MoveNext
            Wend
		    iSeq = iSeq + 1
          else
            iSeq = 1
          end if
      
          StrSQL = " INSERT INTO LoLockOnInfo (WkNo, ContNo, SeqNumber, UpdtTime, UpdtPgCd, UpdtTmnl, LoStatus, LoNearRange, KaccsSend, InputID, LoWorkKind, HiTSUserID, LoReqTime, TerminalCode, WorkDate)"
          StrSQL = StrSQL & " VALUES ( "
          StrSQL = StrSQL & "'" & Trim(Arr_WkNo(i)) & "',"                 'WkNo
          StrSQL = StrSQL & "'" & Trim(Arr_ContNo(i)) & "',"               'ContNo
          StrSQL = StrSQL & "'" & iSeq & "',"                              'Seq
          StrSQL = StrSQL & "'" & Now() & "',"                             'UpdtTime
          StrSQL = StrSQL & "'" & "PREDEF01" & "',"                        'UpdtPgCd
          StrSQL = StrSQL & "'" & USER & "',"                              'UpdtTmnl
          StrSQL = StrSQL & "'" & "1" & "',"                               'LoStatus
          StrSQL = StrSQL & "" & "NULL" & ","                              'LoNearRange
          StrSQL = StrSQL & "'" & "0" & "',"                               'KaccsSend
          'StrSQL = StrSQL & "'" & "0" & "',"                              'SphoneSend
          
          StrSQL = StrSQL & "'" & Trim(Arr_LOID(i)) & "', "         'InputID
          
          If Len(Trim(Arr_LOID(i))) = 2 then
            'StrSQL = StrSQL & "'', "                                      'DriverID
            'StrSQL = StrSQL & "'', "                                      'HeadID
          else
            'StrSQL = StrSQL & "'"', "              'DriverID
            'StrSQL = StrSQL & "'" & Trim(CheckDriver(Arr_LOID(i))) & "', "'HeadID
          end if
          StrSQL = StrSQL & "'" & Trim(Arr_ImpExp(i)) & "', "              'LoWorkKind
          StrSQL = StrSQL & "'" & USER & "', "                             'HiTSUserID
          StrSQL = StrSQL & "'" & Now() & "', "                            'LoReqTime
          StrSQL = StrSQL & "'" & Trim(Arr_RecTerminalCode(i)) & "', "     'TerminalCode 
          If Trim(Arr_WorkDate(i)) <> "" Then
            StrSQL = StrSQL & "'" & Trim(Arr_WorkDate(i)) & "'"            'WorkDate
          else
            StrSQL = StrSQL & "NULL" 
          end if
          StrSQL = StrSQL & ")"
          
          'response.Write StrSQL
          
          ObjConnLO.Execute(StrSQL)
          if err <> 0 then
			Set ObjRSLO = Nothing				
			jampErrerPDB ObjConnLO,"2","b107","01","","104","SQL:<BR>"& strSQL
			ObjRSLO.Close
			DisConnDBH ObjConnLO, ObjRSLO	'DB�ؒf
			Exit Function
		  end if
		  
		  
		  If Len(Trim(Arr_LOID(i))) = 2 Then
		    Arr_GroupDrivers = Split(v_GroupDrivers,",")
		    For icnt = 0 To UBound(Arr_GroupDrivers)
		      If Trim(Arr_GroupDrivers(icnt)) <> "" Then
		        'StrSQL = " UPDATE LomDriver SET "
                'StrSQL = StrSQL & "SpAlarmOrder='1', "                                          'SpAlarmOrder
                'StrSQL = StrSQL & "SpAlarmCancel='' "                                           'SpAlarmCancel
                'StrSQL = StrSQL & "WHERE LoDriverID ='" & Trim(Arr_GroupDrivers(icnt))  & "'"       
                'ObjConnLO.Execute(StrSQL)
	            'If err <> 0 Then
	            '  Set ObjRSLO = Nothing				
	            '  jampErrerPDB ObjConnLO,"2","b107","01","","104","SQL:<BR>"& strSQL
	            '  Exit For
	            'End If
	            fUpdateDriverAlarm(Arr_GroupDrivers(icnt))
	          End If
		    Next
		  Else
		    'StrSQL = " UPDATE LomDriver SET "
            'StrSQL = StrSQL & "SpAlarmOrder='1', "                                          'SpAlarmOrder
            'StrSQL = StrSQL & "SpAlarmCancel='' "                                           'SpAlarmCancel
            'StrSQL = StrSQL & "WHERE LoDriverID ='" & Trim(Arr_LOID(i))  & "'"       
            'ObjConnLO.Execute(StrSQL)
	        'If err <> 0 Then
	        '  Set ObjRSLO = Nothing				
	        '  jampErrerPDB ObjConnLO,"2","b107","01","","104","SQL:<BR>"& strSQL
	        'End If
	        fUpdateDriverAlarm(Arr_LOID(i))
		  End If
		  
		  ObjRSLO.Close
		elseif Trim(Arr_DriverID(i)) <> "" then 
		  StrSQL = "SELECT * FROM LoLockOnInfo WHERE WkNo ='" & Arr_WkNo(i)  & "'"&_
                                               " AND ContNo='" & Arr_ContNo(i) & "'"&_
                                               " AND SeqNumber='" & Arr_SeqNumber(i) & "'"
          ObjRSLO.Open StrSQL, ObjConnLO
          If ObjRSLO.recordcount > 0 then
            if Trim(ObjRSLO("LoStatus")) <> "2" then
              call LfUpdLODriverID(Arr_WkNo(i), Arr_ContNo(i), Arr_SeqNumber(i), Arr_LOID(i), v_GroupDrivers)
            end if
          end if
		  ObjRSLO.close
		   
		end if
      end if
    end if
  Next
    
  DisConnDBH ObjConnLO, ObjRSLO	'DB�ؒf
                                                           
End Function

Function CheckGroupID(GroupID)
  Dim cnt
  CheckGroupID = false
  For x = 0 to UBound(Arr_GroupID)
    if Trim(GroupID) = Arr_GroupID(x) then
      CheckGroupID = true
      Exit For
    else
      CheckGroupID = false
    end if
  Next
End Function

Function LfReleaseLOInfo()
    dim StrSQL
    dim ObjConnLO, ObjRSLO
    dim ErrFlg
    dim iSeq
	
    ConnDBH ObjConnLO, ObjRSLO	
	WriteLogH "", "", "", ""
	
	For i = 1 to v_DataCnt2-1
      If UCase(Trim(Arr_Check(i))) = "ON" Then
       If Trim(Arr_DriverID(i)) <> "" then
        'QUERY VALUES FOR UPDATE
        StrSQL = "SELECT * FROM LoLockOnInfo WHERE WkNo ='" & Arr_WkNo(i)  & "'"&_
                                             " AND ContNo='" & Arr_ContNo(i) & "'"&_
                                             " AND SeqNumber='" & Arr_SeqNumber(i) & "'"                
        ObjRSLO.Open StrSQL, ObjConnLO
        If ObjRSLO.recordcount > 0 then
          
            StrSQL = " UPDATE LoLockOnInfo SET "
            StrSQL = StrSQL & "WkNo='" & Trim(Arr_WkNo(i)) & "',"        'WkNo
            StrSQL = StrSQL & "ContNo='" & Trim(Arr_ContNo(i)) & "',"    'ContNo
            StrSQL = StrSQL & "UpdtTime='" & Now() & "',"                'UpdtTime
            StrSQL = StrSQL & "UpdtPgCd='" & "PREDEF01" & "',"           'UpdtPgCd
            StrSQL = StrSQL & "UpdtTmnl='" & USER & "',"                 'UpdtTmnl
            StrSQL = StrSQL & "LoStatus='" & "c" & "',"                  'LoStatus
            StrSQL = StrSQL & "LoEndReason='" & "c"  & "',"              'LoEndReason
            StrSQL = StrSQL & "KaccsSend='" & "0"  & "',"                'KaccsSend
            StrSQL = StrSQL & "LoEndTime='" & Now() & "' "               'LoEndTime
            
            StrSQL = StrSQL & "WHERE WkNo ='" & Trim(Arr_WkNo(i))  & "'"&_         
                             " AND ContNo='" & Trim(Arr_ContNo(i)) & "'"&_
                             " AND SeqNumber='" & Arr_SeqNumber(i) & "'" 
            ObjConnLO.Execute(StrSQL)
            if err <> 0 then
			  Set ObjRSLO = Nothing				
			  jampErrerPDB ObjConnLO,"2","b107","01","","104","SQL:<BR>"& strSQL
			  ObjRSLO.Close
			  DisConnDBH ObjConnLO, ObjRSLO	'DB�ؒf
			  Exit Function
		    end if
		    
		    If Trim(ObjRSLO("LoStatus")) = "2" then
		      StrSQL = " UPDATE LomDriver SET "
              If Trim(Arr_ImpExp(i)) = "O" Then
                StrSQL = StrSQL & "SpAlarmOrder='', "                   'SpAlarmOrder
                StrSQL = StrSQL & "SpAlarmCancel='2' "                  'SpAlarmCancel
              Else
                StrSQL = StrSQL & "SpAlarmOrder='', "                   'SpAlarmOrder
                StrSQL = StrSQL & "SpAlarmCancel='1' "                  'SpAlarmCancel
              End If
              StrSQL = StrSQL & "WHERE LoDriverID ='" & Trim(Arr_DriverID(i))  & "'"       
              ObjConnLO.Execute(StrSQL)
		      If err <> 0 Then
			    Set ObjRSLO = Nothing				
			    jampErrerPDB ObjConnLO,"2","b107","01","","104","SQL:<BR>"& strSQL
			    ObjRSLO.Close
			    DisConnDBH ObjConnLO, ObjRSLO	'DB�ؒf
			    Exit Function
		      End If
            End If
		    
		  'Resending After Lock On Release  
		  if Trim(ObjRSLO("LoStatus")) = "2" then  
		    StrSQL = " UPDATE hITCommonInfo SET "
		    StrSQL = StrSQL & "UpdtTime='" & Now() & "',"                           'UpdtTime
            StrSQL = StrSQL & "UpdtPgCd='" & "PREDEF01" & "',"                      'UpdtPgCd
            StrSQL = StrSQL & "UpdtTmnl='" & USER & "',"                            'UpdtTmnl
            StrSQL = StrSQL & "HeadID=NULL, "                                       'HeadID
            StrSQL = StrSQL & "Status='0' "                                         'Status
            StrSQL = StrSQL & "WHERE WkContrlNo ='" & Trim(Arr_WrkCtrlNo(i))  & "'"          
		  end if
		  
		  ObjConnLO.Execute(StrSQL)
          if err <> 0 then
		    Set ObjRSLO = Nothing				
		    jampErrerPDB ObjConnLO,"2","b107","01","","104","SQL:<BR>"& strSQL
		  end if
		  
	    end if
	    ObjRSLO.Close
	   end if
      end if
      
    Next
    
    DisConnDBH ObjConnLO, ObjRSLO	'DB�ؒf
    
End function


Function LfUpdLODriverID(StrWkNo, StrContNo, StrSeq, StrDriverID, StrGroupDrivers)
    dim StrSQL
    dim ObjConnLO, ObjRSLO
    dim ErrFlg
    dim iSeq
    dim Arr_GroupDrivers
    dim icnt
	
    ConnDBH ObjConnLO, ObjRSLO	
	WriteLogH "", "", "", ""
    
    if Trim(StrWkNo) <> "" and Trim(StrContNo) <> "" then    
    StrSQL = "SELECT * FROM LoLockOnInfo WHERE WkNo ='" & StrWkNo  & "'"&_
                                             " AND ContNo='" & StrContNo & "'"                
    ObjRSLO.Open StrSQL, ObjConnLO
    If ObjRSLO.recordcount > 0 then
      StrSQL = " UPDATE LoLockOnInfo SET "
      StrSQL = StrSQL & "InputID='" & Trim(UCase(StrDriverID)) & "', "                'InputID
      
      If Len(Trim(StrDriverID)) = 2 then
        StrSQL = StrSQL & "LODriverID='', "              'DriverID
        'StrSQL = StrSQL & "LOHeadID='', "   'HeadID
      else
        StrSQL = StrSQL & "LODriverID='', "              'DriverID
        'StrSQL = StrSQL & "LOHeadID='" & Trim(CheckDriver(StrDriverID)) & "', "      'HeadID      
      end if
      
      StrSQL = StrSQL & "UpdtTime='" & Now() & "',"                                   'UpdtTime
      StrSQL = StrSQL & "UpdtPgCd='" & "PREDEF01" & "',"                              'UpdtPgCd
      StrSQL = StrSQL & "UpdtTmnl='" & USER & "', "                                   'UpdtTmnl
      StrSQL = StrSQL & "LoReqTime='" & Now() & "',"                                  'LoReqTime
      StrSQL = StrSQL & "KaccsSend='" & "0"  & "',"                                   'KaccsSend
      If Trim(Arr_WorkDate(i)) <> "" then
        StrSQL = StrSQL & "WorkDate='" & Trim(Arr_WorkDate(i)) & "' "                 'WorkDate
      else
        StrSQL = StrSQL & "WorkDate=NULL "
      end if 
      
      
      StrSQL = StrSQL & "WHERE WkNo ='" & Trim(StrWkNo)  & "'"&_         
                        " AND ContNo='" & Trim(StrContNo) & "'"&_ 
                        " AND SeqNumber='" & Trim(StrSeq) & "'"
      ObjConnLO.Execute(StrSQL)
      
      if err <> 0 then
		Set ObjRSLO = Nothing				
		jampErrerPDB ObjConnLO,"2","b107","01","","104","SQL:<BR>"& strSQL
		ObjRSLO.Close
	    DisConnDBH ObjConnLO, ObjRSLO	'DB�ؒf
	    Exit Function
	  end if
	  
	  If LEN(Trim(StrDriverID)) = 2 Then
	    If Trim(StrGroupDrivers) <> "" Then
	      Arr_GroupDrivers = Split(StrGroupDrivers,",")
		  For icnt = 0 To UBound(Arr_GroupDrivers)
		    If Trim(Arr_GroupDrivers(icnt)) <> "" Then
		      StrSQL = " UPDATE LomDriver SET "
              StrSQL = StrSQL & "SpAlarmOrder='1', "                                          'SpAlarmOrder
              StrSQL = StrSQL & "SpAlarmCancel='' "                                           'SpAlarmCancel
              StrSQL = StrSQL & "WHERE LoDriverID ='" & Trim(Arr_GroupDrivers(icnt))  & "'"       
              ObjConnLO.Execute(StrSQL)
	          If err <> 0 Then
	            Set ObjRSLO = Nothing				
	            jampErrerPDB ObjConnLO,"2","b107","01","","104","SQL:<BR>"& strSQL
	            Exit For
	          End If
	        End If
		  Next
		End If
	  Else
	    StrSQL = " UPDATE LomDriver SET "
        StrSQL = StrSQL & "SpAlarmOrder='1', "                                          'SpAlarmOrder
        StrSQL = StrSQL & "SpAlarmCancel='' "                                           'SpAlarmCancel
        StrSQL = StrSQL & "WHERE LoDriverID ='" & Trim(StrDriverID)  & "'"       
        ObjConnLO.Execute(StrSQL)
	    If err <> 0 Then
	      Set ObjRSLO = Nothing				
	      jampErrerPDB ObjConnLO,"2","b107","01","","104","SQL:<BR>"& strSQL
	    End If	
	  End If
	  
	 end if
	 
	 ObjRSLO.Close  
    end if
    DisConnDBH ObjConnLO, ObjRSLO	'DB�ؒf
    
End function

Function CheckDriver(strDriverID)
    dim StrSQL
    dim ObjConnLO, ObjRSLO
    dim Arr_OwnGroup()
    dim cntRec
    Redim Arr_OwnGroup(0)
    cntRec = 0
    CheckDriver = ""
    ConnDBH ObjConnLO, ObjRSLO
    
    If Trim(strDriverID) <> "" Then
      StrSQL = "SELECT * FROM LomDriver WHERE "
      StrSQL = StrSQL + " LomDriver.LoDriverID ='" & Trim(strDriverID)  & "'"
      StrSQL = StrSQL + " AND LomDriver.HiTSUserID ='" & USER  & "'"
      'Response.Write StrSQL
      ObjRSLO.Open StrSQL, ObjConnLO
      
      If ObjRSLO.recordcount > 0 then
        If Trim(ObjRSLO("AcceptStatus")) = "3" then
          CheckDriver = ""
          LimitErrMsg = "������"
        Else
          CheckDriver = Trim(ObjRSLO("LoDriverHeadID"))
          ObjRSLO.Close
          DisConnDBH ObjConnLO, ObjRSLO	'DB�ؒf
          Exit Function
        End If
      else
        CheckDriver = ""
      end if
      ObjRSLO.Close  
      
      StrSQL = " SELECT TOP 1 * FROM LomDriver "
      StrSQL = StrSQL & " LEFT JOIN LoGroupeDriver ON LomDriver.LoDriverID = LoGroupeDriver.LoDriverID AND LoGroupeDriver.HiTSUserID = '" & USER & "'"
      StrSQL = StrSQL & " RIGHT JOIN "
      StrSQL = StrSQL & " ( "
      StrSQL = StrSQL & " SELECT DISTINCT LomGroup.LoGroupID "
      StrSQL = StrSQL & " FROM LomDriver "
      'StrSQL = StrSQL & " INNER JOIN LoGroupeDriver ON LomDriver.LoDriverID = LoGroupeDriver.LoDriverID "
      StrSQL = StrSQL & " INNER JOIN LomGroup ON LomDriver.HiTSUserID = LomGroup.HiTSUserID "
      StrSQL = StrSQL & " WHERE LomDriver.HiTSUserID = '" & USER & "' "
      StrSQL = StrSQL & " ) OWNGROUP ON LoGroupeDriver.LoGroupID = OWNGROUP.LoGroupID "
      StrSQL = StrSQL + " WHERE "
      StrSQL = StrSQL + " LomDriver.LoDriverID ='" & strDriverID  & "'"
      'Response.Write StrSQL
      ObjRSLO.Open StrSQL, ObjConnLO
      If ObjRSLO.recordcount > 0 then
        If Trim(ObjRSLO("AcceptStatus")) = "3" then
          CheckDriver = ""
          LimitErrMsg = "������"
        Else
          CheckDriver = Trim(ObjRSLO("LoDriverHeadID"))
          ObjRSLO.Close
          DisConnDBH ObjConnLO, ObjRSLO	'DB�ؒf
          Exit Function
        End If
      else
        CheckDriver = ""
      end if

      ObjRSLO.Close  
      
    End If
    DisConnDBH ObjConnLO, ObjRSLO	'DB�ؒf
End function

Function CheckGroup(strGroupID)
    dim StrSQL
    dim ObjConnLO, ObjRSLO
    
    CheckGroup = ""
    v_GroupDrivers = ""
    ConnDBH ObjConnLO, ObjRSLO
    
    If Trim(strGroupID) <> "" Then
      StrSQL = "SELECT LoGroupeDriver.* FROM LomGroup "
      StrSQL = StrSQL + " INNER JOIN LoGroupeDriver ON  LomGroup.LoGroupID=LoGroupeDriver.LoGroupID AND LomGroup.HiTSUserID = LoGroupeDriver.HiTSUserID "
      StrSQL = StrSQL + " WHERE LomGroup.HiTSUserID ='" & USER  & "'"
      StrSQL = StrSQL + "   AND LomGroup.LoGroupID ='" & Trim(strGroupID)  & "'"
      'Response.Write StrSQL 
      ObjRSLO.Open StrSQL, ObjConnLO
      If ObjRSLO.recordcount > 0 then
        CheckGroup = "AA"
        While Not ObjRSLO.EOF
          v_GroupDrivers = v_GroupDrivers & Trim(ObjRSLO("LoDriverID")) & "," 
          ObjRSLO.MoveNext
        Wend
      else
        CheckGroup = ""
        v_GroupDrivers = ""
      end if
    End If
    DisConnDBH ObjConnLO, ObjRSLO	'DB�ؒf
End function


Function getDriverInfo()
    dim StrSQL
    dim ObjConnLO, ObjRSLO
    dim cnt
    dim chrCnt
 
    ConnDBH ObjConnLO, ObjRSLO
    
    StrSQL = "SELECT * "    
    StrSQL = StrSQL & " FROM LomDriver " 
    StrSQL = StrSQL & " WHERE HiTSUserID='" & USER & "'" 
    
    ObjRSLO.Open StrSQL, ObjConnLO
    If ObjRSLO.recordcount > 0 then
       	Redim Arr_DriverInfoText(CInt(ObjRSLO.recordcount))
	    Redim Arr_DriverInfoVal(CInt(ObjRSLO.recordcount))
	    
	    Redim Arr_DriverIDInfo(CInt(ObjRSLO.recordcount))
	    Redim Arr_DriverHeadIDInfo(CInt(ObjRSLO.recordcount))
	    Redim Arr_DriverNameInfo(CInt(ObjRSLO.recordcount))
	    Redim Arr_DriverCompanyInfo(CInt(ObjRSLO.recordcount))
	
	    Arr_DriverInfoText(0) =""
	    Arr_DriverInfoVal(0) =""
	    cnt = 1
        While NOT ObjRSLO.EOF
           'Arr_DriverInfoText(cnt) = Trim(ObjRSLO("ComboSelection"))'Trim(ObjRSLO("LoDriverID")) & String(14-Len(Trim(ObjRSLO("LoDriverID"))),"*") & Trim(ObjRSLO("LoDriverHeadID")) & String(7-Len(Trim(ObjRSLO("LoDriverHeadID"))),"*") & Trim(ObjRSLO("LoDriverName")) & String(18-Len(Trim(ObjRSLO("LoDriverName"))),"*") & Trim(ObjRSLO("LoDriverCompany"))
           Arr_DriverInfoVal(cnt) = CStr(Trim(ObjRSLO("LoDriverID")))
           
           Arr_DriverIDInfo(cnt) = Trim(ObjRSLO("LoDriverID"))
           If LEN(Arr_DriverIDInfo(cnt)) >= 0 then
             For chrCnt= LEN(Arr_DriverIDInfo(cnt)) to 13 
               Arr_DriverIDInfo(cnt) = Arr_DriverIDInfo(cnt) & "&nbsp;"
             Next
           end if
	       Arr_DriverHeadIDInfo(cnt) = Trim(ObjRSLO("LoDriverHeadID"))
	       If LEN(Arr_DriverHeadIDInfo(cnt)) >= 0 then
             For chrCnt = LEN(Arr_DriverHeadIDInfo(cnt)) to 6 
               Arr_DriverHeadIDInfo(cnt) = Arr_DriverHeadIDInfo(cnt) & "&nbsp;"
             Next
           end if
	       Arr_DriverNameInfo(cnt) = Trim(ObjRSLO("LoDriverName"))
	       If LEN(Arr_DriverNameInfo(cnt)) >= 0 then
             For chrCnt=getByteLength(Arr_DriverNameInfo(cnt)) to 17 
               Arr_DriverNameInfo(cnt) = Arr_DriverNameInfo(cnt) & "&nbsp;"
             Next
           end if
           
	       Arr_DriverCompanyInfo(cnt) = Trim(ObjRSLO("LoDriverCompany"))           
           If LEN(Arr_DriverCompanyInfo(cnt)) >= 0 then
             For chrCnt=getByteLength(Arr_DriverCompanyInfo(cnt)) to 21 
               Arr_DriverCompanyInfo(cnt) = Arr_DriverCompanyInfo(cnt) & "&nbsp;"
             Next
           end if
           
           Arr_DriverInfoText(cnt) = Arr_DriverIDInfo(cnt) & Arr_DriverHeadIDInfo(cnt) & Arr_DriverNameInfo(cnt) & Arr_DriverCompanyInfo(cnt)
           'Response.Write Arr_DriverInfoText(cnt) & "<BR/>"
           cnt = cnt + 1
           ObjRSLO.MoveNext
        Wend
    End if
    ObjRSLO.Close  
    DisConnDBH ObjConnLO, ObjRSLO	'DB�ؒf

End Function

Function LfUpdGrpLOInfo()

   dim StrSQL
   dim ObjConnLO, ObjRSLO
   dim ErrFlg
   dim iSeq
   dim Arr_UpdDriverID
   dim UpdFlg
   dim icnt
    
   dim strWkNo
   dim strContNo
   dim Arr_GroupDrivers
   
   ConnDBH ObjConnLO, ObjRSLO	
   WriteLogH "", "", "", ""
	
   ErrFlg = false
   UpdFlg = true
   If v_driverInfoChkFlg = "2" then
     Arr_UpdDriverID = TRIM(MID(v_GroupInfo,1,2))
   else
     Arr_UpdDriverID = TRIM(MID(v_DriverInfo,1,12))	
   end if
   
    For i = 1 to v_DataCnt2-1
      Arr_LOID(i) = Arr_UpdDriverID
      If UCase(Trim(Arr_Check(i))) = "ON" And Trim(Arr_LoStatus(i)) <> "2" And Trim(Arr_WorkDate(i)) <> "" Then
        If Trim(Arr_UpdDriverID) <> "" Then
          
          If v_driverInfoChkFlg = "2" then
            If CheckGroup(Arr_UpdDriverID) = "" then
              Arr_ErrID(i) = UCase(Arr_UpdDriverID)
              Arr_Err(i) = "�o�^��"
              ErrFlg = false
            Else
              ErrFlg = false
            End If 
          Else
            If CheckDriver(Arr_UpdDriverID) = "" then
              If Trim(LimitErrMsg) = "" Then
                Arr_ErrID(i) = "ID�s��"
              Else
                Arr_ErrID(i) = LimitErrMsg
              End if
              Arr_Err(i) = "�o�^�Ȃ�"
              ErrFlg = true
            Else
              ErrFlg = false
            End If 
          End if
        
          If ErrFlg = false then
            If Trim(Arr_DriverID(i)) = "" then
              'QUERY VALUES FOR INSERTION
              StrSQL = "SELECT * FROM LoLockOnInfo WHERE WkNo ='" & Arr_WkNo(i)  & "'"&_
                                                   " AND ContNo='" & Arr_ContNo(i) & "'"
              ObjRSLO.Open StrSQL, ObjConnLO
      
              If ObjRSLO.recordcount > 0 then
                While NOT ObjRSLO.EOF
                  iSeq = CInt(Trim(ObjRSLO("SeqNumber")))
                  ObjRSLO.MoveNext
                Wend
		        iSeq = iSeq + 1
		        UpdFlg = true
              Else
                iSeq = 1
                UpdFlg = true
              End if
            
              If UpdFlg <> false then
                StrSQL = " INSERT INTO LoLockOnInfo (WkNo, ContNo, SeqNumber, UpdtTime, UpdtPgCd, UpdtTmnl, LoStatus, LoNearRange, KaccsSend, InputID, LoWorkKind, HiTSUserID, LoReqTime, TerminalCode, WorkDate)"
                StrSQL = StrSQL & " VALUES ( "
                StrSQL = StrSQL & "'" & Trim(Arr_WkNo(i)) & "',"                       'WkNo
                StrSQL = StrSQL & "'" & Trim(Arr_ContNo(i)) & "',"                     'ContNo
                StrSQL = StrSQL & "'" & iSeq & "',"                                    'Seq
                StrSQL = StrSQL & "'" & Now() & "',"                                   'UpdtTime
                StrSQL = StrSQL & "'" & "PREDEF01" & "',"                              'UpdtPgCd
                StrSQL = StrSQL & "'" & USER & "',"                                    'UpdtTmnl
                StrSQL = StrSQL & "'" & "1" & "',"                                     'LoStatus
                StrSQL = StrSQL & "" & "NULL" & ","                                    'LoNearRange
                StrSQL = StrSQL & "'" & "0" & "',"                                     'KaccsSend
                'StrSQL = StrSQL & "'" & "0" & "',"                                    'SphoneSend
                StrSQL = StrSQL & "'" & Trim(Arr_UpdDriverID) & "', "                  'InputID
                If v_driverInfoChkFlg = "2" then
                  'StrSQL = StrSQL & "'', "                                            'DriverID
                  'StrSQL = StrSQL & "'', "                                            'HeadID
                Else
                  'StrSQL = StrSQL & "'" & Trim(Arr_UpdDriverID(0)) & "', "               'DriverID
                  'StrSQL = StrSQL & "'" & Trim(CheckDriver(Arr_UpdDriverID(0))) & "', "  'HeadID
                  'StrSQL = StrSQL & "'', "                                             'HeadID
                End if
                StrSQL = StrSQL & "'" & Trim(Arr_ImpExp(i)) & "', "                    'LoWorkKind
                StrSQL = StrSQL & "'" & USER & "',"                                    'HiTSUserID
                StrSQL = StrSQL & "'" & Now() & "',"                                   'LoReqTime
                StrSQL = StrSQL & "'" & Trim(Arr_RecTerminalCode(i)) & "', "           'TerminalCode
                If Trim(Arr_WorkDate(i)) <> "" then
                  StrSQL = StrSQL & "'" & Trim(Arr_WorkDate(i)) & "' "                 'WorkDate
                else
                  StrSQL = StrSQL & "NULL"
                end if            
                StrSQL = StrSQL & ")"
                'response.Write StrSQL & "<BR/>"
                ObjConnLO.Execute(StrSQL)
          
                if err <> 0 then
			      Set ObjRSLO = Nothing				
			      jampErrerPDB ObjConnLO,"2","b107","01","","104","SQL:<BR>"& strSQL
			      ObjRSLO.close
			      DisConnDBH ObjConnLO, ObjRSLO	'DB�ؒf
			      Exit Function
		        end if  
		        
		        If v_driverInfoChkFlg = "2" Then
		           Arr_GroupDrivers = Split(v_GroupDrivers,",")
		           For icnt = 0 To UBound(Arr_GroupDrivers)
		             If Trim(Arr_GroupDrivers(icnt)) <> "" Then
		               'StrSQL = " UPDATE LomDriver SET "
                       'StrSQL = StrSQL & "SpAlarmOrder='1', "                                          'SpAlarmOrder
                       'StrSQL = StrSQL & "SpAlarmCancel='' "                                           'SpAlarmCancel
                       'StrSQL = StrSQL & "WHERE LoDriverID ='" & Trim(Arr_GroupDrivers(icnt))  & "'"       
                       'Response.Write StrSQL & "<BR/>"
                       'ObjConnLO.Execute(StrSQL)
	                   'If err <> 0 Then
	                   '  Set ObjRSLO = Nothing				
	                   '  jampErrerPDB ObjConnLO,"2","b107","01","","104","SQL:<BR>"& strSQL
	                   '  Exit For
	                   'End If
	                   fUpdateDriverAlarm(Arr_GroupDrivers(icnt))
	                 End If
		           Next
		        Else
		          'StrSQL = " UPDATE LomDriver SET "
                  'StrSQL = StrSQL & "SpAlarmOrder='1', "                                          'SpAlarmOrder
                  'StrSQL = StrSQL & "SpAlarmCancel='' "                                           'SpAlarmCancel
                  'StrSQL = StrSQL & "WHERE LoDriverID ='" & Trim(Arr_UpdDriverID)  & "'"       
                  'Response.Write StrSQL
                  'ObjConnLO.Execute(StrSQL)
	              'If err <> 0 Then
	              '  Set ObjRSLO = Nothing				
	              '  jampErrerPDB ObjConnLO,"2","b107","01","","104","SQL:<BR>"& strSQL
	              'End If
	              fUpdateDriverAlarm(Arr_UpdDriverID)
		        End If
		      End if
		      ObjRSLO.close  
		    Elseif Trim(Arr_DriverID(i)) <> "" then 
		       StrSQL = "SELECT * FROM LoLockOnInfo WHERE WkNo ='" & Arr_WkNo(i)  & "'"&_
                                               " AND ContNo='" & Arr_ContNo(i) & "'"&_
                                               " AND SeqNumber='" & Arr_SeqNumber(i) & "'"
              ObjRSLO.Open StrSQL, ObjConnLO
              If ObjRSLO.recordcount > 0 then
                 if Trim(ObjRSLO("LoStatus")) <> "2" then
                   call LfUpdLODriverID(Arr_WkNo(i), Arr_ContNo(i), Arr_SeqNumber(i), Arr_UpdDriverID, v_GroupDrivers)
                 end if
              end if
		      ObjRSLO.close 
		    End if
		   
          End if
        End if
      End if
    Next
    
    DisConnDBH ObjConnLO, ObjRSLO	'DB�ؒf
End Function

Function fUpdateDriverAlarm(strDriverID)
    dim StrSQL
    dim ObjConnLO, ObjRSLO
    dim ErrFlg
    dim iSeq
    dim Arr_GroupDrivers
    dim icnt
	
    ConnDBH ObjConnLO, ObjRSLO	
	WriteLogH "", "", "", ""
	
	If strDriverID <> "" Then
	  StrSQL = "SELECT * FROM LomDriver "
      StrSQL = StrSQL & "WHERE LoDriverID ='" & Trim(CSTR(strDriverID))  & "'"              
      ObjRSLO.Open StrSQL, ObjConnLO
      If ObjRSLO.recordcount > 0 then
         StrSQL = " UPDATE LomDriver SET "
         StrSQL = StrSQL & "SpAlarmOrder='1', "                                          'SpAlarmOrder
         StrSQL = StrSQL & "SpAlarmCancel='' "                                           'SpAlarmCancel
         StrSQL = StrSQL & "WHERE LoDriverID ='" & Trim(CSTR(strDriverID))  & "'"       
         'Response.Write "HERE:" & StrSQL
         ObjConnLO.Execute(StrSQL)
	     If err <> 0 Then
	       Set ObjRSLO = Nothing				
	       jampErrerPDB ObjConnLO,"2","b107","01","","104","SQL:<BR>"& strSQL
	     End If
	  End If
	End If              
	ObjRSLO.close 
	DisConnDBH ObjConnLO, ObjRSLO	'DB�ؒf              
End Function

Function getSort2(Key,SortKey,str)
	getSort2 = str	
	if str = "" AND Key<>"" then
		str = " ORDER BY "
	elseif str <> "" AND Key<>"" Then 
		str = str & ","		
	end if

	if Key <> "" then 
			if FieldName2(CInt(Key)) = "DelPermitDate" AND SortKey = "ASC" then 
		        str = str & " DelPermitDate_Sort ASC "
			elseif FieldName2(CInt(Key)) = "CYCut" AND SortKey = "ASC" then 
		        str = str & " CYCut_Sort ASC "
			elseif FieldName2(CInt(Key)) = "FreeTime" AND SortKey = "ASC" then 
				str = str & " FreeTime_Sort ASC "
			elseif FieldName2(CInt(Key)) = "WorkDate" AND SortKey = "ASC" then 
		        str = str & " WorkDate_Sort ASC "
		    else
		        str = str & FieldName2(Key) & " " & SortKey	
		    end if			
	end if	
    getSort2 = str  
end function

Function getImage(SortKey)
	getImage = ""
	if SortKey = "ASC" then
		getImage = CONST_ASC	
	else
		getImage = CONST_DESC
	end if	
end function

function LfgetSCAC(keyCode1,arr())
	dim ObjConnSCAC, ObjRSSCAC, StrSQL
    dim cnt

    cnt = 0         '������
    LfgetSCAC = ""
	
	'�G���[�g���b�v�J�n
	on error resume next	
	'DB�ڑ�	
	ConnDBH ObjConnSCAC, ObjRSSCAC
	
    redim arr(0, 0)
    if trim(keyCode1) <> "" then
        StrSQL = "select Distinct ShipLine from mShipLine where ShipLine = '" & Trim(keyCode1) & "'"
    else
        StrSQL = "select Distinct ShipLine from mShipLine"
    end if
	
    ObjRSSCAC.Open StrSQL, ObjConnSCAC
	
    if not ObjRSSCAC.eof then
        redim arr(ObjRSSCAC.recordcount,0)
        while not ObjRSSCAC.eof
            cnt=cnt+1
            arr(cnt,0)=Trim(ucase(ObjRSSCAC("ShipLine")))
            ObjRSSCAC.movenext
        wend
    end if
	
	DisConnDBH ObjConnSCAC, ObjRSSCAC	
end function

function LfPutPage(rec,page,pagecount,link)
	dim pg, i, j
	dim FirstPage, LastPage	
	dim PageIndex
	dim PageWkNo
	dim intNextFlag
	dim strParam
	PageIndex=0
	PageWkNo=0	
	if rec > 0 then	

		if pagecount<page then
			page=pagecount
		end if
		'�y�[�WIndex��ݒ�
		PageIndex=Fix(page/gcPage)
		if page mod gcPage=0 then
			PageIndex=PageIndex-1
		End If
		PageWkNo=((gcPage*PageIndex)+1)-gcPage
		
		'�擪�y�[�W��0��菬�����ꍇ��1��ݒ�
		if PageWkNo<=0 Then
			PageWkNo=0
		End If

		'�p�����[�^�ݒ�
		
	    strParam="&InOutF=" & v_InOutFlag
		
		'--- �������A���y�[�W�� 
		LastPage=pagecount		
		FirstPage=1
			
		if page>1 then
			response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & FirstPage & strParam & """>�ŏ���</a>"
			response.write "| &nbsp;"
			'Y.TAKAKUWA Upd-S 2015-03-13
			'if PageWkNo<>0 Then
			'	response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & PageWkNo & strParam & """>�O��</a>"
			'Else
			'	response.write "<font style='color:#FFFFFF;'>�O��</font>"
			'End If
			response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & page-1 & strParam & """>�O��</a>"
			'Y.TAKAKUWA Upd-E 2015-03-13
		else
			response.write "<font style='color:#FFFFFF;'>�ŏ���</font>"
			response.write "| &nbsp;"
			response.write "<font style='color:#FFFFFF;'>�O��</font>"
		end if        		
		'--- �C���f�b�N�X
		'�y�[�W��1�y�[�W�ȏ㑶�݂���ꍇ
		if pagecount>1 then
			response.write "| &nbsp;"

			'�w��y�[�W�������[�v
			for i=1 to gcPage
				'�y�[�W���Z�o
				PageWkNo=(gcPage*PageIndex)+i

				'�y�[�W���S�y�[�W���傫���ꍇ�͏������f
				if pagecount< PageWkNo then
					PageWkNo=PageWkNo-1
					exit for
				end if
				'���ݑI������Ă���y�[�W�̏ꍇ
				if PageWkNo=page then
					response.write "&nbsp;" & PageWkNo 
				else
					response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & PageWkNo & strParam & """ >&nbsp;" & PageWkNo & "</a>"
				End If
			Next
			response.write "| &nbsp;"
		End If
		'Y.TAKAKUWA Upd-S 2015-03-13			
		'if page<pagecount-1 then
		if page<pagecount then
		'Y.TAKAKUWA Upd-E 2015-03-13
		    'Y.TAKAKUWA Upd-S 2015-03-13
			'PageWkNo=PageWkNo+1
			'If PageWkNo<=LastPage Then
			'	response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & PageWkNo & strParam & """>����</a>"'
			'Else
			'	response.write "<font style='color:#FFFFFF;'>����</font>"
			'End If
			PageWkNo=page+1
			response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & PageWkNo & strParam & """>����</a>"'
			'Y.TAKAKUWA Upd-E 2015-03-13
			response.write "| &nbsp;"
			response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & LastPage & strParam & """>�Ō��</a>"'            
		else
			response.write "<font style='color:#FFFFFF;'>����</font>"
			response.write "| &nbsp;"
			response.write "<font style='color:#FFFFFF;'>�Ō��</font>"
		end if
	end if
end function

function getByteLength(v_Data)

    dim StrSQL
    dim ObjConnLO, ObjRSLO
	dim ByteLength
    ConnDBH ObjConnLO, ObjRSLO
    
    StrSQL = "SELECT TOP 1 DATALENGTH('" & Trim(v_Data) & "') AS ByteLength "
    ObjRSLO.Open StrSQL, ObjConnLO
    If ObjRSLO.recordcount > 0 then
       ByteLength = Trim(ObjRSLO("ByteLength"))
    end if
    ObjRSLO.Close
    getByteLength = ByteLength
    DisConnDBH ObjConnLO, ObjRSLO	'DB�ؒf
    
end function

'-----------------------------
'   ���l�ϊ� (Long�^)
'-----------------------------
function gfCLng(str1)
    dim str
    str=gfTrim(str1)
    if isnull(str) then
        gfCLng=0
    elseif trim(str)="" then
        gfCLng=0
    elseif not isNumeric(str) then
        gfCLng=0
    elseif len(str)>9 then
        if instr(str,".")>0 and instr(str,".")<10 then
            gfClng=clng(left(str,instr(str,".")-1))
        else
            gfClng=0
        end if
    else
        gfCLng = CLng(fix(str))
    end if
end function

'-----------------------------
'   Trim�@NULL�̏ꍇ����l(Space0)
'-----------------------------
function gfTrim(str)
    if isnull(str) then
        gfTrim=""
    else
        gfTrim=trim(str)
    end if
end function
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE></TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR>
<STYLE>
th.hlist {
	position: relative;
}
th {
    border-width: 1px 1px 1px 1px;
    padding: 4px;
    background-color: #ffcc33;
}
SELECT.chr {
    BACKGROUND-COLOR: #ffffff;
    BORDER-BOTTOM: #ffffff 1px solid;
    BORDER-LEFT: #002f7b 0px solid;
    BORDER-RIGHT: #ffffff 0px solid;
    BORDER-TOP: #ffffff 0px solid;
    COLOR: black;
    FONT-FAMILY: '�l�r �S�V�b�N';
    FONT-SIZE: 12px;
    FONT-WEIGHT: normal;
    PADDING-BOTTOM: 2px;
    PADDING-LEFT: 1px;
    PADDING-RIGHT: 2px;
    PADDING-TOP: 3px;
    TEXT-ALIGN: left
}
table {
    border-width: 0px 1px 1px 0px;
}
DIV.center {
	text-align:center;
}
DIV.BDIV1 {
    position: relative;
    border-width: 0px 0px 1px 0px;
}
DIV.BDIV2 {
    position: relative;
    border-width: 0px 0px 1px 0px;
}
thead tr {
    position: relative;
    top: expression(this.offsetParent.scrollTop);
}
#loading2 {
	font:bold 12px Verdana;
	color:red;
	position:absolute; 
	top:220px; 
	left:390px;
	width:300px;
	height:30px; 
	z-index:69;
	font-size:12pt;
	border:0px;
	vertical-align: middle;
}
</style>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT Language="JavaScript">
function finit(){
	var str;
	
	//�f�[�^���p���ݒ�  
    document.frm.Gamen_Mode.value="<%=v_GamenMode%>";  
    document.frm.InOutSrhFlag.value='';	
	str = readCookie('HitsTbl2')
	if(str!= null && "<%=Num2%>" != "0"){  		  		
		displayColumn(str,"TBInOut")
	}
	
}

function displayColumn(str,TableName){

	var objTable;
	var intMaxRow;
	var disp;
	var colHeader;
    
	//Table�I�u�W�F�N�g�擾
	objTable=document.getElementById(TableName);
	//�s���擾
	intMaxRow=objTable.rows.length;

	var trs=objTable.getElementsByTagName("TR");


	if (TableName=='TBEmpty'){
		//�s�������[�v
		for(var intRowCnt=1; intRowCnt<intMaxRow; intRowCnt++) {

			var tds=trs[intRowCnt].getElementsByTagName("TD");
			for(var intColCnt=0; intColCnt<14; intColCnt++) {

				//�w�b�_�[ID��ݒ�
				if(intColCnt<9){
					colHeader="H1Col0"+(intColCnt+1)
				}else{
					colHeader="H1Col"+(intColCnt+1)
				}


				if(str.charAt(intColCnt) == "0"){
					disp = 'none'
				}else{
					disp = '';
				}
				if (intRowCnt==1){
					document.getElementById(colHeader).style.display = disp;
				}
				tds[intColCnt].style.display =disp;
			}

		}

	}else{
		//�s�������[�v
		for(var intRowCnt=1; intRowCnt<intMaxRow; intRowCnt++) {

			var tds=trs[intRowCnt].getElementsByTagName("TD");
			for(var intColCnt=0; intColCnt<19; intColCnt++) {

				//�w�b�_�[ID��ݒ�
				if(intColCnt<9){
					colHeader="H2Col0"+(intColCnt+1)
				}else{
					colHeader="H2Col"+(intColCnt+1)
				}


				if(str.charAt(intColCnt) == "0"){
				   if(intColCnt== 0 || intColCnt== 3 || intColCnt== 4 || intColCnt== 5 || intColCnt== 6){
				     disp = ''
				   }
				   else{
					disp = 'none'
				   }
				}else{
				    if(intColCnt== 1 || intColCnt== 2){
				       disp="none";
				    }
				    else{
					  disp = '';
					}
				}
				if (intRowCnt==1){
					document.getElementById(colHeader).style.display = disp;
				}
				tds[intColCnt].style.display =disp;
			}

		}	

	}

}


function readCookie(name) {
	var nameEQ = name + "=";
	var ca = document.cookie.split(';');
	for(var i=0;i < ca.length;i++) {
		var c = ca[i];
		while (c.charAt(0)==' ') c = c.substring(1,c.length);
		if (c.indexOf(nameEQ) == 0) return c.substring(nameEQ.length,c.length);
	}
	return null;
}

//�f�[�^�������ꍇ�̕\������
function view(){
	var sortedHeight;
	sortedHeight = 0;
	var vHeight;
	var obj2=document.getElementById("BDIV2");
	var objH1=document.getElementById("DivHeaderRow");
	var rowHeight;
	
	if('<%=Num2%>'!='0'){
	  var rowHeightTbody = getRowHeightTbody();
	  if(rowHeightTbody > 0){
	    rowHeight=rowHeightTbody*10
	  }
	  rowHeight=rowHeight;//+rowHeightThead
    }
    else{
      rowHeight = 0;
      rowHeight=(24*10) + 38;
    }
    
	if((document.body.offsetWidth-50) < 50){
		obj2.style.width=50;
		objH1.style.width = 50;
		obj2.style.overflowX="auto";	 
	}else if((document.body.offsetWidth-50)  < 813){
		obj2.style.width=document.body.offsetWidth-30;
		objH1.style.width = document.body.offsetWidth-30;
		obj2.style.overflowX="auto";
	}else{
		obj2.style.width=document.body.offsetWidth-30;
		objH1.style.width=document.body.offsetWidth-30;
		obj2.style.overflowX="auto";
	}	
	
	
	if('<%=blnSorted2%>'.toUpperCase()=='TRUE'){
	   sortedHeight = 10;
	}
	
	if((document.body.offsetHeight-rowHeight) < 133){ 
	    if(obj2.clientWidth<obj2.scrollWidth)
	    {
	      obj2.style.height = 55+sortedHeight;
		  obj2.style.overflowY = "auto";
	    }
	    else{
	      obj2.style.height = 40+sortedHeight;
		  obj2.style.overflowY = "auto";
		}

	}else if((document.body.offsetHeight-rowHeight) < 410){
	    vHeight = rowHeight + 83;
		obj2.style.height = document.body.offsetHeight-vHeight;
		obj2.style.overflowY = "auto";
	}else{
		if(obj2.clientWidth < obj2.scrollWidth)
        {
           obj2.style.height = rowHeight+17;
           obj2.style.overflowY="auto";
        }
        else{
           obj2.style.height = rowHeight;
           obj2.style.overflowY="auto";
        } 
        	
	}
	//<!-- ARVEE EDIT-S 2014-11-17-->
	if(document.frm.windowHeight){
		document.frm.windowHeight.value = document.body.offsetHeight;
	}
	if(document.frm.TableHeight){
		document.frm.TableHeight.value = rowHeight;
	}
   
    var obj3=document.getElementById("BDIV3");

	if((document.body.offsetWidth-10)  < 880){
		obj3.style.width=document.body.offsetWidth-10;
		obj3.style.overflowX="auto";
	}
	else{
		obj3.style.width=document.body.offsetWidth-10;
		obj3.style.overflowX="auto";
	}
    if((document.body.offsetHeight) > 15 ){
	  obj3.style.height=document.body.offsetHeight-15;
	  obj3.style.overflowY="auto";
	}
	else{
	  obj3.style.height=document.body.offsetHeight;
	  obj3.style.overflowY="auto";	
	}
}

function getRowHeightThead()
{
  var oRows = document.getElementById('TBInOut').getElementsByTagName('thead');
  var rowsH=[];
  var rowsHeight;
  for(var i=0;i<oRows.length;i++){ 
    rowsH[i]=oRows[i].offsetHeight; 
    rowsHeight = rowsH[i];
  } 
  return rowsHeight;
}
function getRowHeightTbody()
{
  var oRows = document.getElementById('TBInOut').getElementsByTagName('td');
  var rowsH=[];
  var rowsHeight;
  for(var i=0;i<oRows.length;i++){ 
    rowsH[i]=oRows[i].offsetHeight; 
    rowsHeight = rowsH[i];
  } 
  return rowsHeight;
}

function OpenCodeWin(strTitle,tableno)
{
	var CodeWin;
	var w=400;
	var h=525;
	var l=0;
	var t=0;
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
	
	CodeWin = window.open("./show.asp?user=<%=Session.Contents("userid")%>&pagetitle="+strTitle+"&show_column="+tableno,"codelist","scrollbars=yes,resizable=yes,width="+w+",height="+h+",top="+t+",left="+l);

	CodeWin.focus();

}

function OpenCodeWin2(tableno){
	var CodeWin;
	var w=400;
	var h=300;
	var l=0;
	var t=0;
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
	
	CodeWin = window.open("./sort.asp?user=<%=Session.Contents("userid")%>&left_menu="+tableno,"codelist","scrollbars=yes,resizable=yes,width="+w+",height="+h+",top="+t+",left="+l);
	CodeWin.focus();
}

//�X�V
function GoRenew(No){
	Fname=document.frm;
	Fname.targetNo.value=No;
	Fname.action="./dmi015.asp";
	newWin = window.open("", "ReEntry", "left=10,top=10,status=yes,resizable=yes,scrollbars=yes");
	Fname.target="ReEntry";
	Fname.elements[No].disabled=false;
	Fname.submit();
	Fname.elements[No].disabled=true;
	Fname.target="_self";
	Fname.action="";
}

//�X�V
function GoRenew2(sakuNo,bookNo,conNo){
  Fname=document.frm;
  Fname.SakuNo.value=sakuNo;
  Fname.BookNo.value=bookNo;
  Fname.CONnum.value=conNo;
  Fname.action="./dmo320.asp";
  newWin = window.open("", "ReEntry", "status=yes,width=500,height=500,left=10,top=10,resizable=yes,scrollbars=yes");
  Fname.target="ReEntry";
  Fname.submit();
  Fname.target="_self";
  Fname.action="";
}

//�R���e�i�ڍ�
function GoConinf2(wkcNo,flag,conNo){
	Fname=document.frm;  
	Fname.SakuNo.value=wkcNo;
	Fname.flag.value=flag;
	Fname.CONnum.value=conNo;
	Fname.action="./dmo900.asp";
	newWin = window.open("", "ConInfo", "left=30,top=10,status=yes,scrollbars=yes,resizable=yes,menubar=yes");
	Fname.target="ConInfo";
	Fname.submit();
	Fname.target="_self";
	Fname.InfoFlag.value="0";
	Fname.action="";
}

//�R���e�i�ڍ�
function GoConinfContNo(conNo){
  Fname=document.frm;
  Fname.CONnum.value=conNo;
  Fname.BookNo.value="";        //CW-021 ADD
  BookInfo(Fname);
  Fname.action="";
}

function LockOnReg(){
	document.frm.Gamen_Mode.value = "I";
    document.frm.submit();
}

function fRelease(){
    var i;
    var updReleaseFlg;
    var obj;
    var obj2;
    
    for(i=1; i <= (parseInt(document.frm.DataCnt2.value)-1); i++){
       obj = eval("document.frm.chkInOut" + i);
	   if (obj.checked==true) {
          obj2 = eval("document.frm.LoStatus" + i)	      
	      if(obj2.value == "1"){
	        if(updReleaseFlg==2){
	          updReleaseFlg = 2;
	        }else{ 
	          updReleaseFlg = 1;
	        }
	      } 
	      if(obj2.value == "2"){
	        updReleaseFlg = 2;
	      } 
	   }
	}
	
	if(updReleaseFlg ==1){
	   var msg = confirm("�R���e�i���b�N�w�����������܂��B��낵���ł����H");
       if(msg == true){
         document.frm.Gamen_Mode.value = "R";
         document.frm.submit();
       }  
	}
	else{
	   if(updReleaseFlg ==2){
	     var msg = confirm("�R���e�i���b�N���̍�Ƃ�����܂�����낵���ł����H");
         if(msg == true){
           document.frm.Gamen_Mode.value = "R";
           document.frm.submit();
         }
	   }
	}
	
}

function chkobj(id)
{
    var obj;

    obj = eval("document.frm." + id);
    return (obj.checked) ? 1 : 0;
}

function fChange(obj1,obj2)
{
  document.getElementById(obj1).checked = true;
  document.getElementById(obj2).checked = false;
}

function fSelect(objChk1,objChk2,obj)
{
  document.getElementById(objChk1).checked = true;
  document.getElementById(objChk2).checked = false;
  document.getElementById(obj).value = "";
}

function fRSearch(){
	document.frm.Gamen_Mode.value = "S";
    document.frm.submit();
}

function fGrpUpd(){
    if(document.frm.driverInfo.value != "" && document.frm.groupInfoChk1.checked == true )
    {
       document.frm.Gamen_Mode.value = "GI";
       document.frm.submit();
    }
    else{
      if(document.frm.groupInfo.value != "" && document.frm.groupInfoChk2.checked == true )
      {
        document.frm.Gamen_Mode.value = "GI";
        document.frm.submit();
      }
    }
}
function OpenWin(i)
{
   switch(i){
     case 1:
       var w=1100;
       var h=700;
       var l=0;
       var t=0;
       if(screen.width){
         l=(screen.width-w)/2;
       }
       if(screen.availWidth){
         l=(screen.availWidth-w)/2;
       }
       if(screen.height){
         t=(screen.height-(h+140))/2;
       }
       if(screen.availHeight){
          t=(screen.availHeight-(h+140))/2;
       }
       Win = window.open('dml000B.asp', 'DriverApproval', 'status=no,width='+w+',height='+h+',top='+t+',left='+l+',resizable=yes,scrollbars=yes,toolbar=yes,menubar=yes,location=yes');
       break;

     case 2:
       var w=1100;
       var h=700;
       var l=0;
       var t=0;
       if(screen.width){
         l=(screen.width-w)/2;
       }
       if(screen.availWidth){
         l=(screen.availWidth-w)/2;
       }
       if(screen.height){
         t=(screen.height-(h+140))/2;
       }
       if(screen.availHeight){
          t=(screen.availHeight-(h+140))/2;
       }
       Win = window.open('dml000D.asp', 'LockOnDriver', 'status=no,width='+w+',height='+h+',top='+t+',left='+l+',resizable=yes,scrollbars=yes,toolbar=yes,menubar=yes,location=yes');
       break;
     case 3:
       var w=1100;
       var l=0;
       var t=0;
       if(screen.width){
         l=(screen.width-w)/2;
       }
       if(screen.availWidth){
         l=(screen.availWidth-w)/2;
       }
       Win = window.open('dml000C.asp', 'LockOnGroup', 'status=no,width='+w+',height='+screen.height+',top='+0+',left='+l+',resizable=yes,scrollbars=yes,toolbar=yes,menubar=yes,location=yes');
       break;
   }
}
function OnScrollDiv(Scrollablediv) {
    document.getElementById('DivHeaderRow').scrollLeft = Scrollablediv.scrollLeft;
}
</SCRIPT>
</HEAD>
<BODY onLoad="finit();view();" onResize="view();">
<form name="frm" method="post">
<div id="BDIV3">

<!--Hidden Values Start-->
<INPUT type=hidden name="Gamen_Mode" size="9" readonly tabindex= -1>
<INPUT type=hidden name="InfoFlag" value="">
<INPUT type=hidden name="SakuNo" value="">
<INPUT type=hidden name="flag" value="">
<INPUT type=hidden name="targetNo" value="">
<INPUT type=hidden name="CONnum" value="" >
<INPUT type=hidden name="BookNo" value="" >
<INPUT type=hidden name="CompF" value="" >
<INPUT type=hidden name="COMPcd0" value="" >
<INPUT type=hidden name="COMPcd1" value="" >
<INPUT type=hidden name="ShipLine" value="" >
<INPUT type=hidden name="ShoriMode" value="EMoutUpd">
<INPUT type=hidden name="Mord" value="1" >
<INPUT type=hidden name="SCACSrhFlag">
<INPUT type=hidden name="InOutSrhFlag">
<!--<INPUT name="windowHeight">
<INPUT name="TableHeight">-->
<!--Hidden Values End-->
<!--Added Start-->
<table border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td width="100%" height="30">
        <div style="margin-left:30px;">
          <table border="0" cellpadding="0" cellspacing="0">
			  <!-- 2016/07/27 H.Yoshikawa Del start -->
		      <!-- <td width="115" align="center" nowrap><a HREF="JavaScript:OpenWin(1)">�h���C�o���F</a></td> -->
			  <!-- 2016/07/27 H.Yoshikawa Del end   -->
		      <td width="115" align="center" nowrap>
		        <a HREF="JavaScript:OpenWin(2)">�h���C�o�ꗗ�ƍ폜</a>
              </td>
		      <td align="center" nowrap style="width: 115px"><a href="JavaScript:OpenWin(3)">�O���[�v�o�^</a></td>
		      <!--Hide and Ordering Buttons Start-->
		      <!--
		      <td><input type="button" value="�\����ݒ�" onClick="OpenCodeWin('�\����I���i�����o���j','2')"></td>
		      <td><input type="button" value="���ёւ�"  onClick="OpenCodeWin2('6');"></td>
		      -->
		      <!--Hide and Ordering Buttons End-->
		      <td width="250" align="center" nowrap>
		      <!--Page Pagination Start-->
		      
		          <%					
						If Num2 > 0 Then						
							abspage = ObjRS2.AbsolutePage
							pagecnt = ObjRS2.PageCount
							call LfPutPage(Num2,abspage,pagecnt,"pagenum2")
						End If									
				  %>
		      <!--Page Pagination End-->
		      </td>
		      <td align="left">
		         <table>
		         <tr>
		           <td>
		              <select name="cmbINOut">
					  <% If v_InOutFlag="1" Then %>
						<OPTION VALUE = ''>���������o</OPTION>
						<OPTION VALUE = '1' SELECTED>�����̂�</OPTION>
						<OPTION VALUE = '2'>���o�̂�</OPTION>
					  <% ELSEIf v_InOutFlag="2" Then %>
						<OPTION VALUE = ''>���������o</OPTION>
						<OPTION VALUE = '1'>�����̂�</OPTION>
						<OPTION VALUE = '2' SELECTED>���o�̂�</OPTION>
					  <% ELSE  %>
						<OPTION VALUE = '' SELECTED>���������o</OPTION>
						<OPTION VALUE = '1'>�����̂�</OPTION>
						<OPTION VALUE = '2'>���o�̂�</OPTION>
					  <% END IF  %>
				     </Select>
		           </td>
		           <td>
		              <input type="button" name="Button"  value="�\���X�V" onClick="fRSearch()">
		           </td>
		         </tr>
		         </table>
		      </td>
		      
		    
		  </table> 
        </div>
      </td>
    </tr>
	<tr>		
		<!--Place Here Start-->
		<td>
		    <!-- Y.TAKAKUWA Add-S 2014-11-17 -->
			<div style="overflow: hidden; background-color:#fff;" id="DivHeaderRow">
					<table border="1" cellpadding="0" cellspacing="0" width="100%" id="TBInOut2">				
						<% If blnSorted2 Then%>
						  <thead style="height:48px;">
						<%Else%>
						  <thead>
						<%End If%>
						    <!--HEADER INFORMATION START-->
							<tr>
								<th id="H2Col01"" class="hlist" align="center" nowrap><%=HeaderTbl2(0)%></th>
								<th id="H2Col02"class="hlist" align="left" nowrap style="display:none;"></th>
								<th id="H2Col03"class="hlist" align="left" nowrap style="display:none;"></th>								
								<th id="H2Col04" class="hlist" nowrap><%=HeaderTbl2(1)%></th>
								<th id="H2Col05" class="hlist" nowrap><%=HeaderTbl2(2)%></th>
								<th id="H2Col06" class="hlist" nowrap><%=HeaderTbl2(3)%></th>
								<th id="H2Col07" class="hlist" nowrap><%=HeaderTbl2(4)%></th>									
								<th id="H2Col08" class="hlist" nowrap><%=HeaderTbl2(5)%></th>
								<th id="H2Col09" class="hlist" nowrap><%=HeaderTbl2(6)%></th>
								<th id="H2Col10" class="hlist" nowrap><%=HeaderTbl2(7)%></th>
								<th id="H2Col11" class="hlist" nowrap><%=HeaderTbl2(8)%></th>
								<th id="H2Col12" class="hlist" nowrap><%=HeaderTbl2(9)%></th>
								<th id="H2Col13" class="hlist" nowrap><%=HeaderTbl2(10)%></th>
								<th id="H2Col14" class="hlist" nowrap><%=HeaderTbl2(11)%></th>
								<th id="H2Col15" class="hlist" nowrap><%=HeaderTbl2(12)%></th>
								<th id="H2Col16" class="hlist" nowrap><%=HeaderTbl2(13)%></th>
								<th id="H2Col17" class="hlist" nowrap><%=HeaderTbl2(14)%></th>
								<th id="H2Col18" class="hlist" nowrap><%=HeaderTbl2(15)%></th>
								<th id="H2Col19" class="hlist" nowrap><%=HeaderTbl2(16)%></th>																																				
							</tr>
						    <!--HEADER INFORMATION END-->
						</thead>
						<TBODY>
						<TR>
							<TD nowrap style="width:50px;border-bottom:none"></TD>
							<TD nowrap style="width:0px;border-bottom:none;display:none;"></TD>
							<TD nowrap style="width:0px;border-bottom:none;display:none;"></TD>
							<TD nowrap style="width:140px;border-bottom:none"></TD>
							<TD nowrap style="width:140px;border-bottom:none"></TD>
							<TD nowrap style="width:60px;border-bottom:none"></TD>
							<TD nowrap style="width:80px;border-bottom:none"></TD>
							<TD nowrap style="width:60px;border-bottom:none"></TD>
							<TD nowrap style="width:100px;border-bottom:none"></TD>
							<TD nowrap style="width:60px;border-bottom:none"></TD>
							<TD nowrap style="width:130px;border-bottom:none"></TD>
							<TD nowrap style="width:40px;border-bottom:none"></TD>
							<TD nowrap style="width:80px;border-bottom:none"></TD>
							<TD nowrap style="width:100px;border-bottom:none"></TD>
							<TD nowrap style="width:100px;border-bottom:none"></TD>
							<TD nowrap style="width:90px;border-bottom:none"></TD>
							<TD nowrap style="width:90px;border-bottom:none"></TD>
							<TD nowrap style="width:90px;border-bottom:none"></TD>
							<TD nowrap style="width:90px;border-bottom:none"></TD>
						</TR>
						</TBODY>
                    </table>
			</div>
			<!-- Y.TAKAKUWA Add-E 2014-11-17 -->
			<div id="BDIV2" onscroll="OnScrollDiv(this)">
			   	<% If Num2>0 Then%>
			   	    <!--Y.TAKAKUWA 2013-03-01 Add-S-->
			   	    <% If blnSorted2 Then%>
					<!--<iframe frameborder="0" style="background-color:transparent;position:absolute;top:expression(this.offsetParent.scrollTop);left:0px;width:400;height:47px;"></iframe>				
					<% Else%>
					<!--<iframe frameborder="0" style="background-color:transparent;position:absolute;top:expression(this.offsetParent.scrollTop);left:0px;width:400;height:37px;"></iframe>									
					<% End If%>
			   	    <!--Y.TAKAKUWA 2013-03-01 Add-E-->
			   		<!--Work List Start-->	
					
						
                    <table border="1" cellpadding="0" cellspacing="0" width="100%" id="TBInOut">						
						<tbody>
						     <!--DETAIL INFORMATION START-->
						   	  
                              
                              <!------------------------------------------------------>
                              <% 
								x = 1 							
								For i=1 To ObjRS2.PageSize
								 	If Not ObjRS2.EOF Then
									x = x + 1
							  %>	
							  <% v_ItemName = "chkInOut" + cstr(i) %>
							  <% if Trim(ObjRS2("Type")) = "FULLOUT" then%>
								<tr bgcolor="#FFCCFF">
								<td id="D2Col01" align="center" style="width:50px;" nowrap valign="middle"><input type="checkbox" name="<%= v_ItemName %>"><BR></td>
							  <% else%>
								<tr bgcolor="#CCFFFF">
								<td id="D2Col01" align="center" style="width:50px;" nowrap valign="middle"><input type="checkbox" name="<%= v_ItemName %>"><BR></td>
							  <% end if%>	
							  
							  <td id="D2Col02" align="center" style="width:150px;display:none;" nowrap valign="middle" style="display:none;"></td>
							  <td id="D2Col03" align="center" style="width:130px;display:none;" nowrap valign="middle" style="display:none;"></td>	
							  <% v_ItemName = "InputID" + cstr(i) %>
							  <td id="D2Col04" nowrap align="center" valign="middle" style="width:140px;">
							  <% If Trim(ObjRS2("LoStatus")) = "2" Then %>
							    <input type="text" name="<%= v_ItemName %>" disabled="disabled" style="background-color:Transparent;">
							  <%Else%>
							    <%If Trim(ObjRS2("WorkDate")) = "" Then  %>
							      <input type="text" name="<%= v_ItemName %>" disabled="disabled" style="background-color:Transparent;">
							    <%Else%>
                                  <input type="text" name="<%= v_ItemName %>" onfocus="this.select();">
                                <%End if%>
                              <% End If %>
                              </td>
                              
                              
                              <%If v_DataCnt2 <> "" Then%>
							     <% If Arr_Err(i) <> "" Then %>
							       <td id="D2Col05" nowrap style="width:100px;" style="color:Red;width:140px;">
							       <%=Arr_ErrID(i)%></td>	
							     <%Else%>
							        <%If Trim(ObjRS2("GroupFlag")) = "0" then %>
							          <td id="D2Col05" nowrap style="color:Red;width:140px;">
							          <%=Trim(ObjRS2("LoDriverName"))%><BR /></td>
							        <%else%>
							          <td id="D2Col05" nowrap style="width:140px;">
							          <%=Trim(ObjRS2("LoDriverName"))%><BR /></td>
							        <%end if%>
							     <%End If%>
							  <%Else%>
							     <%If Trim(ObjRS2("GroupFlag")) = "0" then %>
							       <td id="D2Col05" nowrap style="color:Red;width:140px;">
							       <%=Trim(ObjRS2("LoDriverName"))%><BR /></td>
							     <%else%>
							       <td id="D2Col05" nowrap style="width:140px;">
							       <%=Trim(ObjRS2("LoDriverName"))%><BR /></td>
							     <%end if%>	
							  <%End If%>
							  
							  
                              										
							  <%
									If Trim(ObjRS2("TruckerSubCode4")) = COMPcd Then
										TruckerSubCode = 4
									ElseIf Trim(ObjRS2("TruckerSubCode3")) = COMPcd Then
										TruckerSubCode = 3
									ElseIf Trim(ObjRS2("TruckerSubCode2")) = COMPcd Then
										TruckerSubCode = 2
									ElseIf Trim(ObjRS2("TruckerSubCode1")) = COMPcd Then
										TruckerSubCode = 1
									Else
										TruckerSubCode = 0
									end if										
							  %>									
							  <INPUT type=hidden name='Datatbl<%=i%>' value='<%=Trim(ObjRS2("DeliverTo"))%>,<%=Mid(ObjRS2("WorkDate"),3,8)%>,<%=Trim(ObjRS2("Code1"))%>,<%=Trim(ObjRS2("WkNo"))%>,<%=Trim(ObjRS2("FullOutType"))%>,<%=Trim(ObjRS2("BLContNo"))%>,<%=Trim(Mid(ObjRS2("WorkCompleteDate"),3,14))%>,<%=Trim(ObjRS2("ReturnDateStr"))%>,<%=ObjRS2("ReturnValue")%>,<%=Trim(ObjRS2("Code2"))%>,<%=ObjRS2("Flag1")%>,<%=Trim(ObjRS2("BLNo"))%>,<%="-"%>,<%=TruckerSubCode%>,<%=Trim(ObjRS2("Flag2"))%>,<%=Trim(ObjRS2("ShipLine"))%>,<%=Trim(ObjRS2("ShipName"))%>,<%=Trim(ObjRS2("ContSize"))%>,<%=Trim(ObjRS2("CY"))%>,<%=Mid(ObjRS2("FreeTime"),3,8)%>,<%=TruckerSubCode%>,<%=Trim(ObjRS2("ReturnDateVal"))%>,<%=Trim(ObjRS2("Comment1"))%>,<%=Trim(ObjRS2("Comment2"))%>,<%=Trim(ObjRS2("DeliverTo1"))%>,<%=Left(Trim(ObjRS2("Name1")),8)%>'>
							  
							  
							    <%If v_DataCnt2 <> "" Then%>
							      <% If Arr_Err(i) <> "" Then %>
							        <td id="D2Col06" nowrap style="color:Red;width:60px;">
							        <%=Arr_Err(i)%>
							      <%Else%>
							        <%If Trim(ObjRS2("GroupFlag")) = "0" then %>
							        <td id="D2Col06" nowrap style="color:Red;width:60px;">
							        �o�^��
							        <%else%>
							        <td id="D2Col06" nowrap style="width:60px;">
							        <%=Trim(ObjRS2("LoHeadID"))%>
							        <%end if%>
							        <BR />
							      <%End If%>
							    <%Else%>
							        <%If Trim(ObjRS2("GroupFlag")) = "0" then %>
							        <td id="D2Col06" nowrap style="color:Red;width:60px;">
							        �o�^��
							        <%else%>
							        <td id="D2Col06" nowrap style="width:60px;">
							        <%=Trim(ObjRS2("LoHeadID"))%>
							        <%end if%>
							        <BR />
							    <%End If%>
							  </td>
							  
							  <%If Trim(ObjRS2("WorkDate")) = "" Then  %>
							    <td id="D2Col07" nowrap style="color:Red;width:80px;">������<BR /></td>
							  <%else%>
							    <td id="D2Col07" nowrap style="width:80px;"><%=Trim(ObjRS2("WorkDate"))%><BR /></td>
							  <%End If%>
							  <% if Trim(ObjRS2("Type")) = "FULLOUT" then%>						
							  <td id="D2Col08" nowrap style="width:60px;"><A HREF="JavaScript:GoRenew('<%=i%>');"><%=Trim(ObjRS2("WkNo"))%></A><BR></td>																
							  <% else %>
							  <td id="D2Col08" nowrap style="width:60px;"><A HREF="JavaScript:GoRenew2('<%=ObjRS2("WkNo")%>','<%=ObjRS2("BookNo")%>','<%=ObjRS2("ContNo")%>');"><%=Trim(ObjRS2("WkNo"))%></A><BR></td>																
							  <% end if %>
							  
							  <% if Trim(ObjRS2("Type")) = "FULLOUT" then%>						
							  <td id="D2Col09" nowrap style="width:100px"><A HREF="JavaScript:GoConinf2('<%=Trim(ObjRS2("WkNo"))%>','<%=Trim(ObjRS2("FullOutType"))%>','<%=Trim(ObjRS2("BLContNo"))%>')"><%=Trim(ObjRS2("BLContNo"))%></A><BR></td>																
							  <% else %>
							  <td id="D2Col09" nowrap style="width:100px"><A HREF="JavaScript:GoConinfContNo('<%=Trim(ObjRS2("BLContNo"))%>');"><%=Trim(ObjRS2("BLContNo"))%></A><BR></td>																
							  <% end if %>
							  <td id="D2Col10" nowrap style="width:60px"><%=Trim(ObjRS2("ShipLine"))%><BR></td>
							  <td id="D2Col11" nowrap style="width:130px"><%=Trim(ObjRS2("ShipName"))%><BR></td>
							  <td id="D2Col12" nowrap style="width:40px"><%=Trim(ObjRS2("ContSize"))%><BR></td>
							  <td id="D2Col13" nowrap style="width:80px"><%=Trim(ObjRS2("ReceiveFrom"))%><BR></td>
							  <td id="D2Col14" nowrap style="width:100px"><%=Trim(ObjRS2("CY"))%><BR></td>
							  <% if Trim(ObjRS2("DelPermitDate")) = "-" then%>
							    <td id="D2Col15" nowrap style="width:100px" align="center"><%=Trim(ObjRS2("DelPermitDate"))%><BR></td>
							  <% else%>
							    <td id="D2Col15" nowrap style="width:100px"><%=Trim(ObjRS2("DelPermitDate"))%><BR></td>
							  <% end if%>								
							  <td id="D2Col16" nowrap style="width:90px;"><%=Trim(ObjRS2("FreeTime"))%><BR></td>								
							  <td id="D2Col17" nowrap style="width:90px;"><%=Trim(ObjRS2("CYCut"))%><BR></td>
							  <td id="D2Col18" nowrap style="width:80px;"><%=Trim(ObjRS2("Code2"))%><BR></td>
							  <td id="D2Col19" nowrap style="width:80px;"><%=Trim(ObjRS2("Flag1"))%><BR></td>
							  
							  <% v_ItemName = "WkNo" + cstr(i) %>
							  <INPUT type=hidden name="<%=v_ItemName%>" value="<%=Trim(ObjRS2("WkNo"))%>">
							  <% v_ItemName = "ContNo" + cstr(i) %>
							  <INPUT type=hidden name="<%=v_ItemName%>" value="<%=Trim(ObjRS2("BLContNo"))%>">
							  <% v_ItemName = "BookNo" + cstr(i) %>
							  <INPUT type=hidden name="<%=v_ItemName%>" value="<%=Trim(ObjRS2("BookNo"))%>">
							  <% v_ItemName = "WorkControlNo" + cstr(i) %>
							  <INPUT type=hidden name="<%=v_ItemName%>" value="<%=Trim(ObjRS2("WkContrlNo"))%>">
							  <% v_ItemName = "LODriverID" + cstr(i) %>
							  <INPUT type=hidden name="<%=v_ItemName%>" value="<%=Trim(ObjRS2("LOID"))%>">
							  <% v_ItemName = "LoStatus" + cstr(i) %>
							  <INPUT type=hidden name="<%=v_ItemName%>" value="<%=Trim(ObjRS2("LoStatus"))%>">
							  <% v_ItemName = "ImpExp" + cstr(i) %>
							  <INPUT type=hidden name="<%=v_ItemName%>" value="<%=Trim(ObjRS2("ImpExp"))%>">
							  <% v_ItemName = "SeqNumber" + cstr(i) %>
							  <INPUT type=hidden name="<%=v_ItemName%>" value="<%=Trim(ObjRS2("SeqNumber"))%>">
							  <% v_ItemName = "WorkDate" + cstr(i) %>
							  <INPUT type=hidden name="<%=v_ItemName%>" value="<%=Trim(ObjRS2("WorkDate"))%>">
							  <% v_ItemName = "RecTerminalCode" + cstr(i) %>
							  <INPUT type=hidden name="<%=v_ItemName%>" value="<%=Trim(ObjRS2("RecTerminalCode"))%>">
							  <!--	
							  <% v_ItemName = "Flag2" + cstr(i) %>
							  <INPUT type=hidden name="<%=v_ItemName%>" value="<%=Trim(ObjRS2("Flag2"))%>">
							  <% v_ItemName = "TruckerSubCodeTbl2" + cstr(i) %>
							  <INPUT type=hidden name="<%=v_ItemName%>" value="<%=TruckerSubCode%>">																																
							  <% v_ItemName = "TypeCode" + cstr(i) %>
							  <INPUT type=hidden name="<%=v_ItemName%>" value="<%=Trim(ObjRS2("Type"))%>">
							  -->
							</tr>
						    <% 
									ObjRS2.MoveNext 		
									End If
								Next	
							  ObjRS2.close    
						      DisConnDBH ObjConn2, ObjRS2
						    %>  
						    <!--DETAIL INFORMATION END-->	    									
						</tbody>								
					</table>
					<!--Work List End-->
					<INPUT type=hidden name="DataCnt2" value="<%=x%>">
				<% Else %>
					<table border="1" cellPadding="2" cellSpacing="0">						
					  <TR class=bgw><TD nowrap>��ƈČ��͂���܂���</TD></TR>
					</table>
				<% End If %>		
			</div>
		</td>
		
		
		
		<!--Place Here End-->
	</tr>
	<tr><td>&nbsp;</td></tr>
	<tr>		
		<td>
		    <div style="margin-left:50px;">
			  <table border="0" cellpadding="2" cellspacing="0">
			  <tr>
			    <td><input type="button" name="Button"  value="���͂���ID�œo�^" onClick="LockOnReg()"></td>
			  </tr>
			  </table>	
			</div>		
		</td>
	</tr>
	<tr>
	  <td><HR/></td>
	</tr>
	<tr>		
		<td>
		  <div id="BDIV2">
		    <DIV style="width:100px; padding:10px;background-color:#FFCCFF; text-align:center;">�I�����ēo�^</DIV>
		      <Br />
		      <div style="margin-left:30px;">
		        <table border="0" cellpadding="2" cellspacing="0">
			    <tr>
			      <td> �\�őI��������Ƃɑ΂��A�h���C�o�A�܂��́A�O���[�v�Ŏw�����܂��B</td>
			    </tr>
			    <tr>
			      <td>
			        <div style="margin-left:20px;">
			          <table>
			            <tr>
			              <td><input type="radio" name="groupInfoChk1" value="1" onclick="fChange('groupInfoChk1','groupInfoChk2');"></td>
			              <td width="80">�h���C�o�w��</td>
			              <td>
			                <select style="width:420px; font-family:Monospace; overflow: scroll" size="1" name="driverInfo" onchange="fSelect('groupInfoChk1','groupInfoChk2','groupInfo');">
			                <%If Ubound(Arr_DriverInfoVal) > 0 then %>
			                  <%For i=0 to CInt(Ubound(Arr_DriverInfoVal)) %>
			                     <option value="<%=Arr_DriverInfoVal(i)%>">
			                       <%= Arr_DriverInfoText(i)%> 
			                     </option>
			                  <%Next%>
			                <%else%>
                              <option value=""></option>
                            <%end if%>
                            </select>
			              </td>
			              
			            </tr>
                      </table>
                    </div>
			      </td>
			    </tr>
			    <tr>
			      <td>
			        <div style="margin-left:20px;">
			          <table>
			          <tr>
			            <td><input type="radio" name="groupInfoChk2" value="2" onclick="fChange('groupInfoChk2','groupInfoChk1');"></td>
			            <td width="80">�O���[�v�w��</td>
			            <td>
			              <select style="width:420px; font-family:Monospace;" name="groupInfo" onchange="fSelect('groupInfoChk2','groupInfoChk1','driverInfo');">
			              <option value=""></option>
			              <%For i = 0 to UBound(Arr_GroupID) %>
			                <option value="<%=MID(Arr_GroupID(i),1,2)%>"><%=Arr_GroupID(i)%></option>
			              <%Next%>
                          </select>
			            </td>
			          </tr>
                      </table>
                    </div>
			      </td>
			    </tr>
			    <tr></tr>
			    </table>
			 </div>
		  </div>
		</td>
	</tr>
	<tr>
		<td>
		  <div style="margin-left:50px;">
			<table border="0" cellpadding="2" cellspacing="0">
			<tr>
			<td><input type="button" name="Button"  value="�o�^" onClick="fGrpUpd()"></td>
			</tr>
			</table>
		  </div>
		</td>	
	</tr>
	<tr>
	  <td><HR/></td>
	</tr>
	<tr>
		<td>
		  <DIV style="width:100px; padding:10px;background-color:#FFCCFF; text-align:center;">�I�����ĉ���</DIV>
		  <Br />
		  <div style="margin-left:30px;">
		    <table border="0" cellpadding="2" cellspacing="0">
			<tr>
			<td> �\�őI��������Ƃɑ΂��A�R���e�i���b�N�w����h���C�o�̃R���e�i���b�N���������܂��B</td>
			</tr>
			</table>
		  </div>
		</td>
	</tr>
	<tr>
		<td>
		  <div style="margin-left:50px;">
			<table border="0" cellpadding="2" cellspacing="0">
			<tr>
			<td><input type="button" name="Button"  value="����" onClick="fRelease()"></td>
			</tr>
			</table>
		  </div>
		</td>
	</tr>	
</table>
<!--Added End-->
<br />
</div>
</form>


</BODY>
</HTML>