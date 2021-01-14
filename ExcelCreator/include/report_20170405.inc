<%
	'---------------------------------------------------------
	'���[�̋��ʌŒ�֐�
	'---------------------------------------------------------
	Const gOutFileForder  = "/ExcelCreator/outfiles/"
	Const gTemplateFolder = "/ExcelCreator/excel/"

	'---------------------------------------------------------
	'���[�̋��ʊ֐�
	'---------------------------------------------------------
	Dim wDate
	Dim wTime
	Dim wOutFileName	'���[�̃t�@�C����
	Dim wReportID       '���[ID
	Dim wReportName     '���[��
	Dim wFilePath		'�t�@�C���p�[�X
	Dim wIISFilePath	'IIS�t�@�C���p�[�X
	Dim wInFileName		'�e���v���[�g�̃t�@�C����
	Dim wMsg			'���b�Z�[�W
	Dim wErrNo          '�G���[�ԍ�
	
	'---------------------------------------------------------
	' �o�̓t�@�C�����̓��t�쐬
	'---------------------------------------------------------	
	wDate = gfformatIntDate(Date)
	wTime = FormatDateTime (Time,vbLongTime)
		
	'---------------------------------------------------------
	' �ȉ���wFilePath�AwIISFilePath�AwInFileName�̒l�́A
    ' ���s���̍\���ɏ]���ύX���ĉ�����
	'---------------------------------------------------------
	wIISFilePath = "http://" & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT") & "/ExcelCreator/outfiles/"
	wFilePath  =  Server.Mappath(gOutFileForder)
	
'-----------------------------
'   ���t�ҏW�iyyyymmdd�j
'-----------------------------
Function gfformatIntDate(d)
    dim w
    if isnull(d) then
        exit function
    end if
    gfformatIntDate=""
    w=""
    if isdate(d) then
        w=year(d)*10000 + month(d)*100 + day(d)
    elseif len(d)=10 then
        w=""
        gfformatDate=d
    elseif len(d)=8 then
        w=gfTrim(d)
    elseif len(d)=14 then
        w=left(d, 8)
    end if
    gfformatIntDate= w
End Function	

'---------------------------------------------------------
'�����[(1�����R�[�h�j
'---------------------------------------------------------
Public Function gfReceiveReport(lBookNo,lWorkNo,lContNo)
	Dim sql
	Dim rs,cn
	Dim Xls1
	Dim flname
	Dim nAttrNo				'2016/08/19 H.Yoshikawa Add

	
	On error resume next
	
	err.clear
	'2016/11/29 H.Yoshikawa Upd Start
	'flname = wReportName & wDate & Replace(wTime,":","") & ".xls"
	flname = wReportName & Trim(lBookNo) & "_" & Trim(lContNo) & ".xls"
	'2016/11/29 H.Yoshikawa Upd End
	wInFileName = Server.Mappath(gTemplateFolder & wReportID & ".xls")
    ConnDBH cn, rs
	
	'2016/08/18 H.Yoshikawa Add Start
	'�o�^�S���ҁA���l�擾����
	Dim TruckerName, Comment1,Comment2,Comment3
	sql = "SELECT ITC.WkContrlNo, ITC.RegisterCode, ITC.TruckerSubCode1, ITC.TruckerSubCode2, "&_
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
          "WHERE ITC.ContNo='"&lContNo&"' AND ITC.WkNo='"& lWorkNo &"' AND ITC.WkType='3' AND ITC.Process='R'"
	rs.Open sql, cn
	if not rs.EOF then 
		Select Case Trim(Session.Contents("userid"))
			Case Trim(rs("RegisterCode"))
				TruckerName = Trim(rs("TruckerSubName1"))
			Case Trim(rs("Trucked1"))
				TruckerName = Trim(rs("TruckerSubName2"))
			Case Trim(rs("Trucked2"))
				TruckerName = Trim(rs("TruckerSubName3"))
			Case Trim(rs("Trucked3"))
				TruckerName = Trim(rs("TruckerSubName4"))
			Case Trim(rs("Trucked4"))
				TruckerName = Trim(rs("TruckerSubName5"))
		End Select 
		Comment1 = Trim(rs("Comment1"))
		Comment2 = Trim(rs("Comment2"))
		Comment3 = Trim(rs("Comment3"))
	end if
	rs.Close
	'2016/08/18 H.Yoshikawa Add End
	
	'---------------------------------------------------------
	'SQL���̍쐬
	'---------------------------------------------------------
	sql = gfCreateReportSql(lBookNo,lWorkNo,lContNo)
	'response.write sql
	'response.end
	rs.Open sql, cn
	
	if not rs.EOF then 
		'2016/01/20 T.Okui Add Start
		'�ݒ艷�x���擾
		dim SetTemp,ContType,Sender,VoyCtrl,VslCode
		
		Sender = Trim(rs("Sender"))
		VoyCtrl = Trim(rs("VoyCtrl"))
		VslCode = Trim(rs("VslCode"))
		SetTemp = gfGetSetTemp(Sender,lBookNo,lContNo,VoyCtrl,VslCode)
		'SetTemp = ""
		'2016/01/20 T.Okui Add End
		'2017/02/02 T.Okui Add Start
		
		ContType = gfGetContType(Sender,lContNo)

		'2017/02/02 T.Okui Add End
		'--------------------------------------------------------
		'  ExcelCreator �I�u�W�F�N�g������Excel�t�@�C���o��
		'--------------------------------------------------------        
		 Set Xls1= Server.CreateObject("ExcelCrtOcx.ExcelCrtOcx.1")  

		'---------------------------------------------------------
		'�����[(�I�[�o�[���C)�t�@�C���I�[�v��
		'---------------------------------------------------------
  		Xls1.OpenBook wFilePath & "\" & flname ,wInFileName

		'���^�V�[�g���Ăяo��
		 Xls1.SheetNo = 0
		Xls1.SheetName=Trim(rs("ContNo"))
		'�u���E�U��œ��͂����f�[�^���V�[�g�ɏo��
		Xls1.Cell("**LineCode").Str  = Trim(rs("ShippingLineName")) '�戵�D��
		Xls1.Cell("**VslName").Str   = Trim(rs("VslName"))			'�D��
		Xls1.Cell("**ContNo").Str    = Trim(rs("ContNo"))			'�R���e�i�ԍ�
		Xls1.Cell("**SealNo").Str    = "'" & Trim(rs("SealNo"))			'�Z�[���ԍ�
		Xls1.Cell("**DPort").Str     = Trim(rs("PortName"))			'�g�`��
		Xls1.Cell("**ContSize").Str  = Trim(rs("ContSize"))			'�R���e�i�T�C�Y
		Xls1.Cell("**ContType").Str  = Trim(rs("ContType"))			'�R���e�i�^�C�v
		Xls1.Cell("**ContHeight").Str= Trim(rs("ContHeight"))		'�R���e�i����
		
		'if CLng(rs("OvHeight")) > 0 then							'�n�g
			'Xls1.Cell("**OH").Str= "��"
			Xls1.Cell("**OH").Str= CLng(rs("OvHeight"))			
		'end if
		'if CLng(rs("OvWidthL")) > 0 then							'�n�v�k
			'Xls1.Cell("**OWL").Str= "��"
			Xls1.Cell("**OWL").Str=CLng(rs("OvWidthL"))
		'end if
		'if CLng(rs("OvWidthR")) > 0 then							'�n�v�q
			'Xls1.Cell("**OWR").Str= "��"
			Xls1.Cell("**OWR").Str= CLng(rs("OvWidthR"))
		'end if
		'if CLng(rs("OvLengthF")) > 0 then							'�n�k�e
			'Xls1.Cell("**OLF").Str= "��"
			Xls1.Cell("**OLF").Str= CLng(rs("OvLengthF"))							
		'end if
		'if CLng(rs("OvLengthA")) > 0 then							'�n�k�`
			'Xls1.Cell("**OLA").Str= "��"
			Xls1.Cell("**OLA").Str= CLng(rs("OvLengthA"))
		'end if

		if Trim(rs("CustClear")) = "Y" then 
			Xls1.Cell("**CusOK").Str= "��"							'�ʊ�
		elseif Trim(rs("CustClear")) = "N" then 
			Xls1.Cell("**CusNo").Str= "��"							'�ʊ�
		end if
		
		Xls1.Cell("**WorkNo").Str= Trim(rs("WkNo"))					'��Ɣԍ�					
		Xls1.Cell("**BookingNo").Str= Trim(rs("BookNo"))			'Booking�ԍ�
		Xls1.Cell("**GrossWT").Value= Trim(rs("ContWeight"))			'���d��
		Xls1.Cell("**Voyage").Str= Trim(rs("Voyage"))				'Voyage
		Xls1.Cell("**Fdest").Str= Trim(rs("FDest"))					'�ŏI�ړI�n
		'2010/05/10 Upd-S Tanaka
		'Xls1.Cell("**TareWT").Str= Trim(rs("TareWeight"))		    'TARE�E�F�C�g
		if Trim(rs("TareWeight"))<100 Then
			Xls1.Cell("**TareWT").Value= Trim(rs("TareWeight"))*100		    'TARE�E�F�C�g
		Else
			Xls1.Cell("**TareWT").Value= Trim(rs("TareWeight"))		    'TARE�E�F�C�g
		End If
		'2010/05/10 Upd-E Tanaka
		Xls1.Cell("**NetWT").Value= CInt(Trim(rs("ContWeight"))) - CInt(Trim(rs("TareWeight")))			'2016/10/28 H.Yoshikawa Add
		
		'2017/01/20 T.Okui Upd Start  KACCS���̃f�[�^���g�p����悤�ɕύX
		'Xls1.Cell("**Temp").Str= Trim(rs("SetTemp"))	    		'�Ⓚ���x
		'if Trim(rs("AsDry")) = "1" then
		'	Xls1.Cell("**Temp").Str = ""							'�Ⓚ���x
		'	Xls1.Cell("**AsDry").Str = "DRY USE"						'As Dry
		'else
		'	Xls1.Cell("**Temp").Str= Mid(Trim(rs("SetTemp")),1,5)	'�Ⓚ���x
		'	Xls1.Cell("**AsDry").Str = ""							'As Dry
		'end if

		if  Trim(ContType) = "RF"  then
			if Trim(SetTemp) <> "NO" and Trim(SetTemp) <> "" then
				Xls1.Cell("**Temp").Str = Mid(Trim(SetTemp),1,5)		'�Ⓚ���x
			else
				Xls1.Cell("**Temp").Str = ""
			end if
			
			'As Dry
			if Trim(SetTemp) <> "NO" and Trim(SetTemp) = "" then
				Xls1.Cell("**AsDry").Str = "DRY USE"		
			else
				Xls1.Cell("**AsDry").Str = ""
			end if
			
			if Mid(Trim(SetTemp),6,1) = "C" then
				Xls1.Cell("**TempDegree1").Str= "�K"
				Xls1.Cell("**TempDegree2").Str= "C"							'��C
			elseif Mid(Trim(SetTemp),6,1) = "F" then
				Xls1.Cell("**TempDegree1").Str= "�K"
				Xls1.Cell("**TempDegree2").Str= "F"							'��F
			
			end if				
		end if
		'2017/01/20 T.Okui Upd End
		
		'2016/08/18 H.Yoshikawa Del Start
		'if UCASE(Trim(rs("TempDegree")))="C" Then 				    
		'	Xls1.Cell("**Cel").Str= "��"
		'elseif UCASE(Trim(rs("TempDegree")))="F" Then 
		'	Xls1.Cell("**Far").Str= "��"
		'end if
		'2016/08/18 H.Yoshikawa Del End
		
		if Trim(rs("Ventilation")) <> "" Then						'�ʕ��E
			Xls1.Cell("**Vent").Str= Trim(rs("Ventilation"))					
		end if

		if Trim(rs("IMDG1")) <> "" then 							'IMCO
			Xls1.Cell("**IMCOA").Str= Trim(rs("IMDG1"))
		end if
		if Trim(rs("IMDG2")) <> "" then
			Xls1.Cell("**IMCOB").Str= Trim(rs("IMDG2"))
		end if
		if Trim(rs("IMDG3")) <> "" then 
			Xls1.Cell("**IMCOC").Str= Trim(rs("IMDG3"))
		end if

		if Trim(rs("UNNo1")) <> "" then								'UN 
			Xls1.Cell("**UNA").Str= Trim(rs("UNNo1"))
		end if
		if Trim(rs("UNNo2")) <> "" then
			Xls1.Cell("**UNB").Str= Trim(rs("UNNo2"))
		end if
		if Trim(rs("UNNo3")) <> "" then
			Xls1.Cell("**UNC").Str= Trim(rs("UNNo3"))
		end if

		'2016/08/18 H.Yoshikawa Add Start
		if Trim(rs("IMDG4")) <> "" then 							'IMCO
			Xls1.Cell("**IMCOD").Str= Trim(rs("IMDG4"))
		end if
		if Trim(rs("IMDG5")) <> "" then
			Xls1.Cell("**IMCOE").Str= Trim(rs("IMDG5"))
		end if

		if Trim(rs("UNNo4")) <> "" then
			Xls1.Cell("**UND").Str= Trim(rs("UNNo4"))
		end if
		if Trim(rs("UNNo5")) <> "" then
			Xls1.Cell("**UNE").Str= Trim(rs("UNNo5"))
		end if
		
		if Trim(rs("LqFlag1")) = "1" then
			Xls1.Cell("**LQA").Str= "��"
		end if
		if Trim(rs("LqFlag2")) = "1" then
			Xls1.Cell("**LQB").Str= "��"
		end if
		if Trim(rs("LqFlag3")) = "1" then
			Xls1.Cell("**LQC").Str= "��"
		end if
		if Trim(rs("LqFlag4")) = "1" then
			Xls1.Cell("**LQD").Str= "��"
		end if
		if Trim(rs("LqFlag5")) = "1" then
			Xls1.Cell("**LQE").Str= "��"
		end if
		
		'2016/10/20 H.Yoshikawa Add Start
		Xls1.Cell("**LBLA").Str= Trim(rs("Label1"))
		Xls1.Cell("**LBLB").Str= Trim(rs("Label2"))
		Xls1.Cell("**LBLC").Str= Trim(rs("Label3"))
		Xls1.Cell("**LBLD").Str= Trim(rs("Label4"))
		Xls1.Cell("**LBLE").Str= Trim(rs("Label5"))
		Xls1.Cell("**SLBLA").Str= Trim(rs("SubLabel1"))
		Xls1.Cell("**SLBLB").Str= Trim(rs("SubLabel2"))
		Xls1.Cell("**SLBLC").Str= Trim(rs("SubLabel3"))
		Xls1.Cell("**SLBLD").Str= Trim(rs("SubLabel4"))
		Xls1.Cell("**SLBLE").Str= Trim(rs("SubLabel5"))
		'2016/10/20 H.Yoshikawa Add End
		
		Xls1.Cell("**decider").Str = ""										'�m���						'2016/10/20 H.Yoshikawa Add�i�b��I�ɋ�l�j
		Xls1.Cell("**ReportNo").Str= Trim(rs("ReportNo"))					'�͏o�ԍ��܂��͓o�^�ԍ�
		Xls1.Cell("**Shipper").Str= Trim(rs("PRShipper"))					'�׎喼
		Xls1.Cell("**PRFwd").Str= Trim(rs("PRForwarder"))					'�C�݋Ǝ҉�Ж�
		Xls1.Cell("**PRFwdTan").Str= Trim(rs("PRForwarderTan"))				'�C�݋ƎҒS���Җ�
		Xls1.Cell("**PRFwdTel").Str= Trim(rs("PRForwarderTel"))				'�C�݋ƎҘA����
		'2016/10/20 H.Yoshikawa Del Start
		'Xls1.Cell("**Comment1").Str= Comment1								'���l1
		'Xls1.Cell("**Comment2").Str= Comment2								'���l2
		'Xls1.Cell("**Comment3").Str= Comment3								'���l3
		'2016/10/20 H.Yoshikawa Del End

		nAttrNo = Xls1.Cell("**OutDate").AttrNo
		Xls1.Cell("**OutDate").Value2 CDate(rs("UpdtTime")), nAttrNo									'�o�^���i̫�ϯĂ̑����������p���j
		Xls1.Cell("**OutDate").PosHorz = phEven													'�����ʒu�F�ϓ���
		'2016/11/22 H.Yoshikawa Upd Start
		'Xls1.Cell("**Signature").Str= Trim(Session.Contents("LinUN")) & "�@" & TruckerName			'���O�C�����[�U�� + �o�^�S���Җ�
		'Xls1.Cell("**ContactInfo").Str= Trim(rs("ContactInfo"))									'�o�^�ҘA����
		Xls1.Cell("**Signature").Str= Trim(rs("PRForwarder")) & "�@" & Trim(rs("PRForwarderTan"))	'�C�݋Ǝ҉�Ж� + �C�݋ƎҒS���Җ�
		Xls1.Cell("**ContactInfo").Str= Trim(rs("PRForwarderTel"))										'�C�݋ƎҘA����
		'2016/11/22 H.Yoshikawa Upd Start

		'2016/08/18 H.Yoshikawa Add Start
		
		wMsg = "Excel�t�@�C�����쐬���܂����B�ȉ����쐬�����t�@�C�����_�E�����[�h�ł��܂�"

		if Err.Number<>0 then 
			wMsg=Err.Description
		end if

		wErrNo = Xls1.ErrorNo
		If wErrNo <> 0 Then
			wMsg = "ExcelCreator3�G���[���b�Z�[�W�F" & Xls1.ErrorMessage
		End If
		Xls1.CloseBook

		Set Xls1 = Nothing		
		gfReceiveReport=flname
	else
		gfReceiveReport=""
	end if
	'RecordSet�ڑ�����
	rs.Close
	'DB�ڑ�����
	DisConnDBH cn, rs

End Function 
'---------------------------------------------------------
'�����[(1�����R�[�h�j
'---------------------------------------------------------
Public Function gfCreateReportSql(lBookNo,lWorkNo,lContNo)
	dim sql

	'---------------------------------------------------------
	'SQL���̍쐬F
	'---------------------------------------------------------
	'2016/08/18 H.Yoshikawa Upd Start
	'sql = "SELECT mS.FullName2 As ShippingLineName, CYV.VslName, " &_
	'      "  CYV.ContNo,IsNull(CYV.SealNo,'') as SealNo,CYV.ContSize,CYV.ContHeight,CYV.ContType, " &_
	'      "  CYV.CustClear,CYV.WkNo, CYV.OvHeight, CYV.OvWidthL, CYV.OvWidthR, CYV.OvLengthF, CYV.OvLengthA," &_
	'      "  isNull(mULP.FullName,'') AS PortName, " &_
	'	  "  isNull(mULP2.FullName,'') AS FDest, " &_
	'	  "  CYV.BookNo,CYV.Voyage,CYV.ContWeight,CYV.TareWeight, "&_
	'	  "  CASE WHEN EXC.SetTemp IS NOT NULL AND RTRIM(EXC.SetTemp)<>'' THEN SUBSTRING(EXC.SetTemp,1,LEN(EXC.SetTemp)-1) "&_
	'	  "       ELSE '' "&_
	'	  "  END  SetTemp, "&_
	'	  "  CASE WHEN EXC.SetTemp IS NOT NULL AND RTRIM(EXC.SetTemp)<>'' THEN SUBSTRING(EXC.SetTemp,LEN(EXC.SetTemp),1) "&_
	'	  "       ELSE '' "&_
	'     "  END  TempDegree, "&_ 
	'	  "EXC.Ventilation, EXC.IMDG1, EXC.IMDG2, EXC.IMDG3, EXC.UNNo1, EXC.UNNo2,EXC.UNNo3 "&_
	'      "FROM CYVanInfo AS CYV " &_ 
	'      "LEFT JOIN ExportCont AS EXC ON CYV.ContNo = EXC.ContNo AND CYV.BookNo = EXC.BookNo " &_
	'      "LEFT JOIN Booking AS BOK ON EXC.VslCode = BOK.VslCode AND EXC.VoyCtrl = BOK.VoyCtrl AND EXC.BookNo = BOK.BookNo " &_
	'      "LEFT JOIN mShipLine mS ON CYV.ShipLine=mS.ShipLine " &_
	'      "LEFT JOIN mPort AS mULP ON CYV.DPort = mULP.PortCode " &_
	'      "LEFT JOIN mPort AS mULP2 ON CYV.DelivPlace = mULP2.PortCode " &_
	'	  "WHERE CYV.BookNo='" & lBookNo & "' " &_
	'      "  AND CYV.ContNo='" & lContNo & "' " &_
	'      "  AND CYV.WkNo='" & lWorkNo & "' "
	'2017/01/20 T.Okui Upd Start
	sql = "SELECT mS.FullName2 As ShippingLineName, CYV.VslName, CYV.UpdtTime, " &_
	      "  CYV.ContNo,IsNull(CYV.SealNo,'') as SealNo,CYV.ContSize,CYV.ContHeight,isnull(mTP.FullName, CYV.ContType) AS ContType, " &_
	      "  CYV.CustClear,CYV.WkNo, CYV.OvHeight, CYV.OvWidthL, CYV.OvWidthR, CYV.OvLengthF, CYV.OvLengthA," &_
	      "  isNull(mULP.FullName,'') AS PortName, " &_
		  "  isNull(mULP2.FullName,'') AS FDest, " &_
		  "  CYV.BookNo,CYV.Voyage,CYV.ContWeight,CYV.TareWeight, CYV.AsDry, "&_
		  "  CYV.SetTemp, CYV.Ventilation, CYV.ReportNo, CYV.ContactInfo, CYV.PRShipper, "&_
		  "  CYV.PRForwarder, CYV.PRForwarderTan, CYV.PRForwarderTel, "&_
		  "  CYV.IMDG1, CYV.IMDG2, CYV.IMDG3, CYV.IMDG4, CYV.IMDG5, "&_
		  "  CYV.UNNo1, CYV.UNNo2, CYV.UNNo3, CYV.UNNo4, CYV.UNNo5, "&_
		  "  CYV.Label1, CYV.Label2, CYV.Label3, CYV.Label4, CYV.Label5, "&_
		  "  CYV.SubLabel1, CYV.SubLabel2, CYV.SubLabel3, CYV.SubLabel4, CYV.SubLabel5, "&_
		  "  CYV.LqFlag1, CYV.LqFlag2, CYV.LqFlag3, CYV.LqFlag4, CYV.LqFlag5,BOK.Sender,BOK.VoyCtrl,BOK.VslCode "&_  
	      "FROM CYVanInfo AS CYV " &_ 
	      "LEFT JOIN ExportCont AS EXC ON CYV.ContNo = EXC.ContNo AND CYV.BookNo = EXC.BookNo " &_
	      "LEFT JOIN Booking AS BOK ON EXC.VslCode = BOK.VslCode AND EXC.VoyCtrl = BOK.VoyCtrl AND EXC.BookNo = BOK.BookNo " &_
	      "LEFT JOIN mShipLine mS ON CYV.ShipLine=mS.ShipLine " &_
	      "LEFT JOIN mPort AS mULP ON CYV.DPort = mULP.PortCode " &_
	      "LEFT JOIN mPort AS mULP2 ON CYV.DelivPlace = mULP2.PortCode " &_
	      "LEFT JOIN mContType AS mTP ON CYV.ContType = mTP.ContType " &_
		  "WHERE CYV.BookNo='" & lBookNo & "' " &_
	      "  AND CYV.ContNo='" & lContNo & "' " &_
	      "  AND CYV.WkNo='" & lWorkNo & "' "
	'2017/01/20 T.Okui Upd End
	'2016/08/18 H.Yoshikawa Upd Start
	
	gfCreateReportSql=sql
End Function

'2017/01/20 T.Okui Add Start
'---------------------------------------------------------
'VIEW:oBookCont����ݒ艷�x�擾
'---------------------------------------------------------
Public Function gfGetSetTemp(Sender,lBookNo,lContNo,VoyCtrl,VslCode)
	'Sender����I�y�ЃR�[�h���擾
	dim Operator,SetTemp
    Operator = ""
    SetTemp = ""
    
	select case gfTrim(Sender)
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
		case else
			Operator = ""
	end select
		
	if Operator = "" then
		gfGetSetTemp = ""
		exit function
    end if
	dim rs,sql,cn
	dim ExVoy
	ConnDBH cn,rs
	'VoyCtrl,VslCode����ExVoyage���擾
	sql = "SELECT LdVoyage FROM VslSchedule "
	sql = sql & " WHERE VslCode='"& VslCode &"' "
	sql = sql & "   AND VoyCtrl='"& VoyCtrl &"' "
	
	rs.Open sql, cn
	if not rs.EOF then
		ExVoy   = gfTrim(rs("LdVoyage"))
	else
		gfGetSetTemp = ""
		rs.Close
		exit function
	end if
	rs.Close
	
	dim rs2
    ConnDBH cn, rs2

	'oBookCont����ݒ艷�x�擾
	sql = "SELECT ob.SetTemp FROM " & Trim(Operator) & "_oBookCont ob "
	sql = sql & " INNER JOIN KAC_kMVessel mv on mv.VslCode = ob.VslCode "
	sql = sql & " INNER JOIN KAC_kSchedule kc on kc.VslCode = ob.VslCode and kc.Voyage = ob.Voyage "
	sql = sql & " WHERE ob.BookNo='"& lBookNo &"' "
	sql = sql & "   AND mv.CallSign='"& VslCode &"' "
	sql = sql & "   AND kc.ExVoyage='"& ExVoy &"' "
	sql = sql & "   AND ob.ContNo='"& lContNo &"' "
	rs2.Open sql, cn
		
	if not rs2.EOF then
		SetTemp    = gfTrim(rs2("SetTemp"))
	else
		SetTemp    = "NO"
	end if

	rs2.Close
	
	gfGetSetTemp=SetTemp
End Function
'2017/01/20 T.Okui Add End

'2017/02/02 T.Okui Add Start
'---------------------------------------------------------
'VIEW:oContainer����R���e�i�^�C�v�擾
'---------------------------------------------------------
Public Function gfGetContType(Sender,lContNo)
	'Sender����I�y�ЃR�[�h���擾
	dim Operator,ContType
    Operator = ""
    ContType = ""
    
	select case gfTrim(Sender)
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
		case else
			Operator = ""
	end select


	if Operator = "" then
		gfGetContType = ""
		exit function
    end if
	
	dim cn,rs,sql
    ConnDBH cn, rs

	'oContainer����R���e�i�^�C�v�擾
	sql = "SELECT oc.ContType FROM " & Trim(Operator) & "_oContainer oc "
	sql = sql & " WHERE oc.ContNo='"& lContNo &"' "
	
	rs.Open sql, cn

	if not rs.EOF then
		ContType    = gfTrim(rs("ContType"))
	end if

	rs.Close
	
	gfGetContType=ContType
End Function
'2017/02/02 T.Okui Add End
'---------------------------------------------------------
'�����[(�������R�[�h�j
'---------------------------------------------------------
Public Function gfReceiveReportMultiple()
	 Dim rs
	 Dim sql
	 Dim cn
	 Dim fld,flname,Xls1
	 Dim arySlot()
	 Dim cnt
	 Dim bCopyReturn,bDelReturn
	 Dim iSheetNo
	 Dim nAttrNo										'2016/08/19 H.Yoshikawa Add
	 Dim TruckerName, Comment1,Comment2,Comment3		'2016/08/19 H.Yoshikawa Add

 	 On error resume next
	 
	 err.clear
	'DB�ڑ�
	 ConnDBH cn,rs
	 With Request
		 cnt=1
		 Redim Preserve	arySlot(cnt)
		 '�I�񂾏����m�F����
		 For  Each fld in .Form						
			If Mid(fld,1,8) = "chkInOut" Then 																
				If .Form(fld) = "on" Then
					Redim Preserve	arySlot(cnt+1)
					arySlot(cnt)=Mid(fld,9)									
					cnt = cnt + 1
				End if				
			End If			
		 Next

		 if Ubound(arySlot) > 0 then
			
			'�t�@�C�������쐬����
			flname = wReportName & wDate & Replace(wTime,":","") & ".xls"
			wInFileName = Server.Mappath(gTemplateFolder & wReportID & ".xls")

		 
			'--------------------------------------------------------        
			'  ExcelCreator �I�u�W�F�N�g������Excel�t�@�C���o��
			'--------------------------------------------------------        
			Set Xls1= Server.CreateObject("ExcelCrtOcx.ExcelCrtOcx.1")  

			'---------------------------------------------------------
			'�����[(�I�[�o�[���C)�t�@�C���I�[�v��
			'---------------------------------------------------------
			Xls1.OpenBook wFilePath & "\" & flname ,wInFileName

			For cnt=1 to Ubound(arySlot) - 1

				'layout���R�s�[����
				iSheetNo = Xls1.SheetCount
				bCopyReturn = Xls1.CopySheet(0,iSheetNo,"LineCode" & iSheetNo)

				'���^�V�[�g���Ăяo��
				Xls1.SheetNo = Xls1.SheetCount-1
								
				'2016/08/18 H.Yoshikawa Add Start
				'�o�^�S���ҁA���l�擾����
				sql = "SELECT ITC.WkContrlNo, ITC.RegisterCode, ITC.TruckerSubCode1, ITC.TruckerSubCode2, "&_
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
			          "WHERE ITC.ContNo='"&Trim(.Form("ContNo" & Cstr(arySlot(cnt))))&"' AND ITC.WkNo='"& Trim(.Form("WkNo" & Cstr(arySlot(cnt)))) &"' AND ITC.WkType='3' AND ITC.Process='R'"
				rs.Open sql, cn
				if not rs.EOF then 
					Select Case Trim(Session.Contents("userid"))
						Case Trim(rs("RegisterCode"))
							TruckerName = Trim(rs("TruckerSubName1"))
						Case Trim(rs("Trucked1"))
							TruckerName = Trim(rs("TruckerSubName2"))
						Case Trim(rs("Trucked2"))
							TruckerName = Trim(rs("TruckerSubName3"))
						Case Trim(rs("Trucked3"))
							TruckerName = Trim(rs("TruckerSubName4"))
						Case Trim(rs("Trucked4"))
							TruckerName = Trim(rs("TruckerSubName5"))
					End Select 
					Comment1 = Trim(rs("Comment1"))
					Comment2 = Trim(rs("Comment2"))
					Comment3 = Trim(rs("Comment3"))
				end if
				rs.Close
				'2016/08/18 H.Yoshikawa Add End

				'�f�[�^���擾����
				sql = gfCreateReportSql(.Form("BookNo" & Cstr(arySlot(cnt))),.Form("WkNo" & Cstr(arySlot(cnt))),.Form("ContNo" & Cstr(arySlot(cnt))))
				
				rs.Open sql,cn
				if not rs.EOF then 
					'2016/01/20 T.Okui Add Start
					'�ݒ艷�x�A�R���e�i�^�C�v���擾
					dim SetTemp,ContType,Sender,VoyCtrl,VslCode
					
					Sender = Trim(rs("Sender"))
					VoyCtrl = Trim(rs("VoyCtrl"))
					VslCode = Trim(rs("VslCode"))
					'response.write Sender & "<br>"
					'response.write .Form("BookNo" & Cstr(arySlot(cnt)))& "<br>"
					'response.write .Form("ContNo" & Cstr(arySlot(cnt)))& "<br>"
					'response.write VoyCtrl& "<br>"
					'response.write VslCode& "<br>"
					'response.end
					
					SetTemp = gfGetSetTemp(Sender,.Form("BookNo" & Cstr(arySlot(cnt))),.Form("ContNo" & Cstr(arySlot(cnt))),VoyCtrl,VslCode)
					'SetTemp = ""
					'2017/02/02 T.Okui Add Start

					ContType = gfGetContType(Sender,.Form("ContNo" & Cstr(arySlot(cnt))))
					'2017/02/02 T.Okui Add End
										
					'2016/01/20 T.Okui Add End
					
					'bCopyReturn = Xls1.CopySheet(0,iSheetNo,Trim(rs("ContNo")))
					Xls1.SheetName=Trim(rs("ContNo"))
					'�u���E�U��œ��͂����f�[�^���V�[�g�ɏo��
					Xls1.Cell("**LineCode").Str  = Trim(rs("ShippingLineName")) '�戵�D��
					Xls1.Cell("**VslName").Str   = Trim(rs("VslName"))			'�D��
					Xls1.Cell("**ContNo").Str    = Trim(rs("ContNo"))			'�R���e�i�ԍ�
					Xls1.Cell("**SealNo").Str    = "'" & Trim(rs("SealNo"))		'�Z�[���ԍ�
					Xls1.Cell("**DPort").Str     = Trim(rs("PortName"))			'�g�`��
					Xls1.Cell("**ContSize").Str  = Trim(rs("ContSize"))			'�R���e�i�T�C�Y
					Xls1.Cell("**ContType").Str  = Trim(rs("ContType"))			'�R���e�i�^�C�v
					Xls1.Cell("**ContHeight").Str= Trim(rs("ContHeight"))		'�R���e�i����
					
					'if CLng(rs("OvHeight")) > 0 then							'�n�g
						'Xls1.Cell("**OH").Str= "��"
						Xls1.Cell("**OH").Str= CLng(rs("OvHeight"))			
					'end if
					'if CLng(rs("OvWidthL")) > 0 then							'�n�v�k
						'Xls1.Cell("**OWL").Str= "��"
						Xls1.Cell("**OWL").Str=CLng(rs("OvWidthL"))
					'end if
					'if CLng(rs("OvWidthR")) > 0 then							'�n�v�q
						'Xls1.Cell("**OWR").Str= "��"
						Xls1.Cell("**OWR").Str= CLng(rs("OvWidthR"))
					'end if
					'if CLng(rs("OvLengthF")) > 0 then							'�n�k�e
						'Xls1.Cell("**OLF").Str= "��"
						Xls1.Cell("**OLF").Str= CLng(rs("OvLengthF"))							
					'end if
					'if CLng(rs("OvLengthA")) > 0 then							'�n�k�`
						'Xls1.Cell("**OLA").Str= "��"
						Xls1.Cell("**OLA").Str= CLng(rs("OvLengthA"))
					'end if

					if Trim(rs("CustClear")) = "Y" then 
						Xls1.Cell("**CusOK").Str= "��"							'�ʊ�
					elseif Trim(rs("CustClear")) = "N" then 
						Xls1.Cell("**CusNo").Str= "��"							'�ʊ�
					end if
					
					Xls1.Cell("**WorkNo").Str= Trim(rs("WkNo"))					'��Ɣԍ�					
					Xls1.Cell("**BookingNo").Str= Trim(rs("BookNo"))			'Booking�ԍ�
					Xls1.Cell("**GrossWT").Value= Trim(rs("ContWeight"))		'���d��
					Xls1.Cell("**Voyage").Str= Trim(rs("Voyage"))				'Voyage
					Xls1.Cell("**Fdest").Str= Trim(rs("FDest"))					'�ŏI�ړI�n
					'2010/05/10 Upd-S Tanaka
					'Xls1.Cell("**TareWT").Str= Trim(rs("TareWeight"))		    'TARE�E�F�C�g
					if Trim(rs("TareWeight"))<100 Then
						Xls1.Cell("**TareWT").Value= Trim(rs("TareWeight"))*100		'TARE�E�F�C�g
					Else
						Xls1.Cell("**TareWT").Value= Trim(rs("TareWeight"))			'TARE�E�F�C�g
					End If
					'2010/05/10 Upd-E Tanaka
					Xls1.Cell("**NetWT").Value= CInt(Trim(rs("ContWeight"))) - CInt(Trim(rs("TareWeight")))			'2016/10/28 H.Yoshikawa Add
					
					'Xls1.Cell("**Temp").Str= Mid(Trim(rs("SetTemp")),1,5)		'�Ⓚ���x
'					
'					if Trim(rs("Ventilation")) <> "" Then
'						Xls1.Cell("**VentHole").Str= "CLOSE"					'�ʕ��E
'					end if
'					
'					if Mid(Trim(rs("SetTemp")),6,1) = "C" then
'						Xls1.Cell("**TempC").Str= "��"							'��C
'					elseif Mid(Trim(rs("SetTemp")),6,1) = "F" then
'						Xls1.Cell("**TempF").Str= "��"							'��F
'					end if				
'					
'					if Trim(rs("IMDG1")) <> "" then 							'IMCO
'						Xls1.Cell("**IMCO").Str= "��"
'					elseif Trim(rs("IMDG2")) <> "" then
'						Xls1.Cell("**IMCO").Str= "��"
'					elseif Trim(rs("IMDG3")) <> "" then 
'						Xls1.Cell("**IMCO").Str= "��"
'					end if
'
'					if Trim(rs("UNNo1")) <> "" then								'UN 
'						Xls1.Cell("**UN").Str= "��"
'					elseif Trim(rs("UNNo2")) <> "" then
'						Xls1.Cell("**UN").Str= "��"
'					elseif Trim(rs("UNNo3")) <> "" then
'						Xls1.Cell("**UN").Str= "��"
'					end if

					'2017/01/20 T.Okui Upd Start  KACCS���̃f�[�^���g�p����悤�ɕύX
					'Xls1.Cell("**Temp").Str= Trim(rs("SetTemp"))	    		'�Ⓚ���x
					'if Trim(rs("AsDry")) = "1" then
					'	Xls1.Cell("**Temp").Str = ""							'�Ⓚ���x
					'	Xls1.Cell("**AsDry").Str = "DRY USE"					'As Dry
					'else
					'	Xls1.Cell("**Temp").Str= Mid(Trim(rs("SetTemp")),1,5)	'�Ⓚ���x
					'	Xls1.Cell("**AsDry").Str = ""							'As Dry
					'end if

					if  Trim(ContType) = "RF"  then
						if (Trim(SetTemp) <> "NO" and Trim(SetTemp) <> "") then
							Xls1.Cell("**Temp").Str = Mid(Trim(SetTemp),1,5)		'�Ⓚ���x
						else
							Xls1.Cell("**Temp").Str = ""
						end if
			
						'As Dry
						if Trim(SetTemp) <> "NO" and Trim(SetTemp) = "" then
							Xls1.Cell("**AsDry").Str = "DRY USE"		
						else
							Xls1.Cell("**AsDry").Str = ""
						end if
						
						if Mid(Trim(SetTemp),6,1) = "C" then
							Xls1.Cell("**TempDegree1").Str= "�K"
							Xls1.Cell("**TempDegree2").Str= "C"							'��C
						elseif Mid(Trim(SetTemp),6,1) = "F" then
							Xls1.Cell("**TempDegree1").Str= "�K"
							Xls1.Cell("**TempDegree2").Str= "F"							'��F
						
						end if				
					end if
					'2017/01/20 T.Okui Upd End

					'2016/08/19 H.Yoshikawa Del Start
					'if UCASE(Trim(rs("TempDegree")))="C" Then 				    
					'	Xls1.Cell("**Cel").Str= "��"
					'elseif UCASE(Trim(rs("TempDegree")))="F" Then 
					'	Xls1.Cell("**Far").Str= "��"
					'end if
					'2016/08/18 H.Yoshikawa Del End
					
					if Trim(rs("Ventilation")) <> "" Then						'�ʕ��E
						Xls1.Cell("**Vent").Str= Trim(rs("Ventilation"))					
					end if

					if Trim(rs("IMDG1")) <> "" then 							'IMCO
						Xls1.Cell("**IMCOA").Str= Trim(rs("IMDG1"))
					end if
					if Trim(rs("IMDG2")) <> "" then
						Xls1.Cell("**IMCOB").Str= Trim(rs("IMDG2"))
					end if
					if Trim(rs("IMDG3")) <> "" then 
						Xls1.Cell("**IMCOC").Str= Trim(rs("IMDG3"))
					end if

					if Trim(rs("UNNo1")) <> "" then								'UN 
						Xls1.Cell("**UNA").Str= Trim(rs("UNNo1"))
					end if
					if Trim(rs("UNNo2")) <> "" then
						Xls1.Cell("**UNB").Str= Trim(rs("UNNo2"))
					end if
					if Trim(rs("UNNo3")) <> "" then
						Xls1.Cell("**UNC").Str= Trim(rs("UNNo3"))
					end if
					
					'2016/08/19 H.Yoshikawa Add Start
					if Trim(rs("IMDG4")) <> "" then 							'IMCO
						Xls1.Cell("**IMCOD").Str= Trim(rs("IMDG4"))
					end if
					if Trim(rs("IMDG5")) <> "" then
						Xls1.Cell("**IMCOE").Str= Trim(rs("IMDG5"))
					end if

					if Trim(rs("UNNo4")) <> "" then
						Xls1.Cell("**UND").Str= Trim(rs("UNNo4"))
					end if
					if Trim(rs("UNNo5")) <> "" then
						Xls1.Cell("**UNE").Str= Trim(rs("UNNo5"))
					end if
					
					if Trim(rs("LqFlag1")) = "1" then
						Xls1.Cell("**LQA").Str= "��"
					end if
					if Trim(rs("LqFlag2")) = "1" then
						Xls1.Cell("**LQB").Str= "��"
					end if
					if Trim(rs("LqFlag3")) = "1" then
						Xls1.Cell("**LQC").Str= "��"
					end if
					if Trim(rs("LqFlag4")) = "1" then
						Xls1.Cell("**LQD").Str= "��"
					end if
					if Trim(rs("LqFlag5")) = "1" then
						Xls1.Cell("**LQE").Str= "��"
					end if
					
					'2016/10/20 H.Yoshikawa Add Start
					Xls1.Cell("**LBLA").Str= Trim(rs("Label1"))
					Xls1.Cell("**LBLB").Str= Trim(rs("Label2"))
					Xls1.Cell("**LBLC").Str= Trim(rs("Label3"))
					Xls1.Cell("**LBLD").Str= Trim(rs("Label4"))
					Xls1.Cell("**LBLE").Str= Trim(rs("Label5"))
					Xls1.Cell("**SLBLA").Str= Trim(rs("SubLabel1"))
					Xls1.Cell("**SLBLB").Str= Trim(rs("SubLabel2"))
					Xls1.Cell("**SLBLC").Str= Trim(rs("SubLabel3"))
					Xls1.Cell("**SLBLD").Str= Trim(rs("SubLabel4"))
					Xls1.Cell("**SLBLE").Str= Trim(rs("SubLabel5"))
					'2016/10/20 H.Yoshikawa Add End
					
					Xls1.Cell("**decider").Str = ""										'�m���						'2016/10/20 H.Yoshikawa Add�i�b��I�ɋ�l�j
					Xls1.Cell("**ReportNo").Str= Trim(rs("ReportNo"))					'�͏o�ԍ��܂��͓o�^�ԍ�
					Xls1.Cell("**Shipper").Str= Trim(rs("PRShipper"))					'�׎喼
					Xls1.Cell("**PRFwd").Str= Trim(rs("PRForwarder"))					'�C�݋Ǝ҉�Ж�
					Xls1.Cell("**PRFwdTan").Str= Trim(rs("PRForwarderTan"))				'�C�݋ƎҒS���Җ�
					Xls1.Cell("**PRFwdTel").Str= Trim(rs("PRForwarderTel"))				'�C�݋ƎҘA����
					'2016/10/20 H.Yoshikawa Del Start
					'Xls1.Cell("**Comment1").Str= Comment1								'���l1
					'Xls1.Cell("**Comment2").Str= Comment2								'���l2
					'Xls1.Cell("**Comment3").Str= Comment3								'���l3
					'2016/10/20 H.Yoshikawa Del End

					nAttrNo = Xls1.Cell("**OutDate").AttrNo
					Xls1.Cell("**OutDate").Value2 CDate(rs("UpdtTime")), nAttrNo									'�o�^���i̫�ϯĂ̑����������p���j
					Xls1.Cell("**OutDate").PosHorz = phEven													'�����ʒu�F�ϓ���
					'2016/11/22 H.Yoshikawa Upd Start
					'Xls1.Cell("**Signature").Str= Trim(Session.Contents("LinUN")) & "�@" & TruckerName			'���O�C�����[�U�� + �o�^�S���Җ�
					'Xls1.Cell("**ContactInfo").Str= Trim(rs("ContactInfo"))									'�o�^�ҘA����
					Xls1.Cell("**Signature").Str= Trim(rs("PRForwarder")) & "�@" & Trim(rs("PRForwarderTan"))	'�C�݋Ǝ҉�Ж� + �C�݋ƎҒS���Җ�
					Xls1.Cell("**ContactInfo").Str= Trim(rs("PRForwarderTel"))										'�C�݋ƎҘA����
					'2016/11/22 H.Yoshikawa Upd Start

					'2016/08/19 H.Yoshikawa Add Start

				end if
				'RecordSet�ڑ�����
				rs.Close
			Next

			bDelReturn=Xls1.DelSheet(0,1)
			
			wMsg = "Excel�t�@�C�����쐬���܂����B�ȉ����쐬�����t�@�C�����_�E�����[�h�ł��܂��B"

			wErrNo = Xls1.ErrorNo
			
			if Err.Number<>0 then 
				wMsg=Err.Description
			end if

			If wErrNo <> 0 Then
				wMsg = "ExcelCreator3�G���[���b�Z�[�W�F" & Xls1.ErrorMessage
			End If
			
			
			Xls1.CloseBook

			Set Xls1 = Nothing		
			gfReceiveReportMultiple=flname
		else
			gfReceiveReportMultiple=""
		end if
	'DB�ڑ�����
	DisConnDBH cn, rs
			 
	End With
End Function 
%>