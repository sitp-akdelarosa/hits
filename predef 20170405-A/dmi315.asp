<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmo315.asp				_/
'_/	Function	:���O���������擾		_/
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
<!--#include File="CommonFunc.inc"-->
<%
'�Z�b�V�����̗L�������`�F�b�N
  CheckLoginH
  
'�f�[�^���擾
  dim CONnum,BookNo,User
  BookNo = Trim(Request("BookNo"))
  CONnum = Trim(Request("CONnum"))
  User   = Session.Contents("userid")
  'add-s h.matsuda 2006/03/06
	dim ShipLine,ShoriMode
	ShipLine = Trim(Request("ShipLine"))
	ShoriMode = Trim(Request("ShoriMode"))
  'add-e h.matsuda 2006/03/06

  '2016/10/18 H.Yoshikawa Add Start
  dim WkNo
  WkNo = gfTrim(Request("WkNo"))
  '2016/10/18 H.Yoshikawa Add End

'�f�[�^��DB��茟��
  dim shipFact,shipName,RecTerminal,PlaceDel,LPort,DPort
  dim RHO,SetTemp,Ventilation,IMDG1,IMDG2,IMDG3,UNNo1,UNNo2,UNNo3
  dim ContSize,ContType,ContHeight,Material,TareWeight,SealNo,ContWeight
  dim CMPcd,MrSk,HFrom,TuSk,NextV,OH,OWL,OWR,OLF,OLA,NiwataP,Operator
  dim Hmon,Hday,HedId,Comment1,Comment2,Comment3
  dim SakuNo,UpFlag,compFlag,WkCNo,TruckerFlag
  dim TruckerSubName
  dim ShipLineName											'2016/08/05 H.Yoshikawa Add
  dim Shipper, Forwarder, FwdrTan, FwdrTEL, TrkrTEL			'2016/10/13 H.Yoshikawa Add
  dim VslCode,VoyCtrl										'2016/10/14 H.Yoshikawa Add
  dim CMPcd1,ReportNo,AsDry,IMDG4,IMDG5,UNNo4,UNNo5			'2016/10/18 H.Yoshikawa Add
  dim Label1,Label2,Label3,Label4,Label5					'2016/10/18 H.Yoshikawa Add
  dim SubLabel1,SubLabel2,SubLabel3,SubLabel4,SubLabel5		'2016/10/18 H.Yoshikawa Add
  dim LqFlag1,LqFlag2,LqFlag3,LqFlag4,LqFlag5				'2016/10/18 H.Yoshikawa Add
  dim StrCodes												'2016/10/18 H.Yoshikawa Add
  dim PlaceRec, NiwataNm, NiukeNm, LPortNm, DPortNm			'2016/11/02 H.Yoshikawa Add

  TruckerSubName = Trim(Request("TruckerSubName"))

  CMPcd   =Array(Ucase(User),"","","","")
'�G���[�g���b�v�J�n
  on error resume next
'DB�ڑ�
  dim ObjConn, ObjRS, StrSQL, ErrerM
  ConnDBH ObjConn, ObjRS
  
  dim dummy,ret
  ret=true

'2016/10/18 H.Yoshikawa Add Start
'��Ɣԍ��w��̏ꍇ
if WkNo <> "" then
	Dim ObjRS_CYV, ObjRS_BOK
	Dim WkCtrlNo

	StrSQL = "select * from CYVanInfo where WkNo = '"& gfSQLEncode(WkNo) & "' "
	ObjRS.Open StrSQL, ObjConn
	if err <> 0 then
		DisConnDBH ObjConn, ObjRS2	'DB�ؒf
		jampErrerP "1","b401","01","���O�o�^�F��Ɣԍ����݃`�F�b�N","101","SQL:<BR>"&StrSQL
	end if
	if ObjRS.eof then
		ret=false
		ErrerM="�w�肵����Ɣԍ����V�X�e���ɓo�^����Ă��܂���B<BR>���͂̊ԈႢ���Ȃ����ԍ����m�F���Ă��������B</P>"
	else
		Set ObjRS_CYV = ObjRS.clone
		BookNo = gfTrim(ObjRS_CYV("BookNo"))
		ShipLine = gfTrim(ObjRS_CYV("ShipLine"))
	end if
	ObjRS.Close
	
	if ret then
		'�u�b�L���O�ԍ��̑��݃`�F�b�N
		ret=true
		StrSQL = "SELECT * From Booking AS BOK "&_
		         " WHERE BOK.BookNo='"& BookNo &"'"
		strSQL=strSQL & " AND BOK.shipline='"& ShipLine &"'"
		ObjRS.Open StrSQL, ObjConn
		if err <> 0 then
			DisConnDBH ObjConn, ObjRS	'DB�ؒf
			jampErrerP "1","b401","01","�������F���݃`�F�b�N","101","SQL:<BR>"&StrSQL
		end if
		If ObjRS.eof Then
			ret=false
			ErrerM="�w�肵���u�b�L���ONo���V�X�e���ɓo�^����Ă��܂���B<BR>���͂̊ԈႢ���Ȃ����ԍ����m�F���Ă��������B</P>"
		else
			Set ObjRS_BOK = ObjRS.clone
		End If
		ObjRS.Close
	end if
	
	If ret Then
		'�d���`�F�b�N
		StrSQL = "SELECT Count(ITC.WkContrlNo) AS Num "&_
				"FROM hITCommonInfo AS ITC LEFT JOIN CYVanInfo AS CYV ON (ITC.WkNo = CYV.WkNo) AND (ITC.ContNo = CYV.ContNo) "&_
				"WHERE ITC.ContNo='" & CONnum &"' AND ITC.Process='R' AND ITC.WkType='3' AND CYV.BookNo='"& BookNo &"' "
		ObjRS.Open StrSQL, ObjConn
		if err <> 0 then
			DisConnDBH ObjConn, ObjRS	'DB�ؒf
			jampErrerP "1","b401","01","�������F���ݏd���`�F�b�N","101","SQL:<BR>"&StrSQL
		end if
		If Trim(ObjRS("Num")) <> "0" Then
			ret=false
			ErrerM="�w�肵����Ɣԍ��A�R���e�i�ԍ��͂��łɓo�^����Ă��܂�</P>"
		End If
		ObjRS.Close
	end if

	If ret Then
		'2016/11/08 H.Yoshikawa Add Start
		'TareWeight��Container���擾
		strSQL=" select top 1"
		strSQL=strSQL & " isnull(TareWeight,'') as TareWeight"
		strSQL=strSQL & " from container where contno='" & CONnum & "'"
		strSQL=strSQL & " order by updttime desc"
		ObjRS.Open StrSQL, ObjConn
		if err <> 0 then
			DisConnDBH ObjConn, ObjRS	'DB�ؒf
			jampErrerP "1","b401","01","�������F�R���e�i���`�F�b�N","101","SQL:<BR>"&StrSQL
		end if
		if not ObjRS.EOF then
			TareWeight=Trim(ObjRS("TareWeight"))
		end if
		ObjRS.Close
		'2016/11/08 H.Yoshikawa Add End
	
		'���O�C�����[�U�ɂ���ĉ�ЃR�[�h�\������
		chengeCompCd CMPcd, UpFlag
		compFlag = true
		if err <> 0 then
			jampErrerP "1","b402","01","�������F�f�[�^�ҏW","200",""
		end if

		RecTerminal= gfTrim(ObjRS_BOK("RecTerminal"))		'������

		shipFact   = gfTrim(ObjRS_CYV("ShipLine"))			'�D��
		shipName   = gfTrim(ObjRS_CYV("VslName"))			'�D��
		VslCode    = gfTrim(ObjRS_BOK("VslCode"))			'�D���R�[�h
		NextV      = gfTrim(ObjRS_CYV("Voyage"))			'���q�i�A�oVoy�j
		VoyCtrl    = gfTrim(ObjRS_BOK("VoyCtrl"))			'���q
		'2016/11/02 H.Yoshikawa Upd Start
		'PlaceDel   = gfTrim(ObjRS_BOK("PlaceRec"))			'�׎�n
		'LPort      = gfTrim(ObjRS_BOK("LPort"))			'�ύ`
		PlaceRec   = gfTrim(ObjRS_CYV("PlaceRec"))			'�׎�n
		LPort      = gfTrim(ObjRS_CYV("LPort"))				'�ύ`
		'2016/11/02 H.Yoshikawa Upd End
		DPort      = gfTrim(ObjRS_CYV("DPort"))				'�g�`
		NiwataP    = gfTrim(ObjRS_CYV("DelivPlace"))		'�דn�n
		
		Shipper   = gfTrim(ObjRS_CYV("PRShipper"))			'�׎�
		Forwarder = gfTrim(ObjRS_CYV("PRForwarder"))		'�戵�C�݋Ǝ�
		FwdrTan   = gfTrim(ObjRS_CYV("PRForwarderTan"))		'�戵�C�ݒS����
		FwdrTEL   = gfTrim(ObjRS_CYV("PRForwarderTEL"))		'�C�ݘA����

		ContSize   = gfTrim(ObjRS_CYV("ContSize"))
		ContType   = gfTrim(ObjRS_CYV("ContType"))
		ContHeight = gfTrim(ObjRS_CYV("ContHeight"))
		'TareWeight = gfTrim(ObjRS_CYV("TareWeight"))								'2016/11/08 H.Yoshikawa Del
		'2016/11/16 H.Yoshikawa Upd Start
		if gfTrim(ObjRS_CYV("CustOK")) = "Y" then
			MrSk   = "Y"									'�ۊ�
		else
			MrSk   = "N"									'�ۊ�
		end if
		'2016/11/16 H.Yoshikawa Upd End
		SealNo     = ""										'�V�[���ԍ�
		ContWeight = ""										'�O���X�E�F�C�g
		ReportNo   = gfTrim(ObjRS_CYV("ReportNo"))			'�o�^�ԍ��܂��͓͏o�ԍ�
		HFrom      = gfTrim(ObjRS_CYV("ReceiveFrom"))		'������
		SetTemp    = gfTrim(ObjRS_CYV("SetTemp"))			'�ݒ艷�x
		Ventilation= gfTrim(ObjRS_CYV("Ventilation"))		'VENT
		AsDry      = gfTrim(ObjRS_CYV("AsDry"))
		IMDG1      = gfTrim(ObjRS_CYV("IMDG1"))
		IMDG2      = gfTrim(ObjRS_CYV("IMDG2"))
		IMDG3      = gfTrim(ObjRS_CYV("IMDG3"))
		IMDG4      = gfTrim(ObjRS_CYV("IMDG4"))
		IMDG5      = gfTrim(ObjRS_CYV("IMDG5"))
		Label1     = gfTrim(ObjRS_CYV("Label1"))
		Label2     = gfTrim(ObjRS_CYV("Label2"))
		Label3     = gfTrim(ObjRS_CYV("Label3"))
		Label4     = gfTrim(ObjRS_CYV("Label4"))
		Label5     = gfTrim(ObjRS_CYV("Label5"))
		SubLabel1  = gfTrim(ObjRS_CYV("SubLabel1"))
		SubLabel2  = gfTrim(ObjRS_CYV("SubLabel2"))
		SubLabel3  = gfTrim(ObjRS_CYV("SubLabel3"))
		SubLabel4  = gfTrim(ObjRS_CYV("SubLabel4"))
		SubLabel5  = gfTrim(ObjRS_CYV("SubLabel5"))
		UNNo1      = gfTrim(ObjRS_CYV("UNNo1"))
		UNNo2      = gfTrim(ObjRS_CYV("UNNo2"))
		UNNo3      = gfTrim(ObjRS_CYV("UNNo3"))
		UNNo4      = gfTrim(ObjRS_CYV("UNNo4"))
		UNNo5      = gfTrim(ObjRS_CYV("UNNo5"))
		LqFlag1    = gfTrim(ObjRS_CYV("LqFlag1"))
		LqFlag2    = gfTrim(ObjRS_CYV("LqFlag2"))
		LqFlag3    = gfTrim(ObjRS_CYV("LqFlag3"))
		LqFlag4    = gfTrim(ObjRS_CYV("LqFlag4"))
		LqFlag5    = gfTrim(ObjRS_CYV("LqFlag5"))
		OH         = gfTrim(ObjRS_CYV("OvHeight"))
		OWL        = gfTrim(ObjRS_CYV("OvWidthL"))
		OWR        = gfTrim(ObjRS_CYV("OvWidthR"))
		OLF        = gfTrim(ObjRS_CYV("OvLengthF"))
		OLA        = gfTrim(ObjRS_CYV("OvLengthA"))
		Operator   = gfTrim(ObjRS_CYV("Operator"))			'�I�y���[�^

		TrkrTEL    = gfTrim(ObjRS_CYV("ContactInfo"))			'�S���ҘA����
		TruckerFlag=0
		
		'IT���ʃe�[�u������
		StrSQL = "SELECT *, convert(char(10), WorkDate, 111) as WorkYMD From hITCommonInfo "
		StrSQL = StrSQL & " WHERE WkNo = '"& gfSQLEncode(WkNo) &"' "
		ObjRS.Open StrSQL, ObjConn
		if err <> 0 then
			DisConnDBH ObjConn, ObjRS	'DB�ؒf
			jampErrerP "1","b401","01","�������FIT���ʃe�[�u������","101","SQL:<BR>"&StrSQL&"<BR>Err:"& err.description 
		end if
		If not ObjRS.eof Then
			CMPcd1 = gfTrim(ObjRS("TruckerSubCode1"))		'��ЃR�[�h
			HedId  = gfTrim(ObjRS("HeadID"))				'�w�b�h�h�c
			'�����\���
			dim TmpA
			If IsNull(ObjRS("WorkYMD")) Then	'���t��Null�ł������ꍇ
				Hmon   = Null
				Hday   = Null
			Else
				TmpA   = Split(gfTrim(ObjRS("WorkYMD")), "/")
				Hmon   = TmpA(1)
				Hday   = TmpA(2)
			End If
			Comment1= ""
			Comment2= ""
			Comment3= ""
			TruckerSubName = gfTrim(ObjRS("TruckerSubName1"))		'�S����
		End If
		ObjRS.Close

		'�A�o�R���e�i�e�[�u������
		StrSQL = "SELECT RHO From ExportCont "
		StrSQL = StrSQL & " WHERE VslCode = '"& gfSQLEncode(VslCode) &"'"
		StrSQL = StrSQL & "   AND VoyCtrl = '"& gfSQLEncode(VoyCtrl) &"'"
		StrSQL = StrSQL & "   AND BookNo  = '"& gfSQLEncode(BookNo) &"'"
		StrSQL = StrSQL & "   AND ContNo  = '"& gfSQLEncode(CONnum) &"'"
		ObjRS.Open StrSQL, ObjConn
		if err <> 0 then
			DisConnDBH ObjConn, ObjRS	'DB�ؒf
			jampErrerP "1","b401","01","�������F�A�o�R���e�i�e�[�u������","101","SQL:<BR>"&StrSQL
		end if
		If not ObjRS.eof Then
			RHO = gfTrim(ObjRS("RHO"))		'RHO
		End If
		ObjRS.Close

		'�׎�n�A�ύ`�A�g�`�A�דn�n��FullName��
		StrCodes="'"&PlaceRec&"','"&LPort&"','"&DPort&"','"&NiwataP&"'"			'2016/11/02 H.Yoshikawa Upd (PlaceDel��PlaceRec)
		StrSQL = "SELECT mP.PortCode,mP.FullName From mPort AS mP "&_
		       "WHERE mP.PortCode IN ("& StrCodes &") "
		ObjRS.Open StrSQL, ObjConn
		if err <> 0 then
			DisConnDBH ObjConn, ObjRS	'DB�ؒf
			jampErrerP "1","b401","01","�������F�f�[�^�擾","102","SQL:<BR>"&StrSQL
		end if
		Do Until ObjRS.EOF
			If Not IsNull(ObjRS("FullName")) Then
			  If PlaceRec=Trim(ObjRS("PortCode")) Then				'2016/11/02 H.Yoshikawa Upd (PlaceDel��PlaceRec)
			    NiukeNm=Trim(ObjRS("FullName"))						'2016/11/02 H.Yoshikawa Upd (PlaceDel��NiukeNm)
			  End If
			  If LPort=Trim(ObjRS("PortCode")) Then
			    LPortNm=Trim(ObjRS("FullName"))						'2016/11/02 H.Yoshikawa Upd (LPort��LPortNm)
			  End If
			  '2016/11/02 H.Yoshikawa Add Start
			  If DPort=Trim(ObjRS("PortCode")) Then
			    DPortNm=Trim(ObjRS("FullName"))
			  End If
			  If NiwataP=Trim(ObjRS("PortCode")) Then
			    NiwataNm=Trim(ObjRS("FullName"))
			  End If
			  '2016/11/02 H.Yoshikawa Add End
			'20040701�b��Ή�
			'          If DPort=Trim(ObjRS("PortCode")) Then
			'            DPort=Trim(ObjRS("FullName"))
			'          End If
			'          If NiwataP=Trim(ObjRS("PortCode")) Then
			'            NiwataP=Trim(ObjRS("FullName"))
			'          End If
			End If
			ObjRS.MoveNext
		Loop
		ObjRS.Close

		if gfTrim(shipFact) <> "" then
			'�戵�D�Ж��擾
			strSQL="SELECT FullName FROM mShipLine "
			strSQL=strSQL & " WHERE ShipLine = '" & shipFact & "' "
			ObjRS.Open StrSQL, ObjConn
			if err <> 0 then
				DisConnDBH ObjConn, ObjRS	'DB�ؒf
				jampErrerP "1","b401","01","�������F�f�[�^�擾","102","SQL:<BR>"&StrSQL
			end if
			if not ObjRS.EOF then
				ShipLineName   = Trim(ObjRS("FullName"))		'�D�Ж���
			end if
			ObjRS.Close
		end if
		
		ObjRS_CYV.close
		ObjRS_BOK.close
	end if
else
'2016/10/18 H.Yoshikawa Add End

'�u�b�L���O�ԍ��̑��݃`�F�b�N
  StrSQL = "SELECT Count(BOK.BookNo) AS Num "&_
           "From Booking AS BOK WHERE BOK.BookNo='"& BookNo &"'"
'2006/03/06 add-s h.matsuda
	if ShipLine<>"" and ShoriMode<>"" then
		strSQL=strSQL & " AND BOK.shipline='"& ShipLine &"'"
	End If
'2006/03/06 add-s h.matsuda
  ObjRS.Open StrSQL, ObjConn
  if err <> 0 then
    DisConnDBH ObjConn, ObjRS	'DB�ؒf
    jampErrerP "1","b401","01","�������F���݃`�F�b�N","101","SQL:<BR>"&StrSQL
  end if
  If Trim(ObjRS("Num")) = "0" Then
    ret=false
    ErrerM="�w�肵���u�b�L���ONo���V�X�e���ɓo�^����Ă��܂���B<BR>���͂̊ԈႢ���Ȃ����ԍ����m�F���Ă��������B</P>"
  End If
  ObjRS.Close

  If ret Then
  '�d���`�F�b�N
    StrSQL = "SELECT Count(ITC.WkContrlNo) AS Num "&_
             "FROM hITCommonInfo AS ITC LEFT JOIN CYVanInfo AS CYV ON (ITC.WkNo = CYV.WkNo) AND (ITC.ContNo = CYV.ContNo) "&_
             "WHERE ITC.ContNo='" & CONnum &"' AND ITC.Process='R' AND ITC.WkType='3' AND CYV.BookNo='"& BookNo &"' "
    ObjRS.Open StrSQL, ObjConn
    if err <> 0 then
      DisConnDBH ObjConn, ObjRS	'DB�ؒf
      jampErrerP "1","b401","01","�������F���ݏd���`�F�b�N","101","SQL:<BR>"&StrSQL
    end if
    If Trim(ObjRS("Num")) <> "0" Then
      ret=false
      ErrerM="�w�肵���u�b�L���ONo�A�R���e�i�ԍ��͂��łɓo�^����Ă��܂�</P>"

    End If
    ObjRS.Close
    If ret Then

		strSQL=" select top 1"
		strSQL=strSQL & " isnull(ContSize,'') as ContSize,"
		strSQL=strSQL & " isnull(ContType,'') as ContType,"
		strSQL=strSQL & " isnull(ContHeight,'') as ContHeight,"
		strSQL=strSQL & " isnull(Material,'') as Material,"
		strSQL=strSQL & " isnull(TareWeight,'') as TareWeight"
		strSQL=strSQL & " from container where contno='" & CONnum & "'"
		strSQL=strSQL & " order by updttime desc"
		ObjRS.Open StrSQL, ObjConn
		if err <> 0 then
			DisConnDBH ObjConn, ObjRS	'DB�ؒf
			jampErrerP "1","b401","01","�������F�R���e�i���`�F�b�N","101","SQL:<BR>"&StrSQL
		end if
		if not ObjRS.EOF then
			ContSize=Trim(ObjRS("ContSize"))
			ContType=Trim(ObjRS("ContType"))
			ContHeight=Trim(ObjRS("ContHeight"))
			Material=Trim(ObjRS("Material"))
			TareWeight=Trim(ObjRS("TareWeight"))
		end if
		ObjRS.Close


'2006/03/06 add-s h.matsuda(SQL�����č\�z)
'      StrSQL = "SELECT BOK.RecTerminal, BOK.PlaceRec, BOK.LPort, "&_
'               "CYV.ShipLine,CYV.VslName,CYV.DPort, CYV.Voyage,CYV.DelivPlace,CYV.Operator,"&_
'               "CYV.OvHeight,CYV.OvWidthL,CYV.OvWidthR,CYV.OvLengthF,CYV.OvLengthA,"&_
'               "CON.ContSize, CON.ContType, CON.ContHeight, CON.Material, CON.TareWeight, "&_
'               "EXC.RHO, EXC.SetTemp, EXC.Ventilation, " &_
'               "EXC.IMDG1, EXC.IMDG2, EXC.IMDG3, EXC.UNNo1, EXC.UNNo2,EXC.UNNo3 "&_
'               "From Booking AS BOK LEFT JOIN mVessel AS mV ON BOK.VslCode = mV.VslCode "&_
'               "LEFT JOIN CYVanInfo CYV ON BOK.BookNo=CYV.BookNo AND SenderCode=' ' "&_
'               "LEFT JOIN ExportCont AS EXC ON EXC.ContNo='"& CONnum &"' "&_
'               "AND BOK.BookNo=EXC.BookNo AND BOK.VslCode=EXC.VslCode AND BOK.VoyCtrl=EXC.VoyCtrl "&_
'               "LEFT JOIN Container AS CON ON EXC.ContNo=CON.ContNo "&_
'               "AND EXC.VslCode=CON.VslCode AND EXC.VoyCtrl=CON.VoyCtrl "&_
'               "WHERE BOK.BookNo='"& BookNo &"' "&_
'               "ORDER BY BOK.UpdtTime DESC"
		strSQL="		  SELECT BOK.RecTerminal, BOK.PlaceRec, BOK.LPort,      						"
		strSQL=strSQL & " coalesce(CYV.ShipLine,bok.shipline) shipline,                                 "
		strSQL=strSQL & " coalesce(CYV.VslName,mv.fullname) vslname,                                    "
'2016/11/04 H.Yoshikawa Upd Start
'		strSQL=strSQL & " coalesce(CYV.DPort,bok.dport) dport,                                          "
		strSQL=strSQL & " BOK.DPort, "
'2016/11/04 H.Yoshikawa Upd End
		strSQL=strSQL & " CYV.Voyage,																	"
'2016/11/04 H.Yoshikawa Upd Start
'		strSQL=strSQL & " coalesce(CYV.DelivPlace,bok.delivplace) delivplace,CYV.Operator,              "
		strSQL=strSQL & " BOK.DelivPlace,CYV.Operator,              "
'2016/11/04 H.Yoshikawa Upd End
		strSQL=strSQL & " CYV.OvHeight,CYV.OvWidthL,CYV.OvWidthR,CYV.OvLengthF,CYV.OvLengthA,           "
'2006/04/18 mod-s h.matsuda
'			strSQL=strSQL & " CON.ContSize, CON.ContType, CON.ContHeight, CON.Material, CON.TareWeight,     "
		strSQL=strSQL & " isnull(cyv.ContSize,'" & ContSize & "') as ContSize,							"
		'2016/08/19 H.Yoshikawa Upd Start
		'strSQL=strSQL & " isnull(cyv.ContSize,'" & ContType & "') as ContType,							"
		'strSQL=strSQL & " isnull(cyv.ContSize,'" & ContHeight & "') as ContHeight,						"
		'strSQL=strSQL & " isnull(cyv.ContSize,'" & Material & "') as Material,							"
		'strSQL=strSQL & " isnull(cyv.ContSize,'" & TareWeight & "') as TareWeight,						"
		strSQL=strSQL & " isnull(cyv.ContType,'" & ContType & "') as ContType,							"
		strSQL=strSQL & " isnull(cyv.ContHeight,'" & ContHeight & "') as ContHeight,						"
		strSQL=strSQL & " isnull(cyv.Material,'" & Material & "') as Material,							"
		strSQL=strSQL & " isnull(cyv.TareWeight,'" & TareWeight & "') as TareWeight,						"
		'2016/08/19 H.Yoshikawa Upd End
'2006/04/18 mod-e h.matsuda
		strSQL=strSQL & " EXC.RHO, EXC.SetTemp, EXC.Ventilation,                                        "
		strSQL=strSQL & " EXC.IMDG1, EXC.IMDG2, EXC.IMDG3, EXC.UNNo1, EXC.UNNo2,EXC.UNNo3               "
		'2016/10/14 H.Yoshikawa Upd Start
		strSQL=strSQL & " ,BOK.VslCode, BOK.VoyCtrl "
		'2016/10/14 H.Yoshikawa Upd End
		strSQL=strSQL & " ,cyv.CustOK "									'2016/11/16 H.Yoshikawa Add
		strSQL=strSQL & " From Booking AS BOK                                                           "
		strSQL=strSQL & " LEFT JOIN mVessel AS mV ON BOK.VslCode = mV.VslCode                           "
		strSQL=strSQL & " LEFT JOIN CYVanInfo CYV ON BOK.BookNo=CYV.BookNo AND SenderCode=' '           "
		if ShipLine<>"" and ShoriMode<>"" then
			strSQL=strSQL & " AND BOK.shipline=cyv.shipline                                                 "
		End If
		strSQL=strSQL & " LEFT JOIN ExportCont AS EXC ON                                                "
		strSQL=strSQL & " EXC.ContNo='" & CONnum & "'                                                    "
		strSQL=strSQL & " AND BOK.BookNo=EXC.BookNo AND BOK.VslCode=EXC.VslCode                         "
		strSQL=strSQL & " AND BOK.VoyCtrl=EXC.VoyCtrl                                                   "
		strSQL=strSQL & " LEFT JOIN Container AS CON ON EXC.ContNo=CON.ContNo                           "
		strSQL=strSQL & " AND EXC.VslCode=CON.VslCode AND EXC.VoyCtrl=CON.VoyCtrl                       "
		strSQL=strSQL & " WHERE BOK.BookNo='" & BookNo & "'                                             "
		if ShipLine<>"" and ShoriMode<>"" then
			strSQL=strSQL & " and BOK.ShipLine='"& ShipLine &"'                                             "
		End If
		strSQL=strSQL & " ORDER BY BOK.UpdtTime DESC                                                    "
'2006/03/06 add-e h.matsuda

'CW-324 Change ASC->DESC
'20040227 Change PlaceDel->PlaceRec
		ObjRS.Open StrSQL, ObjConn
		if err <> 0 then
			DisConnDBH ObjConn, ObjRS	'DB�ؒf
			jampErrerP "1","b401","01","�������F�f�[�^�擾","102","SQL:<BR>"&StrSQL
		end if
		shipFact   = Trim(ObjRS("ShipLine"))		'�D��
		shipName   = ""								'�D��				'2016/10/18 H.Yoshikawa Upd (�����l�Ȃ�)
		RecTerminal= Trim(ObjRS("RecTerminal"))		'������

		'20040227 Change      PlaceDel   = Trim(ObjRS("PlaceDel"))		'�׎�n
		'2016/11/02 H.Yoshikawa Upd Start
		'PlaceDel   = Trim(ObjRS("PlaceRec"))		'�׎�n
		PlaceRec   = Trim(ObjRS("PlaceRec"))		'�׎�n
		'2016/11/02 H.Yoshikawa Upd End
		LPort      = Trim(ObjRS("LPort"))			'�ύ`
		DPort      = Trim(ObjRS("DPort"))			'�g�`
		RHO        = Trim(ObjRS("RHO"))
		SetTemp    = ""								'�ݒ艷�x			'2016/10/18 H.Yoshikawa Upd (�����l�Ȃ�)
		Ventilation= ""								'VENT				'2016/10/18 H.Yoshikawa Upd (�����l�Ȃ�)
		IMDG1      = ""													'2016/10/18 H.Yoshikawa Upd (�����l�Ȃ�)
		IMDG2      = ""													'2016/10/18 H.Yoshikawa Upd (�����l�Ȃ�)
		IMDG3      = ""													'2016/10/18 H.Yoshikawa Upd (�����l�Ȃ�)
		UNNo1      = ""													'2016/10/18 H.Yoshikawa Upd (�����l�Ȃ�)
		UNNo2      = ""													'2016/10/18 H.Yoshikawa Upd (�����l�Ȃ�)
		UNNo3      = ""													'2016/10/18 H.Yoshikawa Upd (�����l�Ȃ�)

		ContSize   = ""													'2016/10/18 H.Yoshikawa Upd (�����l�Ȃ�)
		ContType   = ""													'2016/10/18 H.Yoshikawa Upd (�����l�Ȃ�)
		ContHeight = ""													'2016/10/18 H.Yoshikawa Upd (�����l�Ȃ�)
		Material   = Trim(ObjRS("Material"))
		TareWeight = Trim(ObjRS("TareWeight"))
		SealNo     = ""
		ContWeight = ""	'�O���X�E�F�C�g

		'2016/11/16 H.Yoshikawa Upd Start
		'MrSk       = ""	'�ۊ�
		if gfTrim(ObjRS("CustOK")) = "Y" then
			MrSk   = "Y"
		else
			MrSk   = "N"
		end if
		'2016/11/16 H.Yoshikawa Upd End
		HFrom      = ""	'������
		TuSk       = ""	'�ʊ�
		NextV      = ""			'���q									'2016/10/18 H.Yoshikawa Upd (�����l�Ȃ�)
		VoyCtrl    = Trim(ObjRS("VoyCtrl"))			'���q				'2016/10/14 H.Yoshikawa Add
		VslCode    = Trim(ObjRS("VslCode"))			'�D���R�[�h			'2016/10/14 H.Yoshikawa Add
		
		OH         = ""													'2016/10/18 H.Yoshikawa Upd (�����l�Ȃ�)
		OWL        = ""													'2016/10/18 H.Yoshikawa Upd (�����l�Ȃ�)
		OWR        = ""													'2016/10/18 H.Yoshikawa Upd (�����l�Ȃ�)
		OLF        = ""													'2016/10/18 H.Yoshikawa Upd (�����l�Ȃ�)
		OLA        = ""													'2016/10/18 H.Yoshikawa Upd (�����l�Ȃ�)
		NiwataP    = Trim(ObjRS("DelivPlace"))	'�דn�n
		Operator   = Trim(ObjRS("Operator"))	'�I�y���[�^
		Hmon    = ""	'�����\���
		Hday    = ""
		HedId   = ""
		Comment1= ""
		Comment2= ""
		Comment3= ""
		TruckerFlag=0
		'���O�C�����[�U�ɂ���ĉ�ЃR�[�h�\������
		chengeCompCd CMPcd, UpFlag
		compFlag = true
		if err <> 0 then
			DisConnDBH ObjConn, ObjRS	'DB�ؒf
			jampErrerP "1","b402","01","�������F�f�[�^�ҏW","200",""
		end if
		ObjRS.Close
		'20040227 ADD START �׎�n�A�ύ`�A�g�`�A�דn�n��FullName��
		'Dim StrCodes														'2016/10/18 H.Yoshikawa Del�i�擪�ֈړ��j
		StrCodes="'"&PlaceRec&"','"&LPort&"','"&DPort&"','"&NiwataP&"'"			'2016/11/02 H.Yoshikawa Upd (PlaceDel��PlaceRec)
		StrSQL = "SELECT mP.PortCode,mP.FullName From mPort AS mP "&_
		       "WHERE mP.PortCode IN ("& StrCodes &") "
		ObjRS.Open StrSQL, ObjConn
		if err <> 0 then
			DisConnDBH ObjConn, ObjRS	'DB�ؒf
			jampErrerP "1","b401","01","�������F�f�[�^�擾","102","SQL:<BR>"&StrSQL
		end if
		Do Until ObjRS.EOF
			If Not IsNull(ObjRS("FullName")) Then
			  If PlaceRec=Trim(ObjRS("PortCode")) Then				'2016/11/02 H.Yoshikawa Upd (PlaceDel��PlaceRec)
			    NiukeNm=Trim(ObjRS("FullName"))						'2016/11/02 H.Yoshikawa Upd (PlaceDel��NiukeNm)
			  End If
			  If LPort=Trim(ObjRS("PortCode")) Then
			    LPortNm=Trim(ObjRS("FullName"))						'2016/11/02 H.Yoshikawa Upd (LPort��LPortNm)
			  End If
			  '2016/11/02 H.Yoshikawa Add Start
			  If DPort=Trim(ObjRS("PortCode")) Then
			    DPortNm=Trim(ObjRS("FullName"))
			  End If
			  If NiwataP=Trim(ObjRS("PortCode")) Then
			    NiwataNm=Trim(ObjRS("FullName"))
			  End If
			  '2016/11/02 H.Yoshikawa Add End
			'20040701�b��Ή�
			'          If DPort=Trim(ObjRS("PortCode")) Then
			'            DPort=Trim(ObjRS("FullName"))
			'          End If
			'          If NiwataP=Trim(ObjRS("PortCode")) Then
			'            NiwataP=Trim(ObjRS("FullName"))
			'          End If
			End If
			ObjRS.MoveNext
		Loop
		ObjRS.Close
		'2016/11/02 H.Yoshikawa Upd End
'20040227 ADD END
'2016/08/05 H.Yoshikawa Add Start
		if gfTrim(shipFact) <> "" then
			'�戵�D�Ж��擾
			strSQL="SELECT FullName FROM mShipLine "
			strSQL=strSQL & " WHERE ShipLine = '" & shipFact & "' "
			ObjRS.Open StrSQL, ObjConn
			if err <> 0 then
				DisConnDBH ObjConn, ObjRS	'DB�ؒf
				jampErrerP "1","b401","01","�������F�f�[�^�擾","102","SQL:<BR>"&StrSQL
			end if
			if not ObjRS.EOF then
				ShipLineName   = Trim(ObjRS("FullName"))		'�D�Ж���
			end if
			ObjRS.Close
		end if
'2016/08/05 H.Yoshikawa Add End

'2016/10/13 H.Yoshikawa Add Start
		'���[�U�}�X�^�擾
		strSQL="SELECT * FROM mUsers "
		StrSQL= StrSQL & "where UserCode = '" & User & "' "
		ObjRS.Open StrSQL, ObjConn
		if err <> 0 then
			DisConnDBH ObjConn, ObjRS	'DB�ؒf
			jampErrerP "1","b401","01","�������F�f�[�^�擾","102","SQL:<BR>"&StrSQL
		end if
		if not ObjRS.EOF then
			Shipper   = gfTrim(ObjRS("PRShipper"))			'�׎�
			Forwarder = gfTrim(ObjRS("PRForwarder"))		'�戵�C�݋Ǝ�
			FwdrTan   = gfTrim(ObjRS("PRForwarderTan"))		'�戵�C�ݒS����
			FwdrTEL   = gfTrim(ObjRS("PRForwarderTEL"))		'�C�ݘA����
			TrkrTEL   = gfTrim(ObjRS("TelNo"))				'�S���ҘA����
			TruckerSubName = gfTrim(ObjRS("TTName"))		'�S����
		end if
		ObjRS.Close
'2016/10/13 H.Yoshikawa Add End
    End If
  End If
end if										'2016/10/18 H.Yoshikawa Add
'DB�ڑ�����
  DisConnDBH ObjConn, ObjRS
'�G���[�g���b�v����
  on error goto 0
  
  If ret Then
    WriteLogH "b402", "���������O������", "01",BookNo&",���͓��e�̐���:0(������)"
  Else
    WriteLogH "b402", "���������O������", "01",BookNo&",���͓��e�̐���:1(���)"
  End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>���������擾��</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
function setParam(target){
<% If ret Then %>
  //2017/02/07 T.Okui Upd Start
  //2016/08/01 H.Yoshikawa Upd Start
  //window.resizeTo(850,690);
   window.moveTo(0,0);
   window.resizeTo(1200,900);
  //2016/08/01 H.Yoshikawa Upd Start
  //2017/02/07 T.Okui Upd End
  document.dmi315F.action="./dmi320.asp";
  document.dmi315F.submit();
<% Else %>
  window.resizeTo(500,500);
  window.focus();
<% End If %>
}
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY onLoad="setParam(document.dmi315F)">
<!-------------���������擾���--------------------------->
<FORM name="dmi315F" method="POST">
<% If ret Then %>
<INPUT type=hidden name="CONnum"  value="<%=CONnum%>">
<INPUT type=hidden name="BookNo"  value="<%=BookNo%>">
<INPUT type=hidden name="CONsize" value="<%=ContSize%>">
<INPUT type=hidden name="CONtype" value="<%=ContType%>">
<INPUT type=hidden name="CONhite" value="<%=ContHeight%>">
<INPUT type=hidden name="CONsitu" value="<%=Material%>">
<INPUT type=hidden name="CONtear" value="<%=TareWeight%>">
<INPUT type=hidden name="SealNo"  value="<%=SealNo%>">
<INPUT type=hidden name="GrosW"   value="<%=ContWeight%>">
<INPUT type=hidden name="HTo"     value="<%=RecTerminal%>">
<INPUT type=hidden name="ThkSya"  value="<%=shipFact%>">
<INPUT type=hidden name="ShipN"   value="<%=shipName%>">
<INPUT type=hidden name="RHO"     value="<%=RHO%>">
<INPUT type=hidden name="SttiT"   value="<%=SetTemp%>">
<INPUT type=hidden name="VENT"    value="<%=Ventilation%>">
<INPUT type=hidden name="NiukP"   value="<%=PlaceRec%>">					<!-- 2016/11/02 H.Yoshikawa Upd (PlaceDel��PlaceRec) -->
<INPUT type=hidden name="IMDG1"   value="<%=IMDG1%>">
<INPUT type=hidden name="IMDG2"   value="<%=IMDG2%>">
<INPUT type=hidden name="IMDG3"   value="<%=IMDG3%>">
<INPUT type=hidden name="IMDG4"   value="<%=IMDG4%>">						<!-- 2016/10/18 H.Yoshikawa Add -->
<INPUT type=hidden name="IMDG5"   value="<%=IMDG5%>">						<!-- 2016/10/18 H.Yoshikawa Add -->
<INPUT type=hidden name="TumiP"   value="<%=LPort%>">
<INPUT type=hidden name="UNNo1"   value="<%=UNNo1%>">
<INPUT type=hidden name="UNNo2"   value="<%=UNNo2%>">
<INPUT type=hidden name="UNNo3"   value="<%=UNNo3%>">
<INPUT type=hidden name="UNNo4"   value="<%=UNNo4%>">						<!-- 2016/10/18 H.Yoshikawa Add -->
<INPUT type=hidden name="UNNo5"   value="<%=UNNo5%>">						<!-- 2016/10/18 H.Yoshikawa Add -->
<INPUT type=hidden name="AgeP"    value="<%=DPort%>">

<INPUT type=hidden name="CMPcd0"  value="<%=CMPcd(0)%>">
<INPUT type=hidden name="CMPcd1"  value="<%=CMPcd1%>">						<!-- 2016/10/18 H.Yoshikawa Upd �i�����l�ύX�j-->
<INPUT type=hidden name="CMPcd2"  value="<%=CMPcd(2)%>">
<INPUT type=hidden name="CMPcd3"  value="<%=CMPcd(3)%>">
<INPUT type=hidden name="CMPcd4"  value="<%=CMPcd(4)%>">
<!-- 2009/03/10 R.Shibuta Add-S -->
<INPUT type=hidden name="TruckerSubName"  value="<%=TruckerSubName%>">
<!-- 2009/03/10 R.Shibuta Add-E -->
<INPUT type=hidden name="MrSk"    value="<%=MrSk%>">
<INPUT type=hidden name="HedId"   value="<%=HedId%>">
<INPUT type=hidden name="HFrom"   value="<%=HFrom%>">
<INPUT type=hidden name="Hmon"    value="<%=Hmon%>">
<INPUT type=hidden name="Hday"    value="<%=Hday%>">
<INPUT type=hidden name="TuSk"    value="<%=TuSk%>">
<INPUT type=hidden name="NextV"   value="<%=NextV%>">
<INPUT type=hidden name="OH"  value="<%=OH%>">
<INPUT type=hidden name="OWL" value="<%=OWL%>">
<INPUT type=hidden name="OWR" value="<%=OWR%>">
<INPUT type=hidden name="OLF" value="<%=OLF%>">
<INPUT type=hidden name="OLA" value="<%=OLA%>">
<INPUT type=hidden name="NiwataP" value="<%=NiwataP%>">
<INPUT type=hidden name="Operator" value="<%=Operator%>">
<INPUT type=hidden name="Comment1" value="<%=Comment1%>">
<INPUT type=hidden name="Comment2" value="<%=Comment2%>">
<INPUT type=hidden name="Comment3" value="<%=Comment1%>">
<INPUT type=hidden name="Mord" value="0">
<INPUT type=hidden name="SakuNo" value="<%=SakuNo%>">
<INPUT type=hidden name="UpFlag"  value="<%=UpFlag%>">
<INPUT type=hidden name="compFlag" value="<%=compFlag%>">
<INPUT type=hidden name="WkCNo"     value="<%=WkCNo%>">
<INPUT type=hidden name="TruckerFlag" value="<%=TruckerFlag%>">

<%'Add-s 2006/03/06 h.matsuda%>
<INPUT type=hidden name="ShipLineName" value="<%=ShipLineName%>">
<INPUT type=hidden name="shorimode" value="<%=shorimode%>">
<%'Add-e 2006/03/06 h.matsuda%>

<!-- 2016/10/13 H.Yoshikawa Add-S -->
<INPUT type=hidden name="TruckerTel" value="<%=TrkrTEL%>">
<INPUT type=hidden name="Shipper" value="<%=Shipper%>">
<INPUT type=hidden name="Forwarder" value="<%=Forwarder%>">
<INPUT type=hidden name="FwdStaff" value="<%=FwdrTan%>">
<INPUT type=hidden name="FwdTel" value="<%=FwdrTEL%>">
<!-- 2016/10/13 H.Yoshikawa Add-E -->
<!-- 2016/10/14 H.Yoshikawa Add-S -->
<INPUT type=hidden name="ShipCode" value="<%=VslCode%>">
<INPUT type=hidden name="VoyCtrl" value="<%=VoyCtrl%>">
<!-- 2016/10/14 H.Yoshikawa Add-E -->
<!-- 2016/10/18 H.Yoshikawa Add-S -->
<INPUT type=hidden name="ReportNo"  value="<%=ReportNo%>">
<INPUT type=hidden name="AsDry"     value="<%=AsDry%>">
<INPUT type=hidden name="Label1"    value="<%=Label1%>">
<INPUT type=hidden name="Label2"    value="<%=Label2%>">
<INPUT type=hidden name="Label3"    value="<%=Label3%>">
<INPUT type=hidden name="Label4"    value="<%=Label4%>">
<INPUT type=hidden name="Label5"    value="<%=Label5%>">
<INPUT type=hidden name="SubLabel1" value="<%=SubLabel1%>">
<INPUT type=hidden name="SubLabel2" value="<%=SubLabel2%>">
<INPUT type=hidden name="SubLabel3" value="<%=SubLabel3%>">
<INPUT type=hidden name="SubLabel4" value="<%=SubLabel4%>">
<INPUT type=hidden name="SubLabel5" value="<%=SubLabel5%>">
<INPUT type=hidden name="LqFlag1"   value="<%=LqFlag1%>">
<INPUT type=hidden name="LqFlag2"   value="<%=LqFlag2%>">
<INPUT type=hidden name="LqFlag3"   value="<%=LqFlag3%>">
<INPUT type=hidden name="LqFlag4"   value="<%=LqFlag4%>">
<INPUT type=hidden name="LqFlag5"   value="<%=LqFlag5%>">
<!-- 2016/10/18 H.Yoshikawa Add-E -->
<!-- 2016/11/02 H.Yoshikawa Add-S -->
<INPUT type=hidden name="NiukeNm"    value="<%=NiukeNm%>">
<INPUT type=hidden name="LPortNm"    value="<%=LPortNm%>">
<INPUT type=hidden name="DPortNm"    value="<%=DPortNm%>">
<INPUT type=hidden name="NiwataNm"   value="<%=NiwataNm%>">
<!-- 2016/11/02 H.Yoshikawa Add-E -->

<% Else %>
<CENTER>
  <DIV class=alert>
    <%= ErrerM %>
  </DIV>
  <P><INPUT type=button value="����" onClick="window.close()"></P>
</CENTER>
<% End If %>
</FORM>
<!-------------��ʏI���--------------------------->
</BODY></HTML>
