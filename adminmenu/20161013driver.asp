<%@LANGUAGE="VBSCRIPT" CODEPAGE="932"%>
<%
'**********************************************
'  �y�v���O�����h�c�z�@: driver.asp
'  �y�v���O�������́z�@: �h���C�o���F
'
'  �i�ύX�����j
'   2016/07/26    H.Yoshikawa    �쐬�i���O��񂩂�ڐA�j
'
'**********************************************
	
	Option Explicit
	Response.Expires = 0
    On Error Resume Next			'2016/07/28 H.Yoshikawa Add

	call CheckLoginH()
%>
<!--#include File="./Common/common.inc"-->

<%
		'���[�U�f�[�^����
	'�Z�b�V�����̗L�������`�F�b�N
	CheckLoginH
	
	'���[�U�f�[�^����
	dim USER, COMPcd
	dim v_GamenMode
	dim v_SearchType      '2013-06-28  Y.TAKAKUWA  Add
	dim v_DataCnt2
		
	dim Num2	
	dim strOrder2
	dim FieldName2	
	dim ObjRS2,ObjConn2
	
	dim wk
	dim i,x
	dim v_ItemName
	dim v_ItemValue
	dim abspage, pagecnt,reccnt	
	
	dim Arr_DriverID()
	dim Arr_Check()
	
	dim v_DriverInfo
	dim v_driverInfoChkFlg
	
	dim Arr_SendStat()
	'2013-06-27 Y.TAKAKUWA Add-S
	dim v_AdminMailAddress
	dim v_SendDetail
	'2013-06-27 Y.TAKAKUWA Add-E
	
	dim v_Msg			'2016/07/28 H.Yoshikawa  Add
	v_Msg = ""			'2016/07/28 H.Yoshikawa  Add
	
	const gcPage = 10
	
	'2016/07/28 H.Yoshikawa Mod Start
	'USER   = UCase(Session.Contents("userid"))
	'COMPcd = Session.Contents("COMPcd")  	
	USER = Trim(session("user_id"))
	'2016/07/28 H.Yoshikawa Mod End
	
	'----------------------------------------
    ' �ĕ`��O�̍��ڎ擾
   	'----------------------------------------			
	call LfGetRequestItem
	Call getAdminMailAddress()
			
	If v_GamenMode = "AP" then
	  call LfSendEmail()
	end if
	'2013-06-27 Y.TAKAKUWA Add-S
	If v_GamenMode = "D" Then
	  call LfDeleteDriverInfo()
	End If
	'2013-06-27 Y.TAKAKUWA Add-E
	Call getDriverInfo()
	
Function LfGetRequestItem()
   
	If Request.form("Gamen_Mode") = "" then
	  v_GamenMode = Request.QueryString("GamenMode")
	else
	  v_GamenMode = Request.form("Gamen_Mode")
	end if
	
	If Request.Form("SearchApprovalType") = "" Then
	  v_SearchType = Request.QueryString("SearchType")
	Else
	  v_SearchType = Request.Form("SearchApprovalType")
	End If
	if Trim(v_GamenMode) = "PS" then
	  'v_DriverInfo = Request.QueryString("driverInfo")
      v_DataCnt2 = Request.QueryString("DataCnt")
	else
	  v_DriverInfo = Request.Form("driverInfo")
      v_DataCnt2 = Request.form("DataCnt2")
    end if
    If v_DataCnt2 = "" then
      v_DataCnt2 = 0
    end if
	ReDimension(v_DataCnt2)
	
    v_SendDetail = Request.form("Send_Detail") 
    
	For i = 1 to (v_DataCnt2) - 1 
	    Arr_Check(i) = Trim(Request.form("chkInOut" & i))
        Arr_DriverID(i) = TRIM(Request.form("LODriverID" & i))
	Next
End Function

Function ReDimension(index)
   Redim Arr_Check(index)
   Redim Arr_DriverID(index)
End Function

Function getDriverInfo()
    dim StrSQL
 
    ConnectSvr ObjConn2, ObjRS2
    
    StrSQL = "SELECT * FROM LomDriver "
    StrSQL = StrSQL & " WHERE "
    '2016/07/27 H.Yoshikawa Mod Start
    'StrSQL = StrSQL & " HiTSUserID = '" & gfSQLEncode(USER) & "'"
    StrSQL = StrSQL & " 1 = 1 "
    '2016/07/27 H.Yoshikawa Mod End
    '2013-06-28 Y.TAKAKUWA Add-S
    If Trim(v_SearchType) = "S2" Then
      StrSQL = StrSQL & " AND (AcceptStatus = '1' OR AcceptStatus = '2')"
    Else
      'StrSQL = StrSQL & " AND AcceptStatus <> '1' AND AcceptStatus <> '3' " '2013-06-27 Y.TAKAKUWA Del
      StrSQL = StrSQL & " AND (AcceptStatus = '' OR AcceptStatus = NULL OR AcceptStatus='0')"
    End If
    StrSQL = StrSQL & " ORDER BY LomDriver.LoDriverID "
    '2013-06-28 Y.TAKAKUWA Add-E
    ObjRS2.PageSize = 50
	ObjRS2.CacheSize = 50
	ObjRS2.CursorLocation = 3
	ObjRS2.Open StrSQL, ObjConn2

	Num2 = ObjRS2.recordcount	
	
	if Num2 > 50 then 
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
	    ObjRS2.close
	    ObjConn2.close
		Exit Function
	end if			
	'�G���[�g���b�v����
    on error goto 0	

End Function

'2013-06-27 Y.TAKAKUWA Add-S
Function getAdminMailAddress()   
    dim StrSQL
    dim ObjConnLO, ObjRSLO
    dim ErrFlg
    dim iSeq
    
   ConnectSvr ObjConnLO, ObjRSLO	
   StrSQL = "SELECT * FROM mUsers " 
   StrSQL = StrSQL & " WHERE UserCode = '" & gfSQLEncode(USER) & "' "
   StrSQL = StrSQL & " ORDER BY UserCode"
   ObjRSLO.Open StrSQL, ObjConnLO
   While Not ObjRSLO.EOF
     v_AdminMailAddress = ObjRSLO("MailAddress")
     ObjRSLO.MoveNext
   Wend

   ObjRSLO.Close
   ObjConnLO.Close
   
End Function
'2013-06-27 Y.TAKAKUWA Add-E

'2013-06-27 Y.TAKAKUWA Add-S
Function LfDeleteDriverInfo()
  On Error Resume Next			'2016/07/28 H.Yoshikawa Add
  
    dim StrSQL
    dim ObjConnLO, ObjRSLO
    dim ErrFlg
    dim iSeq
	
    ConnectSvr ObjConnLO, ObjRSLO	
	'2016/07/27 H.Yoshikawa Del Start
	''2013-09-26 Y.TAKAKUWA Add-S
    'WriteLogH "b502", "�h���C�o���F�i�񏳔F�j���s", "01", ""
    ''2013-09-26 Y.TAKAKUWA Add-E
	'2016/07/27 H.Yoshikawa Del Start
    
	For i = 1 to v_DataCnt2-1
      If UCase(Trim(Arr_Check(i))) = "ON" Then
        'QUERY VALUES FOR Delete
        StrSQL = "SELECT * FROM LomDriver WHERE LoDriverID ='" & gfSQLEncode(Arr_DriverID(i))  & "'"
        ObjRSLO.Open StrSQL, ObjConnLO
        '2016/07/27 H.Yoshikawa Upd Start
        'If ObjRSLO.recordcount > 0 Then
        if Not ObjRSLO.EOF then
        '2016/07/27 H.Yoshikawa Upd End
            StrSQL = " DELETE FROM LomDriver WHERE "
            StrSQL = StrSQL & "LoDriverID='" & gfSQLEncode(Arr_DriverID(i)) & "'"        
            ObjConnLO.Execute(StrSQL)
            if err <> 0 then
	          v_Msg = "�h���C�o�̍폜�Ɏ��s���܂����B"
		    end if
	    end if
	    ObjRSLO.Close
	    
	    if Trim(v_Msg) = "" then
    		'2016/07/27 H.Yoshikawa Upd Start
		    'StrSQL = "SELECT * FROM LoGroupeDriver WHERE HiTSUserID='" & USER & "' AND LoDriverID ='" & Arr_DriverID(i)  & "'"
		    StrSQL = "SELECT * FROM LoGroupeDriver WHERE LoDriverID ='" & gfSQLEncode(Arr_DriverID(i))  & "'"
    		'2016/07/27 H.Yoshikawa Upd End
		    ObjRSLO.Open StrSQL, ObjConnLO
	        '2016/07/27 H.Yoshikawa Upd Start
	        'If ObjRSLO.recordcount > 0 Then
	        if Not ObjRSLO.EOF then
	        '2016/07/27 H.Yoshikawa Upd End
			'DELETE ALSO IN GROUP
			  StrSQL = " DELETE FROM LoGroupeDriver WHERE "
    		  '2016/07/27 H.Yoshikawa Upd Start
	          'StrSQL = StrSQL & " HiTSUserID='" & USER & "'"
	          'StrSQL = StrSQL & " AND LoDriverID='" & Trim(Arr_DriverID(i)) & "'"  
	          StrSQL = StrSQL & " LoDriverID='" & gfSQLEncode(Arr_DriverID(i)) & "'"  
    		  '2016/07/27 H.Yoshikawa Upd Start
	          ObjConnLO.Execute(StrSQL)
	          if err <> 0 then
		          v_Msg = "�h���C�o�̍폜�Ɏ��s���܂����B"
			  end if
			end if
			ObjRSLO.Close
		end if
		
      end if
    Next

    Set ObjRSLO = Nothing
    ObjConnLO.Close
    Set ObjConnLO = Nothing
    
End function
'2013-06-27 Y.TAKAKUWA Add-E

Function LfSendEmail()

  On Error Resume Next			'2016/07/28 H.Yoshikawa Add

  Dim ObjConnLO, ObjRSLO
  Dim ErrFlg
  Dim StrSQL

  Dim v_LoDriverName
  Dim v_LoDriverID
  Dim v_LoDriverPassword
  Dim v_LoEmailAddress
  Dim v_ErrMsg
  
  '2016/07/27 H.Yoshikawa Del Start
  'WriteLog fs, "", "", "", ""
  ''2013-09-26 Y.TAKAKUWA Add-S
  'WriteLog fs, "driver.asp", "�h���C�o���F�i�񏳔F�j���s", "01",""
  ''2013-09-26 Y.TAKAKUWA Add-E
  '2016/07/27 H.Yoshikawa Del End
  Dim objMail
  Dim rc	
  
  Dim svName 
  Dim mailTo 
  Dim mailFrom 
  Dim mailSubject
  Dim strBody
  Dim attachedFiles
  Dim sendDetailArr
  
  ConnectSvr ObjConnLO, ObjRSLO

  If Trim(v_SendDetail) <> "" Then
    sendDetailArr = Split(v_SendDetail,"/")
  End If
  If Ubound(sendDetailArr) > 0 Then
    If Trim(sendDetailArr(1)) <> "" Then
  'For i = 1 to v_DataCnt2-1                         '2013-06-27 Y.TAKAKUWA Del
  'For i = 1 to UBound(sendDetailArr)                 '2013-06-27 Y.TAKAKUWA Add
    'If UCase(Trim(Arr_Check(i))) = "ON" Then        '2013-06-27 Y.TAKAKUWA Del
      
      '�h���C�o���e�捞-S
      StrSQL = "SELECT * FROM LomDriver WHERE LoDriverID ='" & gfSQLEncode(sendDetailArr(0))  & "'"
      ObjRSLO.Open StrSQL, ObjConnLO
      
      '2016/07/27 H.Yoshikawa Upd Start
      'If ObjRSLO.recordcount > 0 Then
      if Not ObjRSLO.EOF then
      '2016/07/27 H.Yoshikawa Upd End
        v_LoDriverName = Trim(ObjRSLO("LoDriverName")) 
        v_LoDriverID  = Trim(ObjRSLO("LoDriverID"))
        v_LoDriverPassword  = MID(Trim(ObjRSLO("LoDriverPW")),1,1) & String(Len(Trim(ObjRSLO("LoDriverPW"))) -1,"*")
        v_LoEmailAddress  = Trim(ObjRSLO("MailAddress"))
      End If 
      '�h���C�o���e�捞-E
      ObjRSLO.Close			'2016/07/28 H.Yoshikawa Add
      
      svName = "153.150.17.106"
      'svName = "221.186.126.66"
      'svName = "192.168.17.243"
      mailTo = Trim(sendDetailArr(1))
      If Trim(sendDetailArr(2)) <> "" Then
        mailTo = mailTo & vbtab & sendDetailArr(2) 'Trim(v_LoEmailAddress)  '2013-06-27 Y.TAKAKUWA Upd
      End If
      mailFrom = "mrhits@hits-h.com" 
      mailSubject = "HiTS�h���C�o���F"
      attachedFiles = ""

      '���[�����e-S
      strBody = v_LoDriverName & " �l" & vbCrLf
      strBody = strBody & "HiTS���p�h���C�o�Ƃ��ď��F����܂����B" & vbCrLf  
      strBody = strBody & "�@�@�h���C�oID��" & v_LoDriverID  & vbCrLf 
      strBody = strBody & "�@�@�p�X���[�h��" & v_LoDriverPassword & vbCrLf
      strBody = strBody & "���̎菇�Ő�p�A�v�����C���X�g�[�����Ă��������B" & vbCrLf & vbCrLf
      strBody = strBody & "��Android�g�т̏ꍇ" & vbCrLf
      strBody = strBody & "�P�D�u�񋟌��s���̃A�v���v�̃C���X�g�[���������Ă��������B" & vbCrLf
      strBody = strBody & "�@(��j�ݒ聨�A�v���P�[�V�����ݒ�@���ɍ��ڂ�����܂�" & vbCrLf
      strBody = strBody & "�Q�D���L��URL���N���b�N���Đ�p�A�v�����_�E�����[�h���Ă��������B" & vbCrLf
      strBody = strBody & "�R�D�_�E�����[�h���I���܂�����A�C���X�g�[�����s���Ă��������B" & vbCrLf
      strBody = strBody & "�@(��j�ʒm�p�l���ɂ���uHiTS.apk�v���^�b�v���āA�u�C���X�g�[���v���^�b�v���Ă��������B" & vbCrLf
      strBody = strBody & "�S�D�_�E�����[�h�A�v���̈ꗗ�ɁuHiTS�v�A�C�R�����ǉ�����܂��B" & vbCrLf
      strBody = strBody & "�@�A�v�����N�����A�h���C�o�o�^�Őݒ肳�ꂽ�h���C�oID�ƃp�X���[�h����͂���΃��O�C���ł��܂��B" & vbCrLf
      strBody = strBody & "�T�D�C���X�g�[�����I����A�K�v�ɉ����āu�񋟌��s���̃A�v���v�̃C���X�g�[����s���ɖ߂��Ă��������B" & vbCrLf
      strBody = strBody & "�@https://www.hits-h.com/sp/android/download.html" & vbCrLf & vbCrLf
      strBody = strBody & "��iPhone�g�т̏ꍇ" & vbCrLf
      strBody = strBody & "�P�D���L��URL���^�b�v���Ă��������B" & vbCrLf
      strBody = strBody & "�Q�D�A�v���_�E�����[�h�p�̔F�؉�ʂ��\������܂��̂ŁA���L��ID�ƃp�X���[�h����͂��Ă��������B" & vbCrLf
      strBody = strBody & "�@�@���[�U���@: hits �i�S�ď������j" & vbCrLf
      strBody = strBody & "�@�@�p�X���[�h: Logi-app �i�ŏ��̂ݑ啶���j" & vbCrLf 
      strBody = strBody & "�@�@���h���C�o�o�^���̂��̂ł͂���܂���B"& vbCrLf
      strBody = strBody & "�R�D���͌�A�u�C���X�g�[���v���^�b�v���Ă��������B" & vbCrLf
      strBody = strBody & "�S�D�C���X�g�[��������A�z�[����ʂɁuHiTS�v�A�C�R�����ǉ�����܂��B" & vbCrLf
      strBody = strBody & "�@�A�v�����N�����A�h���C�o�o�^�Őݒ肳�ꂽ�h���C�oID�ƃp�X���[�h����͂���΃��O�C���ł��܂��B" & vbCrLf
      strBody = strBody & "�@https://www.hits-h.com/sp/iOS/download.html" & vbCrLf & vbCrLf
      strBody = strBody & "�����̃��[���Ɋւ��Ă̂��₢���킹�́A���LURL�̃y�[�W�ɂ���܂��A����܂ŁA���A�������肢�������܂��B" & vbCrLf
      strBody = strBody & "�@http://www.hits-h.com/request.asp" & vbCrLf
      strBody = strBody & "�����̃��[���ɂ��S������̂Ȃ����́A���̕����Ԉ���Ė{�T�[�r�X�Ƀ��[���A�h���X��o�^���ꂽ�\��������܂��B" &vbCrLf
      strBody = strBody & "�@���萔�����������܂����A���̃��[����j�����Ă��������܂��悤�A���肢�������܂��B"
      strBody = Server.HTMLEncode(strBody)
      '���[�����e-E
      
      If svName <> "" And mailTo <> "" Then
        Set ObjMail = Server.CreateObject("BASP21")
        rc=ObjMail.Sendmail(svName, mailTo, mailFrom, mailSubject, strBody, attachedFiles)

        if rc <> "" then
            StrSQL = " UPDATE LomDriver SET "
            '2013/07/30 Upd-S Fujiyama ���[�����M�G���[�͐��툵���ɂ���
            'StrSQL = StrSQL & "AcceptStatus='2', "                          'AcceptStatus
            StrSQL = StrSQL & "AcceptStatus='1', "                          'AcceptStatus
            '2013/07/30 Upd-E Fujiyama ���[�����M�G���[�͐��툵���ɂ���
            StrSQL = StrSQL & "UpdtTime='" & Now() & "',"                   'UpdtTime
            StrSQL = StrSQL & "UpdtPgCd='" & "PREDEF01" & "',"              'UpdtPgCd
            StrSQL = StrSQL & "UpdtTmnl='" & gfSQLEncode(USER) & "' "                   'UpdtTmnl
            'StrSQL = StrSQL & "MailAddress='" & gfSQLEncode(sendDetailArr(1)) & "' "     'MailAddress
            StrSQL = StrSQL & "WHERE LoDriverID='" & gfSQLEncode(sendDetailArr(0)) & "'"      '2013-06-27 Y.TAKAKUWA Upd
            ObjConnLO.Execute(StrSQL)
            if err <> 0 then
	          v_Msg = "���[�����M���ʂ̍X�V�Ɏ��s���܂����B�i���M���s�j"
	        end if
	        v_ErrMsg = "���s"
        else
          StrSQL = " UPDATE LomDriver SET "
          StrSQL = StrSQL & "AcceptStatus='1', "                          'AcceptStatus
          StrSQL = StrSQL & "UpdtTime='" & Now() & "',"                   'UpdtTime
          StrSQL = StrSQL & "UpdtPgCd='" & "PREDEF01" & "',"              'UpdtPgCd
          StrSQL = StrSQL & "UpdtTmnl='" & gfSQLEncode(USER) & "', "                    'UpdtTmnl
          StrSQL = StrSQL & "MailAddress='" & gfSQLEncode(sendDetailArr(1)) & "' "     'MailAddress
          StrSQL = StrSQL & "WHERE LoDriverID='" & gfSQLEncode(sendDetailArr(0)) & "'"     '2013-06-27 Y.TAKAKUWA Upd
          ObjConnLO.Execute(StrSQL)
          if err.number <> 0 then
	          v_Msg = "���[�����M���ʂ̍X�V�Ɏ��s���܂����B�i���M�����j"
	      end if
	      v_ErrMsg = "����"
        end if
      Else
          StrSQL = " UPDATE LomDriver SET "
          '2013/07/30 Upd-S Fujiyama ���[�����M�G���[�͐��툵���ɂ���
          'StrSQL = StrSQL & "AcceptStatus='2', "                          'AcceptStatus
          StrSQL = StrSQL & "AcceptStatus='1', "                          'AcceptStatus
          '2013/07/30 Upd-E Fujiyama
          StrSQL = StrSQL & "UpdtTime='" & Now() & "',"                   'UpdtTime
          StrSQL = StrSQL & "UpdtPgCd='" & "PREDEF01" & "',"              'UpdtPgCd
          StrSQL = StrSQL & "UpdtTmnl='" & gfSQLEncode(USER) & "' "                    'UpdtTmnl
          'StrSQL = StrSQL & "MailAddress='" & gfSQLEncode(sendDetailArr(1)) & "' "     'MailAddress
          StrSQL = StrSQL & "WHERE LoDriverID='" & gfSQLEncode(sendDetailArr(0)) & "'"      '2013-06-27 Y.TAKAKUWA Upd
          ObjConnLO.Execute(StrSQL)
          if err <> 0 then
	          v_Msg = "���[�����M���ʂ̍X�V�Ɏ��s���܂����B�i���M���s�j"
	      end if
	      v_ErrMsg = "���s"
      End If
      
      '2013-06-28 Y.TAKAKUWA Add-S
      If Trim(sendDetailArr(2)) <> "" Then
        If Trim(v_AdminMailAddress) <> Trim(sendDetailArr(2)) Then
          'Response.Write "Admin updated:" & v_AdminMailAddress
          StrSQL = "SELECT * FROM mUsers WHERE UserCode ='" & gfSQLEncode(USER)  & "'"
          ObjRSLO.Open StrSQL, ObjConnLO
          '2016/07/27 H.Yoshikawa Upd Start
          'If ObjRSLO.recordcount > 0 Then
          if Not ObjRSLO.EOF then
          '2016/07/27 H.Yoshikawa Upd End
            StrSQL = " UPDATE mUsers SET "
            StrSQL = StrSQL & "MailAddress='" & gfSQLEncode(sendDetailArr(2)) & "' "     'MailAddress
            StrSQL = StrSQL & "WHERE UserCode='" & gfSQLEncode(USER) & "'"
            ObjConnLO.Execute(StrSQL)
            if err <> 0 then
	          v_Msg = "���[���A�h���X�̍X�V�Ɏ��s���܂����B"
	        Else
	          v_AdminMailAddress = sendDetailArr(2)
	        end if
          End If
          ObjRSLO.Close			'2016/07/28 H.Yoshikawa Add
        End If
      End If
      '2013-06-28 Y.TAKAKUWA Add-E
    'End If
    
  'Next
  End If
  End If
  'Set objMsg = Nothing
  Set ObjRSLO = Nothing
  ObjConnLO.Close
  Set ObjConnLO = Nothing

End Function

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
		
		'2016/07/29 H.Yoshikawa Upd Start
		'�y�[�WIndex��ݒ�
		'PageIndex=Fix(page/gcPage)
		'if page mod gcPage=0 then
		'	PageIndex=PageIndex-1
		'End If
		'PageWkNo=((gcPage*PageIndex)+1)-gcPage
				
		'�擪�y�[�W��0��菬�����ꍇ��1��ݒ�
		'if PageWkNo<=0 Then
		'	PageWkNo=0
		'End If
		PageWkNo = page - 1
		'2016/07/29 H.Yoshikawa Upd End

		'�p�����[�^�ݒ�
		
	    'strParam="&InOutF=" & v_InOutFlag
		strParam=""
		'--- �������A���y�[�W�� 
		LastPage=pagecount		
		FirstPage=1
			
		if page>1 then
			response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & FirstPage & strParam & "&SearchType=" & v_SearchType & """>�ŏ���</a>"
			response.write "| &nbsp;"
		'2016/07/29 H.Yoshikawa Upd Start
			'if PageWkNo<>0 Then
			if PageWkNo>0 Then
		'2016/07/29 H.Yoshikawa Upd End
				response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & PageWkNo & strParam & "&SearchType=" & v_SearchType & """>�O��</a>"
			Else
				response.write "<font style='color:#FFFFFF;'>�O��</font>"
			End If
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
					response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & PageWkNo & strParam  & "&SearchType=" & v_SearchType & """ >&nbsp;" & PageWkNo & "</a>"
				End If
			Next
			response.write "| &nbsp;"
		End If
					
		if page<pagecount then
			'2016/07/29 H.Yoshikawa Upd Start
			'PageWkNo=PageWkNo+1
			PageWkNo=page+1
			'2016/07/29 H.Yoshikawa Upd End
			If PageWkNo<=LastPage Then
				response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & PageWkNo & strParam & "&SearchType=" & v_SearchType & """>����</a>"'
			Else
				response.write "<font style='color:#FFFFFF;'>����</font>"
			End If
			response.write "| &nbsp;"
			response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & LastPage & strParam & "&SearchType=" & v_SearchType & """>�Ō��</a>"'            
		else
			response.write "<font style='color:#FFFFFF;'>����</font>"
			response.write "| &nbsp;"
			response.write "<font style='color:#FFFFFF;'>�Ō��</font>"
		end if
	end if
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

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<TITLE>�g���s�r-�h���C�o���F</TITLE>
<link href="./Common/style.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
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
    FONT-SIZE: 10px;
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
    //position: relative;
    //top: expression(this.offsetParent.scrollTop);
}
#loading2 {
	font:bold 10px Verdana;
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
#footer {
 position: fixed;
 top: 100%;
 width: 100%;
}

.cmbUser option{
  height:10px;
}

thead.scrollHead,tbody.scrollBody{
  display:block;
}
tbody.scrollBody{
  height:30px;
  overflow-y:scroll;
}

/*������*/
td,th{
  table-layout:fixed;
}
.Col01{
  width:30px;
}
.Col02{
  width:100px;
}
.Col03{
  width:100px;
}
.Col04{
  width:80px;
}
.Col05{
  width:100px;
}
.Col06{
  width:60px;
}
.Col07{
  width:200px;
}

</STYLE>
<SCRIPT Language="JavaScript">
function finit(){
	//�f�[�^���p���ݒ�  
    document.frm.Gamen_Mode.value="<%=v_GamenMode%>";
    if("<%=v_Searchtype %>"=="S1"){
      document.getElementById("chk1").checked=true;
    }
    else{
      if("<%=v_Searchtype %>"=="S2"){
         document.getElementById("chk2").checked=true;
      }
    }
    if ("<%=v_Msg%>" != ""){
      alert("<%=v_Msg%>");
      return false;
    }

}

//�f�[�^�������ꍇ�̕\������
function view(){
	var sortedHeight;
	sortedHeight = 0;
	var vHeight;
	var obj2=document.getElementById("BDIV2");
	var rowHeight;
	
	if('<%=Num2%>'!='0'){
	  var rowHeightThead = getRowHeightThead();
	  var rowHeightTbody = getRowHeightTbody();
	  
	
	  if(rowHeightThead > 0){
	    rowHeightThead=rowHeightThead
	  }
	  if(rowHeightTbody > 0){
	    rowHeight=rowHeightTbody*15
	  }
	  rowHeight=rowHeight+rowHeightThead
    }
    else{
      rowHeight = 0;
      rowHeight=23*15;
    }
    
	if((document.body.offsetWidth-50) < 50){
		obj2.style.width=50;
		obj2.style.overflowX="auto";	 
	}else if((document.body.offsetWidth-50)  < 813){
		//obj2.style.width=document.body.offsetWidth-200;
		obj2.style.width=document.body.offsetWidth-220;
		obj2.style.overflowX="auto";
	}else{
		obj2.style.width=document.body.offsetWidth-220;
		obj2.style.overflowX="auto";
	}	
	
	if((document.body.offsetHeight-rowHeight) < 100){ 
	    if(obj2.clientWidth<obj2.scrollWidth)
	    {
	      obj2.style.height = 40;
		  obj2.style.overflowY = "auto";
	    }
	    else{
	      obj2.style.height = 25;
		  obj2.style.overflowY = "auto";
		}
	}else if((document.body.offsetHeight-rowHeight) < 395){
	    vHeight = rowHeight + 68;
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
    
	var obj3=document.getElementById("BDIV3");

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

function LockOnReg(){
	document.frm.Gamen_Mode.value = "I";
    document.frm.submit();
}


function fRSearch(){
	document.frm.Gamen_Mode.value = "S";
    document.frm.submit();
}

function fDelete(){
    var i;
    var chkFlag;
    chkFlag = 0;

    for(i=1;i<=parseInt(document.frm.DataCnt2.value)-1; i++){
      if(document.getElementById("checkbox" + i.toString()).checked == true){
         chkFlag = 1;
      }
    }
    
    if(chkFlag == 1){
      
      var show_modal = showModalDialog("dmlModal.asp?ActionType=D", window, "dialogWidth:400px; dialogHeight:200px; center:1; scroll: no;");
      if(show_modal){
        document.frm.Gamen_Mode.value = "D"
        document.frm.submit();
      }
    }
}

function fApproval(){
    var i;
    var chkFlag;
    var selectedCnt;
    var gTitle;
    var rowCnt;
    var show_modal;
    //show_modal = null;
    rowCnt = 0;
    chkFlag = 0;
    
    for(i=1;i<=parseInt(document.frm.DataCnt2.value)-1; i++){
      if(document.getElementById("checkbox" + i.toString()).checked == true){
         chkFlag = 1;
         selectedCnt = i;
         rowCnt = rowCnt + 1;
      }
    }

    if(rowCnt > 1){
       alert("�����I���͋�����Ă��܂���");
       return;
    }
    
    if(chkFlag == 1){
      if(document.getElementById("chk1").checked==true){
        gTitle = "S1";
      }
      else{
        gTitle = "S2";
      }
      show_modal = showModalDialog("dmlModal.asp?ActionType=S&SendTo=" + document.getElementById("InputMailAddress" + selectedCnt.toString()).value.toString() + "&DriverID=" + document.getElementById("InputDriverID" + selectedCnt.toString()).value.toString() + "&AdminMailAddress=" + '<%=v_AdminMailAddress%>' + "&GamenTitle=" + gTitle.toString(), window, "dialogWidth:400px; dialogHeight:250px; center:1; scroll: no;");
      if(show_modal != false && (typeof(show_modal) != 'undefined' && show_modal != null)){
        document.frm.Gamen_Mode.value = "AP"
        document.frm.Send_Detail.value = show_modal;
        document.frm.submit();
      }
      
      
      
    }
}
</script>


</HEAD>
<body class="bckcolor" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="finit();" onResize="">
<form name="frm" method="post">
<!-------------�������烁�C�����--------------------------->
<table class="main2" align="center" border="0" cellspacing="0" cellpadding="0" style="width:1020px;">
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <%
			DisplayHeader2("�h���C�o���F���")
		%>
		  <INPUT type="hidden" name="Gamen_Mode" size="9" readonly tabindex= -1>
		  <INPUT type=hidden name="Send_Detail" size="9" readonly tabindex= -1>
      </table>
 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
      <td style="width:90px">&nbsp;</td>
      <td>
         <div id="BDIV3">
            <!--Detail Start-->
            <table border="0" cellpadding="0" cellspacing="0" width="100%">
              <tr nowrap>
                <td nowrap>
                  <div style="margin-left:13px">
                  <table>
                    <tr>
                      <td nowrap><input type=radio name="SearchApprovalType" id="chk1" value="S1" checked=true onclick="fRSearch();"/>���F�҂��h���C�o�ꗗ</td>
                    </tr>
                    <tr>
                      <td nowrap><input type=radio name="SearchApprovalType" id="chk2" value="S2" onclick="fRSearch();"/>���F�h���C�o�ꗗ</td>
                    </tr>
                  </table>
                  </div>
                </td>
              </tr>
              <tr><td>&nbsp;</td></tr>
              <tr nowrap>
                <td nowrap><BR/></td>
              </tr>
              <tr align=right nowrap>
                <td width="100%" height="30" align=right nowrap>
                  <div style="margin-left:20px">
                  <table border="0" cellpadding="0" cellspacing="0">
                    <tr>
		               <td width="100%" align="center" nowrap>
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
		            </tr>
		          </table> 
		          </div>
                </td>
              </tr>
	          <tr>		
		        <td nowrap>
		          <div style="margin-left:20px">
			      <div id="BDIV2">
			   	    <% If Num2>0 Then%>
			   		<!--Driver List Start-->	
					<table border="1" cellpadding="0" cellspacing="0" width=100% id="TBInOut" height="300px">
						<thead class="scrollHead">
						   <!--HEADER INFORMATION START-->
							<tr>
								<th id="H2Col01" class="Col01" class="hlist" align="center" nowrap>�I��</th>
								<th id="H2Col02" class="Col02" class="hlist" nowrap>����</th>
								<th id="H2Col03" class="Col03" class="hlist" nowrap>�h���C�oID</th>
								<th id="H2Col04" class="Col04" class="hlist" nowrap>�p�X���[�h</th>
								<th id="H2Col05" class="Col05" class="hlist" nowrap>��Ж�</th>
								<th id="H2Col06" class="Col06" class="hlist" nowrap>�g�єԍ�</th>
								<th id="H2Col07" class="Col07" class="hlist" nowrap>���[���A�h���X</th>
							</tr>
						    <!--HEADER INFORMATION END-->
						</thead>
						<tbody class="scrollBody" height="300px">
						    <!--DETAIL INFORMATION START-->
                            <% 
								x = 1
								For i=1 To ObjRS2.PageSize
								 	If Not ObjRS2.EOF Then
									x = x + 1
							%>
							<tr bgcolor="#CCFFFF">	
							  <% v_ItemName = "chkInOut" + cstr(i) %>
							<td id="D2Col01" class="Col01" align="center" width="30" align="center" nowrap>
							  <input type="checkbox" name="<%= v_ItemName %>" id="checkbox<%=cstr(i)%>"><BR>
							</td>
							<td id="D2Col02" class="Col02" align="center" valign="middle" nowrap>
                              <%=gfHTMLEncode(ObjRS2("LoDriverName"))%><BR />
                            </td>
							<td id="D2Col03" class="Col03" align="center" valign="middle" nowrap>
                              <%=gfHTMLEncode(ObjRS2("LoDriverID"))%><BR />
                            </td>
							<td id="D2Col04" class="Col04" align="center" valign="middle" nowrap>
                              <%=MID(gfHTMLEncode(ObjRS2("LoDriverPW")),1,1) & String(Len(gfHTMLEncode(ObjRS2("LoDriverPW")))-1,"*")%>
                              <BR />
                            </td>
                            
							<td id="D2Col05" class="Col05" align="center" valign="middle" nowrap>
                              <%=gfHTMLEncode(ObjRS2("LoDriverCompany"))%><BR />
                            </td>
                            
                            <td id="D2Col06" class="Col06" align="center" valign="middle" nowrap>
                              <%=gfHTMLEncode(ObjRS2("PhoneNum"))%><BR />
                            </td>
                            
							<td id="D2Col07" class="Col07" align="center" valign="middle" nowrap>
                              <a href="mailto:<%=gfHTMLEncode(ObjRS2("MailAddress"))%>"><%=gfHTMLEncode(ObjRS2("MailAddress"))%></a>
                              <BR />
                            </td>
                            <% v_ItemName = "LODriverID" + cstr(i) %>
							<INPUT type=hidden name="<%=v_ItemName%>" id="InputDriverID<%=CStr(i)%>" value="<%=gfHTMLEncode(ObjRS2("LODriverID"))%>">
							
							<% v_ItemName = "MailAddress" + cstr(i) %>
							<INPUT type=hidden name="<%=v_ItemName%>" id="InputMailAddress<%=CStr(i)%>" value="<%=gfHTMLEncode(ObjRS2("MailAddress"))%>">
							
							</tr>
						    <% 
									ObjRS2.MoveNext 		
									End If
								Next	
							  ObjRS2.close    
						      ObjConn2.close
						    %>  
						    <!--DETAIL INFORMATION END-->	    									
						</tbody>								
					</table>
					<!--Work List End-->
					<INPUT type=hidden name="DataCnt2" value="<%=x%>">
					
				    <% Else %>
					  <table border="0" cellPadding="2" cellSpacing="0" id="NODATA">						
					    <TR class=bgw><TD nowrap style="color:Red;">�h���C�o�[�̓o�^������܂���</TD></TR>
					  </table>
				    <% End If %>		
			      </div>
			      </div>
		        </td>
	          </tr>
	          <tr><td>&nbsp;</td></tr>  
	          <tr>
				  <td>
				    <div>�����F���[���𑗐M���Ă��A�g�ђ[�����̖��f���[���ݒ蓙�Ŏ�M�ł��Ȃ��ꍇ������܂��B</div>
			        <div style="margin-left:12px; ">���F���[�����͂��Ȃ��ꍇ�́A�u���F�h���C�o�ꗗ�v�\������đ����Ă��������B</div>
			        <div style="margin-left:12px; margin-top:20px;">���F���[�����^�s�Ǘ��҂��w��h���C�o�֓]�����邱�ƂŁA�w��URL���A�v�����_�E�����[�h���邱�Ƃ��\�ƂȂ�܂��B</div>
				  </td>
	          </tr>
	          <tr><td>&nbsp;</td></tr>
	          <tr>		
		        <td>
				    <div>
					  <table border="0" cellpadding="2" cellspacing="0">
					  <tr>
					    <%
					      If v_SearchType="S2" Then
					         v_ItemValue = "���F���[�����đ�"
					      Else
					         v_ItemValue = "�I�������h���C�o�����F"
					      End If
					    %>
					    <%If Num2>0 then%>
					    <td><input type="button" name="Approval" onclick="fApproval();" value="<%=v_ItemValue%>"></td>
					    <%else%>
					    <td><input type="button" name="Approval" onclick="fApproval();" value="<%=v_ItemValue%>" disabled></td>
					    <%end if%>
					  </tr>
					  <%If v_SearchType <> "S2" then%>
					  <tr><td><br /></td></tr>
					  <tr>
					    <%If Num2>0 then%>
					    <td><input type="button" name="Delete" onclick="fDelete();" value="�I�������h���C�o�����F�����폜"></td>
					    <%else%>
					    <td><input type="button" name="Delete" onclick="fDelete();" value="�I�������h���C�o�����F�����폜" disabled></td>
					    <%end if%>
					  </tr>
					  <%End If%>
					  </table>
					</div>		
		        </td>
	          </tr> 
	          <tr><td>&nbsp;</td></tr>
	          <tr><td><center><a href="menu.asp">����</a></center></td></tr>   
	          <tr><td>&nbsp;</td></tr>
            </table>
          </div>
      </td>
      <td style="width:30px">&nbsp;</td>   
      </tr></table>  
    </td>
 </tr>
	<%
		DisplayFooter
	%>
</table>
</form>
</body>
</HTML>
