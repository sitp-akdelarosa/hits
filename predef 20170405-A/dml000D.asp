<%
'**********************************************
'  �y�v���O�����h�c�z�@: dml000D
'  �y�v���O�������́z�@: �h���C�o�ꗗ�ƍ폜
'
'  �i�ύX�����j
'   2013-03-15   Y.TAKAKUWA   �쐬
'   2013-09-26   Y.TAKAKUWA   ���M���O�̋@�\��ǉ�
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
<%
	'**********************************************

	'�Z�b�V�����̗L�������`�F�b�N
	CheckLoginH
	
	'���[�U�f�[�^����
	dim USER, COMPcd  			
	dim v_GamenMode
	dim v_DataCnt2
		
	dim Num2	
	dim strOrder2
	dim FieldName2	
	dim ObjRS2,ObjConn2
	
	dim wk
	dim i,x
	dim v_ItemName
	dim abspage, pagecnt,reccnt	
	
	dim Arr_DriverID()
	dim Arr_Check()
	dim Arr_HiTSUserID()
	
	dim v_DriverInfo
	dim v_driverInfoChkFlg
	
	'Search Condition Start
	dim SDriverName
	dim SDriverCompany
	dim SDriverID
	'Search Condition End
	
	'Option Condition Start
	dim v_LogOnUser
	'Option Condition End
		
	dim gSQLStr				'2016/07/29 H.Yoshikawa Add
	
	const gcPage = 10
	
	USER   = UCase(Session.Contents("userid"))
	COMPcd = Session.Contents("COMPcd")  	
	
	'----------------------------------------
    ' �ĕ`��O�̍��ڎ擾
   	'----------------------------------------			
	call LfGetRequestItem
		
	'�o�^
	if v_GamenMode = "I" then		
		'call LfUpdLOInfo()
	end if

	'Delete Driver
	If v_GamenMode = "D" then
	  call LfDeleteLoDriverInfo()
	end if
	
	'Delete Driver to own group
	If v_GamenMode = "DO" then
	  call LfDeleteLoGroupDriverInfo()
	end if
	
	Call getDriverInfo()
	
	'2016/07/29 H.Yoshikawa Add Start
	If v_GamenMode = "DL" then
	  'CSV�_�E�����[�h
	  call LfDownLoadCSV()
	end if
	'2016/07/29 H.Yoshikawa Add End
	
Function LfGetRequestItem()
   
	If Request.form("Gamen_Mode") = "" then
	  v_GamenMode = Request.QueryString("GamenMode")
	else
	  v_GamenMode = Request.form("Gamen_Mode")
	end if
	
	v_LogOnUser = ""												'2016/07/29 H.Yoshikawa Add
	if Trim(v_GamenMode) = "PS" then
	  SDriverName = Request.QueryString("SDriverName")
	  SDriverCompany = Request.QueryString("SDriverCompany")
	  SDriverID = Request.QueryString("SDriverID")
	  'v_DriverInfo = Request.QueryString("driverInfo")
	  'v_LogOnUser = Request.QueryString("LogOnUser")				'2016/07/29 H.Yoshikawa Del
      v_DataCnt2 = Request.QueryString("DataCnt")
	else
	  SDriverName = Request.form("SDriverName")
	  SDriverCompany = Request.form("SDriverCompany")
	  SDriverID = Request.form("SDriverID")
	  v_DriverInfo = Request.Form("driverInfo")
	  'v_LogOnUser = Request.form("selectCompany")					'2016/07/29 H.Yoshikawa Del
      v_DataCnt2 = Request.form("DataCnt2")
    end if
    If v_DataCnt2 = "" then
      v_DataCnt2 = 0
    end if
	ReDimension(v_DataCnt2)

	For i = 1 to (v_DataCnt2) - 1 
	    Arr_Check(i) = Trim(Request.form("chkInOut" & i))
        Arr_DriverID(i) = TRIM(Request.form("LODriverID" & i))
        Arr_HiTSUserID(i) = TRIM(Request.form("HiTSUserID" & i))
	Next
End Function

Function ReDimension(index)
   Redim Arr_Check(index)
   Redim Arr_DriverID(index)
   Redim Arr_HitsUserID(index)
End Function

Function getDriverInfo()
    dim StrSQL
 
    ConnDBH ObjConn2, ObjRS2
    
    '2013-09-26 Y.TAKAKUWA Add-S
    WriteLogH "b503", "�h���C�o�ꗗ�\��", "01",""
    '2013-09-26 Y.TAKAKUWA Add-E
    
    StrSQL = "SELECT DISTINCT LomDriver.LoDriverID, LomDriver.LoDriverHeadID, LomDriver.LoDriverName, LomDriver.LoDriverPW, LomDriver.LoDriverCompany, LomDriver.MailAddress, LomDriver.HiTSUserID "
    StrSQL = StrSQL & ", LomDriver.PhoneNum "			'2016/07/29 H.Yoshikawa Add
    StrSQL = StrSQL & " FROM LomDriver "
    '2016/07/29 H.Yoshikawa Del Start
    'If Trim(v_LogOnUser) = "2" Then
    '  StrSQL = StrSQL  & " INNER JOIN LoGroupeDriver ON LomDriver.LoDriverID = LoGroupeDriver.LoDriverID AND LoGroupeDriver.HiTSUserID = '" & USER & "'"
    '  StrSQL = StrSQL  & " RIGHT JOIN "
    '  StrSQL = StrSQL  & " ( "
    '  StrSQL = StrSQL  & "   SELECT DISTINCT LoGroupeDriver.LoGroupID FROM LomGroup "
    '  'StrSQL = StrSQL  & "   FROM LomDriver "
    '  'StrSQL = StrSQL  & "   INNER JOIN LoGroupeDriver ON LomDriver.LoDriverID = LoGroupeDriver.LoDriverID "
    '  StrSQL = StrSQL & "    INNER JOIN LoGroupeDriver ON  LomGroup.LoGroupID=LoGroupeDriver.LoGroupID AND LomGroup.HiTSUserID = LoGroupeDriver.HiTSUserID "
    '  StrSQL = StrSQL  & "   WHERE LomGroup.HiTSUserID = '" & USER & "'"
    '  StrSQL = StrSQL  & " ) OWNGROUP ON LoGroupeDriver.LoGroupID = OWNGROUP.LoGroupID "
    'End If
    '2016/07/29 H.Yoshikawa Del End
    
    StrSQL = StrSQL & " WHERE LomDriver.AcceptStatus='1' "
    
    if Trim(SDriverName) <> "" or Trim(SDriverCompany) <> "" or Trim(SDriverID) <> "" then
       'StrSQL = StrSQL  & " WHERE "
       if Trim(SDriverName) <> "" then
         StrSQL = StrSQL  & "AND LomDriver.LoDriverName LIKE '%" & Trim(SDriverName) & "%'"
       end if
       
       if Trim(SDriverCompany) <> "" then
         'if Trim(SDriverName) <> "" then
            StrSQL = StrSQL  & " AND "  
         'end if
         StrSQL = StrSQL  & " LomDriver.LoDriverCompany LIKE '%" & Trim(SDriverCompany) & "%'"
       end if
       if Trim(SDriverID) <> "" then
         'if Trim(SDriverName) <> "" Or Trim(SDriverCompany) <> "" then
            StrSQL = StrSQL  & " AND "  
         'end if
         StrSQL = StrSQL  & " LomDriver.LoDriverID LIKE '%" & Trim(SDriverID) & "%'"
       end if
    end if
    
    'if Trim(SDriverName) = "" and Trim(SDriverCompany) = "" and Trim(SDriverID) = "" then
    '  StrSQL = StrSQL  & " WHERE "
    'end if
    
    '2016/07/29 H.Yoshikawa Upd Start
    'if Trim(v_LogOnUser) = "1" Or Trim(v_LogOnUser) = "" then
    '  'if Trim(SDriverID) <> "" Or Trim(SDriverName) <> "" Or Trim(SDriverCompany) <> "" then
    '  '  StrSQL = StrSQL  & " AND "  
    '  'end if
    '  StrSQL = StrSQL  & "AND LomDriver.HiTSUserID = '" & USER & "'"
    'elseIf Trim(v_LogOnUser) = "2" then
    '  'if Trim(SDriverID) <> "" Or Trim(SDriverName) <> "" Or Trim(SDriverCompany) <> "" then
    '  '  StrSQL = StrSQL  & " AND "  
    '  'end if
    '  StrSQL = StrSQL  & "AND LomDriver.HiTSUserID<>'" & USER & "'"
    'end if
    if Session.Contents("UType") <> "0" then
    	StrSQL = StrSQL  & "AND LomDriver.HiTSUserID = '" & USER & "'"
    end if
	'2016/07/29 H.Yoshikawa Upd End

	'Response.Write StrSQL
	
	gSQLStr = StrSQL				'2016/07/29 H.Yoshikawa Add
	
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

Function LfDeleteLoDriverInfo()
    dim StrSQL
    dim ObjConnLO, ObjRSLO
    dim ErrFlg
    dim iSeq
	
    ConnDBH ObjConnLO, ObjRSLO	
	WriteLogH "", "", "", ""
	
	For i = 1 to v_DataCnt2-1
      If UCase(Trim(Arr_Check(i))) = "ON" Then
        'QUERY VALUES FOR Delete
        StrSQL = "SELECT * FROM LomDriver WHERE LoDriverID ='" & Arr_DriverID(i)  & "'"
                                                        
        ObjRSLO.Open StrSQL, ObjConnLO
        If ObjRSLO.recordcount > 0 then
            StrSQL = " DELETE FROM LomDriver WHERE "
            StrSQL = StrSQL & "LoDriverID='" & Trim(Arr_DriverID(i)) & "'"        
            ObjConnLO.Execute(StrSQL)
            if err <> 0 then
			  Set ObjRSLO = Nothing				
			  jampErrerPDB ObjConnLO,"2","b107","01","","104","SQL:<BR>"& strSQL
		    end if
		    
		    'DELETE ALSO IN GROUP
		    StrSQL = " DELETE FROM LoGroupeDriver WHERE "
		    if Session.Contents("UType") <> "0" then								'2016/07/29 H.Yoshikawa Add
              StrSQL = StrSQL & " HiTSUserID='" & USER & "'" 
              StrSQL = StrSQL & " AND LoDriverID='" & Trim(Arr_DriverID(i)) & "'"  
            '2016/07/29 H.Yoshikawa Add Start
            else
              StrSQL = StrSQL & " LoDriverID='" & Trim(Arr_DriverID(i)) & "'"  
            end if
            '2016/07/29 H.Yoshikawa Add End
            ObjConnLO.Execute(StrSQL)
            if err <> 0 then
			  Set ObjRSLO = Nothing				
			  jampErrerPDB ObjConnLO,"2","b107","01","","104","SQL:<BR>"& strSQL
		    end if
	    end if
	    ObjRSLO.Close
      end if
    Next
    
    DisConnDBH ObjConnLO, ObjRSLO	'DB�ؒf
    
End function


Function LfDeleteLoGroupDriverInfo
    dim StrSQL
    dim ObjConnLO, ObjRSLO
    dim ErrFlg
    dim iSeq
    dim Arr_GroupID()
    Redim Arr_GroupID(0)
	iSeq = 0
	
    ConnDBH ObjConnLO, ObjRSLO	
	WriteLogH "", "", "", ""
	
	'QUERY OWN DRIVER GROUP-S 
	StrSQL = "SELECT DISTINCT LoGroupID FROM LomDriver "
    StrSQL = StrSQL & "INNER JOIN LomGroup ON LomDriver.HiTSUserID = LomGroup.HiTSUserID "
    StrSQL = StrSQL & "WHERE LomDriver.HiTSUserID = '" & USER & "' "
    'Response.Write StrSQL
	'QUERY OWN DRIVER GROUP-E
	ObjRSLO.Open StrSQL, ObjConnLO
	If ObjRSLO.recordcount > 0 then
	    Redim Arr_GroupID(ObjRSLO.recordcount)
	    iSeq = 0
	    While Not ObjRSLO.EOF
	      Arr_GroupID(iSeq) = Trim(ObjRSLO("LoGroupID")) 
	      iSeq = iSeq + 1
	      ObjRSLO.MoveNext
	    Wend
	    
	end if
	ObjRSLO.Close
	
	For x = 0 to UBound(Arr_GroupID)    
	  For i = 1 to v_DataCnt2-1
	    If UCase(Trim(Arr_Check(i))) = "ON" Then
          'QUERY VALUES FOR Delete
          StrSQL = "SELECT * FROM LoGroupeDriver WHERE HiTSUserID ='" & USER & "'" &_
                                                 " AND LoGroupID ='" & Trim(Arr_GroupID(x)) & "'" &_
                                                 " AND LoDriverID ='" & Trim(Arr_DriverID(i)) & "'" 
        
          'Response.Write StrSQL                                               
          ObjRSLO.Open StrSQL, ObjConnLO
          If ObjRSLO.recordcount > 0 Then
            StrSQL = " DELETE FROM LoGroupeDriver WHERE "
            StrSQL = StrSQL & "      HiTSUserID ='" & USER  & "'"&_
                               " AND LoGroupID ='" & Trim(Arr_GroupID(x)) & "'" &_
                               " AND LoDriverID ='" & Trim(Arr_DriverID(i)) & "'"
            ObjConnLO.Execute(StrSQL)
            If err <> 0 Then
			  Set ObjRSLO = Nothing				
			  jampErrerPDB ObjConnLO,"2","b107","01","","104","SQL:<BR>"& strSQL
		    End If
	      End If
	      ObjRSLO.Close
	    End If
      Next
    Next
  
	DisConnDBH ObjConnLO, ObjRSLO	'DB�ؒf

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
		
	    'strParam="&InOutF=" & v_InOutFlag
		strParam=""
		'--- �������A���y�[�W�� 
		LastPage=pagecount		
		FirstPage=1
			
		if page>1 then
			response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & FirstPage & strParam & "&GamenMode=PS" & "&SDriverName=" & SDriverName & "&SDriverCompany=" & SDriverCompany & "&SDriverID=" & SDriverID & "&LogOnUser=" & v_LogOnUser & "&DataCnt=" & v_DataCnt2 & """>�ŏ���</a>"
			response.write "| &nbsp;"
			if PageWkNo<>0 Then
				response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & PageWkNo & strParam & "&GamenMode=PS" & "&SDriverName=" & SDriverName & "&SDriverCompany=" & SDriverCompany & "&SDriverID=" & SDriverID & "&LogOnUser=" & v_LogOnUser & "&DataCnt=" & v_DataCnt2 & """>�O��</a>"
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
					response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & PageWkNo & strParam & "&GamenMode=PS" & "&SDriverName=" & SDriverName & "&SDriverCompany=" & SDriverCompany & "&SDriverID=" & SDriverID & "&LogOnUser=" & v_LogOnUser & "&DataCnt=" & v_DataCnt2 & """ >&nbsp;" & PageWkNo & "</a>"
				End If
			Next
			response.write "| &nbsp;"
		End If
					
		if page<pagecount-1 then
			PageWkNo=PageWkNo+1
			If PageWkNo<=LastPage Then
				response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & PageWkNo & strParam & "&GamenMode=PS" & "&SDriverName=" & SDriverName & "&SDriverCompany=" & SDriverCompany & "&SDriverID=" & SDriverID & "&LogOnUser=" & v_LogOnUser & "&DataCnt=" & v_DataCnt2 & """>����</a>"'
			Else
				response.write "<font style='color:#FFFFFF;'>����</font>"
			End If
			response.write "| &nbsp;"
			response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & LastPage & strParam & """>�Ō��</a>"'            
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

'2016/07/29 H.Yoshikawa Add Start
function LfDownLoadCSV
    dim ObjConnLO, ObjRSLO

	on error resume next
	err.clear
	
    ConnDBH ObjConnLO, ObjRSLO	
    ObjRSLO.Open gSQLStr, ObjConnLO
	'WriteLogH "", "", "", ""

	if not ObjRSLO.eof then
		Response.Addheader "Content-Disposition", "attachment ; filename=driver.csv"
		Response.Contenttype = "application/x-binary"
'		Response.Buffer = False
		'���x���o��
		response.write("�h���C�oID,����,��Ж�,�g�єԍ�,���[���A�h���X,�w���Ǘ���")
		response.write(vbcrlf)
		
		'�f�[�^�o��
		while not ObjRSLO.eof
			response.write("""" & Trim(ObjRSLO("LoDriverID")) & """")
			response.write(",""" & Trim(ObjRSLO("LoDriverName")) & """")
			response.write(",""" & Trim(ObjRSLO("LoDriverCompany")) & """")
			response.write(",""" & Trim(ObjRSLO("PhoneNum")) & """")
			response.write(",""" & Trim(ObjRSLO("MailAddress")) & """")
			response.write(",""" & Trim(ObjRSLO("HiTSUserID")) & """")
			response.write(vbcrlf)
			ObjRSLO.movenext
		wend
	end if
	
	DisConnDBH ObjConnLO, ObjRSLO	'DB�ؒf
	
	Response.End

end function
'2016/07/29 H.Yoshikawa Add End
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE></TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR>
<STYLE>
body {
    background-image:url('../gif/back.gif');
    margin:0;
    padding:0;
}

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
#footer {
 position: fixed;
 top: 100%;
 width: 100%;
}

</style>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT Language="JavaScript">

function finit(){
	//�f�[�^���p���ݒ�  
    document.frm.Gamen_Mode.value="<%=v_GamenMode%>";  
    //alert("<%=v_LogOnUser %>");	
    //2016/07/29 H.Yoshikawa Del Start
    //if("<%=v_LogOnUser %>"=="1"){
    //  document.getElementById("chk1").checked=true;
    //}
    //else{
    //  if("<%=v_LogOnUser %>"=="2"){
    //     document.getElementById("chk2").checked=true;
    //  }
    //}
    //2016/07/29 H.Yoshikawa Del End
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

//	if((document.body.offsetWidth-10)  < 880){
//		obj3.style.width=document.body.offsetWidth-10;
//		obj3.style.overflowX="auto";
//	}
//	else{
//		obj3.style.width=document.body.offsetWidth-10;
//		obj3.style.overflowX="auto";
//	}
//    if((document.body.offsetHeight) > 15 ){
//	  obj3.style.height=document.body.offsetHeight-15;
//	  obj3.style.overflowY="auto";
//	}
//	else{
//	  obj3.style.height=document.body.offsetHeight;
//	  obj3.style.overflowY="auto";	
//	}
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

function fDel()
{
  var chkFlag;
  chkFlag = 0;
  for(i=1; i <= (parseInt(document.frm.DataCnt2.value)-1); i++){
    obj = eval("document.frm.chkInOut" + i);
    if (obj.checked==true) {
       chkFlag = 1;
	}
  }
  
  if(chkFlag==1){
  var msg = confirm("�I�������h���C�o���폜���܂��B��낵���ł����H",1,4,0);
    if(msg == true){
      document.frm.Gamen_Mode.value = "D";
      document.frm.submit();
    }
  }

}
function ClearSItem()
{
  document.frm.SDriverName.value=""
  document.frm.SDriverCompany.value=""
  document.frm.SDriverID.value=""
  document.frm.Gamen_Mode.value = "S";
  document.frm.submit();
}

function refreshParent() 
{
    if('<%=v_GamenMode%>'=='D'){
      //window.opener.location.reload(true);
      alert("���̉�ʂɔ��f����ɂ́A���̉�ʍ��́u�R���e�i���b�N�v���j���[���N���b�N���čĕ`�悵�Ă�������");
    }
}

//2016/07/29 H.Yoshikawa Add Start
function fcsv(){
	if(document.frm.DataCnt2.value < 0){
		alert("�Y���f�[�^�����݂��܂���B");
		return;
	}
	
	document.frm.Gamen_Mode.value = "DL";
	document.frm.submit();
}
//2016/07/29 H.Yoshikawa Add End

</SCRIPT>
<script type="text/vbscript">
Public Sub Delete_onclick()
  Dim chkFlag
  Dim x
  Dim i
  
  chkFlag = 0
  x=MsgBox(document.frm.DataCnt2.value,0)
  for i = 1 to CInt(document.frm.DataCnt2.value-1)
     If document.frm.elements("chkInOut" + CStr(i)).checked then
       chkFlag = 1
     end if
  Next
  
  if chkFlag=1 then
    x=MsgBox("�I�������h���C�o���폜���܂��B��낵���ł����H",4,"Confirm")
    if x = vbYes then
      document.frm.Gamen_Mode.value = "D"
      document.frm.submit()
    end if
  end if

End Sub

Public Sub Delete2_onclick()
  Dim chkFlag
  Dim x
  Dim i
  
  chkFlag = 0
  
  for i = 1 to CInt(document.frm.DataCnt2.value-1)
     If document.frm.elements("chkInOut" + CStr(i)).checked then
       chkFlag = 1
     end if
  Next
  
  if chkFlag=1 then
    x=MsgBox("�I���������Ѓh���C�o�����ׂĂ̎��ЃO���[�v���珜�O���܂��B" & vbCrLf & "                  �i�h���C�o��񎩑͎̂c��܂��j�B" & vbCrLf & "                          ��낵���ł����H",4,"Confirm")
    if x = vbYes then
      document.frm.Gamen_Mode.value = "DO"
      document.frm.submit()
    end if
  end if

End Sub

</script>

</HEAD>
<BODY onLoad="finit();view();refreshParent();" onResize="view();">
<form name="frm" method="post">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
   <td rowspan="2">
	 <IMG height=73 src="Image/Driver.png" width=507>
   </td>
   <td height="19" bgcolor="#000099" align="right">
	 <IMG height=19 src="../gif/logo_hits_ver2_1.gif">
   </td>
</tr>
<tr>
   <td align="right" width="100%" height="45">
		
   </td>
</tr>
</table>

<table width="100%" height="82%" border="0" cellspacing="0" cellpadding="0">
<tr><td width="120" nowrap>&nbsp;</td><td>
  <div id="BDIV3" style="width: 300px;">

  <!--Hidden Values Start-->
  <INPUT type=hidden name="Gamen_Mode" size="9" readonly tabindex= -1>
  <!--Hidden Values End-->
  <!--Added Start-->

  <table border="0" cellpadding="0" cellspacing="0">
  	<!-- 2016/07/29 H.Yoshikawa Del Start -->
  	<!--
    <tr nowrap><td nowrap>
      
      <table>
        <tr><td nowrap><input type="radio" name="selectCompany" id="chk1" value="1" checked=true onclick="ClearSItem();">���Џ��F�h���C�o��\��</td></tr>
        <tr><td nowrap><input type="radio" name="selectCompany" id="chk2" value="2" onclick="ClearSItem();">���Џ��F�h���C�o��\��</td></tr>
      </table>
    </td></tr>
    -->
  	<!-- 2016/07/29 H.Yoshikawa Del End -->
    <tr nowrap><td nowrap>&nbsp;</td></tr>
    <tr nowrap><td nowrap>
    <div style="margin-left:30px;">
    <!--Search Conditions Start-->
      <table>
        <tr>
          <td nowrap>���O����</td><td nowrap><input type="text" name="SDriverName" value="<%=SDriverName%>" onfocus="this.select();"></td>
        </tr>
        <tr>
          <td nowrap>��Ж�����</td><td nowrap><input type="text" name="SDriverCompany" value="<%=SDriverCompany%>"  onfocus="this.select();"></td>
        </tr>
        <tr>
          <td nowrap>�h���C�o�h�c����</td><td nowrap><input type="text" name="SDriverID" value="<%=SDriverID%>"  onfocus="this.select();"></td>
          <td width=100 align=right nowrap><input type="button" name="Button" value="�\���X�V" onClick="fRSearch();"></td>
          <td width=150 align=right nowrap>��������v������</td>
        </tr>
      </table>
    <!--Search Conditions End-->
    </div>
    </td></tr>
    <tr nowrap><td nowrap><BR/></td></tr>
    <tr align=right nowrap>
      <td width="100%" height="30" align=right nowrap>
        
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
 
      </td>
    </tr>

	<tr>		
		<!--Place Here Start-->
		<td nowrap>
			<div id="BDIV2">
			   	<% If Num2>0 Then%>
			   		<!--Work List Start-->	
					<table border="1" cellpadding="0" cellspacing="0" width=100% id="TBInOut">				
						<thead>
						   <!--HEADER INFORMATION START-->
							<tr>
								<th id="H2Col01" class="hlist" align="center" nowrap>�I��</th>
								<th id="H2Col02" class="hlist" nowrap>����</th>
								<th id="H2Col03" class="hlist" nowrap>�h���C�oID</th>
								<%If v_LogOnUser <> "2" then %>								
								<th id="H2Col04" class="hlist" nowrap>�p�X���[�h</th>
								<%End If%>
								  <th id="H2Col05" class="hlist" nowrap>��Ж�</th>
								<%'If v_LogOnUser <> "2" then %>
								<!--
								  <th id="H2Col06" class="hlist" nowrap>���[���A�h���X</th>	
								-->
								<%'End If%>																																	
							</tr>
						    <!--HEADER INFORMATION END-->
						</thead>																
						<tbody>
						    <!--DETAIL INFORMATION START-->
                            <% 
								x = 1 							
								For i=1 To ObjRS2.PageSize
								 	If Not ObjRS2.EOF Then
									x = x + 1
							%>
							<tr bgcolor="#CCFFFF">	
							  <% v_ItemName = "chkInOut" + cstr(i) %>
							<td id="D2Col01" align="center" width="30" align="center" nowrap>
							  <input type="checkbox" name="<%= v_ItemName %>"><BR>
							</td>
							<td id="D2Col02" align="center" valign="middle" nowrap>
                              <%=Trim(ObjRS2("LoDriverName"))%><BR />
                            </td>
							<td id="D2Col03" align="center" valign="middle" nowrap>
                              <%=Trim(ObjRS2("LoDriverID"))%><BR />
                            </td>
                            <%If v_LogOnUser <> "2" then %>
							<td id="Td1" align="center" valign="middle" nowrap>
                              <%=MID(Trim(ObjRS2("LoDriverPW")),1,1) & String(Len(Trim(ObjRS2("LoDriverPW")))-1,"*")%>
                              <BR />
                            </td>
                            <%end if%>
							<td id="D2Col04" align="center" valign="middle" nowrap>
                              <%=Trim(ObjRS2("LoDriverCompany"))%><BR />
                            </td>
                            <%'If v_LogOnUser <> "2" then %>
							<!--
							<td id="D2Col05" align="center" valign="middle" nowrap>
                              <a href="mailto:<%=Trim(ObjRS2("MailAddress"))%>"><%=Trim(ObjRS2("MailAddress"))%></a>
                              <BR />
                            </td>
                            -->
                             <%'end if%>
                            <% v_ItemName = "LODriverID" + cstr(i) %>
							<INPUT type=hidden name="<%=v_ItemName%>" value="<%=Trim(ObjRS2("LODriverID"))%>">
							
							<% v_ItemName = "HiTSUserID" + cstr(i) %>
							<INPUT type=hidden name="<%=v_ItemName%>" value="<%=Trim(ObjRS2("HiTSUserID"))%>">
							
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
				    
					<table border="1" cellPadding="2" cellSpacing="0" id="NODATA">						
					  <TR class=bgw><TD nowrap>�h���C�o�[�̓o�^������܂���</TD></TR>
					</table>
					
				<% End If %>		
			</div>
		</td>
		<!--Place Here End-->
	</tr>
	<tr><td>&nbsp;</td></tr>
	<tr>		
		<td>
		    <div>
			  <table border="0" cellpadding="2" cellspacing="0">
			  <tr>
			    <!-- 2016/07/29 H.Yoshikawa Del Start -->
			    <!--
			    <%If Trim(v_LogOnUser)="2" then%>
			      <td><input type="button" name="Delete2"  value="�I�������h���C�o���폜"></td>
			    <%else%>
			      <td><input type="button" name="Delete"  value="�I�������h���C�o���폜"></td>
			    <%end if%>
			    -->
			    <!-- 2016/07/29 H.Yoshikawa Del End -->
			    <!-- 2016/07/29 H.Yoshikawa Add Start -->
			    <td><input type="button" name="Delete3"  value="�I�������h���C�o���폜" onclick="fDel();"></td>
			    <td><input type="button" name="CSV"  value="CSV�o��" onclick="fcsv();"></td>
			    <!-- 2016/07/29 H.Yoshikawa Add End -->
			  </tr>
			  </table>	
			</div>		
		</td>
	</tr>    
  </table>
  <!--Added End-->
  </div>
</td></tr>  
</table>
</form>
  
<div id="footer">
<MAP name=map>
<AREA coords=22,0,0,22,105,22,105,0 href="http://www.hits-h.com/index.asp" target="_parent" shape=POLY>
</MAP>
<TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
  <TR>
    <TD align=right vAlign=bottom>
      <IMG border=0 height=23 src="Image/b-home.gif" useMap="#map" width=105></TD></TR>
  <TR><TD colspan=2 bgColor=#000099 height=10></TD></TR>
</TABLE>
</div>  
</BODY>
</HTML>