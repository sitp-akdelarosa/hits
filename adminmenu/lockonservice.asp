<%@LANGUAGE="VBSCRIPT" CODEPAGE="932"%>
<%
'**********************************************
'  �y�v���O�����h�c�z�@: 
'  �y�v���O�������́z�@: 
'
'  �i�ύX�����j
'
'**********************************************
	
	Option Explicit
	Response.Expires = 0

	call CheckLoginH()
%>
<!--#include File="./Common/common.inc"-->

<%
		'���[�U�f�[�^����
	dim USER, COMPcd  			
	dim v_GamenMode
	dim v_DataCnt2
	dim v_Msg
	
		
	dim Num2	
	dim strOrder2
	dim FieldName2	
	dim ObjRS2,ObjConn2
	dim ObjConnLO, ObjRSLO
	    
	dim wk
	dim i,x,j
	dim v_ItemName
	dim abspage, pagecnt,reccnt	
	
	dim Arr_DriverID()
	dim Arr_Check()
	dim Arr_HiTSUserID()
	dim Arr_NumOfRedCard()
	dim Arr_AcceptStatus()
	
	dim Arr_Users()
	dim v_CompanyName
	
	dim v_DriverInfo
	dim v_driverInfoChkFlg
	
	'Search Condition Start
	dim SSearchType
	dim SUserID
	dim SDriverName
	dim SDriverCompany
	dim SDriverID
	'Search Condition End
		
	const gcPage = 10
	
	USER   = UCase(Session.Contents("userid"))
	COMPcd = Session.Contents("COMPcd")  	
	
	'----------------------------------------
    ' �ĕ`��O�̍��ڎ擾
   	'----------------------------------------			
	call LfGetRequestItem
		
		
	If v_GamenMode = "S" Then
      Call getCompanyName()
	End If
		
	'�X�V
	If v_GamenMode = "U" Then		
		call LfUpdDriverLimit()
	End If
	
	Call getAllUser()
	Call getDriverInfo()
	
	
Function LfGetRequestItem()
   
	If Request.form("Gamen_Mode") = "" Then
	  v_GamenMode = Request.QueryString("GamenMode")
	Else
	  v_GamenMode = Request.form("Gamen_Mode")
	End If
	
	if Trim(v_GamenMode) = "PS" then
	  SDriverName = Request.QueryString("SDriverName")
	  SDriverCompany = Request.QueryString("SDriverCompany")
	  SDriverID = Request.QueryString("SDriverID")
	  SSearchType = Request.QueryString("searchType")
      v_DataCnt2 = Request.QueryString("Data_Cnt")
      SUserID = Request.QueryString("SUserID")
      v_CompanyName = Request.QueryString("FullName")
	else
	  SDriverName = Request.form("SDriverName")
	  SDriverCompany = Request.form("SDriverCompany")
	  SDriverID = Request.form("SDriverID")
	  SUserID = Request.Form("cmbUser")
	  v_DriverInfo = Request.Form("driverInfo")
	  SSearchType = Request.form("searchType")
      v_DataCnt2 = Request.form("Data_Cnt")
      v_CompanyName = Request.form("FullName")
    end if
    If v_DataCnt2 = "" then
      v_DataCnt2 = 0
    end if
    
    If Trim(SSearchType) = "" Then
      SSearchType = "1"
    End If
    
	ReDimension(v_DataCnt2)

	For i = 1 to (v_DataCnt2) - 1 
	    Arr_Check(i) = Trim(Request.form("chkInOut" & i))
        Arr_DriverID(i) = TRIM(Request.form("LODriverID" & i))
        Arr_HiTSUserID(i) = TRIM(Request.form("HiTSUserID" & i))
        Arr_NumOfRedCard(i) = TRIM(Request.form("NumOfRedCard" & i))
        Arr_AcceptStatus(i) = TRIM(Request.Form("AcceptStatusFlag" & i))
	Next
	
End Function

Function ReDimension(index)
   Redim Arr_Check(index)
   Redim Arr_DriverID(index)
   Redim Arr_HitsUserID(index)
   Redim Arr_NumOfRedCard(index)
   Redim Arr_AcceptStatus(index)
End Function

Function getAllUser()   
   dim StrSQL
   dim recCnt
   recCnt = 0
   Redim Arr_Users(recCnt)
   
   StrSQL = "SELECT * FROM mUsers " 
   StrSQL = StrSQL & " WHERE UserCode <> '' "
   StrSQL = StrSQL & " ORDER BY UserCode"
   ConnectSvr ObjConnLO, ObjRSLO
   
   ObjRSLO.Open StrSQL, ObjConnLo, 0, 1, 1
   
   While Not ObjRSLO.EOF
     Redim Preserve Arr_Users(recCnt)
     Arr_Users(recCnt) = ObjRSLO("UserCode")
     recCnt = recCnt + 1
     ObjRSLO.MoveNext
   Wend

   ObjRSLO.Close
   ObjConnLo.Close
End Function

Function getCompanyName
  dim StrSQL
  
  ConnectSvr ObjConnLO, ObjRSLO
  StrSQL = " SELECT * FROM mUsers "
  StrSQL = StrSQL & " WHERE UserCode='" & Trim(SUserID) & "'"
  ObjRSLO.Open StrSQL, ObjConnLo, 0, 1, 1
   
  While Not ObjRSLO.EOF
    v_CompanyName = Trim(ObjRSLO("FullName"))
    ObjRSLO.MoveNext
  Wend
  
  ObjRSLO.Close
  ObjConnLo.Close
  
End Function

Function getDriverInfo()
    dim StrSQL
 
    ConnectSvr ObjConn2, ObjRS2
    
    StrSQL = "SELECT DISTINCT LomDriver.*, mUsers.UserCode, mUsers.FullName, mUsers.TelNo "
    StrSQL = StrSQL & " FROM LomDriver "
    StrSQL = StrSQL & " LEFT JOIN mUsers ON LomDriver.HiTSUserID = mUsers.UserCode AND (LomDriver.AcceptStatus='1' OR LomDriver.AcceptStatus='3') "
    if Trim(SDriverName) <> "" or Trim(SDriverCompany) <> "" or Trim(SDriverID) <> "" or Trim(SUserID) <> "" then
       StrSQL = StrSQL  & " WHERE "
       
       If Trim(SUserID) <> "" Then
         StrSQL = StrSQL  & "LomDriver.HiTSUserID = '" & Trim(SUserID) & "' "
       End If
       
       if Trim(SDriverName) <> "" then
         If Trim(SUserID) <> "" Then
           StrSQL = StrSQL  & " AND "
         End If
         StrSQL = StrSQL  & "LomDriver.LoDriverName LIKE '%" & Trim(SDriverName) & "%' "
       end if
       
       if Trim(SDriverCompany) <> "" then
         if Trim(SDriverName) <> "" Or Trim(SUserID) <> "" then
            StrSQL = StrSQL  & " AND "  
         end if
         StrSQL = StrSQL  & "LomDriver.LoDriverCompany LIKE '%" & Trim(SDriverCompany) & "%' "
       end if
       if Trim(SDriverID) <> "" then
         if Trim(SDriverName) <> "" Or Trim(SDriverCompany) <> "" Or Trim(SUserID) <> "" then
            StrSQL = StrSQL  & " AND "  
         end if
         StrSQL = StrSQL  & "LomDriver.LoDriverID LIKE '%" & Trim(SDriverID) & "%' "
       end if
    end if
    
    if Trim(SDriverName) = "" and Trim(SDriverCompany) = "" and Trim(SDriverID) = "" And Trim(SUserID) = "" then
      StrSQL = StrSQL  & " WHERE "
    end if
   
    if Trim(SDriverID) <> "" Or Trim(SDriverName) <> "" Or Trim(SDriverCompany) <> "" Or Trim(SUserID) <> "" then
        StrSQL = StrSQL  & " AND "  
    end if
    
    If Trim(SSearchType) = "1" Then
      StrSQL = StrSQL & " LomDriver.AcceptStatus='3'"
    ElseIf Trim(SSearchType) = "2" Then
      StrSQL = StrSQL & " (LomDriver.AcceptStatus='3' OR LomDriver.AcceptStatus='1') " 
    End If 
	
	'Response.Write StrSQL
	
    ObjRS2.PageSize = 100
	ObjRS2.CacheSize = 100
	ObjRS2.CursorLocation = 3
	ObjRS2.Open StrSQL, ObjConn2, 0, 1, 1

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
	    ObjRS2.close
	    ObjConn2.close
		Exit Function
	end if			
	'�G���[�g���b�v����
    on error goto 0	

End Function

Function LfUpdDriverLimit()
    dim StrSQL
    dim ErrFlg
    dim iSeq
	
    ConnectSvr ObjConnLO, ObjRSLO	
	
	For i = 1 to v_DataCnt2-1
      If UCase(Trim(Arr_Check(i))) = "ON" Then
         If Trim(Arr_AcceptStatus(i)) <> "3" Then
           StrSQL = " UPDATE LomDriver SET "
           StrSQL = StrSQL & " AcceptStatus='3', "
           StrSQL = StrSQL & " NumOfRedCard='" & CInt(Trim(Arr_NumOfRedCard(i))) + 1 & "', "
           StrSQL = StrSQL & " UpdtTime='" & Now() & "' "
           StrSQL = StrSQL & " WHERE "
           StrSQL = StrSQL & " LoDriverID='" & Trim(Arr_DriverID(i)) & "'"           
           ObjConnLO.Execute StrSQL
           If err.number <> 0 then
             v_Msg = "�ύX�ł��܂���B"
             ObjConnLO.rollbacktrans
             ObjRSLO.Close
             ObjConnLO.Close
             Exit Function
		   End If
		 End If
	  Else
	     StrSQL = " UPDATE LomDriver SET "
         StrSQL = StrSQL & " AcceptStatus='1', "
         StrSQL = StrSQL & " UpdtTime='" & Now() & "' "
         StrSQL = StrSQL & " WHERE "
         StrSQL = StrSQL & " LoDriverID='" & Trim(Arr_DriverID(i)) & "'"           
         ObjConnLO.Execute StrSQL
         if err.number <> 0 then
           v_Msg = "�ύX�ł��܂���B"
           ObjConnLO.rollbacktrans
           ObjRSLO.Close
           ObjConnLO.Close
           Exit Function
		 end if
      End If
    Next
    ObjConnLO.Close
    
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
			response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & FirstPage & strParam & "&GamenMode=PS" & "&SDriverName=" & SDriverName & "&SDriverCompany=" & SDriverCompany & "&SDriverID=" & SDriverID & "&searchType=" & SSearchType & "&SUserID=" & SUserID & "&FullName=" & v_CompanyName & "&Data_Cnt=" & v_DataCnt2 & """>�ŏ���</a>"
			response.write "| &nbsp;"
			if PageWkNo<>0 Then
				response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & PageWkNo & strParam & "&GamenMode=PS" & "&SDriverName=" & SDriverName & "&SDriverCompany=" & SDriverCompany & "&SDriverID=" & SDriverID & "&searchType=" & SSearchType & "&SUserID=" & SUserID & "&FullName=" & v_CompanyName & "&Data_Cnt=" & v_DataCnt2 & """>�O��</a>"
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
					response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & PageWkNo & strParam & "&GamenMode=PS" & "&SDriverName=" & SDriverName & "&SDriverCompany=" & SDriverCompany & "&SDriverID=" & SDriverID & "&searchType=" & SSearchType & "&SUserID=" & SUserID & "&FullName=" & v_CompanyName & "&Data_Cnt=" & v_DataCnt2 & """ >&nbsp;" & PageWkNo & "</a>"
				End If
			Next
			response.write "| &nbsp;"
		End If
					
		if page<pagecount-1 then
			PageWkNo=PageWkNo+1
			If PageWkNo<=LastPage Then
				response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & PageWkNo & strParam & "&GamenMode=PS" & "&SDriverName=" & SDriverName & "&SDriverCompany=" & SDriverCompany & "&SDriverID=" & SDriverID & "&searchType=" & SSearchType & "&SUserID=" & SUserID & "&FullName=" & v_CompanyName & "&Data_Cnt=" & v_DataCnt2 & """>����</a>"'
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
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<TITLE>�g���s�r-���b�N�I���T�[�r�X����</TITLE>
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
    position: relative;
    top: expression(this.offsetParent.scrollTop);
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

</STYLE>
<SCRIPT Language="JavaScript">
function finit(){
	//�f�[�^���p���ݒ�
	  
    document.frm.Gamen_Mode.value="<%=v_GamenMode%>";  
    document.frm.cmbUser.value="<%=SUserID%>";	
    document.frm.FullName.value="<%=v_CompanyName%>"
    if("<%=SSearchType %>"=="1"){
      document.getElementById("chk1").checked=true;
    }
    else{
      if("<%=SSearchType %>"=="2"){
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
    
	if((document.body.offsetWidth) < 240){
		obj2.style.width=62;
		obj2.style.overflowX="auto";	 
	}else if((document.body.offsetWidth)<1037){
		obj2.style.width=document.body.offsetWidth-140;
		obj2.style.overflowX="auto";
	}//else{
	//	obj2.style.width=document.body.offsetWidth-380;
	//	obj2.style.overflowX="auto";
	//}	

	if((document.body.offsetHeight-rowHeight) < 130){ 
	    if(obj2.clientWidth<obj2.scrollWidth)
	    {
	      obj2.style.height = 62;
		  obj2.style.overflowY = "auto";
	    }
	    else{
	      obj2.style.height = 47;
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
function ClearSItem()
{
  document.frm.SDriverName.value="";
  document.frm.SDriverCompany.value="";
  document.frm.SDriverID.value="";
}

function ClearSItem2()
{
  document.frm.cmbUser.value="";
  document.frm.SDriverName.value="";
  document.frm.SDriverCompany.value="";
  document.frm.SDriverID.value="";
  document.frm.Gamen_Mode.value = "S";
  document.frm.submit();
}

function fSearch(){
	document.frm.Gamen_Mode.value = "S";
	ClearSItem();
    document.frm.submit();
}

function fSearchDriver(){
   	document.frm.Gamen_Mode.value = "SD";
    document.frm.submit();

}
</script>

<script type="text/vbscript">

Public Sub Update_onclick()
  Dim chkFlag
  Dim x
  Dim i
  
  chkFlag = 1

  if chkFlag=1 then
    x=MsgBox("        ��̓��e�œo�^���܂��B" & vbCrLf & "             ��낵���ł����H" ,4,"Confirm")
    if x = vbYes then
      document.frm.Gamen_Mode.value = "U"
      document.frm.submit()
    end if
  end if

End Sub
</script>


</HEAD>
<body class="bckcolor" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="finit();view();" onResize="view();">
<form name="frm" action="lockonservice.asp" method="post">
<!-------------�������烍�O�C�����͉��--------------------------->
<table class="main2" align="center" border="0" cellspacing="0" cellpadding="0" style="width:1020px;">
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <%
			DisplayHeader2("���b�N�I���T�[�r�X����")
		%>
		  <INPUT type="hidden" name="Gamen_Mode" size="9" maxlength="1"  readonly tabindex= -1>
		  <INPUT type=hidden name="FullName" readonly tabindex=-1>        
      </table>
 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
      <td style="width:90px">&nbsp;</td>
      <td>
         <div id="BDIV3">
            <!--Detail Start-->
            <table border="0" cellpadding="0" cellspacing="0" width="100%">
              <tr>
                <td>
                  <DIV style="width:230px; padding:10px;background-color:#FFCCFF; text-align:center;">�h���C�o�ꗗ����̐��� on/off</DIV>
                </td>
              </tr>
              <tr><td>&nbsp;</td></tr>
              <tr nowrap>
                <td nowrap>
                  <div style="margin-left:13px">
                  <table>
                    <tr>
                      <td nowrap><input type="radio" name="searchType" id="chk1" value="1" checked=true onclick="ClearSItem2();">�������h���C�o�̂ݕ\��</td>
                      <td nowrap>&nbsp;&nbsp;&nbsp;&nbsp;</td>
                      <td nowrap><input type="radio" name="searchType" id="chk2" value="2" onclick="ClearSItem2();">�S�h���C�o��\��</td>
                    </tr>
                  </table>
                  </div>
                </td>
              </tr>
              <tr><td>&nbsp;</td></tr>
              <tr>
                <td>
                  <div style="margin-left:20px">
                  <table>
                    <td nowrap style="width:150px">���F���[�U�w��I��</td><td>&nbsp;&nbsp;&nbsp;&nbsp;</td> 
	                  <td nowrap>
	                    <select name="cmbUser" style="width:116px;" onchange="fSearch();" class="cmbUser">
	                      <OPTION VALUE = '' SELECTED></OPTION>
	                      <% If UBound(Arr_Users) > 0 Then%>
	                        <% For j=0 to UBound(Arr_Users) %>
	                          <OPTION VALUE = '<%=Arr_Users(j)%>'><%=Arr_Users(j) %></OPTION>
	                        <% Next %>
	                      <% End If %>
			            </select>
	                  </td>
	                  <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
	                  <td nowrap>
	                    <%=v_CompanyName%>
	                  </td>
                    </td>
                  </table>
                  </div>
                </td>
              </tr>
              
              <tr nowrap>
                <td nowrap>
                  <div style="margin-left:20px">
                  <!--Search Conditions Start-->
                    <table>
                      <tr>
                        <td nowrap style="width:150px">�h���C�o���w��</td><td>&nbsp;&nbsp;&nbsp;&nbsp;</td><td nowrap><input type="text" name="SDriverName" value="<%=SDriverName%>" onfocus="this.select();"></td>
                      </tr>
                      <tr>
                        <td nowrap style="width:150px">�h���C�o�h�c�w��</td><td>&nbsp;&nbsp;&nbsp;&nbsp;</td><td nowrap><input type="text" name="SDriverID" value="<%=SDriverID%>" onfocus="this.select();"></td>
                      </tr>
                      <tr>
                        <td nowrap style="width:150px">��Ж��w��</td><td>&nbsp;&nbsp;&nbsp;&nbsp;</td><td nowrap><input type="text" name="SDriverCompany" value="<%=SDriverCompany%>" onfocus="this.select();"></td>
                        <td width=100 align=right nowrap><input type="button" name="Button" value="�\���X�V" onClick="fSearchDriver();"></td>
                        <td width=150 align=right nowrap>��������v������</td>
                      </tr>
                    </table>
                  <!--Search Conditions End-->
                  </div>    
                </td>
              </tr>
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
					<table border="1" cellpadding="0" cellspacing="0" width=100% id="TBInOut">				
					  <thead>
					    <!--HEADER INFORMATION START-->
					    <tr>
					      <th id="H2Col01" class="hlist" align="center" nowrap>������</th>
					      <th id="H2Col02" class="hlist" nowrap>�h���C�o��</th>
					      <th id="H2Col03" class="hlist" nowrap>�h���C�oID</th>								
					      <th id="H2Col04" class="hlist" nowrap>��Ж�</th>
					      <th id="H2Col05" class="hlist" nowrap>���F���[�U</th>
						  <th id="H2Col06" class="hlist" nowrap>���F���[�U��Ж�</th>
						  <th id="H2Col07" class="hlist" nowrap>���F���[�U<BR/>�d�b�ԍ�</th>
						  <th id="H2Col08" class="hlist" nowrap>����<BR />���{����</th>		
						  <th id="H2Col09" class="hlist" nowrap>����<BR />��</th>																									
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
						  <td id="D2Col01" align="center" nowrap style="width:60px;">
						    <% If Trim(ObjRS2("AcceptStatus")) = "3" Then %>
						      <input type="checkbox" name="<%= v_ItemName %>" checked="true"><BR />
						    <% Else %>
						      <input type="checkbox" name="<%= v_ItemName %>" ><BR />
						    <% End If %>
						  </td>
						  <td id="D2Col02" align="center" valign="middle" nowrap>
                            <%=Trim(ObjRS2("LoDriverName"))%><BR />
                          </td>
						  <td id="D2Col03" align="center" valign="middle" nowrap>
                            <%=Trim(ObjRS2("LoDriverID"))%><BR />
                          </td>
						  <td id="D2Col04" align="center" valign="middle" nowrap>
                            <%=Trim(ObjRS2("LoDriverCompany"))%><BR />
                          </td>

                          <td id="D2Col05" align="center" valign="middle" nowrap>
                            <%=Trim(ObjRS2("UserCode"))%><BR />
                          </td>
                          <td id="D2Col06" align="center" valign="middle" nowrap>
                            <%=Trim(ObjRS2("FullName"))%><BR />
                          </td>
                          <td id="D2Col07" align="center" valign="middle" nowrap style="width:100px;">
                            <%=Trim(ObjRS2("TelNo"))%><BR />
                          </td>
                          <td id="D2Col08" align="center" valign="middle" nowrap style="width:130px;">
                            <% If Trim(ObjRS2("AcceptStatus")) = "3" Then %>
                               <%=Trim(Year(ObjRS2("UpdtTime"))) & "-" & Right("0" & Trim(Month(ObjRS2("UpdtTime"))),2) & "-" & Right("0" & Trim(Day(ObjRS2("UpdtTime"))),2) & " " & Trim(FormatDateTime(ObjRS2("UpdtTime"),4))%>                            
                            <% End If%><BR />
                          </td>
                          <td id="D2Col09" align="center" valign="middle" nowrap style="width:40px;">
                            <%=Trim(ObjRS2("NumOfRedCard"))%>
                            <BR />
                          </td>

                          <% v_ItemName = "LODriverID" + cstr(i) %>
						  <INPUT type=hidden name="<%=v_ItemName%>" value="<%=Trim(ObjRS2("LODriverID"))%>">	
						  <% v_ItemName = "HiTSUserID" + cstr(i) %>
						  <INPUT type=hidden name="<%=v_ItemName%>" value="<%=Trim(ObjRS2("HiTSUserID"))%>">
						  <% v_ItemName = "NumOfRedCard" + cstr(i) %>
						  <INPUT type=hidden name="<%=v_ItemName%>" value="<%=Trim(ObjRS2("NumOfRedCard"))%>">
						  <% v_ItemName = "AcceptStatusFlag" + cstr(i) %>
						  <INPUT type=hidden name="<%=v_ItemName%>" value="<%=Trim(ObjRS2("AcceptStatus"))%>">
						  
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
					<!--Driver List End-->
					<INPUT type="hidden" name="Data_Cnt" value="<%=x%>">
					
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
	              <div style="margin-left:20px">
	              <p style="color:#993300;">���u�������v�̃`�F�b�N�����邱�ƂŁA�h���C�o�̓��b�N�I�����ł��Ȃ��Ȃ�܂��B<BR />
                        &nbsp;&nbsp;&nbsp;&nbsp;�܂��A���̃h���C�o�ւ̎w�����o���Ȃ��Ȃ�܂��B<BR />
                        &nbsp;&nbsp;&nbsp;&nbsp;��������ꍇ�̓`�F�b�N���O���܂��B
                  </p>
                  </div>
	            </td>
	          </tr>
	          <tr><td>&nbsp;</td></tr>
	          <tr>		
		        <td>
		          <div>
		            <div style="margin-left:20px">
			        <table border="0" cellpadding="2" cellspacing="0">
			          <tr>
			            <td><input type="button" name="Update"  value="��̓��e�œo�^"></td>
			          </tr>
			        </table>
			        </div>	
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
