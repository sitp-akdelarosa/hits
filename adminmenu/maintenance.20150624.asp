<%@LANGUAGE="VBSCRIPT" CODEPAGE="932"%>
<%
'**********************************************
'  �y�v���O�����h�c�z�@: maintenance.asp
'  �y�v���O�������́z�@: ���m�点�����e�i���X
'
'  �i�ύX�����j
'
'**********************************************
	Option Explicit
	Response.Expires = 0

	call CheckLoginH()
	
%>
<!--#include File="./Common/common.inc"-->
<!--#include File="./Common/upload.inc"-->
<!--#include File="./Common/sendmail.inc"-->
<!--#include File="../inform/common.inc"-->
<SCRIPT src="./Common/function.js" type=text/javascript></SCRIPT>
<%
	'��ʍ��ڕϐ�
	dim obj
	dim buf
	dim i
	dim v_Mode
	dim totalByte
	dim v_Data_Cnt	
	dim v_ItemName
	dim v_FileName
	dim v_MailSig
	dim v_SendFlag
	dim v_Msg
	dim v_FocusItem
	dim strDir
	
	dim Arr_Files
	dim Arr_MailSig
	dim Arr_DelFlag
	dim Arr_Name
	
	redim Arr_Files(0)
	redim Arr_MailSig(0)
	redim Arr_DelFlag(0)
	
	Set obj=server.createobject("basp21")
	on error resume next
	totalByte = Request.TotalBytes
	buf	= Request.BinaryRead(totalByte)	
	
	v_Mode = ""
	v_Data_Cnt = 0
	
	'----------------------------------------
    ' �ĕ`��O�̍��ڎ擾
    '----------------------------------------	
	call LfRequestItem() 	
	
	if v_Mode = "I" and v_FileName <> "" then		
		strDir = LfgetDirPath
		if strDir <> "" then
			if not gfUploadFile3(v_Filename,"txtFileUpload",strDir) then
				v_Msg = "�A�b�v���[�h�͎��s���܂����B"			
			end if	 
		end if
	end if
	
	if v_Mode = "U" then
		call LfsetMessageTxt()
	end if
	
	if v_Mode = "M" then
		call LfsetSendFlag()
		call LfSendMail()
	end if
	
	if v_Mode = "D" then		
		call LfDeleteFiles()
	end if
	
	call LfReadDir()
	call LfgetMessageTxt()
	call LfgetSendFlag()
		
'-----------------------------
'   �`��O�̉�ʍ��ڂ��擾
'-----------------------------
function LfRequestItem()	
	v_Mode = obj.Form(buf,"Gamen_Mode")
	v_Data_Cnt = obj.Form(buf,"Data_Cnt")
	v_MailSig = obj.Form(buf,"txtMailMsg")	
	'v_MailSig = Replace(v_MailSig, chr(10), "<br />")	
	'v_MailSig = Replace(v_MailSig, chr(13), "<br />")
		
	v_SendFlag = obj.Form(buf,"SendFlag")		
	v_FileName = obj.FormFileName(buf,"txtFileUpload")
	
	if CInt(v_Data_Cnt) > 0 then
		redim preserve Arr_DelFlag(v_Data_Cnt)
		redim preserve Arr_Files(v_Data_Cnt)
		
		for i = 1 to v_Data_Cnt
			Arr_DelFlag(i) = obj.Form(buf,"DelFlag" & i) 
			Arr_Files(i) = obj.Form(buf,"FileNames" & i) 
		next
	end if
end function

function LfgetDirPath()
	dim param(2)
	dim param2(2)	
	dim v_IniFile
	
	LfgetDirPath = ""
	
	call getUploadIni(param,v_Inform)
	
	v_IniFile = param(0) & "inform.ini"
	
	if v_IniFile <> "" then
		call getInformIni(param2,v_IniFile)
	end if
	
	LfgetDirPath = param2(0)
end function
'-----------------------------
'   (�t�@�C��)���擾
'-----------------------------
function LfReadDir()
	dim param(2)
	dim ObjFSO,ObjTS,myfile
	dim cnt
	dim strDir
	dim filecount	
	dim tempName
	dim work 
	dim nName
	
	
	cnt = 0	
	strDir = LfgetDirPath

	'''ini�t�@�C���̒l�̓ǂݍ���
	getIni param
	
	if strDir <> "" then
		Set ObjFSO = Server.CreateObject("Scripting.FileSystemObject")
		Set ObjTS = ObjFSO.GetFolder(strDir)		
		filecount = ObjTS.files.count		
		
		redim Arr_Files(filecount)	'2009/07/22 C.Pestano Add
		
		for each myfile in ObjTS.Files						
			cnt = cnt + 1
			
			if(DateDiff("d",myfile.DateLastModified,Date)<=CInt(param(1))) then '''�����|�쐬��<=�������
				Arr_Files(cnt)= Left(myfile.DateLastModified,4) & "�N" & Mid(myfile.DateLastModified,6,2) & "��" & Mid(myfile.DateLastModified,9,2) & "��" & "|" & Mid(myfile.DateLastModified,12,8) & "|" & myfile.Name & "|1"
			else		'''�����|�쐬��>�������
				Arr_Files(cnt)= Left(myfile.DateLastModified,4) & "�N" & Mid(myfile.DateLastModified,6,2) & "��" & Mid(myfile.DateLastModified,9,2) & "��" & "|" & Mid(myfile.DateLastModified,12,8) & "|" & myfile.Name & "|0"
			end if						
		next			
		
		Redim Arr_Name(cnt,3)
		
		for tempName = 1 to filecount
			for nName = (tempName + 1) to UBOUND(Arr_Files)
				if strComp(Arr_Files(tempName),Arr_Files(nName),1)<0 then 
    	            work = Arr_Files(tempName) 
        	 	    Arr_Files(tempName) = Arr_Files(nName) 
                	Arr_Files(nName) = work
	            end if 
			next
		next
	end if
	
	v_Data_Cnt = cnt
	
	Set ObjTS = nothing
	Set ObjFSO = nothing
end function

'--------------------------------
'   ���m�点mail���M�̖{�����擾
'--------------------------------
function LfgetMessageTxt()
	dim ObjFSO,ObjTS,tmpStr,cnt
	
	cnt = 0
	
	Set ObjFSO = Server.CreateObject("Scripting.FileSystemObject")
	'--- �t�@�C�����J���i�ǂݎ���p�j ---
	Set ObjTS = ObjFSO.OpenTextFile(Server.Mappath("./ini/mail.txt"),1,false,-1)

	'--- �t�@�C���f�[�^�̓Ǎ��� ---
	Do Until ObjTS.AtEndofStream		
		cnt = cnt + 1		
		redim preserve Arr_MailSig(cnt)
		Arr_MailSig(cnt) = ObjTS.ReadLine & Chr(10)
		if Arr_MailSig(cnt) = "" then
			Arr_MailSig(cnt) = Chr(10) 
		end if		
	Loop
	
	ObjTS.Close
	Set ObjTS = Nothing
	Set ObjFSO = Nothing
end function

'--------------------------------
'   ���m�点mail���M�̖{����ύX
'--------------------------------
function LfsetMessageTxt()
	dim ObjFSO,ObjTS,tmpStr
	dim cnt
	
	Set ObjFSO = Server.CreateObject("Scripting.FileSystemObject")
	'--- �t�@�C�����J���i�ǂݎ���p�j ---
	Set ObjTS = ObjFSO.OpenTextFile(Server.Mappath("./ini/mail.txt"),2,false,-1)

	'tmpStr = Split(v_MailSig,"<br />")
	
	'For i = 1 to UBOUND(tmpStr)	
		ObjTS.WriteLine(gfHTMLEncode(v_MailSig))	
	'Next
	
	ObjTS.Close
	Set ObjTS = Nothing
	Set ObjFSO = Nothing
end function

'---------------------------
'   ���M��w���ύX
'---------------------------
function LfsetSendFlag()
	dim ObjFSO,ObjTS
	dim Arr_Temp
	dim cnt
	dim strTemp
	
	cnt = 0
	redim Arr_Temp(0)	 	
   
	'--- �t�@�C�����J���i�ǂݎ���p�j ---
	Set ObjFSO = Server.CreateObject("Scripting.FileSystemObject")	
	Set ObjTS = ObjFSO.OpenTextFile(Server.Mappath("./ini/admin.ini"),1,false)
	
	'--- �t�@�C���f�[�^�̓Ǎ��� ---
	Do Until ObjTS.AtEndofStream		
		cnt = cnt + 1
		strTemp = ObjTS.ReadLine
		redim preserve Arr_Temp(cnt)		
		if Mid(Trim(strTemp),1,14) = v_Mail then		
			strTemp = v_Mail & "=" & v_SendFlag
		end if				
		Arr_Temp(cnt) = strTemp
	Loop
	
	ObjTS.Close	
	
	'--- �t�@�C�����J���i�ǂݎ���p�j ---
	Set ObjTS = ObjFSO.OpenTextFile(Server.Mappath("./ini/admin.ini"),2,false)
	
	for i = 1 to UBOUND(Arr_Temp)	
		ObjTS.WriteLine(Arr_Temp(i))
	next
	
	ObjTS.Close
	
	Set ObjTS = Nothing
	Set ObjFSO = Nothing
end function

'---------------------------
'   ���M��w����擾
'---------------------------
function LfgetSendFlag()
	dim param(2)	
	
	call getUploadIni(param,v_Mail)
	
	v_SendFlag = param(0)	
end function

'---------------------------
'   �폜(�t�@�C��)
'---------------------------
function LfDeleteFiles()
	dim ObjFSO,ObjTS	
	dim wkfilename
	dim param(2)
	dim param2(2)	
	dim v_IniFile
	
	call getUploadIni(param,v_Inform)
	
	v_IniFile = param(0) & "inform.ini"
	
	if v_IniFile <> "" then
		call getInformIni(param2,v_IniFile)
	end if	
	
	Set ObjFSO = Server.CreateObject("Scripting.FileSystemObject")		
	
	for i = 1 to UBOUND(Arr_DelFlag)
		if Arr_DelFlag(i) = "1" and param2(0) <> "" then		
			wkfilename = param2(0) & Arr_Files(i)
			gfDeleteFile(wkfilename)
		end if
	next
			
	ObjTS.Close
	Set ObjTS = Nothing
	Set ObjFSO = Nothing
end function

'--------------------------
'   ���[�����M
'--------------------------
function LfSendMail()
	dim ObjFSO,ObjTS,filetxt	
	dim conn, rs, sql
	dim msgcnt,msg,cnt
	dim param(2)
	dim l_MailFrom
	dim l_MailTo
	dim l_MailSubject
	dim l_MailBody
	dim l_MailServer	
	dim l_LogPath
	dim newfolder
	'----------------------------------------
    ' �c�a�ڑ�
    '----------------------------------------        
    ConnectSvr conn, rs	
	'2009/08/04 M.Marquez Upd-S
	'if v_SendFlag = "0" then
	'	sql = "Select MailAddress from mUsers"		
	'else
	'	sql = "Select MailAddress from mUsers WHERE (MailSend IS NOT NULL AND MailSend <> '')"
	'end if
	if v_SendFlag = "0" then
		sql = "Select MailAddress from mUsers WHERE (MailAddress IS NOT NULL AND MailAddress <> '')"
	else
		sql = "Select MailAddress from mUsers WHERE (MailSend IS NOT NULL AND MailSend <> '') AND (MailAddress IS NOT NULL AND MailAddress <> '')"
	end if
	'2009/08/04 M.Marquez Upd-E
	cnt = 0
	msgcnt = 0
	
	rs.Open sql, conn, 0, 1, 1

	on error resume next
		
	call getUploadIni(param,v_MailFrom)
	l_MailFrom = Replace(gfHTMLEncode(param(0)), chr(10), "") 
	call getUploadIni(param,v_MailSubject)
	l_MailSubject = Replace(gfHTMLEncode(param(0)), chr(10), "") 
	call getUploadIni(param,v_MailServer)
	l_MailServer = Replace(gfHTMLEncode(param(0)), chr(10), "") 
	
'	l_LogPath =  Server.Mappath("./log/maillog.txt")	
'	Set ObjFSO = Server.CreateObject("Scripting.FileSystemObject")		
'	If Not ObjFSO.FolderExists(Server.Mappath("./log/")) Then
'		ObjFSO.CreateFolder(Server.Mappath("./log/"))
'		Set filetxt = ObjFSO.CreateTextFile(Server.Mappath("./log/maillog.txt"), True) 
'	end if

	l_MailBody = ""
	call LfgetMessageTxt()
	
	for i = 1 to UBOUND(Arr_MailSig)
'2009/09/11 Upd-S Fujiyama
'		l_MailBody = l_MailBody & Arr_MailSig(i) & vbNewLine
		l_MailBody = l_MailBody & Arr_MailSig(i)
'2009/09/11 Upd-E Fujiyama
	next									
	
'2009/09/11 Del-S Fujiyama
'	l_MailBody = Replace(l_MailBody, chr(10), "")
'2009/09/11 Del-E Fujiyama

	while not rs.eof		
		if gfTrim(rs("MailAddress")) <> "" then
			cnt = cnt + 1			
			l_MailTo = gfHTMLEncode(gfTrim(rs("MailAddress")))
			msg = gfSendMail(l_MailTo,l_MailFrom,l_MailSubject,l_MailBody,l_MailServer)		
			
			if msg <> "" then
				msgcnt = msgcnt	+ 1						
			end if			
		end if
		rs.movenext
    wend					
	
	rs.close
	conn.Close

	if cnt = 0 then
		v_Msg = "�f�[�^������܂���B"
	elseif msgcnt = 0 then
		v_Msg = "����ɑ��M����܂����B"
	else
		v_Msg = "�y�G���[�z���M�ł��܂���ł����B"
	end if
	
end function
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<TITLE>�g���s�r-���m�点�����e�i���X</TITLE>
<link href="./Common/style.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<style type="text/css">
DIV.listA{
	height:135px;
	width:400px; 
	overflow-y:scroll;
	overflow-x:auto;	
	background-color:#FFFFFF;
	border-top: #666666 2px solid;
	border-left: #666666 2px solid;
	border-bottom: #CCCCCC 2px solid;
	border-right: #CCCCCC 2px solid;
}
.chrReadOnly4
{
    font-family: Tahoma,Arial,"MS Gothic";
	font-size: 13px;
	BACKGROUND-COLOR: #FFFFFF;
    BORDER-BOTTOM: #FFFFFF 0px solid;
    BORDER-LEFT: #FFFFFF 0px solid;
    BORDER-RIGHT: #FFFFFF 0px solid;
    BORDER-TOP: #FFFFFF 0px solid;   
    TEXT-ALIGN: left;
}
</style>
<SCRIPT Language="JavaScript">
//��ʍ��ڂɐݒ�
function finit(){	
	document.frm.Gamen_Mode.value = "<%=v_Mode%>"; 
	document.frm.Data_Cnt.value = "<%=v_Data_Cnt%>";
	document.frm.SendFlag.value = "<%=v_SendFlag%>";  
	
	if (document.frm.Gamen_Mode.value != ""){
        // �G���[���̃��b�Z�[�W�ƃt�H�[�J�X
        if ("<%=v_Msg%>" != ""){
            alert("<%=v_Msg%>");

            //�t�H�[�J�X�ʒu�ݒ�
            for( var i=0; i < document.frm.elements.length; i++ ){
                 if ((document.frm.elements[i].type == "file" || document.frm.elements[i].type == "select") &&
                     document.frm.elements[i].name == "<%=v_FocusItem%>"){
                     document.frm.elements[i].focus();  
                     return false;
                 }    
            }
            return false;
		}
	}else{
		document.frm.txtFileUpload.focus();  	
	}	
}

// ���m�点�t�@�C���A�b�v���[�h�{�^����������
function fUpload(){
	if(gfCHKNull(document.frm.txtFileUpload) == false){
		document.frm.txtFileUpload.focus();
        return false;
    }
	
	document.frm.Gamen_Mode.value = "I";
	document.frm.submit();
}

// �X�V�{�^����������
function fUpd(){
	document.frm.Gamen_Mode.value = "U";
	document.frm.submit();
}

// ���M�{�^����������
function fSetSend(){
	if (gfCHKNull(document.frm.SendFlag) == false){
		document.frm.SendFlag.focus();
        return false;
    }
	 
	document.frm.Gamen_Mode.value = "M";
	document.frm.submit();
}

// �폜�{�^����������
function fDel(){
	var i,cnt;
	var obj;
	
	cnt = 0;
	
	for(i = 1; i <= "<%=v_Data_Cnt%>"; i++){
		obj = eval("document.frm.DelFlag" + i);
        if (obj.value == "1") {  
            cnt++;
        }
    }
	
    if(cnt == 0) {
        window.alert("�t�@�C����I�����Ă��������B");
        return false;
    }
	
	document.frm.Gamen_Mode.value = "D";
	document.frm.submit();	
}

function fHighlight(obj,delflag){
	if(obj.className == "chrReadOnly4"){
		obj.className = "highlight"
		delflag.value = "1"
	}else{
		obj.className = "chrReadOnly4"
		delflag.value = "0"
	}	
}
</script>
</HEAD>
<body class="bckcolor" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  onLoad="finit();">
<form name="frm" action="maintenance.asp" method="post" enctype="multipart/form-data">						  
<!-------------�������烍�O�C�����͉��--------------------------->
<table class="main2" align="center" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
		<%
			DisplayHeader2("���m�点�����e�i���X")
		%>
		<INPUT type="hidden" name="Gamen_Mode" size="9" maxlength="1"  readonly tabindex= -1>
    	<INPUT type="hidden" name="Data_Cnt" size="9" readonly tabindex= -1>
      </table>
      <center>				
        <table class="square" border="0" cellspacing="4" cellpadding="0">
          <tr>
           <td>
		  	<table cellspacing="3" cellpadding="4">
	          <tr>
    	       <td>
				<table width="500" cellspacing="0" cellpadding="5">
				  <tr>
				   <td>
					  <table width="100%">	               
						<tr>						  		 
						  <td align="center">
							<table width="100%">
							<tr>
							  <td align="left"> 
								  <table width="100%" border="0" cellspacing="1" cellpadding="3">
									<tr> 
									  <td nowrap align=left valign=middle colspan="2">�P�D���m�点�t�@�C���֌W</td>				  
									</tr>
									<tr>
										<td colspan="2">&nbsp;&nbsp;&nbsp;&nbsp;�i�P�j���݂̂��m�点�t�@�C���ꗗ</td>
									</tr>
									<tr>
									  <td></td>							  
									  <td width="600">							  
										<div class="listA">
										<TABLE align="left" cellspacing=1 cellPadding=1 width="100%" class="chrReadOnly4">								
										<% for i = 1 to UBOUND(Arr_Files) %>						
										  <tr>					    																
											<td nowrap class="chrReadOnly4" onClick="fHighlight(this,document.frm.DelFlag<%=cstr(i)%>);">
											<% v_ItemName = "DelFlag" & cstr(i) %>
											<input type="hidden" name="<%= v_ItemName %>" size="2">	
											<% 
												v_ItemName = "FileNames" & cstr(i) 
												Arr_Name = split(Arr_Files(i),"|")												
											%>																											
											<input type="hidden" name="<%= v_ItemName %>" value="<%=Arr_Name(2)%>">
											<%=Arr_Name(2)%>
											</td>		    										 
										  </tr>	
										<% next %>
										</TABLE>
										</div> 
									  </td>
									  <td colspan="3" valign="bottom" align="center">
										&nbsp;<input type="button" value="   �폜  " onClick="fDel();">
									  </td>
									</tr>
									<tr>
									  <td></td>	
									  <td align="left" colspan="2"><input type="file" name="txtFileUpload" size="60"></td>
									</tr>
									<tr> 
									  <td align="center" colspan="3">
										<input type="button" value="���m�点�t�@�C���A�b�v���[�h" onClick="fUpload();">
									  </td>				  
									</tr>
									<tr><td>&nbsp;</td></tr>
									<tr> 
									  <td nowrap align=left valign=middle colspan="2">�Q�D���m�点mail���M</td>				  
									</tr>								
									<tr>
										<td colspan="2">&nbsp;&nbsp;&nbsp;&nbsp;�i�P�j�{���ύX</td>
									</tr>
									<tr>
										<td width="50"></td>
										<td>
									<!--<input type="text" name="txtMailMsg1">	-->
									<textarea name="txtMailMsg" cols=52 rows=7><% for i = 1 to UBOUND(Arr_MailSig) %><%=Arr_MailSig(i)%><% next%></textarea>
										</td>
										<td valign="bottom">							  	
										<input type="button" value="   �X�V  " onClick="fUpd();">
									  </td>
									</tr>
									<tr>
										<td colspan="2">&nbsp;&nbsp;&nbsp;&nbsp;�i�Q�jmail���M</td>								
									</tr>											
									<tr>													
										<td colspan="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;���M��w��</td>								
									</tr>
									<tr>
										<td colspan="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<select name ="SendFlag">
										<option value=""></option>
										<option value="0">���ׂĂɑ��M</option>
										<option value="1">��]�҂ɂ̂ݑ��M</option>										
										</select>
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<input type="button" value=" ���M " onClick="fSetSend();">
										</td>	
									</tr>
								  </table>
								  <br>
								  <center>						  
								  <br>
								  <a href="menu.asp">����</a>			
								  </center>  					  	
							  </td>
							</tr>
							</table>
						</td>						
					   </tr>	        	 	
					 </table>				 
				  </td>
				</tr>
			  </table>
	  	    </td>
    	   </tr>
      	  </table>
	  	 </td>
        </tr>
      </table>
	  </center>	  
	  <table border=0><tr><td height=20></td></tr></table>
    </td>	
 </tr> 
 	<%
		DisplayFooter
	%> 
</table>
</form>
</body>
</HTML>
