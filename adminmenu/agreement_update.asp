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
<!--#include File="./Common/upload.inc"-->
<SCRIPT src="./Common/function.js" type=text/javascript></SCRIPT>
<%
	'��ʍ��ڕϐ�
	dim v_Mode
	dim obj
	dim buf
	dim totalByte
	dim v_UploadFile
	dim v_Msg
	dim v_FocusItem		
	
	Set obj=server.createobject("basp21")
	on error resume next
	totalByte = Request.TotalBytes
	buf	= Request.BinaryRead(totalByte)
	
	v_Mode = obj.Form(buf,"Gamen_Mode")
	v_UploadFile = obj.FormFileName(buf,"txtFileUpload")
	
	if v_Mode = "U" and v_UploadFile <> "" then	
		'2009/07/24 C.Pestano Add-S	
		if Mid(v_UploadFile, InstrRev(v_UploadFile, "\")+1) <> "info.html" then
			v_Msg = "�t�@�C����������������܂���B"
			v_FocusItem = "txtFileUpload"
		end if
		'2009/07/24 C.Pestano Add-E
		
		if not gfUploadTempFile(v_UploadFile,"txtFileUpload") and v_Msg = "" then
			v_Msg = "�A�b�v���[�h�͎��s���܂����B"
			v_FocusItem = "txtFileUpload"
		else
			call LfCheckFile()		
		end if				
	end if
	
function LfCheckFile()
	dim ObjFSO,ObjTS	
	dim cnt
	dim strTemp
	dim wkfilename	

	cnt = 0
	
	wkfilename = Server.MapPath("../") & "\adminmenu\temp\" & Mid(v_UploadFile, InstrRev(v_UploadFile, "\")+1)
	
	'--- �t�@�C�����J���i�ǂݎ���p�j ---
	Set ObjFSO = Server.CreateObject("Scripting.FileSystemObject")	 
	Set ObjTS = ObjFSO.OpenTextFile(Server.Mappath("./temp/info.html"),1,false)
	
	'--- �t�@�C���f�[�^�̓Ǎ��� ---
	Do Until ObjTS.AtEndofStream		
		strTemp = ObjTS.ReadLine
						
		if InStr(1,strTemp,"<HTML",1) <> 0 then							
			cnt = cnt + 1	
		end if
		if InStr(1,strTemp,"</HTML",1) <> 0 then
			cnt = cnt + 1	
		end if
		
		if InStr(1,strTemp,"<BODY",1) <> 0 then
			cnt = cnt + 1	
		end if
		
		if InStr(1,strTemp,"</BODY",1) <> 0 then
			cnt = cnt + 1			
		end if
		
		if InStr(1,strTemp,"</TITLE",1) <> 0 then
			cnt = cnt + 1
		end if
		
		if InStr(1,strTemp,"<TITLE",1) <> 0 then
			cnt = cnt + 1			
		end if	
	Loop	
	
	ObjTS.Close
		
	if cnt = 6 then		
		ObjFSO.CopyFile Server.Mappath("./temp/info.html"), Server.Mappath("../") & "\",True
		ObjFSO.DeleteFile (Server.Mappath("./temp/info.html"))
		'2009/07/24 C.Pestano Add-S
		v_Msg = "�A�b�v���[�h���܂���"
		v_FocusItem = "txtFileUpload"	
		'2009/07/24 C.Pestano Add-E
		'ObjFSO.MoveFile "C:\HITS\adminmenu\temp\info.html", "C:\HITS\" 
	else
		v_Msg = "���͒l������������܂���B"
		v_FocusItem = "txtFileUpload"
		ObjFSO.DeleteFile (Server.Mappath("./temp/info.html"))
	end if
	
	
	Set ObjTS = Nothing
	Set ObjFSO = Nothing
end function

function gfUploadTempFile(fname1, upload)
	dim ret
	dim wkfilename 		
	dim serverpath
	
	gfUploadTempFile = True
	
	serverpath = Server.MapPath("../")
	
	wkfilename	= serverpath & "\adminmenu\temp\" & Mid(fname1, InstrRev(fname1, "\")+1)
	
	ret	= obj.FormSaveAs(buf,upload,wkfilename)
	
	if ret > 0 then
		gfUploadTempFile = true
	else		
		gfUploadTempFile = false
	end if
end function
	
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<TITLE>�g���s�r-���p�K��̍X�V</TITLE>
<link href="./Common/style.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">

function finit(){	
	document.frm.Gamen_Mode.value = "<%=v_Mode%>";
	
	if (document.frm.Gamen_Mode.value == "U"){
        // �G���[���̃��b�Z�[�W�ƃt�H�[�J�X
        if ("<%=v_Msg%>" != ""){
            alert("<%=v_Msg%>");

            //�t�H�[�J�X�ʒu�ݒ�
            for( var i=0; i < document.frm.elements.length; i++ ){
                 if ((document.frm.elements[i].type == "file") &&
                     document.frm.elements[i].name == "<%=v_FocusItem%>"){
                     document.frm.elements[i].focus();  
                     return;
                 }    
            }
            return;
		}
	}
	
	document.frm.txtFileUpload.focus();
}

function fUpload(){
	var obj;
	
	if(gfCHKNull(document.frm.txtFileUpload) == false){
		document.frm.txtFileUpload.focus();
        return false;
    }
	
	obj = eval("document.frm.txtFileUpload");
	
	if (obj.value != ""){
            var ext = obj.value;
            ext = ext.substring(ext.length-4,ext.length);
            ext = ext.toLowerCase();
            if (ext != 'html'){
                window.alert("���͒l������������܂���B");
                obj.focus();
                return false;
            }
	}			
	
	document.frm.Gamen_Mode.value="U";
	document.frm.submit();
}

</script>

</HEAD>
<body class="bckcolor" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="finit();">
<form name="frm" action="agreement_update.asp" method="post" enctype="multipart/form-data">
<!-------------�������烍�O�C�����͉��--------------------------->
<table class="main2" align="center" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
         <%
			DisplayHeader2("���p�K��̍X�V")
    	  %>
		  <INPUT type="hidden" name="Gamen_Mode" size="9" maxlength="1"  readonly tabindex= -1>
      </table>
      <center>
		<table border=0><tr><td height=50></td></tr></table>
        <table class="square" border="0" cellspacing="4" cellpadding="0">
          <tr>
           <td>
		  	<table border="0" cellspacing="3" cellpadding="4">
	          <tr>
    	       <td>
				<table width="500" border="0" cellspacing="0" cellpadding="5">
				  <tr>
				   <td>
					  <table width=100%>
						<tr>						  	 
						  <td align="center">
							<table width="100%">
							<tr>
							  <td align="left" width="100%"> 						
								  <table width="100%" cellspacing="2" cellpadding="3">
									<tr> 
									  <td align=left valign=middle nowrap colspan=3>�����p�K��t�@�C���̃A�b�v���[�h</td>				  
									</tr>
									<tr>
									 <td valign="bottom" height="10"></td>
									</tr>
									<tr> 									  									  
									  <td align="center" colspan=3>
										<table width="89%" class="box">
											<tr>
												<td>
												�ȉ��̏����ōX�V�t�@�C����p�ӂ����̃{�^���ŃA�b�v���[�h<BR />
												���s���Ă��������B<BR /><BR />
												
												&nbsp;&nbsp;&nbsp;&nbsp;�`���E�E�EHTML<BR />
												&nbsp;&nbsp;&nbsp;&nbsp;�t�@�C�����E�E�Einfo.html<BR />
												<BR />
												</td>
											</tr>									
										</table>
									  </td>									  
									</tr>
									<tr>
									 <td valign="bottom" height="5"><BR></td>
									</tr>	
									<tr>
									  <td></td>	
									  <td align="center" colspan="2">
									  <input type="file" name="txtFileUpload" size="48" onFocus="document.frm.txtFileUpload.select();"></td>
									</tr>
								  </table>
								  <br>								  
								  <center>
								  <input type="button" value=" �A�b�v���[�h " onClick="fUpload();">
								  <br>						  
								  <br>
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
