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
	'��ʍ��ڕϐ�
	dim v_Mode
	dim v_Data_Cnt
	dim v_Message
	dim i
	
	v_Mode = request.form("Gamen_Mode")
	v_Data_Cnt = request.form("Data_Cnt")
	v_Message = request.form("txtMessage")	
	
	If v_Mode = "U" then
		call LfsetMessage()
	End If
	
	call LfgetMessage()	
	
function LfsetMessage()
	dim ObjFSO,ObjTS
	dim Arr_Temp
	dim cnt
	dim strTemp
	
	cnt = 0
	redim Arr_Temp(0)	 
	
	v_Message = Replace(v_Message, chr(10), " ")	
	v_Message = Replace(v_Message, chr(13), " ")
    
	'--- �t�@�C�����J���i�ǂݎ���p�j ---
	Set ObjFSO = Server.CreateObject("Scripting.FileSystemObject")	
	Set ObjTS = ObjFSO.OpenTextFile(Server.Mappath("../denbun.html"),1,false)
	
	'--- �t�@�C���f�[�^�̓Ǎ��� ---
	Do Until ObjTS.AtEndofStream		
		cnt = cnt + 1
		strTemp = ObjTS.ReadLine
		redim preserve Arr_Temp(cnt)		
		if Mid(Trim(strTemp),1,11) = "var Denbun=" then		
			strTemp = "    var Denbun=""" & gfHTMLEncode(Trim(v_Message)) & "�@�@�@�@�@�@�@�@     �@ �@        """
		end if				
		Arr_Temp(cnt) = strTemp
	Loop
	
	ObjTS.Close	
	
	'--- �t�@�C�����J���i�ǂݎ���p�j ---
	Set ObjTS = ObjFSO.OpenTextFile(Server.Mappath("../denbun.html"),2,false)
	
	for i = 1 to UBOUND(Arr_Temp)	
		ObjTS.WriteLine(Arr_Temp(i))
	next
	
	ObjTS.Close
	
	Set ObjTS = Nothing
	Set ObjFSO = Nothing
end function

function LfgetMessage()
	dim ObjFSO,ObjTS	
	dim cnt
	dim strTemp
	
	cnt = 0
	    
	'--- �t�@�C�����J���i�ǂݎ���p�j ---
	Set ObjFSO = Server.CreateObject("Scripting.FileSystemObject")	
	Set ObjTS = ObjFSO.OpenTextFile(Server.Mappath("../denbun.html"),1,false)
	
	'--- �t�@�C���f�[�^�̓Ǎ��� ---
	Do Until ObjTS.AtEndofStream		
		cnt = cnt + 1
		strTemp = ObjTS.ReadLine		

		if Mid(Trim(strTemp),1,11) = "var Denbun=" then				
			v_Message = Mid(Trim(strTemp),13,Len(Trim(strTemp))-13)			
		end if	
	Loop
	
	ObjTS.Close	
	Set ObjTS = Nothing
	Set ObjFSO = Nothing
end function


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<TITLE>�g���s�r-�e���b�v�X�V</TITLE>
<link href="./Common/style.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">

function finit(){	
	document.frm.Gamen_Mode.value = "<%=v_Mode%>";
	document.frm.Data_Cnt.value = "<%=v_Data_Cnt%>";
	document.frm.txtMessage.value = "<%=v_Message%>";
	document.frm.txtMessage.focus();	
}

function fUpd(){
	document.frm.Gamen_Mode.value = "U";
	document.frm.submit();
}

</script>
</HEAD>
<body class="bckcolor" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad=finit();>
<form name="frm" action="update.asp" method="post">
<!-------------�������烍�O�C�����͉��--------------------------->
<table class="main2" align="center" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <%
			DisplayHeader2("�e���b�v�X�V")
		%>
		  <INPUT type="hidden" name="Gamen_Mode" size="9" maxlength="1"  readonly tabindex= -1>
		  <INPUT type="hidden" name="Data_Cnt" size="9" readonly tabindex= -1>           
      </table>
      <center>
		<table border=0><tr><td height=65></td></tr></table>
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
								  <table width="100%" border="0" cellspacing="2" cellpadding="3">
									<tr> 
									  <td nowrap align=left valign=middle colspan="2">�P�D�e���b�v�X�V</td>				  
									</tr>
									
									<tr> 
									  <td width="100%" border="0" colspan="2"><textarea name="txtMessage" cols=70 rows=5></textarea></td>
									  <td valign="bottom">							  	
										<input type="button" value="   �X�V  " onClick="fUpd();">							  </td>
									</tr>
									<tr> 									  
									  <td nowrap align=left valign=middle>�i�T���v���j</td>								  			  
									</tr>									
									<tr>
									  <td nowrap align=left valign=middle>�g���s�r���Q�̂����p���肪�Ƃ��������܂��B</td>
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
    </td>
 </tr>
	<%
		DisplayFooter
	%>
</table>
</form>
</body>
</HTML>
