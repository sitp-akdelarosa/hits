<%@LANGUAGE="VBSCRIPT" CODEPAGE="932"%>
<%
'**********************************************
'  �y�v���O�����h�c�z�@: warehouse.asp
'  �y�v���O�������́z�@: ���u�ꏊ�R�[�h�����e�i���X
'
'  �i�ύX�����j
'
'**********************************************
	
	Option Explicit
	Response.Expires = 0

	call CheckLoginH()
%>
<!--#include File="./Common/common.inc"-->
<SCRIPT src="./Common/function.js" type=text/javascript></SCRIPT>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<TITLE>���u�ꏊ�R�[�h�����e�i���X</TITLE>
<link href="./Common/style.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">

function finit(){
	document.frm.txtOpeCode.focus();
}

function fDisplay(){	
	var obj = document.frm;  
    var w=900;
    var h=550;
    var l=0;
    var t=0;
	
	if (gfCHKNull(obj.txtOpeCode) == false){
		obj.txtOpeCode.focus();
        return false;
    }
	
	ret = CheckEisuji(obj.txtOpeCode.value);  
	if(ret == false){
    	alert("�I�y���[�^�R�[�h�͔��p�p�����œ��͂��Ă��������B");
		obj.txtOpeCode.focus();
	    return false;
	}	
	
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
	
    var win=window.open("./warehouse2.asp?code=" + obj.txtOpeCode.value,"","status=no,width="+w+",height="+h+",top="+t+",left="+l);
}
</script>
</HEAD>
<body class="bckcolor" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="finit();">
<form name="frm" action="warehouse.asp" method="post">						  
<!-------------�������烍�O�C�����͉��--------------------------->
<table class="main2" align="center" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <%
			DisplayHeader2("���u�ꏊ�R�[�h�����e�i���X")
    	  %>
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
								  <table width="100%" border="0">
									<tr> 
									  <td colspan="2" align=left valign=middle nowrap>�I�y�[���[�^�w��</td>
									</tr>
									<tr>
										<td height="5"></td>
									</tr>
									<tr> 									  
									  <td nowrap width="140" align="center">�I�y���[�^�R�[�h</td>
									  <td>
									  	<input type="text" name="txtOpeCode" size="5" maxlength="3">
									  </td>
									  <td rowspan="2" valign="top">
									  	<input type="button" value="�}�X�^�\��" onclick="fDisplay();">
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
    </td>
 </tr>
	<%
		DisplayFooter
	%>
</table>
</form>
</body>
</HTML>
