<%
'***********************************************************
'  �y�v���O�����h�c�z�@: 
'  �y�v���O�������́z�@: 
'
'  �i�ύX�����j
'2017/01/19 T.Okui ���j���[(10)�u���F�h���C�o�ꗗ�E�폜�v�ǉ�
'***********************************************************
Option Explicit
Response.Expires = 0

call CheckLoginH()

%>
<!--#include File="./Common/Common.inc"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<TITLE>�g���s�r-�Ǘ��җp���j���[ </TITLE>
<link href="./Common/style.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
</HEAD>
<body class="bckcolor" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="menu" action="menu.asp" method="post">
<!-------------�������烍�O�C�����͉��--------------------------->
<table class="main2" align="center" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <%
			DisplayHeader
		%>
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
					  <td></td>		 
					  <td align="center">
					  <table>
					<tr>
					  <td nowrap align="center" class="menu">
					  <dl>
					  <B>�Ǘ��җp���j���[</B>
					  </dl>
					  <center>
					  <table border="0" cellspacing="2" cellpadding="3">
						<tr> 
						  <td nowrap align=left valign=middle><a href="upload.asp">�i�P�j�l���A�b�v���[�h</a></td>				  
						</tr>
						<tr> 
						  <td nowrap align=left valign=middle><a href="maintenance.asp">�i�Q�j���m�点�����e�i���X</a></td>
						</tr>
						<tr> 
						  <td nowrap align=left valign=middle><a href="update.asp">�i�R�j�e���b�v�X�V</a></td>				  
						</tr>
						<tr> 
						  <td nowrap align=left valign=middle><a href="agreement_update.asp">�i�S�j���p�K��̍X�V</a></td>				  
						</tr>
						<tr> 
						  <td nowrap align=left valign=middle><a href="accesstotal.asp">�i�T�j���p�����\��</a></td>				  
						</tr>
						<tr> 
						  <td nowrap align=left valign=middle><a href="place.asp">�i�U�j���u�ꏊ�R�[�h�����e�i���X</a></td>				  
						</tr>
						<tr> 
						  <td nowrap align=left valign=middle><a href="lockonservice.asp">�i�V�j���b�N�I���T�[�r�X����</a></td>
						</tr>
						<tr> 
						  <td nowrap align=left valign=middle><a href="settings.asp">�i�W�j�e��p�����[�^�ݒ�</a></td>
						</tr>
						<!-- 2016/07/27 H.Yoshikawa add start -->
						<tr> 
						  <td nowrap align=left valign=middle><a href="driver.asp">�i�X�j�h���C�o���F</a></td>
						</tr>
						<!-- 2016/07/27 H.Yoshikawa add end   -->
						<!-- 2017/01/20 T.Okui add start -->
						<tr> 
						  <td nowrap align=left valign=middle><a href="driverlist.asp">�i10�j���F�h���C�o�ꗗ�E�폜</a></td>
						</tr>
						<!-- 2017/01/20 T.Okui add end -->
					  </table>	
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

