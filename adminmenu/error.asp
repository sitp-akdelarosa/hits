<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:error.asp				_/
'_/	Function	:�G���[���				_/
'_/	Date		:2003/06/18				_/
'_/	Code By		:SEIKO Electric.Co ��d			_/
'_/	Modify		:					_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTP�R���e���c�^�C�v�ݒ�
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<%
'�G���[���擾
  dim ObjConn, ObjRS,WinFlag,dispId,wkID,wkName,errerCd,etc
  WinFlag= Session.Contents("WinFlag")
  dispId = Session.Contents("dispId")
  wkID   =  Session.Contents("wkID")
  wkName =  Session.Contents("wkName")
  errerCd=  Session.Contents("errerCd")
  etc    =  Session.Contents("etc")
'�Z�b�V�����N���A
  Session.Contents.Remove("WinFlag")
  Session.Contents.Remove("dispId")
  Session.Contents.Remove("wkID")
  Session.Contents.Remove("wkName")
  Session.Contents.Remove("errerCd")
  Session.Contents.Remove("etc")

'�G���[���b�Z�[�W�擾
  dim ErrerM1,ErrerM2
  dim ObjFSO,ObjTS,tmpStr,tmp
  ObjFSO = Server.CreateObject("Scripting.FileSystemObject")
  ObjTS = ObjFSO.OpenTextFile(Server.Mappath("./ini/ADMINERROR.ini"),1,false)
  '--- �t�@�C���f�[�^�̓Ǎ��� ---
  Do Until ObjTS.AtEndofStream
    tmpStr = ObjTS.ReadLine
    If Left(tmpStr,3) = errerCd Then
      tmp=Split(tmpStr,":",3,1)
      ErrerM1 = tmp(1)
      ErrerM2 = tmp(2)
      Exit Do
    End If
  Loop
  ObjTS.Close
  ObjTS = Nothing
  ObjFSO = Nothing

'�{�^���\������
  dim Button
  If WinFlag = 0 Then
    Button="'���O�C����ʂɖ߂�' onClick='submit()'"
  ElseIf WinFlag = 1 Then
    Button="'����' onClick='window.close()'"
  Else
    Button="'�߂�' onClick='window.history.back()'"
  End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<TITLE>�G���[</TITLE>
<link href="./Common/style.css" rel="stylesheet" type="text/css">
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT language=JavaScript>
<!--
// -->
</SCRIPT>
<!--#include File="./Common/common.inc"-->
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY class="bckcolor" leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!-------------�G���[���--------------------------->
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
		  	<table border="0" cellspacing="3" cellpadding="4">
	          <tr>
    	       <td>
				<table width="500" border="0" cellspacing="0" cellpadding="5">
				  <tr>
				   <td>
					  <table width=100%>	               
						<tr>						  		 
						  <td align="center">
							<table width="100%" border=0>
							<tr>
							  <td align="left" width="100%"> 
								  <table width="100%" border="0" cellspacing="2" cellpadding="3">									
									<tr><td colspan=2 align="center" class="menu">�G���[</td></tr>
									<tr><td>�G���[���ID�F���ID</td><td>�F<%=dispId%>�F<%=wkId%></td></tr>
									<tr><td>��Ɩ�</TD><TD>�F<%=wkName%></td></tr>
									<tr><td>�G���[�R�[�h</TD><TD>�F<%=errerCd%></td></tr>
									<tr><td>���b�Z�[�W</TD><TD>�F<%=ErrerM1%><BR></td></tr>
									<tr><td>�Ώ�</td><td>�F<%=ErrerM2%><BR></td></tr>
									<tr><td colspan=2><%=etc%></td></tr>
									<tr><td colspan=2 height="20"></td></tr>
									<tr><td colspan=2 align=center>
									<form action="./login.asp" target="_top">										
										<input type=button value=<%=Button%>>
									</form>
									</td></tr>
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
<!-------------��ʏI���--------------------------->
</BODY></HTML>
