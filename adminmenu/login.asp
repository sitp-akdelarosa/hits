<%
'**********************************************
'  �y�v���O�����h�c�z�@: login.asp
'  �y�v���O�������́z�@: ���O�C��
'
'  �i�ύX�����j
'
'**********************************************
Option Explicit
Response.Expires = 0
%>

<!--#include File="./Common/common.inc"-->
<SCRIPT src="./Common/function.js" type=text/javascript></SCRIPT>
<%
    '--- �ϐ��錾 ---
    dim wArgc                   ' �p�����[�^�ő吔
    dim wRtnB                   ' �֐��߂�l����p
    dim msg
	dim conn, rs
	dim sql
	
	msg = ""
    
	session("username") = empty
	
	' �w������̎擾(���[�U�[�h�c)
    Dim strInputUserID, strInputPassWord
    strInputUserID = UCase(Trim(Request.Form("txtUserID")))
    strInputPassWord = UCase(Trim(Request.Form("txtPass")))
	
	If strInputUserID <> "" and strInputPassWord <> "" then
       'session("Loginid") = strInputUserID
        '----------------------------------------
        ' �c�a�ڑ�
        '----------------------------------------        
        ConnectSvr conn, rs

        '----------------------------------------
        ' ���[�U���擾
        '----------------------------------------
        session("user_id")   = empty
        session("username") = empty

        sql="SELECT FullName,UserType FROM mUsers WHERE UserCode = '" & gfSQLEncode(strInputUserID) & "' And Password = '" & gfSQLEncode(strInputPassWord) & "' AND UserType = '0'"
        rs.Open sql, conn, 0, 1, 1
		
		on error resume next
		
        If rs.eof or err.number<>0 then
            msg="���͂��ꂽ���e�ɊԈႢ������܂��B"
        Else
			' ���O�C�������Z�b�V�����ϐ��ɐݒ�
			session("username") = Trim(rs("FullName"))            
			session("user_id") = strInputUserID            '2016/07/28 H.Yoshikawa Add
        End If
		
        rs.close
		conn.Close
    End If
	
    If session("username") <> "" then
        response.redirect "menu.asp"
    End If    
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<TITLE>�g���s�r-�Ǘ��җp���</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<link href="./Common/style.css" rel="stylesheet" type="text/css">
<SCRIPT Language="JavaScript">

function finit(){
    document.frm.txtUserID.focus();
}


function Check(){
	var obj = document.frm;  

	ret = CheckEisuji(obj.txtUserID.value);
  
	if(ret == false){
    	alert("�Ǘ���ID�͔��p�p�����œ��͂��Ă��������B");
		obj.txtUserID.focus();
	    return false;
	}
	
	if(obj.txtUserID.value == ""){
    	alert("�K�{���͍��ڂł��B");
		obj.txtUserID.focus();
	    return false;	
	}
	
	if(obj.txtPass.value == ""){
    	alert("�K�{���͍��ڂł��B");
		obj.txtPass.focus();
	    return false;	
	}
	
    return true;
}
</SCRIPT>
</HEAD>

<body class="bckcolor" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="finit();">
<form name="frm" action="login.asp" method="post">
<SCRIPT src="./Common/KeyDown.js" type=text/javascript></SCRIPT>
<!-------------�������烍�O�C�����͉��--------------------------->
<table class="main" align="center" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
		<%
			DisplayHeader
		%>
      </table>
      <center>
	  	<BR><BR>
		<table border=0><tr><td height=50></td></tr></table>
        <table class="square" cellspacing="4" cellpadding="0">
          <tr>
           <td>
		  	<table border="0" cellspacing="3" cellpadding="4">
	          <tr>
    	       <td>
	        <table width="500" border="0" cellspacing="0" cellpadding="5">
	          <tr>
	           <td>
	              <table width="100%">	                
					<tr>
			   		<td></td>		 
	                <td align="center">
					<table width="100%">
						<tr>
							<td align="center"><B>�Ǘ��җp��ʂւ̃��O�C���B</B></td>
						</tr>
						<tr>
						  <td nowrap align="center"> 
						  	  <BR>															  
							  <table border="0" cellspacing="2" cellpadding="3">
								<tr> 
								  <td nowrap align=left valign=middle><B>�Ǘ���ID</B></td>
								  <td nowrap>
									<table border=0 cellpadding=0 cellspacing=0>
									  <tr>
										<td width=100>
											<input type=text name="txtUserID" value="" size=10 maxlength=5>
										</td>										
									  </tr>
									</table>
								  </td>
								</tr>
								<tr> 
								  <td nowrap align=left valign=middle><B>�p�X���[�h</B></td>
								  <td nowrap>
									<table border=0 cellpadding=0 cellspacing=0>
									  <tr>
										<td width=100>
											<input type=password name="txtPass" size=10 maxlength=8>
										</td>									
									  </tr>
									</table>
								  </td>
								</tr>
								<tr>
									<td colspan="2">
									<%  if msg<>"" then%>
			   						  <font color="red"><%=msg%></font>
		    					    <%  end if%>
	  							    </td>
								</tr>
							  </table>
						  	  <br>
						  	  <input type="submit" value=" ���O�C�� " onClick="return Check();">
						 </td>
						</tr>
					</table>
		 			</td>
		  			<td></td>
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