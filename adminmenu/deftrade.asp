<%@LANGUAGE="VBSCRIPT" CODEPAGE="932"%>
<%
'**********************************************
'  �y�v���O�����h�c�z�@: deftrade.asp
'  �y�v���O�������́z�@: �m�莖�Ǝ҃}�X�^�����e�i���X
'
'  �i�ύX�����j
'  2017/04/05 H.Yoshikawa �V�K�쐬
'**********************************************
	
	Option Explicit
	Response.Expires = 0

	call CheckLoginH()	
%>
<!--#include File="./Common/common.inc"-->
<%
	dim conn, rs, sql
	dim v_Mode
	dim v_Data_Cnt
	dim v_SDefCode
	dim Arr_DefCode		'�m�莖�Ǝ҃R�[�h
	dim Arr_DefName		'�m�莖�ƎҖ�
	dim cnt,i
	dim v_Msg
	dim v_FocusItem
	dim v_ItemName
	
	redim Arr_DefCode(0)
	redim Arr_DefName(0)
	
	const l_ProgramID = "DefTrade"
	
	cnt = 0	
		
	'----------------------------------------
    ' �ĕ`��O�̍��ڎ擾
    '----------------------------------------	
	call LfRequestItem() 
	
	if v_Mode = "U" then
		call LfUpdData()
	end if 
	
	call LfSearchData() 
	
	
'-----------------------------
'   �`��O�̉�ʍ��ڂ��擾
'-----------------------------
function LfRequestItem()	
	v_Mode = gfTrim(request.form("Gamen_Mode"))
	v_Data_Cnt = gfTrim(request.form("Data_Cnt"))
	v_SDefCode = ucase(gfTrim(request.form("SDefCode")))
	if v_Data_Cnt = "" then
		v_Data_Cnt = 0
	end if
	
	for i = 1 to CInt(v_Data_Cnt)
		redim preserve Arr_DefCode(v_Data_Cnt)
		redim preserve Arr_DefName(v_Data_Cnt)
		Arr_DefCode(i) = ucase(gfTrim(request.form("DefCode" & i)))
		Arr_DefName(i) = ucase(gfTrim(request.form("DefName" & i)))
	next
	
end function

function LfSearchData()
	Dim emptyNum
	
	'----------------------------------------
	' �c�a�ڑ�
	'----------------------------------------        
	ConnectSvr conn, rs
	
	cnt = 0
	
	'������������̏ꍇ�A�ŏ��ɕ\��
	if v_SDefCode <> "" then
		sql = "SELECT * FROM mDefTrade"
		sql = sql & " WHERE DefCode like '%" & gfSQLEncode(v_SDefCode) & "%'"
		sql = sql & " ORDER BY DefCode "		
		rs.Open sql, conn, 0, 1, 1

		on error resume next
		while not rs.eof
			cnt = cnt + 1			
			redim preserve Arr_DefCode(cnt)
			redim preserve Arr_DefName(cnt)
			Arr_DefCode(cnt) = gfTrim(rs("DefCode"))	'�m�莖�Ǝ҃R�[�h
			Arr_DefName(cnt) = gfTrim(rs("DefName"))	'�m�莖�ƎҖ�
			rs.movenext
		wend
		rs.close
	end if
	
	'�S�����i������������̏ꍇ�́A�w��ԍ��ȊO�̃f�[�^�j
	sql = "SELECT * FROM mDefTrade"
	if v_SDefCode <> "" then
		sql = sql & " WHERE DefCode not like '%" & gfSQLEncode(v_SDefCode) & "%'"
	end if
	sql = sql & " ORDER BY DefCode "

	rs.Open sql, conn, 0, 1, 1

	on error resume next
	while not rs.eof
		cnt = cnt + 1			
		redim preserve Arr_DefCode(cnt)
		redim preserve Arr_DefName(cnt)		
		Arr_DefCode(cnt) = gfTrim(rs("DefCode"))	'�m�莖�Ǝ҃R�[�h
		Arr_DefName(cnt) = gfTrim(rs("DefName"))	'�m�莖�ƎҖ�
		rs.movenext
	wend
	rs.close
	
	'�V�K�p��t�B�[���h�ǉ�
	emptyNum = 10
	redim preserve Arr_DefCode(cnt+emptyNum)
	redim preserve Arr_DefName(cnt+emptyNum)
	for i = 1 to emptyNum
		Arr_DefCode(cnt+i) = ""
		Arr_DefName(cnt+i) = ""
	next
	v_Data_Cnt = cnt +emptyNum
	
	conn.Close
end function

function LfUpdData()
	'----------------------------------------
	' �c�a�ڑ�
	'----------------------------------------        
	ConnectSvr conn, rs
	conn.begintrans

	sql = "DELETE FROM mDefTrade "
	
	conn.execute sql
	if err.number<>0 then				'--- �G���[
		conn.rollbacktrans
		v_Msg = "�}�X�^�̍폜�Ɏ��s���܂����B"
		return false
	end if
	
	for i = 1 to CInt(v_Data_Cnt) 
		if gfTrim(Arr_DefCode(i)) <> "" and gfTrim(Arr_DefName(i)) <> "" then
			sql = "INSERT INTO mDefTrade(DefCode,UpdtTime,UpdtPgCd,UpdtTmnl,DefName)"
			sql = sql & " VALUES("
			sql = sql & "'" & gfSQLEncode(Arr_DefCode(i)) & "',"		
			sql = sql & "current_timestamp,"
			sql = sql & "'" & gfSQLEncode(l_ProgramID) & "',"		
			sql = sql & "'" & gfSQLEncode(ucase(Request.ServerVariables("SERVER_NAME"))) & "',"		
			sql = sql & "'" & gfSQLEncode(Arr_DefName(i)) & "')"
					
			conn.execute sql
		
			if err.number<>0 then				'--- �G���[
				conn.rollbacktrans
				v_Msg = "�}�X�^�̒ǉ��Ɏ��s���܂����B"
				v_FocusItem = "DefCode" & i
				return false
			end if
		end if
	next
	conn.committrans
	
	v_Msg = "�X�V���܂����B"

	conn.Close
end function
%>


<SCRIPT src="./Common/function.js" type=text/javascript></SCRIPT>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<TITLE>�g���s�r-�m�莖�Ǝ҃}�X�^�����e�i���X</TITLE>
<link href="./Common/style.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">

function finit(){
	var i;

    // �G���[���̃��b�Z�[�W�ƃt�H�[�J�X
    if ("<%=v_Msg%>" != ""){
        alert("<%=v_Msg%>");

        //�t�H�[�J�X�ʒu�ݒ�
        for(i=0; i < document.frm.elements.length; i++ ){
             if ((document.frm.elements[i].type == "text") &&
                 document.frm.elements[i].name == "<%=v_FocusItem%>"){
                 document.frm.elements[i].focus();  
                 return false;
             }    
        }
        return false;
	}else{
		document.frm.SDefCode.focus();
	}
}

function fSearch(){
	
	//2017/05/08 H.Yoshikawa Upd Start
	//if(document.frm.SDefCode.value.length == 0){
    //	alert("��������o�^�ԍ�����͂��Ă��������B");
	//	document.frm.SDefCode.focus();
    //    return false;
    //}
    //2017/05/08 H.Yoshikawa Upd End

	document.frm.Gamen_Mode.value = "S";
	document.frm.submit();
}

function fUpd(){
	var obj;
	var obj2;
	var obj3;
	var i,x;
	var ret;
	var datacnt;
	
	datacnt = document.frm.Data_Cnt.value;
	for (i = 1; i <= datacnt; i++) {
		obj = eval("document.frm.DefCode" + i);	
		obj2 = eval("document.frm.DefName" + i);	
		
		//�����ꂩ����͂̏ꍇ�A����������K�{
		if(obj.value.length != 0 && obj2.value.length == 0){
    		alert("�m�莖�ƎҖ�����͂��Ă��������B");
			obj2.focus();
		    return false;
		}
		if(obj.value.length == 0 && obj2.value.length != 0){
    		alert("�o�^�ԍ�����͂��Ă��������B");
			obj.focus();
		    return false;
		}

		//��������󗓂̏ꍇ�́A���̍s��
		if(obj.value.length == 0 && obj2.value.length == 0){
			continue;
		}
		
		//�p���`�F�b�N
		//2017/05/08 H.Yoshikawa Upd Start
		//ret = CheckEisuji(obj.value);
		ret = CheckEisujiPlus(obj.value, "-");
		//2017/05/08 H.Yoshikawa Upd End
  		if(ret == false){
			//2017/05/08 H.Yoshikawa Upd Start
    		//alert("�o�^�ԍ��͔��p�p�����œ��͂��Ă��������B");
    		alert("�o�^�ԍ��͔��p�p�����܂��̓n�C�t���œ��͂��Ă��������B");
			//2017/05/08 H.Yoshikawa Upd End
			obj.focus();
		    return false;
		}
		
		//�o�^�ԍ��F18��		2017/05/08 H.Yoshikawa Upd(12����18��)
		//2017/08/08 H.Yoshikawa Del Start
		//if(obj.value.length != 18){
    	//	alert("�o�^�ԍ���18���œ��͂��Ă��������B");
		//	obj.focus();
		//    return false;
		//}
		//2017/08/08 H.Yoshikawa Del End
		
		//�m�莖�ƎҖ��F100�o�C�g�ȉ�
		maxlen = obj2.maxLength;
		maxlenZen = maxlen / 2 ;
		retA=getByte(obj2.value);
		if(retA[0]>maxlen){
		  alertStr="�S�p������" + maxlenZen + "�����ȓ��ɂ��邩\n";
		  alertStr=alertStr+"���p������"+maxlen+"�����ȓ��ɂ��Ă��������B";
		  alert("�m�莖�ƎҖ��́A" + maxlen + "�o�C�g�ȓ��œ��͂��Ă��������B\n" + maxlen + "�o�C�g�ȓ��ɂ���ɂ�"+alertStr);
		  obj2.focus();
		  return false;
		}

		//�d���`�F�b�N
		for(x = 1; x <= datacnt; x++){
			obj3 = eval("document.frm.DefCode" + x);
			if(obj.value == obj3.value && i != x){
				alert("�o�^�ԍ����d�����Ă��܂��B");
				obj3.focus();
				return false;
			}
		}
	}
	
	if(confirm("�X�V���܂��B��낵���ł����H") == false){
		return false;
	}
	document.frm.Gamen_Mode.value = "U";
	document.frm.submit();
}

</script>
</HEAD>
<body class="bckcolor" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="javascript:finit();">
<form name="frm" action="deftrade.asp" method="post">		
<SCRIPT src="./Common/KeyDown.js" type=text/javascript></SCRIPT>				  
<!-------------�������烍�O�C�����͉��--------------------------->
<table class="main2" align="center" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign=top align="right" >
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <%
			DisplayHeader2("�m�莖�Ǝ҃}�X�^�����e�i���X")
    	  %>
		  <INPUT type="hidden" name="Gamen_Mode" size="9" maxlength="1"  readonly tabindex= -1 value="<%=gfHTMLEncode(v_Mode)%>">
    	  <INPUT type="hidden" name="Data_Cnt" size="9" readonly tabindex= -1 value="<%=gfHTMLEncode(v_Data_Cnt)%>">
      </table>

		<table border=0><tr><td height="40"></td></tr></table>
        <table class="square" border="0" cellspacing="4" cellpadding="0" style="margin-right:10px;">
          <tr>
           <td>
		  	<table border="0" cellspacing="3" cellpadding="4">
	          <tr>
    	       <td>
				<table width="720" border="0" cellspacing="0" cellpadding="5">
				  <tr>
				   <td>
					  <table width="100%">
						<tr>			   		  
						  <td align="center">
							<table width="100%">
							<tr>
							  <td align="left" width="100%"> 						
								  <table width="100%" border="0" cellpadding="0" cellspacing="0">
									<tr> 
									  <td colspan="4" align="left" valign="middle" nowrap>�擪�\���w��</td>
									</tr>
									<tr>
										<td height="10"></td>										
									</tr>
									<tr>
									  <td>&nbsp;</td>	 
									  <td>
										  <table>
										  <tr>
										  <td>
											  <table border="1" cellpadding="0" cellspacing="0">
											  <tr>
											  <td bgcolor="#FFCC33">�o�^�ԍ�</td>
											  <td>
												<input type="text" name="SDefCode" size="28" maxlength="18" value="<%=gfHTMLEncode(v_SDefCode)%>">	<!-- 2017/05/08 H.Yoshikawa Upd�isize:15��28�Amaxlength:12��18�j-->
											  </td>
											  </tr>
											  </table>
										  </td>
										  <td style="font-size:12px;vertical-align:middle;">
										  �@���n�C�t���t���œ���	<!-- 2017/05/08 H.Yoshikawa Upd�i�n�C�t���Ȃ��A�n�C�t���t���j-->
										  </td>
										  <td width="10">&nbsp;</td>									 								   
										  <td>
											<input type="button" value="����" onClick="fSearch();">
										  </td>
										  </tr>
										  </table>
									  </td>
									  <td>
									  </td>
									  <td width="10">&nbsp;</td>									 								   
									</tr>									
									<tr>
										<td height="10"></td>										
									</tr>
									<% if v_Data_Cnt > 0 then%>									
									<tr>
										<td colspan="4">�}�X�^���</td>										
									</tr>
									<tr>
										<td height="10"></td>										
									</tr>									
									<tr>
										<td></td>
										<td colspan="3" style="font-size:12px;">���o�^�ԍ��́A�n�C�t���t���œ��͂��Ă��������B</td>		<!-- 2017/05/08 H.Yoshikawa Upd�i�n�C�t���Ȃ��A�n�C�t���t���j-->
									</tr>
									<tr>
										<td height="10"></td>										
									</tr>									
									<tr>
									  <td width="10">&nbsp;</td>									  
									  <td nowrap colspan="3">
									  <table border="0" cellspacing="0" cellPadding="0">
										<tr>											
											<th width="165" class="menutitle">�o�^�ԍ�</th>
											<th width="495" class="menutitle">�m�莖�Ǝ�</th>																						
										</tr>										
									  </table>									
									  </td>									  
									 </tr>
									 <tr>
										<td>&nbsp;</td>	
										<td colspan="3">		
											<div style="width:687px;height:350px; overflow-y:scroll;">
											<table border="0" cellspacing=0 cellPadding=0>																														
											<% for i=1 to UBOUND(Arr_DefCode) %>
												<tr>																						
													<% v_ItemName = "DefCode" + cstr(i) %>
													<td class="data2">
													<input type="text" name="<%= v_ItemName %>" maxlength="18" value="<%=gfHTMLEncode(Arr_DefCode(i))%>" onFocus="document.frm.<%= v_ItemName %>.select();" style="ime-mode: disabled; width:170px;">	<!-- 2017/05/08 H.Yoshikawa Upd�isize:15��28�Amaxlength:12��18�j-->
													</td>
													<% v_ItemName = "DefName" + cstr(i) %>
													<td class="data2">
													<input type="text" name="<%= v_ItemName %>" maxlength="100" value="<%=gfHTMLEncode(Arr_DefName(i))%>" onFocus="document.frm.<%= v_ItemName %>.select();" style="ime-mode: auto; width:500px;">														
													</td>	 
												</tr>	
											<% next %>
											</table>
											</div>
									    </td>
									 </tr>
									 <tr>
										<td height="10"></td>										
									</tr>		
									 <tr>									  
										<td colspan=4 align="center">						
											<input type="button" value="�}�X�^�X�V" onClick="fUpd();">											
									  	</td>
									  </tr>									 
									 <% end if%>
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

    </td>
 </tr>
	<%
		DisplayFooter
	%>
</table>
</form>
</body>
</HTML>
