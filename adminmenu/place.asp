<%@LANGUAGE="VBSCRIPT" CODEPAGE="932"%>
<%
'**********************************************
'  【プログラムＩＤ】　: place.asp
'  【プログラム名称】　: 蔵置場所コードメンテナンス
'
'  （変更履歴）
'
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
	dim v_OpeCode
	dim Arr_Terminal
	dim Arr_StockPlace		'蔵置場所コード
	dim cnt,i
	dim v_Msg
	dim v_FocusItem
	dim v_ItemName
	
	redim Arr_Terminal(0)
	redim Arr_StockPlace(0)
	
	const l_ProgramID = "place"
	
	cnt = 0	
		
	'----------------------------------------
    ' 再描画前の項目取得
    '----------------------------------------	
	call LfRequestItem() 
	
	if v_Data_Cnt = "" then
		v_Data_Cnt = 0
	end if
	
	if v_Mode = "U" then
		call LfUpdPlaceCode()
	end if 
	
	if (v_Mode = "S" or v_Mode = "U") and v_OpeCode <> "" then
		call LfGetOperator() 
	end if 	
	
'-----------------------------
'   描画前の画面項目を取得
'-----------------------------
function LfRequestItem()	
	v_Mode = request.form("Gamen_Mode")
	v_Data_Cnt = request.form("Data_Cnt")
	v_OpeCode = ucase(request.form("txtOpeCode"))
	
	for i = 1 to CInt(v_Data_Cnt)
		redim preserve Arr_Terminal(v_Data_Cnt)
		redim preserve Arr_StockPlace(v_Data_Cnt)
		Arr_Terminal(i) = ucase(request.form("txtTerminal" & i))
		Arr_StockPlace(i) = ucase(request.form("txtStockPlace" & i))		
	next
	
end function

function LfGetOperator()
	'----------------------------------------
	' ＤＢ接続
	'----------------------------------------        
	ConnectSvr conn, rs
	
	sql = "SELECT Terminal,StockPlace From mPlaceCode PC"
	sql = sql & " INNER JOIN mOperator OP ON PC.Operator = OP.OpeCode"
	sql = sql & " WHERE PC.Operator = '" & gfSQLEncode(v_OpeCode) & "'"	
	sql = sql & " ORDER BY Terminal,StockPlace"		
	rs.Open sql, conn, 0, 1, 1

	on error resume next
	while not rs.eof
		cnt = cnt + 1			
		redim preserve Arr_Terminal(cnt)
		redim preserve Arr_StockPlace(cnt)		
		Arr_Terminal(cnt) = gfTrim(rs("Terminal"))		'ターミナルコード		
		Arr_StockPlace(cnt) = gfTrim(rs("StockPlace"))	'蔵置場所コード	
		rs.movenext
	wend
	
	redim preserve Arr_Terminal(cnt+1)
	redim preserve Arr_StockPlace(cnt+1)		
	Arr_Terminal(cnt+1) = ""
	Arr_StockPlace(cnt+1) = ""

	v_Data_Cnt = cnt + 1	
		
	if v_Data_Cnt = 0 then
		v_Mode = ""
		v_Msg = "データがありません。"
		v_FocusItem = "txtOpeCode"
	end if
	
	rs.close
	conn.Close
end function

function LfUpdPlaceCode()
	'----------------------------------------
	' ＤＢ接続
	'----------------------------------------        
	ConnectSvr conn, rs
	
	sql = "DELETE FROM mPlaceCode WHERE Operator = '" & gfSQLEncode(v_OpeCode) & "'"
	
	conn.execute sql
	
	for i = 1 to CInt(v_Data_Cnt) 
		if gfTrim(Arr_Terminal(i)) <> "" then
			sql = "INSERT INTO mPlaceCode(Operator,Terminal,UpdtTime,UpdtPgCd,UpdtTmnl,StockPlace)"
			sql = sql & " VALUES("
			sql = sql & "'" & gfSQLEncode(v_OpeCode) & "',"
			sql = sql & "'" & gfSQLEncode(Arr_Terminal(i)) & "',"		
			sql = sql & "current_timestamp,"
			sql = sql & "'" & gfSQLEncode(l_ProgramID) & "',"		
			sql = sql & "'" & gfSQLEncode(ucase(Request.ServerVariables("SERVER_NAME"))) & "',"		
			sql = sql & "'" & gfSQLEncode(Arr_StockPlace(i)) & "')"
					
			conn.execute sql
		
			if err.number<>0 then				'--- エラー
				conn.rollbacktrans
				v_Msg = "変更できません。"
			end if
		end if
	next		
	
	conn.Close
end function
%>


<SCRIPT src="./Common/function.js" type=text/javascript></SCRIPT>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<TITLE>ＨｉＴＳ-蔵置場所コードメンテナンス</TITLE>
<link href="./Common/style.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">

function finit(){
	document.frm.Gamen_Mode.value = "<%=v_Mode%>"; 
	document.frm.Data_Cnt.value = "<%=v_Data_Cnt%>";
	document.frm.txtOpeCode.value = "<%=v_OpeCode%>";
	
	if (document.frm.Gamen_Mode.value != "" || document.frm.Gamen_Mode.value == ""){
        // エラー等のメッセージとフォーカス
        if ("<%=v_Msg%>" != ""){
            alert("<%=v_Msg%>");

            //フォーカス位置設定
            for( var i=0; i < document.frm.elements.length; i++ ){
                 if ((document.frm.elements[i].type == "text") &&
                     document.frm.elements[i].name == "<%=v_FocusItem%>"){
                     document.frm.elements[i].focus();  
                     return false;
                 }    
            }			
            return false;
		}else{
			document.frm.txtOpeCode.focus();
		}
	}
	
	if(document.frm.Gamen_Mode.value != "" && "<%=v_Data_Cnt%>" != 0){
		<%	
			if v_Data_Cnt > 0 then
				for i=1 to v_Data_Cnt
					response.write "document.frm.txtStockPlace" & i & ".value ='" & Arr_StockPlace(i) & "';" & vbcrlf    '蔵置場所コード					
					response.write "document.frm.txtTerminal" & i & ".value ='" & Arr_Terminal(i) & "';" & vbcrlf      'ターミナルコード					
				next
				response.write "document.frm.txtTerminal1.focus();"
			end if
		%>
	}	
}

function fSearch(){
	var ret;
	
	ret = CheckEisuji(document.frm.txtOpeCode.value);
  
	if(ret == false){
    	alert("オペレータコードは半角英数字で入力してください。");
		document.frm.txtOpeCode.focus();
	    return false;
	}
	
	if (gfCHKNull(document.frm.txtOpeCode) == false){
		document.frm.txtOpeCode.focus();
        return false;
    }

	document.frm.Gamen_Mode.value = "S";
	document.frm.submit();
}

function fUpd(){
	var obj;
	var obj2;
	var obj3;
	var i,x
	var ret;
	
	for (i = 1; i <= <%=v_Data_Cnt%>; i++) {
		obj = eval("document.frm.txtTerminal" + i);	
		obj2 = eval("document.frm.txtStockPlace" + i);	
		
		ret = CheckEisuji(obj.value);
  
		if(ret == false){
    		alert("ターミナルコードは半角英数字で入力してください。");
			obj.focus();
		    return false;
		}
		
		ret = CheckEisuji(obj2.value);
  
		if(ret == false){
    		alert("蔵置場所コードは半角英数字で入力してください。");
			obj2.focus();
		    return false;
		}
	 
		for(x = 1; x <= "<%=v_Data_Cnt%>"; x++){
			obj2 = eval("document.frm.txtTerminal" + x);
			if(obj.value == obj2.value & i != x){
				alert("ターミナルコードは存在しています。");
				obj2.focus();
				return false;
			}
		}
	}
	
	document.frm.Gamen_Mode.value = "U";
	document.frm.submit();
}

function fReset(){
	document.frm.Gamen_Mode.value = "";
	document.frm.Data_Cnt.value = "0";
	document.frm.submit();
}
</script>
</HEAD>
<body class="bckcolor" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="finit();">
<form name="frm" action="place.asp" method="post">		
<SCRIPT src="./Common/KeyDown.js" type=text/javascript></SCRIPT>				  
<!-------------ここからログイン入力画面--------------------------->
<table class="main2" align="center" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <%
			DisplayHeader2("蔵置場所コードメンテナンス")
    	  %>
		  <INPUT type="hidden" name="Gamen_Mode" size="9" maxlength="1"  readonly tabindex= -1>
    	  <INPUT type="hidden" name="Data_Cnt" size="9" readonly tabindex= -1>
      </table>
      <center>
		<table border=0><tr><td height="40"></td></tr></table>
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
								  <table width="100%" border="0" cellpadding="0" cellspacing="0">
									<tr> 
									  <td colspan="4" align="left" valign="middle" nowrap>オペーレータ指定</td>
									</tr>
									<tr>
										<td height="10"></td>										
									</tr>
									<tr>
									  <td width="60">&nbsp;</td>	 
									  <td>
										  <table border="1" cellpadding="0" cellspacing="0">
										  <tr>
										  <td bgcolor="#FFCC33">オペレータコード</td>
										  <% if v_Mode = "" then%>
											  <td>
											<input type="text" name="txtOpeCode" size="5" maxlength="3" onFocus="document.frm.txtOpeCode.select();">
											  </td>										  
										  <% else %>
											  <td>
											<input type="text" name="txtOpeCode" size="5" maxlength="3" class="chrReadOnly3" readonly tabindex= -1>
											  </td>
										  <% end if %>
										  </tr>										  
										  </table>										  										  
									  </td>									  
									  <td>
										  <table border="0" cellpadding="0" cellspacing="0">
										  <tr>								  
											  <td>
												<% if v_Mode = "" then%>
													<input type="button" value="マスタ表示" onClick="fSearch();">
												<% end if %>
											  </td>
										  </tr>	
										  </table>
									  </td>
									  <td width="30">&nbsp;</td>									 								   
									</tr>									
									<tr>
										<td height="10"></td>										
									</tr>
									<% if v_Data_Cnt > 0 then%>									
									<tr>
										<td colspan="4">マスタ情報</td>										
									</tr>
									<tr>
										<td height="10"></td>										
									</tr>									
									<tr>
									  <td width="60">&nbsp;</td>									  
									  <td nowrap colspan="3">
									  <table border="0" cellspacing="0" cellPadding="0">
										<tr>											
											<th width="139" class="menutitle">ターミナルコード</th>
											<th width="141" class="menutitle">蔵置場所コード</th>																						
										</tr>										
									  </table>									
									  </td>									  
									 </tr>
									 <tr>
										<td width="60">&nbsp;</td>	
										<td colspan="3">		
											<div style="width:311px;height:120px; overflow-y:scroll;">
											<table border="0" cellspacing=0 cellPadding=0>																														
											<% for i=1 to UBOUND(Arr_Terminal) %>
												<tr>																						
													<% v_ItemName = "txtTerminal" + cstr(i) %>
													<td width="144" class="data2">
													<input type="text" name="<%= v_ItemName %>" size="2" maxlength="2" value="<%=Arr_Terminal(i)%>" onFocus="document.frm.<%= v_ItemName %>.select();">
													</td>
													<% v_ItemName = "txtStockPlace" + cstr(i) %>
													<td class="data2">
													<input type="text" name="<%= v_ItemName %>" size="20" maxlength="10" value="<%=Arr_StockPlace(i)%>" onFocus="document.frm.<%= v_ItemName %>.select();">														
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
										<td width="60">&nbsp;</td>	
										<td align="center">						
											<input type="button" value="マスタ更新" onClick="fUpd();">											
									  	</td>
									  </tr>									 
									 <% end if%>
								  </table>
								  <br>
								  <center>
								  <br>
								  <% if v_Mode = "" then %>
								  	<a href="menu.asp">閉じる</a>
								  <% else %>
								  	<a href="Javascript:fReset();">閉じる</a>
								  <% end if %>
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
