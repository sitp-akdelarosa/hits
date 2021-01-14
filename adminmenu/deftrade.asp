<%@LANGUAGE="VBSCRIPT" CODEPAGE="932"%>
<%
'**********************************************
'  【プログラムＩＤ】　: deftrade.asp
'  【プログラム名称】　: 確定事業者マスタメンテナンス
'
'  （変更履歴）
'  2017/04/05 H.Yoshikawa 新規作成
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
	dim Arr_DefCode		'確定事業者コード
	dim Arr_DefName		'確定事業者名
	dim cnt,i
	dim v_Msg
	dim v_FocusItem
	dim v_ItemName
	
	redim Arr_DefCode(0)
	redim Arr_DefName(0)
	
	const l_ProgramID = "DefTrade"
	
	cnt = 0	
		
	'----------------------------------------
    ' 再描画前の項目取得
    '----------------------------------------	
	call LfRequestItem() 
	
	if v_Mode = "U" then
		call LfUpdData()
	end if 
	
	call LfSearchData() 
	
	
'-----------------------------
'   描画前の画面項目を取得
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
	' ＤＢ接続
	'----------------------------------------        
	ConnectSvr conn, rs
	
	cnt = 0
	
	'検索条件ありの場合、最初に表示
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
			Arr_DefCode(cnt) = gfTrim(rs("DefCode"))	'確定事業者コード
			Arr_DefName(cnt) = gfTrim(rs("DefName"))	'確定事業者名
			rs.movenext
		wend
		rs.close
	end if
	
	'全検索（検索条件ありの場合は、指定番号以外のデータ）
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
		Arr_DefCode(cnt) = gfTrim(rs("DefCode"))	'確定事業者コード
		Arr_DefName(cnt) = gfTrim(rs("DefName"))	'確定事業者名
		rs.movenext
	wend
	rs.close
	
	'新規用空フィールド追加
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
	' ＤＢ接続
	'----------------------------------------        
	ConnectSvr conn, rs
	conn.begintrans

	sql = "DELETE FROM mDefTrade "
	
	conn.execute sql
	if err.number<>0 then				'--- エラー
		conn.rollbacktrans
		v_Msg = "マスタの削除に失敗しました。"
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
		
			if err.number<>0 then				'--- エラー
				conn.rollbacktrans
				v_Msg = "マスタの追加に失敗しました。"
				v_FocusItem = "DefCode" & i
				return false
			end if
		end if
	next
	conn.committrans
	
	v_Msg = "更新しました。"

	conn.Close
end function
%>


<SCRIPT src="./Common/function.js" type=text/javascript></SCRIPT>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<TITLE>ＨｉＴＳ-確定事業者マスタメンテナンス</TITLE>
<link href="./Common/style.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">

function finit(){
	var i;

    // エラー等のメッセージとフォーカス
    if ("<%=v_Msg%>" != ""){
        alert("<%=v_Msg%>");

        //フォーカス位置設定
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
    //	alert("検索する登録番号を入力してください。");
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
		
		//いずれかを入力の場合、もう一方も必須
		if(obj.value.length != 0 && obj2.value.length == 0){
    		alert("確定事業者名を入力してください。");
			obj2.focus();
		    return false;
		}
		if(obj.value.length == 0 && obj2.value.length != 0){
    		alert("登録番号を入力してください。");
			obj.focus();
		    return false;
		}

		//いずれも空欄の場合は、次の行へ
		if(obj.value.length == 0 && obj2.value.length == 0){
			continue;
		}
		
		//英数チェック
		//2017/05/08 H.Yoshikawa Upd Start
		//ret = CheckEisuji(obj.value);
		ret = CheckEisujiPlus(obj.value, "-");
		//2017/05/08 H.Yoshikawa Upd End
  		if(ret == false){
			//2017/05/08 H.Yoshikawa Upd Start
    		//alert("登録番号は半角英数字で入力してください。");
    		alert("登録番号は半角英数字またはハイフンで入力してください。");
			//2017/05/08 H.Yoshikawa Upd End
			obj.focus();
		    return false;
		}
		
		//登録番号：18桁		2017/05/08 H.Yoshikawa Upd(12桁⇒18桁)
		//2017/08/08 H.Yoshikawa Del Start
		//if(obj.value.length != 18){
    	//	alert("登録番号は18桁で入力してください。");
		//	obj.focus();
		//    return false;
		//}
		//2017/08/08 H.Yoshikawa Del End
		
		//確定事業者名：100バイト以下
		maxlen = obj2.maxLength;
		maxlenZen = maxlen / 2 ;
		retA=getByte(obj2.value);
		if(retA[0]>maxlen){
		  alertStr="全角文字を" + maxlenZen + "文字以内にするか\n";
		  alertStr=alertStr+"半角文字を"+maxlen+"文字以内にしてください。";
		  alert("確定事業者名は、" + maxlen + "バイト以内で入力してください。\n" + maxlen + "バイト以内にするには"+alertStr);
		  obj2.focus();
		  return false;
		}

		//重複チェック
		for(x = 1; x <= datacnt; x++){
			obj3 = eval("document.frm.DefCode" + x);
			if(obj.value == obj3.value && i != x){
				alert("登録番号が重複しています。");
				obj3.focus();
				return false;
			}
		}
	}
	
	if(confirm("更新します。よろしいですか？") == false){
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
<!-------------ここからログイン入力画面--------------------------->
<table class="main2" align="center" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign=top align="right" >
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <%
			DisplayHeader2("確定事業者マスタメンテナンス")
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
									  <td colspan="4" align="left" valign="middle" nowrap>先頭表示指定</td>
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
											  <td bgcolor="#FFCC33">登録番号</td>
											  <td>
												<input type="text" name="SDefCode" size="28" maxlength="18" value="<%=gfHTMLEncode(v_SDefCode)%>">	<!-- 2017/05/08 H.Yoshikawa Upd（size:15⇒28、maxlength:12⇒18）-->
											  </td>
											  </tr>
											  </table>
										  </td>
										  <td style="font-size:12px;vertical-align:middle;">
										  　※ハイフン付きで入力	<!-- 2017/05/08 H.Yoshikawa Upd（ハイフンなし、ハイフン付き）-->
										  </td>
										  <td width="10">&nbsp;</td>									 								   
										  <td>
											<input type="button" value="検索" onClick="fSearch();">
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
										<td colspan="4">マスタ情報</td>										
									</tr>
									<tr>
										<td height="10"></td>										
									</tr>									
									<tr>
										<td></td>
										<td colspan="3" style="font-size:12px;">※登録番号は、ハイフン付きで入力してください。</td>		<!-- 2017/05/08 H.Yoshikawa Upd（ハイフンなし、ハイフン付き）-->
									</tr>
									<tr>
										<td height="10"></td>										
									</tr>									
									<tr>
									  <td width="10">&nbsp;</td>									  
									  <td nowrap colspan="3">
									  <table border="0" cellspacing="0" cellPadding="0">
										<tr>											
											<th width="165" class="menutitle">登録番号</th>
											<th width="495" class="menutitle">確定事業者</th>																						
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
													<input type="text" name="<%= v_ItemName %>" maxlength="18" value="<%=gfHTMLEncode(Arr_DefCode(i))%>" onFocus="document.frm.<%= v_ItemName %>.select();" style="ime-mode: disabled; width:170px;">	<!-- 2017/05/08 H.Yoshikawa Upd（size:15⇒28、maxlength:12⇒18）-->
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
											<input type="button" value="マスタ更新" onClick="fUpd();">											
									  	</td>
									  </tr>									 
									 <% end if%>
								  </table>
								  <br>
								  <center>
								  <br>
								  	<a href="menu.asp">閉じる</a>
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
