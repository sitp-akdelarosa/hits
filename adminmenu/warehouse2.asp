<%@LANGUAGE="VBSCRIPT" CODEPAGE="932"%>
<%
'**********************************************
'  【プログラムＩＤ】　: warehouse2.asp
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
<SCRIPT src="./Common/function.js" type=text/javascript></SCRIPT>
<%
	dim conn, rs, sql
	dim v_OpeCode
	dim Arr_Terminal
	dim Arr_StockPlace
	dim cnt,i
	
	redim Arr_Terminal(0)
	redim Arr_StockPlace(0)
	
	v_OpeCode = ucase(request.querystring("code"))
	cnt = 0
	
	if v_OpeCode <> "" then
		'----------------------------------------
        ' ＤＢ接続
        '----------------------------------------        
        ConnectSvr conn, rs
		
		sql = "SELECT Terminal,StockPlace From mPlaceCode PC"
		sql = sql & " INNER JOIN mOperator OP ON PC.Operator = OP.OpeCode"
		sql = sql & " WHERE PC.Operator = '" & gfSQLEncode(v_OpeCode) & "'"		
        rs.Open sql, conn, 0, 1, 1

		on error resume next
		while not rs.eof
			cnt = cnt + 1			
			redim preserve Arr_Terminal(cnt)
			redim preserve Arr_StockPlace(cnt)		
			Arr_Terminal(cnt) = gfTrim(rs("Terminal"))		
			Arr_StockPlace(cnt) = gfTrim(rs("StockPlace"))
			rs.movenext
        wend
		
        rs.close
		conn.Close
	end if 

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<TITLE>蔵置場所コードメンテナンス</TITLE>
<link href="./Common/style.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
</HEAD>
<body class="bckcolor" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frm" action="warehouse.asp" method="post">						  
<!-------------ここからログイン入力画面--------------------------->
<table class="main2" align="center" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <%
			DisplayHeader2("蔵置場所コードメンテナンス")
    	  %>
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
							<table width="100%">
							<tr>
							  <td align="left" width="100%"> 						
								  <table width="100%" border="0">									
									<tr>
										<td>オペレータ</td>										
									</tr>
									<tr>
										<td height="10"></td>				  
									</tr>
									<tr>									
									  <td align="center">オペレータコード</td>
									  <td><%=v_OpeCode%></td>									  			  
									</tr>
									<tr>
										<td height="10"></td>				  
									</tr>
									<tr>
										<td>マスタ情報</td>										
									</tr>									
									<tr>									  
									  <td colspan="2">
									  <table width="100%" align="left" cellspacing=0 cellPadding=0>
											<tr>
												<td width="40"></td>
												<th width="150" class="menutitle">ターミナルコード</th>
												<th width="160" class="menutitle">蔵置場所コード</th>
												<td width="18"></td>										
											</tr>
									  </table>
									  </td>
									  <td width="150"></td>	
									 </tr>
									 <tr>
									  <td colspan="2">
									  <div style="width:350px;height:120px; overflow-y:scroll;">
										<table width="100%" align="left" cellspacing=0 cellPadding=0 border=0>											
											<% for i = 1 to UBOUND(Arr_Terminal) %>						
											<tr>
												<td width="40"></td>												
												<td width="150" class="data"><%=Arr_Terminal(i)%></td>
												<td width="156" class="data"><%=Arr_StockPlace(i)%></td>	 
											</tr>	
											<% next %>	
										</table>										
									  </div>		
									  </td>	
									  <td width="150"></td>			  
									</tr>
								  </table>
								  <br>
								  <center>
								  <br>
								  <a href="Javascript:window.close();">閉じる</a>			
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
