<%@ LANGUAGE="VBScript" %>

<html>
<head>
<title>搬出許可照会結果</title>
<SCRIPT LANGUAGE="JavaScript">
<!---
//--->
function ClickInquiry() {
}
</SCRIPT>

</head>

<body >
<IMG border=0 height=42 src="image/title01.gif" width=311>
<br><br>
<center>
<p><IMG border=0 height=67 src="image/title12.gif" width=472><p>

<%
Set conn = Server.CreateObject("ADODB.Connection")
'conn.Open "HakataDB", "sa", "hakata"	'D20040314
conn.Open "HakataDB", "sa", ""		'I20040314
Set rsd = Server.CreateObject("ADODB.Recordset")
rsd.Open "sUseDB", conn, 0, 1, 2
if rsd.eof then
	Response.Write "システムエラー:使用DB管理テーブルにレコードがありません。"
	Response.Write "</body>"
	Response.Write "</html>"
	Response.End
else
	rem テーブルの接尾語("1"/"2")を取得
	dbsuffix = rsd("EnableDB")
	wOutUpdtTime = rsd("OutUpdtTime" & dbsuffix) 
%>
	★&nbsp; 現在の情報は&nbsp; <u><b><%=Month(wOutUpdtTime)%>                              月<%=Day(wOutUpdtTime)%>                                       日
										<%=FormatDateTime(wOutUpdtTime, vbShortTime)%></b></u>&nbsp; のものです。<br><br> 
	   (&nbsp; 次回更新予定は&nbsp; <b><%=Month(rsd("OutPUpdtTime"))%>                              月<%=Day(rsd("OutPUpdtTime"))%>                                       日
										<%=FormatDateTime(rsd("OutPUpdtTime"), vbShortTime)%></b>&nbsp; です。&nbsp;) 
<%
end if
rsd.Close
	
if Request.Form("blnumber") <> "" and  Len(TRIM(Request.Form("blnumber"))) <= 4 and Request.Form("container") = "" then
else
	contval=Ucase(TRIM(Request.Form("container")))
	blval=Ucase(TRIM(Request.Form("blnumber")))
	tsubmit="照    会"
	%>
	<!--#include file="ComnForm.inc"-->				
	<%
end if

if Request.Form("container") = "" AND  Request.Form("blnumber") = "" then
	Response.Write "<br><p>コンテナ番号かＢＬ番号を入力してください。</p><br>"
elseIF Request.Form("container") <> "" AND  Request.Form("blnumber") <> "" then
	Response.Write "<br><p>コンテナ番号かＢＬ番号のどちらか一方を入力してください。</p><br>"
else
	if Request.Form("blnumber") <> "" then
		rem B/L照会処理
		
		rem B/L番号で搬出許可コンテナを検索する
		sblno = TRIM(Request.Form("blnumber"))
'2000/11/8 start
		if Len(sblno) <= 4 then						'４桁未満の入力
			dim iblcnt
			dim slblno
			iblcnt = 0
			slblno = "%" & sblno 
			sql = "SELECT RTrim([BLNo]) AS BL  FROM sOutBLCont" & dbsuffix  & " GROUP BY RTrim([BLNo]), BLNo "
			sql = sql  & "HAVING (((RTrim([BLNo])) Like '" & slblno & "'))"
			Set rs4 = Server.CreateObject("ADODB.Recordset") 
			rs4.Open sql, conn, 0, 1, 1 
			if rs4.eof then 
				'ＢＬ番号再表示
				contval=""
				blval=sblno
				tsubmit="照    会"
				%>
				<!--#include file="ComnForm.inc"-->				
				<%
				Response.Write "<br><p>該当B／Lが存在しません。</p><br>"
				Response.Write "<br><p><A href=""index.asp"">メニューに戻る</A></p>"
				Response.Write "</body>"
				Response.Write "</html>"
				Response.End
			else
				do while not rs4.eof 
					iblcnt = CInt(iblcnt) + 1
					if CInt(icnt) >= 2 then 
						exit do 
					end if 
					sblno = trim(rs4("BL"))		'ＢＬ番号再設定
					rs4.MoveNext 
				loop
				if CInt(iblcnt) >= 2 then
					'ＢＬ番号再表示
					contval=""
					blval=mid(slblno,2)
					tsubmit="照    会"
					%>
					<!--#include file="ComnForm.inc"-->				
					<%
					Response.Write "<br><p>ＢＬ番号が複数存在しています。</p><br><br>"
					Response.Write "<p><A href=""index.html"">メニューに戻る</A></p>"
					Response.Write "</body>"
					Response.Write "</html>"
					Response.End
				end if
			end if
			'ＢＬ番号再表示
			contval=""
			blval=sblno
			tsubmit="照    会"
			%>
			<!--#include file="ComnForm.inc"-->				
			<%
		end if	
'2000/11/8  end
		scont = ""										
		sql = "SELECT ContNo FROM sOutBLCont" & dbsuffix & " WHERE BLNo='" & sblno & "'"

		Set rs3 = Server.CreateObject("ADODB.Recordset")
		rs3.Open sql, conn, 0, 1, 1
		do while not rs3.eof
			if scont <> "" then
				scont = scont & " OR "
			end if
			scont = scont & "ContNo='" & rs3("ContNo") & "'"
			rs3.MoveNext
		loop
		rs3.close
	else
'2000/11/8  start
		dim CntNo
		CntNo = TRIM(Request.Form("container"))
		if IsNumeric(CntNo) then
			dim ictcnt
			dim slctno
			ictcnt = 0
			slctno = "%" & TRIM(Request.Form("container"))
			sql = "SELECT RTrim([ContNo]) AS CT  FROM sOutContainer" & dbsuffix  & " GROUP BY RTrim([ContNo]), ContNo "
			sql = sql  & "HAVING (((RTrim([ContNo])) Like '" & slctno & "'))"
			Set rs5 = Server.CreateObject("ADODB.Recordset")
			rs5.Open sql, conn, 0, 1, 1
			if rs5.eof then
				Response.Write "<br><p>在庫コンテナにはありません。</p><br><br>"
				Response.Write "<p><A href=""index.asp"">メニューに戻る</A></p>"
				Response.Write "</body>"
				Response.Write "</html>"
				Response.End
			else
				do while not rs5.eof
					ictcnt = CInt(ictcnt) + 1
					if CInt(icnt) >= 2 then
						exit do
					end if
					CntNo = trim(rs5("CT"))		'コンテナ番号再設定
					rs5.MoveNext
				loop
				if CInt(ictcnt) >= 2 then
					Response.Write "<br><p>コンテナ番号が複数存在します。</p><br><br>"
					Response.Write "<p><A href=""index.asp"">メニューに戻る</A></p>"
					Response.Write "</body>"
					Response.Write "</html>"
					Response.End
				end if
			end if
		end if
'2000/11/8  end
		scont = "ContNo='" & CntNo & "'"
	end if
	if scont = "" then
		if Request.Form("blnumber") <> "" then
			Response.Write "<br><p>該当B／Lが存在しません。</p><br>"
		else
			Response.Write "<br><p>在庫コンテナにはありません。</p><br>"
		end if
	else
		rem 同一コンテナ複数ＢＬのチェック
		Set rs2 = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT VslCode, Voyage FROM sOutContainer" & dbsuffix & " WHERE (" & scont & ")" & svwhere
		sql = sql & " GROUP BY VslCode, Voyage; "
		rs2.Open sql, conn, 0, 1, 1
		icnt = 0
		do while not rs2.eof
			icnt = icnt + 1
			if icnt >= 2 then
				exit do
			end if
			rs2.MoveNext
		loop
		rs2.close
		if icnt >= 2 then
			Response.Write "<br><p>同一のＢＬ番号が複数存在します。コンテナ番号検索を行ってください</p><br>"
		else
			Set rs = Server.CreateObject("ADODB.Recordset")
			sql = "SELECT * FROM sOutContainer" & dbsuffix & " WHERE (" & scont & ")" & svwhere
			sql = sql & " order by ContNo "
			rs.Open sql, conn, 0, 1, 1
			if rs.eof then
				if Request.Form("blnumber") <> "" then
					Response.Write "<br><p>コンテナ情報が存在しません。</p><br>"
				else
					Response.Write "<br><p>在庫コンテナにはありません。</p><br>"
				end if
			else
				dim iGaito
				iGaito = "1"				'該当データあり
%> 
				<table border="1" style="HEIGHT: 98px; WIDTH: 739px"> 
  					<tr> 
    					<td bgcolor="#f4a460" align="middle"><b>コンテナ番号</b></td> 
		    			<td bgcolor="#ff6699" align="middle"><b>搬出</b></td> 
		    			<td bgcolor="#f4a460" align="middle"><b>サイズ</b></td> 
    					<td bgcolor="#f4a460" align="middle"><b>場所</b></td>  
    					<td bgcolor="#f4a460" align="middle"><b>フリータイム</b></td>  
    					<td bgcolor="#f4a460" align="middle"><b>搬出可能日</b></td> 
    					<td bgcolor="#f4a460" align="middle"><b>税関手続き</b></td> 
	    				<td bgcolor="#f4a460" align="middle"><b>ＤＯ</b></td> 
    					<td bgcolor="#f4a460" align="middle"><b>ＯＬＴ運送許可期間</b></td> 
					</tr> 
<%
				lineno = 1
				do while not rs.eof
					rem 搬出可能か
					soute = "　"
					do 
						if rs("FullEmpty") <> "F"  then
							soute = "○"
							exit do
						end if

						if isnull(rs("DelOKDate")) then
							soute = "×"
							exit do
						end if
	
						if Date < (rs("DelOKDate"))  then
							soute = "×"
							exit do
						end if

						if not isnull(rs("DemFTDate")) then
							if Date > rs("DemFTDate") then
								soute = "×"
								exit do
							end if
						end if

						if isnull(rs("OLTFrom"))  then	
							soute = "○"
							exit do
						else
							if rs("OLTFrom") <= Date  And  Date <= rs("OLTTo") then
							else
								soute = "×"
								exit do
							end if
						end if
						soute = "○"
						exit do
					loop
				
					'場所
					dim sPlace
					sPlace = "" 
					if trim(rs("Terminal")) = "KA"  then
						sPlace = "香椎"
					else
'''						sPlace = "箱崎"
						if trim(rs("Terminal")) = "IC"  then
   							sPlace = "ＩＣＣＴ"
						else
							sPlace = "箱崎"
						end if	
					end if	

					'フリータイムFreetime
					dim sFreeTime
					sFreeTime = ""
					if not isnull(rs("DemFTDate")) then
				    	sFreeTime = FormatDateTime(rs("DemFTDate"), vbShortDate)
					else
						sFreeTime = "<br>"
					end if

%>
					<tr> 

<%
						if soute = "○" then %>
			    			<td align="middle" bgcolor="#00ffff"><%=rs("ContNo")%></td>
				    		<td align="middle" bgcolor="#00ffff"><font color="#ff0000"><%=soute%></font></td> 
				    		<td align="middle" bgcolor="#00ffff"><%=rs("ContSize")%></td>
				    		<td align="middle" bgcolor="#00ffff"><%=sPlace%></td>
				    		<td align="middle" bgcolor="#00ffff"><%=sFreeTime%></td>
<%
						else %>
			    			<td align="middle" ><%=rs("ContNo")%></td>
				    		<td align="middle" ><font color="#ff0000"><%=soute%></font></td> 
				    		<td align="middle" ><%=rs("ContSize")%></td>
				    		<td align="middle" ><%=sPlace%></td>
				    		<td align="middle" ><%=sFreeTime%></td>
<%
						end if%>

			    		<td align="middle">
<%
							'搬出可能日
							'if soute = "×" then
								if not isnull(rs("DelOkDate")) then
				    				Response.Write FormatDateTime(rs("DelOkDate"), vbShortDate)
								else
									Response.Write "<br>"
								end if
							'else
							'	Response.Write "<br>"
							'end if
%>
			    		</td> 
<%
							if soute = "×" then
								if rs("FullEmpty") <> "F"  then
				    				sdo = "○"
								else
									if trim(rs("DsListNo")) <> "Y" then
										sZei = "×"
									else
										if not isnull(rs("OLTFrom")) then
											if rs("OLTFrom") <= Date  And  Date <= rs("OLTTo") then
			    								sZei = "○"
											else
			    								sZei = "×"
											end if
										else
											sZei = "○"
										end if
									end if
								end if
							else
								sZei = "　"
							end if
%>
				    	<td align="middle"><%=sZei%></font></td> 
<%
							'Ｄｏ
							if soute = "×" then
								if rs("FullEmpty") <> "F"  then
			    					sdo = "○"
								else
									if  rs("DOStatus") = "Y" then
			    						sdo = "○"
									else
				    					sdo = "×"
									end if
								end if

							else
								sdo = "　"
							end if
%>
				    	<td align="middle"><%=sdo%></font></td> 

			    		<td align="middle">
<%
							'ＯＬＴ運送許可期間
							if soute = "×" then
				   				if not isnull(rs("OLTFrom")) then
				   					sfrom = FormatDateTime(rs("OLTFrom"), vbShortDate)
			   						sto = FormatDateTime(rs("OLTTo"), vbShortDate)
			    					sfto = sfrom & "～" & sto 
					   				Response.Write sfto
								else
									Response.Write "<br>"
				   				end if
							else
								Response.Write "<br>"
							end if
%>	
			    		</td> 
					</tr> 
<%
					lineno = lineno + 1
					rs.MoveNext
				loop
				do while lineno <= 3
%>
					<tr> 
				    	<td align="middle"><br></td> 
				    	<td align="middle"><br></td> 
				    	<td align="middle"><br></td> 
			    		<td align="middle"><br></td> 
			    		<td align="middle"><br></td> 
				    	<td align="middle"><br></td> 
				    	<td align="middle"><br></td> 
				    	<td align="middle"><br></td> 
				    	<td align="middle"><br></td> 
					</tr> 
<%
					lineno = lineno + 1
				loop
%>
			</table> 
<%
			end if
			rs.close
		end if	
		conn.close
	end if
end if
%>
<br> 
<p><A href="index.asp">メニューに戻る</A></p>

<%
	if iGaito = "1" then
%>
	<table border="1" width="637" style="HEIGHT: 133px; WIDTH: 569px" bgColor=#ffff99>
  <TBODY> 
  		<tr> 
			<td align="left"><b>〔搬出〕</b></td> 
			<td align="left"><b>　搬出可：○　　　不可：×</b></td> 
		</tr>

  		<tr> 
    		<td align="left"><b>〔フリータイム〕</b></td>  
			<td align="left"><b>　搬出可能期間</b></td> 
		</tr>

  		<tr> 
    		<td align="left"><b>〔搬出可能日〕</b></td> 
			<td align="left"><b>　この日以降搬出可能</b></td> 
		</tr>

  		<tr> 
    		<td align="left"><b>〔税関手続き〕</b></td> 
			<td align="left"><b>　ＯＬＴ許可取得を含め輸入税関手続き終了を示す</b></td> 
		</tr> 

  		<tr> 
			<td align="left"><b>〔ＤＯ〕</b></td> 
			<td align="left"><b>　デリバリオーダ</b></td> 
		</tr> 
<%	end if	%>
</center></TBODY></TABLE> 
</body> 
</html> 
