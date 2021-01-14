<% @LANGUAGE = VBScript %>
<%
%><% Option Explicit %>
<%
	'HTTPコンテンツタイプ設定
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="CommonFunc.inc"-->
<!--#include File="Common.inc"-->
<%
  Dim s_VslCode
  Dim s_VslName
  Dim ObjRS2,ObjConn2
  Dim StrSQL
  Dim Num2
  Dim abspage,pagecnt
  Dim x,i
  Dim openerForm
  Dim openerFieldVslNm
  Dim openerFieldVslCd
  Dim openerFieldVoy
  Dim openerFieldDspVoy
  Dim DispKbn

  const gcPage = 10
  const rownum = 20

  openerForm = Request.QueryString("tgt")
  openerFieldVslNm = Request.QueryString("fldvn")
  openerFieldVslCd = Request.QueryString("fldvc")
  openerFieldVoy = Request.QueryString("fldvy")
  openerFieldDspVoy = Request.QueryString("flddspvy")
  DispKbn = Request.QueryString("dspkbn")

  if gfTrim(openerForm) = "" then
  	openerForm = gfTrim(Request.Form("openerForm"))
  end if
  if gfTrim(openerFieldVslNm) = "" then
  	openerFieldVslNm = gfTrim(Request.Form("openerFieldVslNm"))
  end if
  if gfTrim(openerFieldVslCd) = "" then
  	openerFieldVslCd = gfTrim(Request.Form("openerFieldVslCd"))
  end if
  if gfTrim(openerFieldVoy) = "" then
  	openerFieldVoy = gfTrim(Request.Form("openerFieldVoy"))
  end if
  if gfTrim(openerFieldDspVoy) = "" then
  	openerFieldDspVoy = gfTrim(Request.Form("openerFieldDspVoy"))
  end if
  if gfTrim(DispKbn) = "" then
  	DispKbn = gfTrim(Request.Form("DispKbn"))
  end if

  s_VslCode = gfTrim(Request.Form("S_VslCode"))
  s_VslName = gfTrim(Request.Form("S_VslName"))
  
'セッションの有効性をチェック
  CheckLoginH
  
  '船名、次航リスト取得
  ConnDBH ObjConn2, ObjRS2

  StrSQL = "SELECT distinct po.VslCode, po.VoyCtrl, mv.FullName as VslName, vs.DsVoyage, vs.LdVoyage "
  StrSQL = StrSQL & " FROM VslPort po "
  StrSQL = StrSQL & " INNER JOIN mVessel mv on mv.VslCode = po.VslCode "
  StrSQL = StrSQL & " INNER JOIN VslSchedule vs on vs.VslCode = po.VslCode and vs.VoyCtrl = po.VoyCtrl "
  StrSQL = StrSQL & " WHERE convert(char(10), po.ETD, 111) >= convert(char(10), GETDATE(), 111) "
  'StrSQL = StrSQL & "   AND po.PortCode = 'JPHKT'"
  if s_VslCode <> "" then
  	  StrSQL = StrSQL & "   AND po.VslCode like '" & gfSQLEncode(s_VslCode) & "%'"
  end if
  if s_VslName <> "" then
  	  StrSQL = StrSQL & "   AND mv.FullName like '%" & gfSQLEncode(s_VslName) & "%'"
  end if
  StrSQL = StrSQL & " ORDER BY po.VslCode, po.VoyCtrl "
  ObjRS2.PageSize = rownum
  ObjRS2.CacheSize = rownum
  ObjRS2.CursorLocation = 3
  ObjRS2.Open StrSQL, ObjConn2

  Num2 = ObjRS2.recordcount	

  if Num2 > rownum then 
	If CInt(Request("pagenum2")) = 0 Then
		ObjRS2.AbsolutePage = 1
	Else
		If CInt(Request("pagenum2")) <= ObjRS2.PageCount Then
			ObjRS2.AbsolutePage = CInt(Request("pagenum2"))
		Else
			ObjRS2.AbsolutePage = 1
		End If
	End If		 
  end if

  if err <> 0 then
	DisConnDBH ObjConn2, ObjRS2	'DB切断
	jampErrerP "2","b301","01","船名・次航検索","102","SQL:<BR>" & StrSQL & err.description & Err.number
  end if			

  
function LfPutPage(rec,page,pagecount,link)
	dim pg, i, j
	dim FirstPage, LastPage	
	dim PageIndex
	dim PageWkNo
	dim intNextFlag
	PageIndex=0
	PageWkNo=0	
	if rec > 0 then	

		if pagecount<page then
			page=pagecount
		end if
		
		'パラメータ設定
		'--- 総件数、総ページ数 
		LastPage=pagecount		
		FirstPage=1
			
		'前のページ
		PageWkNo = page - 1

		if page>1 then
			response.write "<a href=""#"" onClick=""fPageChg('"& link & "', " & FirstPage & ");"">最初へ</a>"
			response.write "| &nbsp;"
			if PageWkNo>0 Then
				response.write "<a href=""#"" onClick=""fPageChg('"& link & "', " & PageWkNo & ");"">前へ</a>"
			Else
				response.write "<font style='color:#FFFFFF;'>前へ</font>"
			End If
		else
			response.write "<font style='color:#FFFFFF;'>最初へ</font>"
			response.write "| &nbsp;"
			response.write "<font style='color:#FFFFFF;'>前へ</font>"
		end if        		
		'--- インデックス
		'ページが1ページ以上存在する場合
		if pagecount>1 then
			response.write "| &nbsp;"

			'指定ページ数分ループ
			for i=1 to gcPage
				'ページ数算出
				PageWkNo=(gcPage*PageIndex)+i

				'ページが全ページより大きい場合は処理中断
				if pagecount< PageWkNo then
					PageWkNo=PageWkNo-1
					exit for
				end if
				'現在選択されているページの場合
				if PageWkNo=page then
					response.write "&nbsp;" & PageWkNo 
				else
					response.write "<a href=""#"" onClick=""fPageChg('"& link & "', " & PageWkNo & ");"" >&nbsp;" & PageWkNo & "</a>"
				End If
			Next
			response.write "| &nbsp;"
		End If
					
		if page<pagecount then
			'次のページ
			PageWkNo=page+1
			If PageWkNo<=LastPage Then
				response.write "<a href=""#"" onClick=""fPageChg('"& link & "', " & PageWkNo & ");"">次へ</a>"'
			Else
				response.write "<font style='color:#FFFFFF;'>次へ</font>"
			End If
			response.write "| &nbsp;"
			response.write "<a href=""#"" onClick=""fPageChg('"& link & "', " & LastPage & ");"">最後へ</a>"'            
		else
			response.write "<font style='color:#FFFFFF;'>次へ</font>"
			response.write "| &nbsp;"
			response.write "<font style='color:#FFFFFF;'>最後へ</font>"
		end if
	end if
end function

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE></TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT language=JavaScript>
<!--
function fSend(index)
{
  var VslCodes, VslNames, VoyCtrls, DsVoys, LdVoys;
  var SelectVal;
  
  VslCodes = document.getElementsByName("VslCode");
  VslNames = document.getElementsByName("VslName");
  VoyCtrls = document.getElementsByName("VoyCtrl");
  DsVoys = document.getElementsByName("DsVoyage");
  LdVoys = document.getElementsByName("LdVoyage");
  index = index - 1;
<%
  if openerFieldVslCd <> "" then
  	  Response.Write "opener." & openerForm & ".elements[""" & openerFieldVslCd & """].value=VslCodes[index].value;"
  end if
  
  if openerFieldVslNm <> "" then
  	  Response.Write "opener." & openerForm & ".elements[""" & openerFieldVslNm & """].value=VslNames[index].value;"
  end if
  if openerFieldVoy <> "" then
  	  Response.Write "opener." & openerForm & ".elements[""" & openerFieldVoy & """].value=VoyCtrls[index].value;"
  end if
  if openerFieldDspVoy <> "" then
  	if DispKbn = "DS" then
  	  Response.Write "opener." & openerForm & ".elements[""" & openerFieldDspVoy & """].value=DsVoys[index].value;"
  	else
  	  Response.Write "opener." & openerForm & ".elements[""" & openerFieldDspVoy & """].value=LdVoys[index].value;"
  	end if
  end if
%>
  //opener.<%=openerForm%>.elements["<%= openerFieldVslCd %>"].value=VslCodes[index].value;
  //opener.<%=openerForm%>.elements["<%= openerFieldVslNm %>"].value=VslNames[index].value;
  //opener.<%=openerForm%>.elements["<%= openerFieldVoy %>"].value=VoyCtrls[index].value;
  //if("<%= DispKbn %>" == "DS"){
  //	opener.<%=openerForm%>.elements["<%= openerFieldDspVoy %>"].value=DsVoys[index].value;
  //}else{
  //	opener.<%=openerForm%>.elements["<%= openerFieldDspVoy %>"].value=LdVoys[index].value;
  //}

  window.close();
}

function fPageChg(item, pageNo)
{
  document.frm.elements[item].value = pageNo;
  document.frm.submit();
}

function fSearch()
{
  document.frm.pagenum2.value = 0;
  document.frm.submit();
}

-->
</SCRIPT>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<form name="frm" method="post">

<table width="100%" height="82%" border="0" cellspacing="0" cellpadding="0">
<tr><td width="50" nowrap>&nbsp;</td>
<td>
  <div id="BDIV3" style="width: 100%; height: 100%; padding-top:20px;">
  <table width="100%">
  <TR><TD colspan=3>＜検索条件＞</TD></TR>
  <tr>
     <td width="60">船名ｺｰﾄﾞ</td>
     <td width="150">
        <input type=text name="S_VslCode" value="<%=gfHTMLEncode(s_VslCode)%>" style="ime-mode:none;"/>
     </td>
     <td rowspan=2 width><input type=button name="search" onclick="fSearch();" value=" 検　索 "/>
  </tr>
  <tr>
     <td>船名</td>
     <td>
        <input type=text name="S_VslName" value="<%=gfHTMLEncode(s_VslName)%>" />
     </td>
  </tr>
  </table>
  <table border="0" cellpadding="0" cellspacing="0">
    <tr align=right nowrap>
      <td width="100%" height="30" align=right nowrap>
          <table border="0" cellpadding="0" cellspacing="0">
            <tr>
		      <td width="100%" align="center" nowrap>
		      <!--Page Pagination Start-->
		        <%
				  If Num2 > 0 Then
					abspage = ObjRS2.AbsolutePage
					pagecnt = ObjRS2.PageCount
					call LfPutPage(Num2,abspage,pagecnt,"pagenum2")
				  End If
			     %>
		      <!--Page Pagination End-->
		      </td>
		    </tr>
		  </table> 
      </td>
    </tr>
	<tr>
		<!--Place Here Start-->
		<td nowrap>
			<div id="BDIV2">
			   	<% If Num2>0 Then%>
			   		<!--Work List Start-->	
					<table border="1" cellpadding="0" cellspacing="0" width=100% id="TBInOut">
						<thead>
						   <!--HEADER INFORMATION START-->
							<tr>
								<th class="hlist" align="center" nowrap>選択</th>
								<th class="hlist" align="center" nowrap>船名ｺｰﾄﾞ</th>
								<th class="hlist" align="center" nowrap>船名</th>
								<th class="hlist" align="center" nowrap>揚げ次航</th>
								<th class="hlist" align="center" nowrap>積み次航</th>
							</tr>
						    <!--HEADER INFORMATION END-->
						</thead>
						<tbody>
						    <!--DETAIL INFORMATION START-->
                            <% 
								x = 1
								For i=1 To ObjRS2.PageSize
								 	If Not ObjRS2.EOF Then
									x = x + 1
							%>
							<tr bgcolor="#CCFFFF">	
							<td align="center" valign="middle" width="50"  height="20" nowrap>
								<a href="#" onclick="fSend(<%=i%>);">選択</a>
							</td>
							<td align="left" valign="middle" width="60" nowrap>
                              <%=gfHTMLEncode(ObjRS2("VslCode"))%>
							  <input type="hidden" name="VslCode" value="<%=gfHTMLEncode(ObjRS2("VslCode"))%>"><BR>
                            </td>
							<td align="left" valign="middle" width="200" nowrap>
                              <%=gfHTMLEncode(ObjRS2("VslName"))%>
							  <input type="hidden" name="VslName" value="<%=gfHTMLEncode(ObjRS2("VslName"))%>"><BR>
                            </td>
							<td align="left" valign="middle" width="80" nowrap>
                              <%=gfHTMLEncode(ObjRS2("DsVoyage"))%>
							  <input type="hidden" name="DsVoyage" value="<%=gfHTMLEncode(ObjRS2("DsVoyage"))%>"><BR>
                            </td>
							<td align="left" valign="middle" width="80" nowrap>
                              <%=gfHTMLEncode(ObjRS2("LdVoyage"))%>
							  <input type="hidden" name="LdVoyage" value="<%=gfHTMLEncode(ObjRS2("LdVoyage"))%>">
							  <input type="hidden" name="VoyCtrl" value="<%=gfHTMLEncode(ObjRS2("VoyCtrl"))%>"><BR>
                            </td>
							</tr>
						    <% 
									ObjRS2.MoveNext 		
									End If
								Next	
							  ObjRS2.close    
						      DisConnDBH ObjConn2, ObjRS2
						    %>  
						    <!--DETAIL INFORMATION END-->	    									
						</tbody>								
					</table>
					<!--Work List End-->
					<INPUT type=hidden name="DataCnt2" value="<%=x%>">
				<% Else %>
				    
					<table border="1" cellPadding="2" cellSpacing="0" id="NODATA">						
					  <TR class=bgw><TD nowrap>本船・次航の登録がありません</TD></TR>
					</table>
					
				<% End If %>		
			</div>
		</td>
		<!--Place Here End-->
	</tr>
	<tr><td>&nbsp;</td></tr>
	<tr>
	  <td align="center"><input type="button" name="close" onclick="window.close();" value="閉じる"></td>
	</tr>
  </table>
  <input type="hidden" name="pagenum2"   value=""/>
  <input type="hidden" name="openerForm"  value="<%=openerForm%>"/>
  <input type="hidden" name="openerFieldVslNm" value="<%=openerFieldVslNm%>"/>
  <input type="hidden" name="openerFieldVslCd" value="<%=openerFieldVslCd%>"/>
  <input type="hidden" name="openerFieldVoy"   value="<%=openerFieldVoy%>"/>
  <input type="hidden" name="openerFieldDspVoy"   value="<%=openerFieldDspVoy%>"/>
  <input type="hidden" name="DispKbn"   value="<%=DispKbn%>"/>
  </div>
</td></tr>  
</table>
</form>
</BODY>
</HTML>
