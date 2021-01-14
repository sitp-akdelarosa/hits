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
  Dim s_PortCode
  Dim s_PortName
  Dim ObjRS2,ObjConn2
  Dim StrSQL
  Dim Num2
  Dim abspage,pagecnt
  Dim x,i
  Dim openerForm
  Dim openerFieldNm
  Dim openerFieldCd

  const gcPage = 10
  const rownum = 20

  openerForm = Request.QueryString("tgt")
  openerFieldCd = Request.QueryString("fldcode")
  openerFieldNm = Request.QueryString("fldname")

  if gfTrim(openerForm) = "" then
  	openerForm = gfTrim(Request.Form("openerForm"))
  end if
  if gfTrim(openerFieldCd) = "" then
  	openerFieldCd = gfTrim(Request.Form("openerFieldCd"))
  end if
  if gfTrim(openerFieldNm) = "" then
  	openerFieldCd = gfTrim(Request.Form("openerFieldNm"))
  end if

  s_PortCode = gfTrim(Request.Form("S_PortCode"))
  s_PortName = gfTrim(Request.Form("S_PortName"))
  
'セッションの有効性をチェック
  CheckLoginH
  
  '船名、次航リスト取得
  ConnDBH ObjConn2, ObjRS2

  StrSQL = "SELECT PortCode, FullName From mPort "
  StrSQL = StrSQL & " where 1 = 1 "
  if s_PortCode <> "" then
  	  StrSQL = StrSQL & "   AND PortCode like '" & gfSQLEncode(s_PortCode) & "%'"
  end if
  if s_PortName <> "" then
  	  StrSQL = StrSQL & "   AND FullName like '%" & gfSQLEncode(s_PortName) & "%'"
  end if
  StrSQL = StrSQL & " ORDER BY PortCode "
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
	jampErrerP "2","b301","01","港検索","102","SQL:<BR>" & StrSQL & err.description & Err.number
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
  var PortCodes, PortNames;
  var SelectVal;
  
  PortCodes = document.getElementsByName("PortCode");
  PortNames = document.getElementsByName("PortName");
  index = index - 1;
<%
  if openerFieldCd <> "" then
  	  Response.Write "opener." & openerForm & ".elements[""" & openerFieldCd & """].value=PortCodes[index].value;"
  end if
  
  if openerFieldNm <> "" then
  	  Response.Write "opener." & openerForm & ".elements[""" & openerFieldNm & """].value=PortNames[index].value;"
  end if
%>
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
     <td width="60">港ｺｰﾄﾞ</td>
     <td width="150">
        <input type=text name="S_PortCode" value="<%=gfHTMLEncode(s_PortCode)%>" style="ime-mode:none;"/>
     </td>
     <td rowspan=2 width><input type=button name="search" onclick="fSearch();" value=" 検　索 "/>
  </tr>
  <tr>
     <td>港名</td>
     <td>
        <input type=text name="S_PortName" value="<%=gfHTMLEncode(s_PortName)%>" />
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
								<th class="hlist" align="center" nowrap>港ｺｰﾄﾞ</th>
								<th class="hlist" align="center" nowrap>港名</th>
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
                              <%=gfHTMLEncode(ObjRS2("PortCode"))%>
							  <input type="hidden" name="PortCode" value="<%=gfHTMLEncode(ObjRS2("PortCode"))%>"><BR>
                            </td>
							<td align="left" valign="middle" width="200" nowrap>
                              <%=gfHTMLEncode(ObjRS2("FullName"))%>
							  <input type="hidden" name="PortName" value="<%=gfHTMLEncode(ObjRS2("FullName"))%>"><BR>
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
					  <TR class=bgw><TD nowrap>港の登録がありません</TD></TR>
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
  <input type="hidden" name="openerFieldNm" value="<%=openerFieldNm%>"/>
  <input type="hidden" name="openerFieldCd" value="<%=openerFieldCd%>"/>
  </div>
</td></tr>  
</table>
</form>
</BODY>
</HTML>
