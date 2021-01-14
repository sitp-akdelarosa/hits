<%@Language="VBScript" %>

<!--#include file="../Common.inc"-->

<html>
<head>
<title>Import Status Delivery Request Help</title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript"><!--
function LinkSelect(form, sel)
{
	adrs = sel.options[sel.selectedIndex].value;
	if (adrs != "-" ) parent.location.href = adrs;
}
function OpenCodeWin()
{
	var CodeWin;
	CodeWin = window.open("../codelist.asp?user=<%=Session.Contents("userid")%>","codelist","scrollbars=yes,resizable=yes,width=300,height=330");
	CodeWin.focus();
}
// -->
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="image/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------ここから画面--------------------------->
<table border="0" cellspacing="0" cellpadding="0" width="100%" height=100%>
<tr>
	<td valign=top>
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td rowspan=2><img src="image/sst_help.gif" width="506" height="73"></td>
			<td height="25" bgcolor="000099" align="right"><img src="image/logo_hits_ver2.gif" width="300" height="25"></td>
		</tr>
		<tr>
			<td align="right" width="100%" height="48"> 
<%
'Y.TAKAKUWA Del-S 2015-03-04
'call	DisplayCodeListButton
'Y.TAKAKUWA Del-E 2015-03-04
%>
			</td>
		</tr>
		</table>
		<center>
		<BR><BR><BR>
		<table border="0">
			<tr>
				<td align="center"> 
					<table border="0" cellspacing="2" cellpadding="3">
						<tr> 
							<td colspan="2" bgcolor="#99ccFF"><b><font color="#000000">◆ Intial setting of Import Status Email Delivery Request</font></b></td>
						</tr>
						<tr> 
							<td width="15">　</td>
							<td width="575">Please click "INITIAL REQUEST" from "Switch Display" of the left side of the screen and 
							                input Container No. or BL No. and click "Register".<br>
											The container which has already picked up from the terminal for more than 11 days before cannot be registered.<br>
											</td>
						</tr>
						<tr>
							<td colspan="2">　</td>
						</tr>
						<tr>
							<td colspan="2" bgcolor="#99ccFF"><b><font color="#000000">◆ Display of Import Status Email Delivery Request list</font></b></td>
						</tr>
						<tr>
							<td width="15">　</td>
							<td width="575">Please click "Email Delivery  Request List" for checking  the list of the delivery requested containers. 
											However, the container numbers cannot be displayed on the list which have already picked up more than 11 days before.</td>
						</tr>
						<tr>
							<td colspan="2">　</td>
						</tr>
						<tr> 
							<td colspan="2" bgcolor="#99ccFF"><b><font color="#000000">◆ Deletion of Import Status Email Delivery</font></b></td>
						</tr>
						<tr> 
							<td width="15">　</td>
							<td width="575">Please click the "No." of the container from the list and click "delete".</td>
						</tr>
						<tr>
							<td colspan="2">　</td>
						</tr>
						<tr> 
							<td colspan="2" bgcolor="#99ccFF"><b><font color="#000000">◆ Real -time Delivery  </font></b></td>
						</tr>
						<tr> 
							<td width="15">　</td>
							<td width="575">Please click "Real-time Delivery"  after inputting Container No. or BL No. to obtain the current status of the container.</td>
						</tr>
						<tr>
							<td colspan="2">　</td>
						</tr>
						<tr> 
							<td colspan="2" bgcolor="#99ccFF"><b><font color="#000000">◆ Items Setting of Import Container Status Email Delivery</font></b></td>
						</tr>
						<tr> 
							<td width="15">　</td>
							<td width="575"> Please click "Real-time Delivery"  after inputting Container No. or BL No.to obtain the current status of the container.
</td>
						</tr>
						<tr>
							<td colspan="2">　</td>
						</tr>
						<tr> 
							<td colspan="2" bgcolor="#99ccFF"><b><font color="#000000">◆ Search</font></b></td>
						</tr>
						<tr> 
							<td width="15">　</td>
							<td width="575">It is possible to seach by Container No. or BL No.  
Suffix search is also available. For example, after inputting "555" and click "Search", "CONT0000555" becomes the object of extraction .</td>
						</tr>
						<tr>
							<td colspan="2">　</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
		<table border="0">
			<form>
			<tr><td>　</td></tr>
			<tr><input type="button" value="CLOSE" onClick="window.close()"></td></tr>
			</form>
		</table>
		</center>
	</td>
</tr>
</table>
<!-------------画面終わり--------------------------->
</body>
</html>
