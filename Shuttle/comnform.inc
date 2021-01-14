<FORM ACTION="SYWB002.asp" METHOD="post" NAME="INQUIRY">
<table border="0" bgcolor="#e8ffe8" style="HEIGHT: 62px; WIDTH: 512px">
	<tr>
		<td>
			<p>　　　コンテナ番号</p>
		</td>
		<td>
			<INPUT NAME="container" SIZE="21"  MAXLENGTH="12" value="<%=contval%>">
		</td>
	</tr>
	<tr>
		<td>
			<p>　　　Ｂ／Ｌ番号</p>
		</td>
		<td>
			<INPUT NAME="blnumber" SIZE="30"  MAXLENGTH="20" value="<%=blval%>">
		</td>
		<td>
			<INPUT TYPE="submit" VALUE="  <%=tsubmit%>  " onClick="ClickInquiry()">
		</td>
	</tr>
</table>   
</FORM>
