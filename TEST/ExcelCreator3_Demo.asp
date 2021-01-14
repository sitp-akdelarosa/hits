<%@ LANGUAGE="VbScript" %>
<% Response.AddHeader "Pragma", "no-cache" %>
<HTML>
<HEAD>
<TITLE>ExcelCreator Ver3.0　Webアプリケーションサンプル</TITLE>
</HEAD>
<BODY>
<%
	'---------------------------------------
	'今日の日付を取得
	'---------------------------------------
	wUDate   = FormatDateTime (Date,vbGeneralDate)
%>

<CENTER>
<P align="center">入力したデータをPC（サーバ等）上にある売上伝票ひな型Excelファイルに設定して売上伝票を作成するデモです<BR>
</P>

<FORM method="POST" action="UriageXls.asp">
<DIV align="center">
<TABLE border="0" cellpadding="0" cellspacing="0" width="85%" bordercolor="#C8E2FF" bgcolor="#F4FAFF" style="border-collapse: collapse">
		<TR>
				<TD width="100%">
				<P align="center"><B><FONT size="3" color="#000080"><BR>
				売上伝票</FONT></B></P>
						<DIV align="center">
								<CENTER>
						<TABLE border="1" cellpadding="0" cellspacing="0" width="90%" bordercolor="#FFFFFF" bgcolor="#C8E2FF" style="border-collapse: collapse">
								<TR>
										<TD width="33%" height="40">
												<P align="center"><B><FONT size="2" color="#000080">売上日</FONT></B></P>
										</TD>
										<TD width="33%" height="40">
												<P align="center"><B><FONT size="2" color="#000080">伝票No</FONT></B></P>
										</TD>
										<TD width="33%" height="40">　&nbsp;&nbsp;&nbsp;</TD>
								</TR>
								<TR>
										<TD width="33%" height="40">
												<P align="center"><FONT color="#666666"><INPUT type="text" name="UDate" size="20" value=<%=wUDate%>></FONT></P>
										</TD>
										<TD width="33%" height="40">
												<P align="center"><FONT color="#666666">
												<INPUT type="text" name="UNo" size="20" value="ASW-2003123-A"></FONT></TD>
										<TD width="33%" height="40">
												<P align="center">&nbsp;&nbsp;&nbsp;
										</TD>
								</TR>
						</TABLE>
								</CENTER>
				</DIV>
<BR>
				<DIV align="center">
						<CENTER>
<TABLE border="1" cellpadding="0" cellspacing="0" width="90%" bordercolor="#FFFFFF" bgcolor="#C8E2FF" style="border-collapse: collapse">
		<TR>
				<TD width="10%" align="center" height="40"><B><FONT size="2" color="#000080">得意先名</FONT></B></TD>
				<TD width="50%" height="40">
				<INPUT type="text" name="TName" size="60" value="アドバンスソフト株式会社"> 様</TD>
		</TR>
		<TR>
				<TD width="10%" align="center" height="40"><B><FONT size="2" color="#000080">ご住所</FONT></B></TD>
				<TD width="50%" height="40">
				<INPUT type="text" name="TAddress" size="60" value="東京都千代田区大手町1-2-3456○○ビル４F"></TD>
		</TR>
		<TR>
				<TD width="10%" align="center" height="40"><B><FONT size="2" color="#000080">お名前</FONT></B></TD>
				<TD width="50%" height="40">
				<INPUT type="text" name="Shimei" size="60" value="山田太郎"> 様</TD>
		</TR>
</TABLE>
						</CENTER>
				</DIV>
<BR>
						<DIV align="center">
								<CENTER>
						<TABLE border="1" cellpadding="0" cellspacing="0" width="90%" bgcolor="#C8E2FF" bordercolor="#FFFFFF" style="border-collapse: collapse">
								<TR>
										<TD width="25%" align="center" height="40"><B><FONT size="2" color="#000080">商品コード</FONT></B></TD>
										<TD width="25%" align="center" height="40"><B><FONT size="2" color="#000080">商品名</FONT></B></TD>
										<TD width="25%" align="center" height="40"><B><FONT size="2" color="#000080">数量</FONT></B></TD>
										<TD width="25%" align="center" height="40"><B><FONT size="2" color="#000080">単価</FONT></B></TD>
								</TR>
								<TR>
										<TD width="25%" height="40">
												<P align="center">
												<INPUT type="text" name="SCode" size="20" value="4993857123456"></P>
										</TD>
										<TD width="25%" height="40">
												<P align="center">
												<INPUT type="text" name="SName" size="60" value="パソコンソフト｢ExcelCreator Ver3.0」"></P>
										</TD>
										<TD width="25%" height="40">
												<P align="center">
												<INPUT type="text" name="Suu" size="20" value="10" style="text-align: right"></P>
										</TD>
										<TD width="25%" height="40">
												<P align="center">
												<INPUT type="text" name="Tanka" size="20" value="34000" style="text-align: right"></TD>
								</TR>
						</TABLE>
								</CENTER>
				</DIV>
						<P align="right"><FONT size="2"><BR>   
						</FONT>
				</TD>
		</TR>
</TABLE>
</DIV>
<P align="center">
<INPUT type="submit" value="伝票作成" name="Denpyo">&nbsp; <INPUT type="reset" value="リセット">
</FORM>
</CENTER>

</BODY>

</HTML>