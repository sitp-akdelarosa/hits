<%@ LANGUAGE="VbScript" %>
<% Response.AddHeader "Pragma", "no-cache" %>
<HTML>
<HEAD>
<TITLE>ExcelCreator Ver3.0�@Web�A�v���P�[�V�����T���v��</TITLE>
</HEAD>
<BODY>
<%
	'---------------------------------------
	'�����̓��t���擾
	'---------------------------------------
	wUDate   = FormatDateTime (Date,vbGeneralDate)
%>

<CENTER>
<P align="center">���͂����f�[�^��PC�i�T�[�o���j��ɂ��锄��`�[�ЂȌ^Excel�t�@�C���ɐݒ肵�Ĕ���`�[���쐬����f���ł�<BR>
</P>

<FORM method="POST" action="UriageXls.asp">
<DIV align="center">
<TABLE border="0" cellpadding="0" cellspacing="0" width="85%" bordercolor="#C8E2FF" bgcolor="#F4FAFF" style="border-collapse: collapse">
		<TR>
				<TD width="100%">
				<P align="center"><B><FONT size="3" color="#000080"><BR>
				����`�[</FONT></B></P>
						<DIV align="center">
								<CENTER>
						<TABLE border="1" cellpadding="0" cellspacing="0" width="90%" bordercolor="#FFFFFF" bgcolor="#C8E2FF" style="border-collapse: collapse">
								<TR>
										<TD width="33%" height="40">
												<P align="center"><B><FONT size="2" color="#000080">�����</FONT></B></P>
										</TD>
										<TD width="33%" height="40">
												<P align="center"><B><FONT size="2" color="#000080">�`�[No</FONT></B></P>
										</TD>
										<TD width="33%" height="40">�@&nbsp;&nbsp;&nbsp;</TD>
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
				<TD width="10%" align="center" height="40"><B><FONT size="2" color="#000080">���Ӑ於</FONT></B></TD>
				<TD width="50%" height="40">
				<INPUT type="text" name="TName" size="60" value="�A�h�o���X�\�t�g�������"> �l</TD>
		</TR>
		<TR>
				<TD width="10%" align="center" height="40"><B><FONT size="2" color="#000080">���Z��</FONT></B></TD>
				<TD width="50%" height="40">
				<INPUT type="text" name="TAddress" size="60" value="�����s���c���蒬1-2-3456�����r���SF"></TD>
		</TR>
		<TR>
				<TD width="10%" align="center" height="40"><B><FONT size="2" color="#000080">�����O</FONT></B></TD>
				<TD width="50%" height="40">
				<INPUT type="text" name="Shimei" size="60" value="�R�c���Y"> �l</TD>
		</TR>
</TABLE>
						</CENTER>
				</DIV>
<BR>
						<DIV align="center">
								<CENTER>
						<TABLE border="1" cellpadding="0" cellspacing="0" width="90%" bgcolor="#C8E2FF" bordercolor="#FFFFFF" style="border-collapse: collapse">
								<TR>
										<TD width="25%" align="center" height="40"><B><FONT size="2" color="#000080">���i�R�[�h</FONT></B></TD>
										<TD width="25%" align="center" height="40"><B><FONT size="2" color="#000080">���i��</FONT></B></TD>
										<TD width="25%" align="center" height="40"><B><FONT size="2" color="#000080">����</FONT></B></TD>
										<TD width="25%" align="center" height="40"><B><FONT size="2" color="#000080">�P��</FONT></B></TD>
								</TR>
								<TR>
										<TD width="25%" height="40">
												<P align="center">
												<INPUT type="text" name="SCode" size="20" value="4993857123456"></P>
										</TD>
										<TD width="25%" height="40">
												<P align="center">
												<INPUT type="text" name="SName" size="60" value="�p�\�R���\�t�g�ExcelCreator Ver3.0�v"></P>
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
<INPUT type="submit" value="�`�[�쐬" name="Denpyo">&nbsp; <INPUT type="reset" value="���Z�b�g">
</FORM>
</CENTER>

</BODY>

</HTML>