<%@Language="VBScript" %>

<!--#include file="Common.inc"-->
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript"><!--
function FancBack()
{
        window.history.back();
}

function LinkSelect(form, sel)
{
        adrs = sel.options[sel.selectedIndex].value;
        if (adrs != "-" ) parent.location.href = adrs;
}
// Added and Commented by seiko-denki 2003.07.18
function OpenCodeWin()
{
  var CodeWin;
  CodeWin = window.open("codelist.asp?user=<%=Session.Contents("userid")%>","codelist","scrollbars=yes,resizable=yes,width=300,height=330");
  CodeWin.focus();
}
// End of Addition by seiko-denki 2003.07.18
// -->
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------����������--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
  <td valign=top>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
          <td rowspan=2><img src="gif/helpt.gif" width="506" height="73"></td>
	      <td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
  </tr>
  <tr>
	<td align="right" width="100%" height="48"> 
<%
' Added and Commented by seiko-denki 2003.07.18
	DisplayCodeListButton
'    DispMenu
'	Dim strRoute
'	strRoute = Session.Contents("route")
' End of Addition by seiko-denki 2003.07.18
%>
          </td>
        </tr>
      </table>
      <center>
<!-- commented by seiko-denki 2003.07.18
		<table width=95% cellpadding="0" cellspacing="0" border="0">
		  <tr>
			<td align="right">
			  <font color="#333333" size="-1">
				<%=strRoute%> &gt; �w���v
			  </font>
			</td>
		  </tr>
		</table>
end of comment by seiko-denki 2003.07.18 -->
		<BR>
		<BR>
		<BR>
        <table>
          <tr>
            <td> 
              <table>
                <tr> 
                  <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                  <td nowrap><b>�{�D���Ó���</b></td>
                  <td><img src="gif/hr.gif" width="540"></td>
                </tr>
              </table>
<center>
              <table border="0" cellspacing="2" cellpadding="3">
                <tr> 
                  <td colspan="2" bgcolor="#99ccFF"><b>���DCSV�t�@�C���]���Ƃ́H</b></td>
                </tr>
                <tr> 
                  <td width="15"> </td>
                  <td width="575">���͂�������񂪑����ꍇ�A���x�����͂���͖̂ʓ|�ł��B<br>
                    �����ŁA�{�V�X�e���ł͏��𗅗񂵂��t�@�C�������A���̃t�@�C����]�����邱�Ƃł܂Ƃ߂ē��͂���@�\��p�ӂ��Ă��܂��B<br>
                    �{�V�X�e���ɓ]���ł���t�@�C���̌`���́uCSV�t�@�C���v�Ƃ������ʓI�Ȃ��̂ł��B<br>
                    ���́uCSV�t�@�C���v���쐬���]�����s���菇���ȉ��ɐ������܂��B<br>
                    &nbsp; </td>
                </tr>
                <tr> 
                  <td colspan="2" bgcolor="#99ccFF"><b>���D�K�v�ȃA�v���P�[�V����</b></td>
                </tr>
                <tr> 
                  <td width="15"> </td>
                  <td width="575"> 
                    <dl> 
                      <dt>CSV�t�@�C���̍쐬��Windows�t���̃������ŉ\�ł��B���邢�́AEXCEL�ō쐬����CSV�t�@�C���`���ŕۑ����邱�Ƃ��\�ł��B<br>
                    </dl>
				   
				   </td>
                </tr>
                <tr> 
                  <td colspan="2" bgcolor="#99ccFF"><b><font color="#000000">���DCSV�t�@�C���̍쐬</font></b></td>
                </tr>
                <tr> 
                  <td width="15"> </td>
                    <td width="575"> 
                      <dl> 
                        <dt>�O�q�̃A�v���P�[�V�������g���āA�R�[���T�C���AVoyage No�D�A�`���E�E�E�̏��ɂЂƂЂƂ̒l���J���}�u,�v�ŋ�؂�Ȃ���P�s�ɂP�Z�b�g�̏����L�q���܂��B<br>

 <dt>&nbsp;&nbsp;<font color=#ff0000>�y���Ӂz</font>
 <dd>�����{�D�i����̃R�[���T�C����Voyage No.�j�̍s�͑����ċL�q���Ă��������B
<dd>�����{�D�̓���`�ɑ΂���f�[�^�͂P�s�ŋL�q���Ă��������B�i�Ⴆ�΁A����Long Schedule�ƒ��ݗ\�莞������͂���ꍇ�͂Q�s
�ɕ������P�s�ŋL�q���Ă��������j<p>
                          <table border="1" cellspacing="1" cellpadding="5" width=500>
                            <tr> 
                              <td bgcolor="#FFFFFF" nowrap><font size="1">
										A1284, B3567, JPTYO, 2002/3/12/14/50, 2002/3/12/15/00, 2002/3/12/16/05, 2002/3/12, 2002/3/12
									<br>
										F8976, D7909, JPTYO, 2002/3/18/03/00, 2002/3/18/03/08, 2002/3/18/4/30, 2002/3/18, 2002/3/18
									</font>
								</td>
                            </tr>
                          </table>
                          <br>
                        <dt>1�s���̍��ڂ̏ڍ׎d�l<BR>
                        <dd>
                          <table width="100" border="1" cellspacing="0" cellpadding="2" bgcolor="#FFFFFF">
                            <tr bgcolor="#99aaFF" align="center"> 
                              <td nowrap><b><font color="#333333">����</font></b></td>
                              <td nowrap><b><font color="#333333">��</font></b></td>
                              <td nowrap><b><font color="#333333">���͎d�l</font></b></td>
                            </tr>
                            <tr valign="top"> 
                              <td nowrap>�R�[���T�C��</td>
                              <td nowrap>A1284</td>
                              <td nowrap>���p�啶���p����7���ȓ�</td>
                            </tr>
                            <tr valign="top"> 
                              <td nowrap>Voyage No.</td>
                              <td nowrap>B3567</td>
                              <td nowrap>���p�啶���p����12���ȓ�<br>
                                ����L���܂ށi'-'�A'/'�Ȃǁj</td>
                            </tr>
                            <tr valign="top"> 
                              <td nowrap>�`�� </td>
                              <td nowrap>JPTYO</td>
                              <td nowrap>UNLO�R�[�h�i���p�啶���p�������T���j</td>
                            </tr>
                            <tr valign="top"> 
                              <td nowrap>���ݗ\�莞��<br>
                                �i�N���������j</td>
                              <td nowrap>2002/3/12/14/5 </td>
                              <td nowrap>�E�N�F����4��<br>
                                �E���̑��F����2��('01'��'1'�̗����̕\���ɑΉ�)<br>
                                �E�ȏ�𔼊p�X���b�V���u/�v�ŋ�؂�B<br>
                                �E�l�������ꍇ�̓X���b�V���������c��(�u//�v)</td>
                            </tr>
                            <tr valign="top"> 
                              <td nowrap>���݊�������</td>
                              <td nowrap>�i���l�̌`���j</td>
                              <td nowrap>����</td>
                            </tr>
                            <tr valign="top"> 
                              <td nowrap>���݊������� </td>
                              <td nowrap>�i���l�̌`���j</td>
                              <td nowrap>����</td>
                            </tr>
                            <tr valign="top"> 
                              <td nowrap>����Long Schedule<BR>�i�N�����j</td>
                              <td nowrap>2002/3/12</td>
                              <td nowrap>����</td>
                            </tr>
                            <tr valign="top"> 
                              <td nowrap>����Long Schedule<BR>�i�N�����j</td>
                              <td nowrap>2002/3/12</td>
                              <td nowrap>����</td>
                            </tr>
                          </table>
                          <br>
                        <dt>�t�@�C�����͉��ł����܂��܂��񂪁A�g���q�͒ʏ�u.csv�v�Ƃ��܂��B�ۑ�������R�ł� 
                        <dd><font color="#FF0033">�y��z</font>C:\MyDocument���� abcdef.csv  �Ƃ����t�@�C�����ŕۑ����܂��B 
                      </dl>
                  </td>
                </tr>
                <tr> 
                    <td colspan="2" bgcolor="#99ccFF"><b>���DCSV�t�@�C���̓]��</b></td>
                </tr>
                <tr> 
                  <td width="15"> </td>
                    <td width="575"> 
                      <ul>
                        <li>��ʏ��CSV�t�@�C���]�����N���b�N����Ǝ��̂悤��CSV�t�@�C�����w�肷���ʂ��\������܂��B<br>
                          <table border="0" cellspacing="0" cellpadding="0">
                            <tr> 
                              <td valign="top" nowrap><font color="#FF0033">�y��z</font></td>
                              <td> 
                                <table border="1" cellspacing="1" cellpadding="5">
                                  <tr> 
                                    <td bgcolor="#FFFFFF" align="center"> 
                                      <form>
                                        <table border="1" cellspacing="0" cellpadding="2">
                                          <tr> 
                                            <td bgcolor="#000099" nowrap> <font color="#FFFFFF"><b>CSV�t�@�C����</b></font> 
                                            </td>
                                            <td nowrap> 
                                              <input name=csvfile size=30 accept="text/css">
                                            </td>
                                            <td nowrap> 
                                              <input type=button value="�Q��..." name="�{�^��">
                                            </td>
                                          </tr>
                                        </table>
                                        <input type=button value=" ��  �M " name="�{�^��">
                                      </form>
                                    </td>
                                  </tr>
                                </table>
                              </td>
                            </tr>
                          </table><br>
                        <li>�󗓂ɍ쐬����CSV�t�@�C���̃t���p�X���L�q���܂��B <br>
                          <font color="#FF0033">�y��z</font>�쐬��̏ꍇ�́uC:\MyDocument\abcdef.csv�v�ƋL�q���܂��B<br>
                        <li>����͂���̂��ʓ|�ȏꍇ�́A�m�Q��...�n�{�^���������ƃt�@�C����I�������ʂ��o�܂��̂ŁA�ۑ���̃t�H���_�ƃt�@�C�������ɑI�����Ă������ƂŃt�@�C�����������I�ɓ��͂���܂��B<br>
                        <li>�Ō�Ɂm���M�n�{�^���������܂��B<br>
                        <li>�������ʂ͒ʏ�̉�ʂŕ\������܂��B
                          <p> 
                          <table border="1" cellspacing="0" cellpadding="3">
                            <tr> 
                              <td bgcolor="#FF9933" nowrap valign="top">����</td>
                              <td bgcolor="#FFFFFF">�t�@�C���̍쐬���K���ǂ���ł��Ă��Ȃ��ƃV�X�e���͓��e��ǂݏo�����Ƃ��ł����G���[��\�����܂��B���̏ꍇ�́A�t�@�C���̓��e���������A�C��������ēx���M���s���Ă��������B</td>
                            </tr>
                          </table>
                      </ul>
                    </td>
                </tr>
              </table>

</center>

   </td>
   </tr>
  </table>
 <!---------->
  </center>
    </td>
 </tr>
 <tr>
    <td valign="bottom"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
	      <td valign="bottom" align="right"><a href="index.html"><img src="gif/b-home.gif" border="0" width="270" height="23" usemap="#map"></a></td>
        </tr>
        <tr>
          <td bgcolor="000099" height="10"><img src="gif/1.gif"></td>
  </tr>
</table>
 </td>
 </tr>
 </table>
<!-------------��ʏI���--------------------------->
<map name="map"> 
  <area shape="poly" coords="20,0,152,0,134,22,0,22" href="JavaScript:FancBack()">
  <area shape="poly" coords="154,0,136,22,284,22,284,0" href="http://www.hits-h.com/index.asp">
</map>
</body>
</html>