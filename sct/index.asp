<%@Language="VBScript" %>

<!--#include file="common.inc"-->

<%
    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' �A�o�R���e�i�Ɖ�
    WriteLog fs, "5001","�d�o�n�d���n���Ɖ�(�֌�)","10", ","
%>

<!-- saved from url=(0022)http://internet.e-mail -->
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT language="javascript" type="text/javascript" src="../index.js"></SCRIPT>

<script language="javascript">
function OpenWin(){
	window.moveTo(5, 5);
}
</SCRIPT>

<style type="text/css">
<!--
.style1 {
	color: #FFFFFF;
	font-weight: bold;
}
.style2 {color: #FFFFFF}
-->
</style>
</head>
<body bgcolor="E6E8FF" text="#000000" link="#3300FF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="OpenWin();">

<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="25" bgcolor="000099" align="left"><span class="style1">&nbsp;&nbsp;�֌��i�V�F�R�E�j�̉�ʐ���</span></td>
          <td bgcolor="000099" align="right"><span class="style2">Hits ver2</span>&nbsp;&nbsp;&nbsp;</td>
        </tr>
      </table>
        <table width="530" border=0>
          <tr>
            <td align=left><table cellpadding="0" cellspacing="0">
                <tr>
                  <td width="30" align="right"><img src="../gif/b-help.gif" width="20" height="20" hspace="4" vspace="4"></td>
                  <td align="left" nowrap><b>�R���e�i��{���̉��</b></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td align=center><img src="1.jpg" width="400" height="298" vspace="9">
            <table border="0" cellspacing="0" cellpadding="3">
                <tr align="left">
                  <td bgcolor="065FBD"><span class="style2">HISTORY INQUIRY </span></td>
                  <td>�{�^�����N���b�N����Ɨ�����ʂ��\������܂��B</td>
                </tr>
            </table></td></tr>
          <tr>
            <td align=left>&nbsp;</td>
          </tr>
          <tr>
            <td align=center>
			<table cellspacing="1" cellpadding="2">
                <tr>
                  <td><table width="500" border="1" cellpadding="2" cellspacing="1">
                      <tr align="left">
                        <td bgcolor="#FFCC33">Container</td>
                        <td bgcolor="#FFFFFF">�R���e�i�ԍ�</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#FFCC33">Line_ID</td>
                        <td bgcolor="#FFFFFF">�D�ЃR�[�h</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#FFCC33">ISO_Code</td>
                        <td bgcolor="#FFFFFF"><a href="javascript:document.forms['queryForm'].submit();">�R���e�i�̃^�C�v������ISO�R�[�h</a></td>
                      </tr>
                      <tr align="left">
                        <td nowrap bgcolor="#FFCC33">Sz/Tp/Ht</td>
                        <td bgcolor="#FFFFFF">�T�C�Y�^�^�C�v�^����<BR>
                        �i�⑫�j�^�C�vGP�Ƃ́H�E�E�Egeneral purpose without ven�i���ʃR���e�i�j</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#FFCC33">Damage</td>
                        <td bgcolor="#FFFFFF">�_���[�W���</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#FFCC33">Category</td>
                        <td bgcolor="#FFFFFF">�A���iI�j�^�A�o�iO)</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#FFCC33">Status</td>
                        <td bgcolor="#FFFFFF">��R���e�i�iE)�^������R���e�i�iF)</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#FFCC33">Location</td>
                        <td bgcolor="#FFFFFF">C�F�Q�[�gOUT�A�s�FTRUCK�̏�A�u�F�{�D�̏�AY�F���[�h��<BR>
                        �i��jC OUT OUT�@�E�E�ECOMMUNITY�@OUT�̗��ŃQ�[�g����o���Ӗ�</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#FFCC33">Load_Port</td>
                        <td bgcolor="#FFFFFF">�d�o�`</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#FFCC33">Discharge_Port</td>
                        <td bgcolor="#FFFFFF">�d���`</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#FFCC33">Gross_Weight</td>
                        <td bgcolor="#FFFFFF">���d��</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#FFCC33">Seal_Nbr1</td>
                        <td bgcolor="#FFFFFF">�|</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#FFCC33">Seal_Nbr2</td>
                        <td bgcolor="#FFFFFF">�|</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#FFCC33">IMCO Class</td>
                        <td bgcolor="#FFFFFF">IMCO�N���X�i�댯�i�����j</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#FFCC33">UNDG</td>
                        <td bgcolor="#FFFFFF">UNDG�R�[�h</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#FFCC33">Temperature</td>
                        <td bgcolor="#FFFFFF">���x</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#FFCC33">Over_Dimension</td>
                        <td bgcolor="#FFFFFF">�K�i�O�T�C�Y���</td>
                      </tr>
                    </table>
  &nbsp;<br>
                    Arrival/Departure Schedule�@�R���e�i�̓����A�o���X�P�W���[�� <br>
                    <table border="1" cellspacing="1" cellpadding="2">
                      <tr align="left">
                        <td bgcolor="#589FE5">Location</td>
                        <td bgcolor="#FFFFFF">���[�h���R���e�i�̈ʒu</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#589FE5">Position</td>
                        <td bgcolor="#FFFFFF">�g���[���A���邢�́A�D��</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#589FE5">Voyage/Train</td>
                        <td bgcolor="#FFFFFF">�{�D���q</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#589FE5">Time</td>
                        <td bgcolor="#FFFFFF">����</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#589FE5">intended_Arrival</td>
                        <td bgcolor="#FFFFFF">�����i�\��j</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#589FE5">intended_Departure</td>
                        <td bgcolor="#FFFFFF">�o���i�\��j</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#589FE5">Actual_Arrival</td>
                        <td bgcolor="#FFFFFF">�����i���сj</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#589FE5">Actual_Departure</td>
                        <td bgcolor="#FFFFFF">�o���i���сj</td>
                      </tr>
                  </table></td>
                </tr>
              </table>
                <br>
&nbsp;&nbsp; </td>
          </tr>
          <tr>
            <td align=left><table cellpadding="0" cellspacing="0">
                <tr>
                  <td width="30" align="right"><img src="../gif/b-help.gif" width="20" height="20" hspace="4" vspace="4"></td>
                  <td align="left" nowrap><b>�R���e�i�������̉��</b></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td align=center><img src="2.jpg" width="380" height="538" vspace="9">
              <table cellspacing="1" cellpadding="2">
                  <tr>
                    <td><table cellspacing="2" cellpadding="2">
                      <tr align="left">
                        <td nowrap bgcolor="#065FBD"><span class="style2">Export EXCEL </span></td>
                        <td>���̌��ʂ�EXCEL�f�[�^�ɏo�͂��܂�</td>
                      </tr>
                        <tr align="left">
                          <td bgcolor="#065FBD"><span class="style2">Export to CSV </span></td>
                        <td>���̌��ʂ�CSV�f�[�^�ɏo�͂��܂�</td>
                        </tr>
                        <tr align="left">
                          <td bgcolor="#065FBD"><span class="style2">Export to XML </span></td>
                        <td>���̌��ʂ�XML�f�[�^�ɏo�͂��܂�</td>
                        </tr>
                        <tr align="left">
                          <td height="21" bgcolor="#065FBD"><span class="style2">Print</span></td>
                        <td>���̌��ʂ�������܂�</td>
                        </tr>
                      </table>
                      <br>
                        <table border="1" cellspacing="1" cellpadding="2">
                          <tr align="left">
                            <td bgcolor="#FFCC33">Line</td>
                            <td bgcolor="#FFFFFF">�D��</td>
                          </tr>
                          <tr align="left">
                            <td bgcolor="#FFCC33">OP_Time</td>
                            <td bgcolor="#FFFFFF">��Ǝ��{����</td>
                          </tr>
                          <tr align="left">
                            <td bgcolor="#FFCC33">Operation</td>
                            <td width="300" bgcolor="#FFFFFF"><a href="#" onClick="javascript:winOpen('win2','./operation_list.html',500,480) ">��Ɠ��e�ꗗ</a></td>
                          </tr>
                          <tr align="left">
                            <td bgcolor="#FFCC33">Move_From</td>
                            <td bgcolor="#FFFFFF">�ړ���</td>
                          </tr>
                          <tr align="left">
                            <td bgcolor="#FFCC33">Move_To</td>
                            <td bgcolor="#FFFFFF">�ړ���</td>
                          </tr>
                          <tr align="left">
                            <td bgcolor="#FFCC33">Notes</td>
                            <td bgcolor="#FFFFFF">���l</td>
                          </tr>
                      </table></td>
                  </tr>
            </table></td></tr>
          <tr>
            <td align=center>&nbsp;</td>
          </tr>
          <tr>
            <td align=center><form>
                <input type="button" value="����" onClick="JavaScript:window.close()">
            </form></td>
          </tr>
        </table>
        <form name="queryForm" method="post" action="http://oi.sctcn.com/Default.aspx?Action=Nav&amp;Content=ISO%20CODE%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20&amp;sm=ISO%20CODE" target="_blank">
<input type="hidden" name="data" value="NA">
<input type="hidden" name="OrgMenu" value="">
<input type="hidden" name="targetPage" value="Report_Regular">
<input type="hidden" name="nav" value="ISO CODE                                ">
	</form>
        <br>
    </td>
  </tr>
</table>
</body>
</html>
