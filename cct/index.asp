<%@Language="VBScript" %>

<!--#include file="common.inc"-->

<%
    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' �A�o�R���e�i�Ɖ�
    WriteLog fs, "5001","�d�o�n�d���n���Ɖ�(�Ԙp)","00", ","
%>

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
          <td height="25" bgcolor="000099" align="left"><span class="style1">&nbsp;&nbsp;�Ԙp�i�`�[�����j�̉�ʐ���</span></td>
          <td bgcolor="000099" align="right"><span class="style2">Hits ver2</span>&nbsp;&nbsp;&nbsp;</td>
        </tr>
      </table>
        <table width="530" border=0>
          <tr>
            <td align=left><table cellpadding="0" cellspacing="0">
                <tr>
                  <td width="30" align="right"><img src="../gif/b-help.gif" width="20" height="20" hspace="4" vspace="4"></td>
                  <td align="left" nowrap><b>�R���e�i�ԍ����w�肷���ʂ��\�����ꂽ�ꍇ</b></td>
                </tr>
              </table>
              </td>
          </tr>
          <tr>
            <td align=center><img src="1.jpg" width="400" height="280" vspace="6"><br>
              <img src="2.gif" width="430" height="106" vspace="9" border="1">              <br>
              <table border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td align="left"><ol>
                    <li>�������Ɖ�����R���e�i�ԍ��i�󔒂̏ꍇ�͔��p�p�����œ��͂��܂��j</li>
                    <li>Excell�̃f�[�^�ɏo�͂������ꍇ�́uEXCELL�v��I�т܂��B</li>
                    <li>����������ďƉ�����s���܂��B</li>
                  </ol></td>
                </tr>
            </table></td></tr>
          <tr>
            <td align=left>&nbsp;</td>
          </tr>
          <tr>
            <td align=left><table cellpadding="0" cellspacing="0">
              <tr>
                <td width="30" align="right"><img src="../gif/b-help.gif" width="20" height="20" hspace="4" vspace="4"></td>
                <td align="left" nowrap><b>�Ɖ�ʉ�ʂ̐���</b></td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td align=center><img src="3.jpg" width="480" height="251" vspace="6"><br>
              <table border="1" cellspacing="1" cellpadding="2">
                <tr align="left">
                  <td bgcolor="#FFCC33">handle type</td>
                  <td bgcolor="#FFFFFF">�Ԙp����̔��o�iout�j�A�����iin�j�A�܂��́ASPC�ispecial order�j</td>
                </tr>
                <tr align="left">
                  <td bgcolor="#FFCC33">handle time</td>
                  <td bgcolor="#FFFFFF">��Ƃ��s�Ȃ�ꂽ����</td>
                </tr>
                <tr align="left">
                  <td bgcolor="#FFCC33">carrier type</td>
                  <td bgcolor="#FFFFFF">��ƑΏۂ̗A���@��B�{�D�iVS)�A�g���b�N�iTR�j�A�͂����iBG)</td>
                </tr>
                <tr align="left">
                  <td nowrap bgcolor="#FFCC33">carrier code</td>
                  <td bgcolor="#FFFFFF">�A���@��R�[�h</td>
                </tr>
                <tr align="left">
                  <td bgcolor="#FFCC33">in voyage</td>
                  <td bgcolor="#FFFFFF">�A�����q</td>
                </tr>
                <tr align="left">
                  <td bgcolor="#FFCC33">out voyage</td>
                  <td bgcolor="#FFFFFF">�A�o���q</td>
                </tr>
                <tr align="left">
                  <td bgcolor="#FFCC33">line</td>
                  <td bgcolor="#FFFFFF">�D��</td>
                </tr>
                <tr align="left">
                  <td bgcolor="#FFCC33">type</td>
                  <td bgcolor="#FFFFFF"><a href="#" onClick="javascript:winOpen('win2','./carrier_type.html',560,500) ">�R���e�i�^�C�v�ꗗ</a></td>
                </tr>
                <tr align="left">
                  <td bgcolor="#FFCC33">length</td>
                  <td bgcolor="#FFFFFF">�R���e�i�̒����i40.00(=40ft)�A20.00(=20ft)�Ȃǁj</td>
                </tr>
                <tr align="left">
                  <td bgcolor="#FFCC33">height</td>
                  <td bgcolor="#FFFFFF">�R���e�i�̍���(9.60(=96)�A8.60(=86)�Ȃǁj</td>
                </tr>
                <tr align="left">
                  <td bgcolor="#FFCC33">IsoCode</td>
                  <td bgcolor="#FFFFFF">�R���e�i�̃^�C�v������ISO�R�[�h</td>
                </tr>
                <tr align="left">
                  <td bgcolor="#FFCC33">E/F</td>
                  <td bgcolor="#FFFFFF">��R���e�i�iE)�^������R���e�i�iF)</td>
                </tr>
                <tr align="left">
                  <td bgcolor="#FFCC33">weight</td>
                  <td bgcolor="#FFFFFF">�R���e�i�d��(�O���X)</td>
                </tr>
                <tr align="left">
                  <td bgcolor="#FFCC33">ship seal</td>
                  <td bgcolor="#FFFFFF">-</td>
                </tr>
                <tr align="left">
                  <td bgcolor="#FFCC33">Cusrms seal</td>
                  <td bgcolor="#FFFFFF">-</td>
                </tr>
                <tr align="left">
                  <td bgcolor="#FFCC33">OOG</td>
                  <td bgcolor="#FFFFFF">�I�[�o�[�f�B�����W����</td>
                </tr>
              </table>
            <br></td></tr>
          <tr>
            <td align=center>                  <form>
                    <input type="button" value="����" onClick="JavaScript:window.close()">
            </form></td>
          </tr>
        </table>
        <br>
    </td>
  </tr>
</table>

</body>
</html>
