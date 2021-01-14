<%@Language="VBScript" %>

<!--#include file="common.inc"-->

<%
    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' 輸出コンテナ照会
    WriteLog fs, "5001","仕出地仕向地情報照会(赤湾)","00", ","
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
          <td height="25" bgcolor="000099" align="left"><span class="style1">&nbsp;&nbsp;赤湾（チーワン）の画面説明</span></td>
          <td bgcolor="000099" align="right"><span class="style2">Hits ver2</span>&nbsp;&nbsp;&nbsp;</td>
        </tr>
      </table>
        <table width="530" border=0>
          <tr>
            <td align=left><table cellpadding="0" cellspacing="0">
                <tr>
                  <td width="30" align="right"><img src="../gif/b-help.gif" width="20" height="20" hspace="4" vspace="4"></td>
                  <td align="left" nowrap><b>コンテナ番号を指定する画面が表示された場合</b></td>
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
                    <li>履歴を照会したいコンテナ番号（空白の場合は半角英数字で入力します）</li>
                    <li>Excellのデータに出力したい場合は「EXCELL」を選びます。</li>
                    <li>これを押して照会を実行します。</li>
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
                <td align="left" nowrap><b>照会結果画面の説明</b></td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td align=center><img src="3.jpg" width="480" height="251" vspace="6"><br>
              <table border="1" cellspacing="1" cellpadding="2">
                <tr align="left">
                  <td bgcolor="#FFCC33">handle type</td>
                  <td bgcolor="#FFFFFF">赤湾からの搬出（out）、搬入（in）、または、SPC（special order）</td>
                </tr>
                <tr align="left">
                  <td bgcolor="#FFCC33">handle time</td>
                  <td bgcolor="#FFFFFF">作業が行なわれた日時</td>
                </tr>
                <tr align="left">
                  <td bgcolor="#FFCC33">carrier type</td>
                  <td bgcolor="#FFFFFF">作業対象の輸送機器。本船（VS)、トラック（TR）、はしけ（BG)</td>
                </tr>
                <tr align="left">
                  <td nowrap bgcolor="#FFCC33">carrier code</td>
                  <td bgcolor="#FFFFFF">輸送機器コード</td>
                </tr>
                <tr align="left">
                  <td bgcolor="#FFCC33">in voyage</td>
                  <td bgcolor="#FFFFFF">輸入次航</td>
                </tr>
                <tr align="left">
                  <td bgcolor="#FFCC33">out voyage</td>
                  <td bgcolor="#FFFFFF">輸出次航</td>
                </tr>
                <tr align="left">
                  <td bgcolor="#FFCC33">line</td>
                  <td bgcolor="#FFFFFF">船社</td>
                </tr>
                <tr align="left">
                  <td bgcolor="#FFCC33">type</td>
                  <td bgcolor="#FFFFFF"><a href="#" onClick="javascript:winOpen('win2','./carrier_type.html',560,500) ">コンテナタイプ一覧</a></td>
                </tr>
                <tr align="left">
                  <td bgcolor="#FFCC33">length</td>
                  <td bgcolor="#FFFFFF">コンテナの長さ（40.00(=40ft)、20.00(=20ft)など）</td>
                </tr>
                <tr align="left">
                  <td bgcolor="#FFCC33">height</td>
                  <td bgcolor="#FFFFFF">コンテナの高さ(9.60(=96)、8.60(=86)など）</td>
                </tr>
                <tr align="left">
                  <td bgcolor="#FFCC33">IsoCode</td>
                  <td bgcolor="#FFFFFF">コンテナのタイプを示すISOコード</td>
                </tr>
                <tr align="left">
                  <td bgcolor="#FFCC33">E/F</td>
                  <td bgcolor="#FFFFFF">空コンテナ（E)／実入りコンテナ（F)</td>
                </tr>
                <tr align="left">
                  <td bgcolor="#FFCC33">weight</td>
                  <td bgcolor="#FFFFFF">コンテナ重量(グロス)</td>
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
                  <td bgcolor="#FFFFFF">オーバーディメンジョン</td>
                </tr>
              </table>
            <br></td></tr>
          <tr>
            <td align=center>                  <form>
                    <input type="button" value="閉じる" onClick="JavaScript:window.close()">
            </form></td>
          </tr>
        </table>
        <br>
    </td>
  </tr>
</table>

</body>
</html>
