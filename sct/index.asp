<%@Language="VBScript" %>

<!--#include file="common.inc"-->

<%
    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' 輸出コンテナ照会
    WriteLog fs, "5001","仕出地仕向地情報照会(蛇口)","10", ","
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
          <td height="25" bgcolor="000099" align="left"><span class="style1">&nbsp;&nbsp;蛇口（シェコウ）の画面説明</span></td>
          <td bgcolor="000099" align="right"><span class="style2">Hits ver2</span>&nbsp;&nbsp;&nbsp;</td>
        </tr>
      </table>
        <table width="530" border=0>
          <tr>
            <td align=left><table cellpadding="0" cellspacing="0">
                <tr>
                  <td width="30" align="right"><img src="../gif/b-help.gif" width="20" height="20" hspace="4" vspace="4"></td>
                  <td align="left" nowrap><b>コンテナ基本情報の画面</b></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td align=center><img src="1.jpg" width="400" height="298" vspace="9">
            <table border="0" cellspacing="0" cellpadding="3">
                <tr align="left">
                  <td bgcolor="065FBD"><span class="style2">HISTORY INQUIRY </span></td>
                  <td>ボタンをクリックすると履歴画面が表示されます。</td>
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
                        <td bgcolor="#FFFFFF">コンテナ番号</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#FFCC33">Line_ID</td>
                        <td bgcolor="#FFFFFF">船社コード</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#FFCC33">ISO_Code</td>
                        <td bgcolor="#FFFFFF"><a href="javascript:document.forms['queryForm'].submit();">コンテナのタイプを示すISOコード</a></td>
                      </tr>
                      <tr align="left">
                        <td nowrap bgcolor="#FFCC33">Sz/Tp/Ht</td>
                        <td bgcolor="#FFFFFF">サイズ／タイプ／高さ<BR>
                        （補足）タイプGPとは？・・・general purpose without ven（普通コンテナ）</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#FFCC33">Damage</td>
                        <td bgcolor="#FFFFFF">ダメージ情報</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#FFCC33">Category</td>
                        <td bgcolor="#FFFFFF">輸入（I）／輸出（O)</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#FFCC33">Status</td>
                        <td bgcolor="#FFFFFF">空コンテナ（E)／実入りコンテナ（F)</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#FFCC33">Location</td>
                        <td bgcolor="#FFFFFF">C：ゲートOUT、Ｔ：TRUCKの上、Ｖ：本船の上、Y：ヤード内<BR>
                        （例）C OUT OUT　・・・COMMUNITY　OUTの略でゲートから出た意味</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#FFCC33">Load_Port</td>
                        <td bgcolor="#FFFFFF">仕出港</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#FFCC33">Discharge_Port</td>
                        <td bgcolor="#FFFFFF">仕向港</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#FFCC33">Gross_Weight</td>
                        <td bgcolor="#FFFFFF">総重量</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#FFCC33">Seal_Nbr1</td>
                        <td bgcolor="#FFFFFF">－</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#FFCC33">Seal_Nbr2</td>
                        <td bgcolor="#FFFFFF">－</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#FFCC33">IMCO Class</td>
                        <td bgcolor="#FFFFFF">IMCOクラス（危険品等級）</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#FFCC33">UNDG</td>
                        <td bgcolor="#FFFFFF">UNDGコード</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#FFCC33">Temperature</td>
                        <td bgcolor="#FFFFFF">温度</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#FFCC33">Over_Dimension</td>
                        <td bgcolor="#FFFFFF">規格外サイズ情報</td>
                      </tr>
                    </table>
  &nbsp;<br>
                    Arrival/Departure Schedule　コンテナの到着、出発スケジュール <br>
                    <table border="1" cellspacing="1" cellpadding="2">
                      <tr align="left">
                        <td bgcolor="#589FE5">Location</td>
                        <td bgcolor="#FFFFFF">ヤード内コンテナの位置</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#589FE5">Position</td>
                        <td bgcolor="#FFFFFF">トレーラ、あるいは、船名</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#589FE5">Voyage/Train</td>
                        <td bgcolor="#FFFFFF">本船次航</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#589FE5">Time</td>
                        <td bgcolor="#FFFFFF">日時</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#589FE5">intended_Arrival</td>
                        <td bgcolor="#FFFFFF">到着（予定）</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#589FE5">intended_Departure</td>
                        <td bgcolor="#FFFFFF">出発（予定）</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#589FE5">Actual_Arrival</td>
                        <td bgcolor="#FFFFFF">到着（実績）</td>
                      </tr>
                      <tr align="left">
                        <td bgcolor="#589FE5">Actual_Departure</td>
                        <td bgcolor="#FFFFFF">出発（実績）</td>
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
                  <td align="left" nowrap><b>コンテナ履歴情報の画面</b></td>
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
                        <td>下の結果をEXCELデータに出力します</td>
                      </tr>
                        <tr align="left">
                          <td bgcolor="#065FBD"><span class="style2">Export to CSV </span></td>
                        <td>下の結果をCSVデータに出力します</td>
                        </tr>
                        <tr align="left">
                          <td bgcolor="#065FBD"><span class="style2">Export to XML </span></td>
                        <td>下の結果をXMLデータに出力します</td>
                        </tr>
                        <tr align="left">
                          <td height="21" bgcolor="#065FBD"><span class="style2">Print</span></td>
                        <td>下の結果を印刷します</td>
                        </tr>
                      </table>
                      <br>
                        <table border="1" cellspacing="1" cellpadding="2">
                          <tr align="left">
                            <td bgcolor="#FFCC33">Line</td>
                            <td bgcolor="#FFFFFF">船社</td>
                          </tr>
                          <tr align="left">
                            <td bgcolor="#FFCC33">OP_Time</td>
                            <td bgcolor="#FFFFFF">作業実施日時</td>
                          </tr>
                          <tr align="left">
                            <td bgcolor="#FFCC33">Operation</td>
                            <td width="300" bgcolor="#FFFFFF"><a href="#" onClick="javascript:winOpen('win2','./operation_list.html',500,480) ">作業内容一覧</a></td>
                          </tr>
                          <tr align="left">
                            <td bgcolor="#FFCC33">Move_From</td>
                            <td bgcolor="#FFFFFF">移動元</td>
                          </tr>
                          <tr align="left">
                            <td bgcolor="#FFCC33">Move_To</td>
                            <td bgcolor="#FFFFFF">移動先</td>
                          </tr>
                          <tr align="left">
                            <td bgcolor="#FFCC33">Notes</td>
                            <td bgcolor="#FFFFFF">備考</td>
                          </tr>
                      </table></td>
                  </tr>
            </table></td></tr>
          <tr>
            <td align=center>&nbsp;</td>
          </tr>
          <tr>
            <td align=center><form>
                <input type="button" value="閉じる" onClick="JavaScript:window.close()">
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
