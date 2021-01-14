<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
		' File System Object の生成
		Set fs=Server.CreateObject("Scripting.FileSystemobject")

		' 輸入コンテナ照会
		WriteLog fs, "2001","輸入コンテナ照会","00", ","
%>

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<link href="./index.css" rel="stylesheet" type="text/css">
<SCRIPT language="javascript" type="text/javascript" src="./index.js"></SCRIPT>
<SCRIPT Language="JavaScript">
<!--
function Submit(formName){
		result = true;
		errormessage = "";
		if(document.forms["HitsForm"].cntnrno.value==""){
				result = false;
				errormessage = "コンテナNOを入力して下さい";
		}
		if(!cntnrnoCheck(document.forms["HitsForm"].cntnrno.value)){
		}else{
				result = false;
				errormessage = "コンテナNOを複数入力することはできません";
		}

		if(result == true){
				if(formName == "queryForm"){
						document.forms[formName].data.value = document.forms["HitsForm"].cntnrno.value;
						document.forms[formName].submit();
				}else if(formName == "Form1"){
<!-- 2014/1/8 MOD-S MES aoyagi -->
<!--            document.forms[formName].cont_no.value = document.forms["HitsForm"].cntnrno.value; -->
						document.forms[formName].txtContainer_no.value = document.forms["HitsForm"].cntnrno.value;
<!--      2014/1/8 MOD-E MES aoyagi -->
						document.forms[formName].submit();


				}
		}else{
				 window.alert(errormessage);
		}
}

function cntnrnoCheck(str){
	return str.match(/^.*[,]{1}.*$/);
}

// 2009/10/29 add-s 中国外部データの検索機能追加
function SubmitGaibu(formName, usercode){
		result = true;
	// 2010/5/30 Del-S MES Aoyagi
		//errormessage = "";
	// 2010/5/30 Del-E MES Aoyagi
		if(document.forms["HitsForm"].cntnrno.value==""){
				result = false;
	// 2010/5/30 Add-S MES Aoyagi  2012/7/24 Add-QINGD by MES Suzaki
	 if(usercode == "HUANG"){
		location.href="./gaibuif/expdetail-HUANG.htm"
	　}else if(usercode == "NANSH"){
		location.href="./gaibuif/expdetail-NANSH.htm"
	　}else if(usercode == "TWTPE"){
		location.href="./gaibuif/expdetail-TWTPE.htm"
	　}else if(usercode == "THBKK"){
		location.href="./gaibuif/expdetail-THBKK.htm"
		　}else{
		location.href="./gaibuif/expdetail-QINGD.htm"

	　}
	// 2010/5/30 Add-E MES Aoyagi  
	// 2010/5/30 Del-S MES Aoyagi
				//errormessage = "コンテナNOを入力して下さい";
	// 2010/5/30 Del-E MES Aoyagi
		}
		if(!cntnrnoCheck(document.forms["HitsForm"].cntnrno.value)){
		}else{
				result = false;
				errormessage = "コンテナNOを複数入力することはできません";
		}

	// 中国の情報表示時はBL入力を拒否
		if(document.forms["HitsForm"].blno.value!="" && formName == "GaibuifForm"){
				result = false;
				errormessage = "黄埔/南沙/青島の情報は、BL NOで検索できません。";
		}

		if(result == true){
				if(formName == "GaibuifForm"){
						document.forms[formName].cntnrno.value = document.forms["HitsForm"].cntnrno.value;
						document.forms[formName].usercode.value = usercode;
						document.forms[formName].submit();
				}
		}else{
				 window.alert(errormessage);
		}
}
// 2009/10/29 add-e 中国外部データの検索機能追加
// -->
<%
		DispMenuJava
%>
</SCRIPT>

</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------ここから照会画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
	<tr>
		<td valign=top>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td rowspan=2><img src="gif/impentryt.gif" width="506" height="73"></td>
					<td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
				</tr>
				<tr>
					<td align="right" width="100%" height="48">
<%
' Added and Commented by seiko-denki 2003.07.07
	DisplayCodeListButton
'    DispMenu
'	Dim strScriptName,strRoute
'	strScriptName = Request.ServerVariables("SCRIPT_NAME")
'	strRoute = SetRoute(strScriptName)
'	Session.Contents("route") = strRoute
' End of Addition by seiko-denki 2003.07.07
%>
					</td>
				</tr>
			</table>
			<center>
<!-- commented by seiko-denki 2003.07.07
		<table width=95% cellpadding="0" cellspacing="0" border="0">
			<tr>
			<td align="right">
				<font color="#333333" size="-1">
				<%=strRoute%>
				</font>
			</td>
			</tr>
		</table>
end of comment by seiko-denki 2003.07.07 -->
		<BR>
		<BR>
		<BR>
<table border=0><tr><td align=left>
			<table>
				<tr>
					<td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
					<td nowrap><b>キー入力の場合</b></td>
					<td><img src="gif/hr.gif"></td>
				</tr>
			</table>
<center>
			<table width=500>
				<tr>
					<td colspan="2">参照したいコンテナNo.または、BL No.を入力し、『輸入照会』ボタンをクリックして下さい。複数入力する場合には","で区切って入力して下さい。<br>
					</td>
				</tr>
				<tr>
					<td width="20">&nbsp;</td>
					<td>複数コンテナNo.入力 例）FYTU2334999,HYKU9882272,DYTU3998821</td>
				</tr>
			</table>
			<form action="impcntnr.asp" name="HitsForm">
				<table border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td><table border="1" cellspacing="1" cellpadding="3" bgcolor="#ffffff">
							<tr>
								<td bgcolor="#000099" nowrap><font color="#FFFFFF"><b>コンテナNo.</b></font></td>
								<td nowrap><table border=0 cellpadding=0 cellspacing=0>
										<tr>
											<td><input type=text name=cntnrno size=20 maxlength="100">
											</td>
											<td align=left valign=middle nowrap><font size=1 color="#ee2200">[ 必須入力 ]</font><BR>
													<font size=1 color="#2288ff">[ 半角英数 ]</font> </td>
										</tr>
								</table></td>
							</tr>
							<tr>
								<td align="center" colspan="2">または </td>
							</tr>
							<tr>
								<td bgcolor="#000099"><font color="#FFFFFF"><b>BL No.</b></font></td>
								<td nowrap><table border=0 cellpadding=0 cellspacing=0>
										<tr>
											<td><input type="text" name=blno size=20 maxlength="100">
											</td>
											<td align=left valign=middle nowrap><font size=1 color="#ee2200">[ 必須入力 ]</font><BR>
													<font size=1 color="#2288ff">[ 半角英数 ]</font> </td>
										</tr>
								</table></td>
							</tr>
						</table></td>
						<td align="center" valign="top"><font size="-1">(※1)</font><br>
							<img src="gif/ya.gif" width="37" height="19" hspace="4"></td>
						<td valign="top"><table border="1" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF">
<!-- mod-s by MES 2015/06/08 表示方法変更対応 -->
							<tr>
								<td align="center" nowrap bgcolor="#000099" colspan="2"><font color="#FFFFFF"><b>仕出港情報</b></font></td>
							</tr>
							<tr>
								<td align="center" nowrap bgcolor="#ffff99"><font color="#000000"><b>中国</b></font></td>
								<td align="center" nowrap bgcolor="#ffff99"><font color="#000000"><b>東南アジア</b></font></td>
							</tr>
							<tr>
								<td align="center"><table border="0" cellspacing="2">
									<tr>
										 <td nowrap align="center"><a href="javascript:Submit('Form1')" class="splinkG" onClick="javascript:winOpen('win1','./cct/index.html',560,500) ">&nbsp;赤湾&nbsp;</a></td>
										 <td nowrap align="center"><a href="javascript:SubmitGaibu('GaibuifForm', 'HUANG')" class="splinkY">&nbsp;黄埔&nbsp;</a></td>
										 <td nowrap align="center"><a href="javascript:SubmitGaibu('GaibuifForm', 'QINGD')" class="splinkB">&nbsp;青島&nbsp;</a></td>
										 <td nowrap align="center"><a href="#" class="dummylink" onClick="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</a></td>
									</tr>
									<tr>              

<!-- 2015/11/30 Del-S MES Aoyagi -->
										<td nowrap align="center"><a href="http://iport.sctcn.com/en-us/" target="_blank" class="splinkG" onClick="javascript:winOpen('win1','./sct/index.htm',560,500)">&nbsp;蛇口&nbsp;</a></td> 
<!-- 2015/11/30 Del-E MES Aoyagi -->

<!-- 2015/11/30 Del-S MES Aoyagi
										<td nowrap align="center"><a href="javascript:Submit('queryForm')" class="splinkG" onClick="javascript:winOpen('win1','./sct/index.asp',560,500)">&nbsp;蛇口&nbsp;</a></td> 
2015/11/30 Del-E MES Aoyagi -->
										<td nowrap align="center"><a href="javascript:SubmitGaibu('GaibuifForm', 'NANSH')" class="splinkY">&nbsp;南沙&nbsp;</a></td>
										<td nowrap align="center"><a href="#" class="dummylink" onClick="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</a></td>
										<td nowrap align="center"><a href="#" class="dummylink" onClick="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</a></td>
									</tr>
								</table></td>
								<td align="center"><table border="0" cellspacing="2">
<!--
									<tr>
										 <td nowrap align="center"><a href="javascript:SubmitGaibu('GaibuifForm', 'TWTPE')" class="splinkR">&nbsp;台北&nbsp;</a></td>
										 <td nowrap align="center"><a href="#" class="dummylink" onClick="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</a></td>
									</tr>
-->
									<tr>
										 <td nowrap align="center"><a href="javascript:SubmitGaibu('GaibuifForm', 'THBKK')" class="splinkLG">&nbsp;バンコク&nbsp;</a></td>
									</tr>
								</table></td>
							</tr>
						</table></td>
<!-- mod-s by MES 2015/06/08 表示方法変更対応 -->
					</tr>
				</table>
							<br>
							<input type=submit value="   輸入照会   ">
							<br>
							<br>
							<table width="500" border="0" cellspacing="0" cellpadding="0">
								<tr>
									<td>何も入力しないで『輸入照会』ボタンを押すと照会結果の サンプルを表示します。 <br>
			(※1)コンテナNo.を入力後、右の赤ボタンをクリックすると当該港内での位置情報等が表示されます。</td>
								</tr>
							</table>
			</form>
<!-- add by nics 2015.03.18 -->
		<center>
				<form>
				<input type="button" value="中央ふ頭" id="cam" style="width: 150px;" onclick="javascript:location.href='<%=SUBDIR%>impentry.asp'">
				</form>
		</center>
<!-- end of add by nics 2015.03.18 -->

<!-- 2015/11/30 Del-S MES Aoyagi
<!-- 2011/11/15 URL修正 by Nics Start 
<form name="queryForm" method="get" target="_blank" action="http://iport.sctcn.com/portal/page/portal/PG_IPort/Tab_OI/">
		<input type="hidden" name="p_parametertype" value="ContainerInfo">
		<input type="hidden" name="p_parametervalue" id="data">
2015/11/30 Del-E MES Aoyagi -->
<!-- 2011/11/15 URL修正 by Nics
<form name="queryForm" method="post" target="_blank" action="http://oi.sctcn.com/Default.aspx?Action=Nav&Content=CONTAINER%20INFO.%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20&sm=CONTAINER%20INFO.">
		<input type="hidden" name="data">		
		<input type="hidden" name="OrgMenu" value="">
		<input type="hidden" name="targetPage" value="CONTAINER_INFO">
		<input type="hidden" name="nav" value="CONTAINER INFO.                         ">
-->
<!-- 2011/11/15 URL修正 by Nics End -->
</form>

<!-- 2014/1/8 DEL-S MES aoyagi
<form name="Form1" method="post" action="http://www.cwcct.com/cct/conhis/con_his_info_show.aspx" id="Form1" target="_blank">
		<input type="hidden" name="Image1.x" value="0" />
		<input type="hidden" name="Image1.y" value="0" />
	2014/1/8 DEL-E MES aoyagi -->
<!-- 2014/1/8 ADD-S MES aoyagi -->
<form name="Form1" method="post" action="http://uport.cwcct.com/Portal/Ship/EN/Public/Pub_cntr_history_show.aspx" id="Form1" target="_blank">
<!-- 2014/1/8 ADD-E MES aoyagi -->
<!--
		<input type="hidden" name="__EVENTTARGET" value="" />
		<input type="hidden" name="__EVENTARGUMENT" value="" /> 
		<input type="hidden" name="__VIEWSTATE" value="dDwtMzMwNTk0MTUxOztsPEltYWdlMTs+Po9koK7lFKyndTfCh4n1g7KjLvsH" />
-->
<!-- 2014/1/8 DEL-S MES aoyagi
		<input type="hidden" name="cont_no" id="cont_no"/>
		<input type="hidden" name="wyex" value="wyE" />
	2014/1/8 DEL-E MES aoyagi -->

<!-- 2014/1/8 ADD-S MES aoyagi -->
		<input type="hidden" name="txtContainer_no" id="txtContainer_no" />
		<input type="hidden" name="rdoDisplay" id="rdoHTML" value="HTML" />
<!-- 2014/1/8 ADD-E MES aoyagi -->

</form>

<!-- 2009/10/29 add-s 中国外部データの検索機能追加 -->
<form name="GaibuifForm" method="get" action="./gaibuif/expcntnr.asp" id="GaibuifForm">
		<input type="hidden" name="cntnrno" id="cntnrno"/>
		<input type="hidden" name="portcode" id="usercode"/>
</form>
<!-- 2009/10/29 add-e 中国外部データの検索機能追加 -->


</center>
					<table>
						<tr> 
							<td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
							<td nowrap><b>CSVファイル入力の場合</b></td>
							<td><img src="gif/hr.gif"></td>
						</tr>
					</table>
<center>
					<table border="0" cellspacing="1" cellpadding="2">
						<tr> 
							<td> 
								<p>検索条件をファイル転送する場合はここをクリック</p>
							</td>
							<td>…</td>
							<td><a href="impcsv.asp">CSVファイル転送</a></td>
						</tr>
						<tr> 
							<td>CSVファイル転送についての説明はここをクリック</td>
							<td>…</td>
							<td><a href="help02.asp">ヘルプ</a></td>
						</tr>
					</table>

				<br>
<!-- commented by nics 2015.03.18
				<form>
				<input type="button" value="中央ふ頭" style="width: 150px" onclick="javascript:location.href='<%=SUBDIR%>impentry.asp'">
				</form>
end of comment by nics 2015.03.18 -->
</center>

			</td>
		</tr>
		</table>

			<br>
		</td>
	</tr>
	<tr>
		<td valign="bottom">
<%
		DispMenuBar
%>
		</td>
	</tr>
</table>
<!-------------照会画面終わり--------------------------->
<%
		DispMenuBarBack "../index.asp" 'http://www.hits-h.com/index.asp
%>
</body>
</html>
