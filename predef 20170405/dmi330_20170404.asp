<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi330.asp				_/
'_/	Function	:事前実搬入入力確認画面			_/
'_/	Date		:2003/05/29				_/
'_/	Code By		:SEIKO Electric.Co 大重			_/
'_/	Modify		:C-002	2003/08/06	備考欄追加	_/
'_/	Modify		:3th	2003/01/31	3次変更	_/
'_/	Modify		:20170118 T.Okui 設定温度を各社ビューから取得_/
'_/	Modify		:	20170207 T.Okui 全体レイアウト変更         _/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTPコンテンツタイプ設定
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->
<!--#include File="CommonFunc.inc"-->
<%
'セッションの有効性をチェック
  CheckLoginH

'データを取得
  dim CMPcd,Hmon,Hday
  CMPcd = Array(Request("CMPcd0"),Request("CMPcd1"),Request("CMPcd2"),Request("CMPcd3"),Request("CMPcd4"))

'表示文言生成
'3th  If Request("Hmon") = 0 Then 
'3th    Hmon = " "
'3th  Else
'3th    Hmon = Right("0"&Request("Hmon"),2)
'3th  End If

'3th  If Request("Hday") = 0 Then 
'3th    Hday = " "
'3th  Else
'3th    Hday = Right("0"&Request("Hday"),2)
'3th  End If

  dim Mord,ret
  dim ErrMsg, StrCodes, NiukeNm, LPortNm, DPortNm, NiwataNm					'2016/11/03 H.Yoshikawa Add

  Mord = Request("Mord")
  If Mord=2 Then
    ret = true
  Else
  'DB接続
    dim ObjConn, ObjRS, StrSQL
    ConnDBH ObjConn, ObjRS
  'ヘッドIDのチェック
    checkHdCd ObjConn, ObjRS, CMPcd, ret
   '2016/11/03 H.Yoshikawa Add Start
   ErrMsg = ""
   if ret = false then
   		ErrMsg = "指定された会社コードは存在しません。<BR>"
   end if
   
   '港コードのチェック
	StrCodes="'"&gfSQLEncode(Request("NiukP"))&"','"&gfSQLEncode(Request("TumiP"))&"','"&gfSQLEncode(Request("AgeP"))&"','"&gfSQLEncode(Request("NiwataP"))&"'"
	StrSQL = "SELECT mP.PortCode,mP.FullName From mPort AS mP "&_
	       "WHERE mP.PortCode IN ("& StrCodes &") "
	ObjRS.Open StrSQL, ObjConn
	if err <> 0 then
		DisConnDBH ObjConn, ObjRS	'DB切断
		jampErrerP "1","b401","01","実搬入：港データ取得","103","SQL:<BR>"&StrSQL
	end if
	Do Until ObjRS.EOF
		If Not IsNull(ObjRS("FullName")) Then
		  If gfTrim(Request("NiukP"))=gfTrim(ObjRS("PortCode")) Then
		    NiukeNm=gfTrim(ObjRS("FullName"))
		  End If
		  If gfTrim(Request("TumiP"))=gfTrim(ObjRS("PortCode")) Then
		    LPortNm=gfTrim(ObjRS("FullName"))
		  End If
		  If gfTrim(Request("AgeP"))=gfTrim(ObjRS("PortCode")) Then
		    DPortNm=gfTrim(ObjRS("FullName"))
		  End If
		  If gfTrim(Request("NiwataP"))=gfTrim(ObjRS("PortCode")) Then
		    NiwataNm=gfTrim(ObjRS("FullName"))
		  End If
		End If
		ObjRS.MoveNext
	Loop
	ObjRS.Close
	if NiukeNm = "" then
		ErrMsg = ErrMsg & "荷受地のコードが正しくありません。<BR>"
		ret = false
	end if
	if LPortNm = "" then
		ErrMsg = ErrMsg & "積港のコードが正しくありません。<BR>"
		ret = false
	end if
	if DPortNm = "" then
		ErrMsg = ErrMsg & "揚港のコードが正しくありません。<BR>"
		ret = false
	end if
	if NiwataNm = "" then
		ErrMsg = ErrMsg & "荷渡地のコードが正しくありません。<BR>"
		ret = false
	end if
    if ErrMsg <> "" then
    	ErrMsg = ErrMsg & "「戻る」ボタンを押下し、再入力してください。<BR>"
    end if
	'2016/11/03 H.Yoshikawa Add End

  'DB接続解除
    DisConnDBH ObjConn, ObjRS
  'エラートラップ解除
    on error goto 0
  End If
  dim tmpstr
  If Request("UpFlag") <> 5 Then
    tmpstr=CMPcd(Request("UpFlag"))&"/"
  Else
    tmpstr="/"
  End If
  tmpstr=tmpstr&Request("HedId")&"/"&Hmon & Hday&"/"&Request("CONsize")&"/"&Request("CONtype") &_
           "/"&Request("CONhite")&"/"&Request("CONsitu")&"/"&Request("CONtear")&"/"&Request("MrSk") &_
           "/"&Request("SealNo")&"/"&Request("GrosW")&"/"&Request("Hfrom")&"/"&Request("TuSk")&"/"&Request("OH") &_
           "/"&Request("OWL")&"/"&Request("OWR")&"/"&Request("OLF")&"/"&Request("OLA")
  If ret Then
    tmpstr=tmpstr&",入力内容の正誤:0(正しい)"
  Else
    tmpstr=tmpstr&",入力内容の正誤:1(誤り)"
  End If
'3th Change Start
'  WriteLogH "b402", "実搬入事前情報入力","13",tmpstr
  If Mord="0" Then
    WriteLogH "b402", "実搬入事前情報入力","02",tmpstr
  Else
    WriteLogH "b402", "実搬入事前情報入力","13",tmpstr
  End If
'3th Cange End
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<STYLE type="text/css">
DIV.bgb{
	text-align:center;
	margin-left:4px;
}
DIV.bgy{
	text-align:center;
	margin-left:4px;
}
</STYLE>
<TITLE>搬入票作成情報入力(確認)</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--

function setParam(target){
  len = target.elements.length;
  for (i=0; i<len-5; i++) target.elements[i].readOnly = true;
  bgset(target);
}

//登録
function GoEntry(){
  target=document.dmi330F;
  target.action="./dmi340.asp";
  target.submit();
}
//戻る
function GoBackT(){
  target=document.dmi330F;
  target.action="./dmi320.asp";
  target.submit();
}
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="setParam(document.dmi330F)">
<!-------------実搬入情報入力確認画面--------------------------->
<FORM name="dmi330F" method="POST">
<TABLE border=0 cellPadding=2 cellSpacing=0 width="100%">
  <TR>
    <TD><B>搬入票作成情報入力確認</B></TD>
    <TD>

    </TD></TR>
  <TR>
    <TD width="500" colspan=2 valign=top>
    <TABLE border=0 cellPadding=2 cellSpacing=0 width="100%">
	  <TR>
	  <TD>
        <DIV style="height:330px;width:500px;border: 1px solid black; margin:5px;">
	<TABLE border=0 cellPadding=2 cellSpacing=0 width="100%">
	  <TR>
    	    <TD><DIV STYLE="FONT-WEIGHT:BOLD;">BOOKING情報</DIV></TD>
    	    <TD></TD></TR>
	  <TR>
	    <TD><DIV class=bgb>ブッキング番号</DIV></TD>
	    <TD><INPUT type=text name="BookNo" value="<%=Request("BookNo")%>"></TD></TR>
	  <TR>
	    <TD><DIV class=bgb>コンテナ番号</DIV></TD>
	    <TD><INPUT type=text name="CONnum" value="<%=Request("CONnum")%>"></TD></TR>
	  <TR>
	    <TD><DIV class=bgb>*シール番号</DIV></TD>
	    <TD><INPUT type=text name="SealNo" value="<%=Request("SealNo")%>"></TD></TR>
	  <TR>
	    <TD><DIV class=bgb>取扱船社</DIV></TD>
	    <TD><INPUT type=hidden name="ThkSya" value="<%=Request("ThkSya")%>" size=27>			<!-- 2016/08/08 H.Yoshikawa Upd (text→hidden) -->
	    <INPUT type=text name="ShipLineName" value="<%=Request("ShipLineName")%>" size=40>		<!-- 2016/08/08 H.Yoshikawa Add -->
	    </TD></TR>
	  <TR>
	    <TD><DIV class=bgb>*本船名</DIV></TD>
	    <TD><INPUT type=text name="ShipN" value="<%=Request("ShipN")%>">
	    	<INPUT type=hidden name="ShipCode" value="<%=Request("ShipCode")%>">				<!-- 2016/08/08 H.Yoshikawa Add -->
	    </TD></TR>
	  <TR>
	    <TD><DIV class=bgb>*次航</DIV></TD>
	    <TD><INPUT type=text name="NextV" value="<%=Request("NextV")%>">
	        <INPUT type=hidden name="VoyCtrl" value="<%=Request("VoyCtrl")%>">					<!-- 2016/10/17 H.Yoshikawa Add -->
	    </TD></TR>
	  <TR>
	    <TD><DIV class=bgb>*荷受地</DIV></TD>
	    <TD><INPUT type=text name="NiukP" value="<%=Request("NiukP")%>" size=8>					<!-- 2016/11/03 H.Yoshikawa Upd(size追加)-->
	    	<INPUT type=text name="NiukeNm" value="<%=NiukeNm%>" size=30>						<!-- 2016/11/03 H.Yoshikawa Add -->
	    </TD></TR>
	  <TR>
	    <TD><DIV class=bgb>*積港</DIV></TD>
	    <TD><INPUT type=text name="TumiP" value="<%=Request("TumiP")%>" size=8>					<!-- 2016/11/03 H.Yoshikawa Upd(size追加)-->
	    	<INPUT type=text name="LPortNm" value="<%=LPortNm%>" size=30>						<!-- 2016/11/03 H.Yoshikawa Add -->
	    </TD></TR>
	  <TR>
	    <TD><DIV class=bgb>*揚港</DIV></TD>
	    <TD><INPUT type=text name="AgeP" value="<%=Request("AgeP")%>" size=8>					<!-- 2016/11/03 H.Yoshikawa Upd(size追加)-->
	    	<INPUT type=text name="DPortNm" value="<%=DPortNm%>" size=30>						<!-- 2016/11/03 H.Yoshikawa Add -->
	    </TD></TR>
	  <TR>
	    <TD><DIV class=bgb>*荷渡地</DIV></TD>
	    <TD><INPUT type=text name="NiwataP" value="<%=Request("NiwataP")%>" size=8>				<!-- 2016/11/03 H.Yoshikawa Upd(size追加)-->
	    	<INPUT type=text name="NiwataNm" value="<%=NiwataNm%>" size=30>						<!-- 2016/11/03 H.Yoshikawa Add -->
	    </TD></TR>
	<!-- 2016/08/08 H.Yoshikawa Add Start -->
	  <TR>
	    <TD><DIV class=bgb>*荷主</DIV></TD>
	    <TD><INPUT type=text name="Shipper" value="<%=Request("Shipper")%>" size=40></TD>
	  </TR>
	  <TR>
	    <TD><DIV class=bgb>搬入先</DIV></TD>
	    <TD><INPUT type=text name="HTo" value="<%=Request("HTo")%>" size=30></TD></TR>
	  <TR>
	    <TD><DIV class=bgb>ターミナルオペレータ</DIV></TD>
	    <TD><INPUT type=text name="OpeName" value="<%=gfConvertOperator(Request("Operator"))%>"></TD>
	    <INPUT type=hidden name="Operator" value="<%=Request("Operator")%>">
	    </TR>
	  </TABLE>
	  </DIV>
	  </TD>
	  </TR>
	  <TR>
	  <TD>
	  <DIV style="height:160px;width:500px;border: 1px solid black; margin:5px;">
	  <TABLE border=0 cellPadding=2 cellSpacing=0 width="100%">
	  <TR>
	    <TD><DIV STYLE="FONT-WEIGHT:BOLD;">登録情報</DIV></TD>
	    <TD></TD>
	  </TR>
	  <TR>
	    <TD><DIV class=bgb>登録会社コード</DIV></TD>
	    <TD>
	        <INPUT type=text name="CMPcd0" value="<%=CMPcd(0)%>" size=7>
	        </TD></TR>
	  <TR>
	  <TR>
	    <TD><DIV class=bgb>指示先会社コード</DIV></TD>
	    <TD>
	        <INPUT type=text name="CMPcd1" value="<%=CMPcd(1)%>" size=5>
	        <INPUT type=text name="CMPcd2" value="<%=CMPcd(2)%>" size=5>
	        <INPUT type=text name="CMPcd3" value="<%=CMPcd(3)%>" size=5>
	        <INPUT type=text name="CMPcd4" value="<%=CMPcd(4)%>" size=5></TD></TR>
	  <TR>
	    <TD><DIV class=bgb>ヘッドＩＤ</DIV></TD>
	    <TD><INPUT type=text name="HedId" value="<%=Request("HedId")%>"></TD></TR>
	  <TR>
	    <TD><DIV class=bgb>搬入元</DIV></TD>
	    <TD><INPUT type=text name="HFrom" value="<%=Request("Hfrom")%>" size=30></TD></TR>
	  <TR>
	    <TD><DIV class=bgb>搬入予定日</DIV></TD>
	    <TD><INPUT type=text name="Hmon" value="<%=Request("Hmon")%>" size=2>月
	        <INPUT type=text name="Hday" value="<%=Request("Hday")%>" size=2>日</TD></TR>
	  </TABLE>
	  </DIV>
	   </TD>
	  </TR>
	  <TR>
	  <TD>
	  
	  
	  
	  <DIV style="height:100px;width:500px;border: 1px solid black; margin:5px;">
	  <TABLE border=0 cellPadding=2 cellSpacing=0 width="100%">
	  	<TR>
	    <TD colspan=2><DIV><SPAN  STYLE="FONT-WEIGHT:BOLD;">搬入票署名欄情報</SPAN><SPAN STYLE="color:red;">※この欄が搬入票署名欄に印字されます</SPAN></DIV></TD>
	    </TR>
	  	<!--<TR><TD colspan=2 style="color:red;">この欄が搬入票署名欄へ印字されます！！</TD></TR>
	  	<TR><TD><DIV class=bgb>*取扱海貨社名<BR>*（担当者）<BR>*（連絡先）</DIV></TD>
	    	<TD><INPUT type=text name="Forwarder" value="<%=Request("Forwarder")%>" style="margin-bottom:2px;" size=40><BR>
	    		<INPUT type=text name="FwdStaff" value="<%=Request("FwdStaff")%>" ><BR>
	    		<INPUT type=text name="FwdTel" value="<%=Request("FwdTel")%>" ></TD>
	    </TR>-->
	  	<TR><TD><DIV class=bgb>*取扱海貨社名</DIV></TD>
	    	<TD><INPUT type=text name="Forwarder" value="<%=Request("Forwarder")%>" style="margin-bottom:2px;" size=40></TD></TR>
	    <TR>
	    	<TD><DIV class=bgb>*（担当者）</DIV></TD>
	    	<TD><INPUT type=text name="FwdStaff" value="<%=Request("FwdStaff")%>" ></TD></TR>
	    <TR>
	    	<TD><DIV class=bgb>*（連絡先）</DIV></TD>
	    	<TD><INPUT type=text name="FwdTel" value="<%=Request("FwdTel")%>" ></TD></TR>
	  	</TR>
	  </TABLE>
	  </DIV>
	  </TD></TR>
	  <TR>
	  <TD>
	  <DIV style="height:55px;width:500px;border: 1px solid black; margin:5px;">
	  <TABLE border=0 cellPadding=2 cellSpacing=0 width="100%">
	  	<TR>
	   <TD><DIV class=bgy>*登録担当者</DIV></TD>
	   <TD><INPUT type=text name="TruckerSubName" value="<%=Request("TruckerSubName")%>"></TD>
	<!-- 2009/03/10 R.Shibuta Add-E -->
	  </TR>
<!-- 2016/08/18 H.Yoshikawa Add-S -->
	  <TR>
	   <TD><DIV class=bgy>*登録者連絡先</DIV></TD>
	   <TD><INPUT type=text name="TruckerTel" value="<%=Request("TruckerTel")%>" onBlur="CheckLen(this,true,true,false)"></TD>
	  </TR>
	<!-- 2016/08/17 H.Yoshikawa Add End   -->
	</TABLE>
	</DIV>
	</TD>
	</TR>
	
	</TABLE>
	</TD>
    <TD width=300 valign=top>
	<TABLE border=0 cellPadding=2 cellSpacing=0 width="100%">
	  <TR>
	    <TD width=300>
		    <DIV style="height:140px;width:300px;border: 1px solid black; margin:5px;">
		    <TABLE cellpadding=1 cellspacing=0>
		    <TR>
		    	<TD><DIV STYLE="FONT-WEIGHT:BOLD;">コンテナ情報</DIV></TD>
		    	<TD></TD>
		    </TR>
		    <TR>
		    	<TD>
		        	<DIV class=bgb>*サイズ</DIV></TD>										
		    	<TD><INPUT type=text name="CONsize" value="<%=Request("CONsize")%>" size=5>
		    	</TD>
		    </TR>
		    <TR>
		    	<TD>
		        	<DIV class=bgb>*タイプ</DIV></TD>										
		    	<TD>
		        	<INPUT type=text name="CONtype" value="<%=Request("CONtype")%>" size=5>
		    	</TD>
		    </TR>
		    <TR>
		    	<TD>
		        	<DIV class=bgb>*高さ</DIV></TD>										
		    	<TD>
		        	<INPUT type=text name="CONhite" value="<%=Request("CONhite")%>" size=5>
		    	</TD>
		    </TR>
		    <!-- 20170118 T.OKui Upd Start -->
		    <TR>
		    	<TD><DIV class=bgb>設定温度</DIV></TD>
		    	<TD><INPUT type=text name="SttiT" value="<%=Request("SttiT")%>" size=6>&nbsp;<%=Request("TempDegree")%>
		    	<!--
		    	<INPUT type=checkbox <% if gfTrim(Request("AsDry")) = "1" then %>checked<% end if %> disabled>AS DRY
		    	<INPUT type=hidden name="AsDry" value="<%=gfTrim(Request("AsDry"))%>" >
		    	-->
		    	</TD>
		    </TR>
		    <!-- 20170118 T.OKui Upd Start -->
		    <TR>
		    	<TD><DIV class=bgb>ベンチレーション</DIV></TD>
		    	<TD><INPUT type=text name="VENT" value="<%=Request("VENT")%>" size=5>&nbsp;%（開口）</TD>
		    </TR>
		    <!-- 2017/03/02 T.Okui Del Start -->
		    <!--
		    <TR>
		    	<TD><DIV class=bgb>丸関</DIV></TD>
		    	<TD><INPUT type=text name="MrSk" value="<%=Request("MrSk")%>" size=5></TD>
		    </TR>
		    -->
		    <!-- 2017/03/02 T.Okui Del End -->
	  </TABLE>
	  </DIV>
	  </TD>
	  <TD valign=top>
	  
	  
	  <DIV style="height:140px;width:300px;border: 1px solid black; margin:5px;">
		  	<TABLE cellpadding=1 cellspacing=0>
		  		<TR>
			    <TD><DIV STYLE="FONT-WEIGHT:BOLD;">コンテナ規格外貨物情報</DIV></TD>
			    <TD></TD></TR>
		  		<TR>
			    <TD><DIV class=bgb>オーバーハイ（上部）</DIV></TD>
			    <TD><INPUT type=text name="OH"  value="<%=Request("OH")%>"  size=5 >&nbsp;cm</TD></TR>
			    <TR>
			    <TD><DIV class=bgb>オーバーワイド（右）</DIV></TD>
			    <TD><INPUT type=text name="OWR" value="<%=Request("OWR")%>" size=5 >&nbsp;cm</TD></TR>
			    <TR>
			    <TD><DIV class=bgb>オーバーワイド（左）</DIV></TD>
			    <TD><INPUT type=text name="OWL" value="<%=Request("OWL")%>" size=5 >&nbsp;cm</TD></TR>
			    <TR>
			    <TR>
			    <TD><DIV class=bgb>オーバーレングス（前）</DIV></TD>
			    <TD><INPUT type=text name="OLF" value="<%=Request("OLF")%>" size=5 >&nbsp;cm</TD></TR>
			    <TR>
			    <TD><DIV class=bgb>オーバーレングス（後）</DIV></TD>
			    <TD><INPUT type=text name="OLA" value="<%=Request("OLA")%>" size=5 >&nbsp;cm</TD></TR>
			</TR>
		  </TABLE>
	   </DIV>
		
	  	</TD>
	  </TR>  
	  <TR>
	  <TD colspan=2>
	  <DIV style="border: 1px solid black; margin:5px;height:125px;">
	  <TABLE>
		  <TR>
		    <TD width="115"><DIV STYLE="FONT-WEIGHT:BOLD;">重量情報</DIV></TD>
		    <TD></TD></TR>
	  <TR>
	    <TD><DIV class=bgb>*コンテナ総重量</DIV></TD>											<!-- 2016/08/08 H.Yoshikawa Upd (グロスウェイト→コンテナグロス) -->
	    <TD><INPUT type=text name="GrosW" value="<%=Request("GrosW")%>" size=9>&nbsp;kg</TD></TR>
	  <TR>
	    <TD>
	        <DIV class=bgb>テアウェイト</DIV></TD>										<!-- 2016/08/08 H.Yoshikawa Upd (材質削除) -->
	    <TD>
	        <!-- <INPUT type=text name="CONsitu" value="<%=Request("CONsitu")%>" size=5> -->	<!-- 2016/08/08 H.Yoshikawa Del -->
	        <INPUT type=text name="CONtear" value="<%=Request("CONtear")%>" size=7>kg
	    </TD></TR>
	  <TR>
	    <TD><DIV class=bgb>計量方法（確認）</DIV></TD>
		<TD style="padding-top:0px;">
			<INPUT type=hidden name="SolasChk" value="<%=gfTrim(Request("SolasChk"))%>" >	    <!-- 2016/08/08 H.Yoshikawa Add -->
			<INPUT type=checkbox <% if gfTrim(Request("SolasChk")) = "1" then %>checked<% end if %> disabled>ここに入力したコンテナ総重量はSOLAS条約に基づく方法で計測された数値です。
		</TD>
	  </TR>
	  <TR>
	    <TD><DIV class=bgb>届出番号　登録番号</DIV></TD>
	    <!--<TD><INPUT type=text name="ReportNo" value="<%=Request("ReportNo")%>" size=35 >　<span style="color:red">※ハイフンなしで入力してください</span></TD></TR>-->
	    <TD><INPUT type=text name="ReportNo" value="<%=Request("ReportNo")%>" size=17 maxlength=12 >　<span style="color:red">※ハイフンなしで入力してください</span></TD></TR>
	  
	  <!-- <TR>
	    <TD><DIV class=bgb>通関</DIV></TD>
	    <TD><INPUT type=text name="TuSk" value="<%=Request("TuSk")%>" size=5></TD></TR> -->
	  </TABLE>
	  </DIV>
	  </TD>
	  </TR>
	  <TR>
	  <TD colspan=2>
	  <DIV style="border: 1px solid black; margin:5px;height:185px;">
	  <TABLE>
		  <TR>
		    <TD><DIV STYLE="FONT-WEIGHT:BOLD;">危険品貨物情報</DIV></TD>
		    <TD></TD></TR>
	  <TR>
	    <TD><DIV class=bgb>ＲＨＯ</DIV></TD>
	    <TD><INPUT type=text name="RHO" value="<%=Request("RHO")%>" size=5></TD></TR>
	  <TR>
		    <TD></TD>
		    <TD>
	        	<TABLE cellpadding=0 cellspacing=0 style="font-weight:bold;"><TR>
		        	<TD align=center width="47px">1</TD>
		        	<TD align=center width="50px">2</TD>
		        	<TD align=center width="50px">3</TD>
		        	<TD align=center width="50px">4</TD>
		        	<TD align=center width="50px">5</TD>
		        </TR></TABLE>
		    </TD></TR>
	  <TR>
	    <TD><DIV class=bgb>ＩＭＤＧコード</DIV></TD>
	    <TD>
	        <TABLE cellpadding=0 cellspacing=0><TR>
	        	<TD width="50px"><INPUT type=text name="IMDG1" value="<%=Request("IMDG1")%>" size=6 ></TD>	<!-- 2016/08/09 H.Yoshikawa Upd （size5→6、readOnly属性削除) -->
	        	<TD width="50px"><INPUT type=text name="IMDG2" value="<%=Request("IMDG2")%>" size=6 ></TD>	<!-- 2016/08/09 H.Yoshikawa Upd （size5→6、readOnly属性削除) -->
	        	<TD width="50px"><INPUT type=text name="IMDG3" value="<%=Request("IMDG3")%>" size=6 ></TD>	<!-- 2016/08/09 H.Yoshikawa Upd （size5→6、readOnly属性削除) -->
	        	<TD width="50px"><INPUT type=text name="IMDG4" value="<%=Request("IMDG4")%>" size=6 ></TD>	<!-- 2016/08/09 H.Yoshikawa Add -->
	        	<TD width="50px"><INPUT type=text name="IMDG5" value="<%=Request("IMDG5")%>" size=6 ></TD>	<!-- 2016/08/09 H.Yoshikawa Add -->
	        </TR></TABLE>
	    </TD>
	  </TR>
	  <!-- 2016/08/01 H.Yoshikawa Add End -->
	  <TR>
	    <TD style="padding-top:0px;"><DIV class=bgb>ＵＮコード</DIV></TD>
	    <TD style="padding-top:0px;">
	        <TABLE cellpadding=0 cellspacing=0><TR>
		        <TD width="50px"><INPUT type=text name="UNNo1" value="<%=Request("UNNo1")%>" size=6 ></TD>					<!-- 2016/08/06 H.Yoshikawa Upd （readOnly属性削除) -->
		        <TD width="50px"><INPUT type=text name="UNNo2" value="<%=Request("UNNo2")%>" size=6 ></TD>					<!-- 2016/08/06 H.Yoshikawa Upd （readOnly属性削除) -->
		        <TD width="50px"><INPUT type=text name="UNNo3" value="<%=Request("UNNo3")%>" size=6 ></TD>					<!-- 2016/08/06 H.Yoshikawa Upd （readOnly属性削除) -->
		        <TD width="50px"><INPUT type=text name="UNNo4" value="<%=Request("UNNo4")%>" size=6 ></TD>					<!-- 2016/08/06 H.Yoshikawa Add -->
		        <TD width="50px"><INPUT type=text name="UNNo5" value="<%=Request("UNNo5")%>" size=6 ></TD>					<!-- 2016/08/06 H.Yoshikawa Add -->
	        </TR></TABLE>
	    </TD>
	  </TR>
	  <!-- 2016/08/01 H.Yoshikawa Add Start -->
	  <!-- 2016/08/09 H.Yoshikawa Add Start -->
	  <TR>
	    <TD style="padding-top:0px;"><DIV class=bgb>サブラベル１</DIV></TD>
	    <TD style="padding-top:0px;">
	        <TABLE cellpadding=0 cellspacing=0><TR>
		        <TD width="50px"><INPUT type=text name="Label1" value="<%=Request("Label1")%>" size=6 ></TD>
		        <TD width="50px"><INPUT type=text name="Label2" value="<%=Request("Label2")%>" size=6 ></TD>
		        <TD width="50px"><INPUT type=text name="Label3" value="<%=Request("Label3")%>" size=6 ></TD>
		        <TD width="50px"><INPUT type=text name="Label4" value="<%=Request("Label4")%>" size=6 ></TD>
		        <TD width="50px"><INPUT type=text name="Label5" value="<%=Request("Label5")%>" size=6 ></TD>
	        </TR></TABLE>
	    </TD>
	  </TR>
	  <TR>
	    <TD style="padding-top:0px;"><DIV class=bgb>サブラベル２</DIV></TD>
	    <TD style="padding-top:0px;">
	        <TABLE cellpadding=0 cellspacing=0><TR>
		        <TD width="50px"><INPUT type=text name="SubLabel1" value="<%=Request("SubLabel1")%>" size=6 ></TD>
		        <TD width="50px"><INPUT type=text name="SubLabel2" value="<%=Request("SubLabel2")%>" size=6 ></TD>
		        <TD width="50px"><INPUT type=text name="SubLabel3" value="<%=Request("SubLabel3")%>" size=6 ></TD>
		        <TD width="50px"><INPUT type=text name="SubLabel4" value="<%=Request("SubLabel4")%>" size=6 ></TD>
		        <TD width="50px"><INPUT type=text name="SubLabel5" value="<%=Request("SubLabel5")%>" size=6 ></TD>
	        </TR></TABLE>
	    </TD>
	  </TR>
	  
	  <TR>
	    <TD style="padding-top:0px;"><DIV class=bgb>少量危険品</DIV></TD>
	    <TD style="padding-top:0px;">
	        <TABLE cellpadding=0 cellspacing=0><TR>
		        <TD width="50px" align=center><INPUT type=checkbox <% if gfTrim(Request("LqFlag1")) = "1" then %>checked<% end if %> disabled></TD>
		        <TD width="50px" align=center><INPUT type=checkbox <% if gfTrim(Request("LqFlag2")) = "1" then %>checked<% end if %> disabled></TD>
		        <TD width="50px" align=center><INPUT type=checkbox <% if gfTrim(Request("LqFlag3")) = "1" then %>checked<% end if %> disabled></TD>
		        <TD width="50px" align=center><INPUT type=checkbox <% if gfTrim(Request("LqFlag4")) = "1" then %>checked<% end if %> disabled></TD>
		        <TD width="50px" align=center><INPUT type=checkbox <% if gfTrim(Request("LqFlag5")) = "1" then %>checked<% end if %> disabled></TD>
		    	<INPUT type=hidden name="LqFlag1" value="<%=gfTrim(Request("LqFlag1"))%>" >	    			<!-- 2016/08/08 H.Yoshikawa Add -->
		    	<INPUT type=hidden name="LqFlag2" value="<%=gfTrim(Request("LqFlag2"))%>" >	    			<!-- 2016/08/08 H.Yoshikawa Add -->
		    	<INPUT type=hidden name="LqFlag3" value="<%=gfTrim(Request("LqFlag3"))%>" >	    			<!-- 2016/08/08 H.Yoshikawa Add -->
		    	<INPUT type=hidden name="LqFlag4" value="<%=gfTrim(Request("LqFlag4"))%>" >	    			<!-- 2016/08/08 H.Yoshikawa Add -->
		    	<INPUT type=hidden name="LqFlag5" value="<%=gfTrim(Request("LqFlag5"))%>" >	    			<!-- 2016/08/08 H.Yoshikawa Add -->
	        </TR></TABLE>
	    </TD>
	  </TR>
	  
	  
	</TABLE></DIV></TD>
  </TR>
  <TR>
	 
	<TD colspan=2 valign="TOP">
	<TABLE border=0 cellpadding=2 cellSpacing=0 width="100%">
	  <TR>
	    <TD align="left" valign="top">
	    　＜注意事項＞<BR>
	    　本画面の誤記・記入漏れは正常なる輸送を阻害しますので、入力済みの項目も含めて必ずご確認ください。<BR>
        　誤記・記入漏れにより発生する損害・費用・罰金等は全て本画面の入力者が負担し、船社(含むターミナル)<BR>
        　は責任を負いません。
	    </TD>
	  </TR>
	</TABLE>
	</TD>
  </TR>
  
  
  
  <TR>
  <TD colspan=2 valign="TOP">
  <TABLE border=0 cellPadding=2 cellSpacing=0 width="100%">
	  <TR>
	    <TD colspan=3 align=left valign=bottom>
	    	<INPUT type=checkbox <% if gfTrim(Request("AgreeChk")) = "1" then %>checked<% end if %> disabled>本画面の入力内容をゲートでの搬入票の代わりとして使用することに同意します。
	    	<BR>　※チェックがない場合は仮登録状態であり、予約受付は完了していません。
	    	<INPUT type=hidden name="AgreeChk" value="<%=gfTrim(Request("AgreeChk"))%>" >
		</TD>
	  </TR>
	<!-- 2016/08/18 H.Yoshikawa Add-E -->
	</TABLE>
	</TD>
	</TR>    
  <TR>
  <TD colspan=4 align=center valign=bottom>
<% If Mord=1 AND Request("UpFlag")<>1 Then %>
    <DIV class=bgw>指示元への回答　　　Yes　　　　　</DIV><P>
<% ElseIf Mord =2 Then %>
    <DIV class=bgw>指示元への回答　　　No　　　　　</DIV><P>
    <DIV class=alert><B>＜注意＞</B>回答をNoで指定の場合は入力したデータは反映されません。</DIV><P>
<% End If %>
       <INPUT type=hidden name="SakuNo"   value="<%=Request("SakuNo")%>">
       <INPUT type=hidden name="UpFlag"   value="<%=Request("UpFlag")%>">
       <INPUT type=hidden name="compFlag" value="<%=Request("compFlag")%>">
       <INPUT type=hidden name="Mord"     value="<%=Mord%>"><%'CW-028 ADD%>
       <INPUT type=hidden name="WkCNo"    value="<%=Request("WkCNo")%>" >
       <INPUT type=hidden name="partFlg"  value="<%=Request("partFlg")%>" >
       <INPUT type=hidden name="TruckerFlag" value="<%=Request("TruckerFlag")%>" >
       <INPUT type=hidden name="kariflag" value="<%=Request("kariflag")%>">					<!-- 2016/10/12 H.Yoshikawa Add -->
   
<% If Not ret Then %>
       <P><DIV class=alert>
       	<!-- 2016/11/03 H.Yoshikawa Upd Start 
        指定された会社コードは存在しません。<BR>
       「戻る」ボタンを押下し、再入力してください。-->
        <%=ErrMsg%>
        <!-- 2016/11/03 H.Yoshikawa Upd End -->
       </DIV></P></TD></TR><TR><TD>
<% Else %>
		<BR><BR>
       </TD></TR><TR><TD>
       <INPUT type=button value="ＯＫ" onClick="GoEntry()"  style="position:relative;left:220px;">
<% End If %>
       <INPUT type=button value="戻る" onClick="GoBackT()"  style="position:relative;left:220px;">
      </TD><TD>
      <% If Mord<>"0" Then %>
      <TABLE border=1 cellPadding=3 cellSpacing=0 align="left" style="position:relative;left:3px;">
          <TR bgcolor="#f0f0f0"><TD>作業番号</TD><TD><%=Request("SakuNo")%></TD></TR>
      </TABLE>
      </TD>
<% End If %>
      </TR>
<!--2017/02/06 T.Okui Del Start-->
  <% if 1=0 then%>
  <TR>
	<TD valign=top>
	<TABLE border=0 cellPadding=2 cellSpacing=0>
	  <!--<TR>
	    <TD width="90px"><DIV class=bgb>備考１</DIV></TD>
	    <TD><INPUT type=text name="Comment1" value="<%=Request("Comment1")%>" size=73></TD>
	  </TR>
	  <TR>
	    <TD><DIV class=bgb>備考２</DIV></TD>
	    <TD><INPUT type=text name="Comment2" value="<%=Request("Comment2")%>" size=73></TD></TR>
	  <TR>
	    <TD><DIV class=bgb>備考３</DIV></TD>
	    <TD><INPUT type=text name="Comment3" value="<%=Request("Comment3")%>" size=73></TD></TR>
	  <TR>-->
	<!-- 2009/03/10 R.Shibuta Add-S -->
	  <TR>
	   <TD><DIV class=bgy>登録担当者</DIV></TD>
	   <TD><INPUT type=text name="TruckerSubName" value="<%=Request("TruckerSubName")%>"></TD>
	<!-- 2009/03/10 R.Shibuta Add-E -->
	  </TR>
	<!-- 2016/08/08 H.Yoshikawa Add-S -->
	  <TR>
	   <TD><DIV class=bgy>*登録者連絡先</DIV></TD>
	   <TD><INPUT type=text name="TruckerTel" value="<%=Request("TruckerTel")%>" onBlur="CheckLen(this,true,true,false)"></TD>
	  </TR>
	</TABLE>
	</TD>
	<TD valign=top>
	<!-- 2016/08/01 H.Yoshikawa Add-S -->
	<TABLE border=0 cellPadding=2 cellSpacing=0>
	  <TR>
	    <TD align="left" valign="top" >
	    　＜注意事項＞<BR>
	    　本画面の誤記・記入漏れは正常なる輸送を阻害しますので、<BR>
	    　入力済みの項目も含めて必ずご確認ください。<BR>
        　誤記・記入漏れにより発生する損害・費用・罰金等は<BR>
        　全て本画面入力者が負担し、船社(含むターミナル)は<BR>
        　責任を負いませんので、予めご承知おき願います。
	    </TD>
	  </TR>
	</TABLE>
	<!-- 2016/08/01 H.Yoshikawa Add-E -->
	</TD>
  </TR>
  <TR>
	<TD colspan=2 align=center>
	<BR>
	<TABLE border=0 cellPadding=2 cellSpacing=0>
	  <TR>
	    <TD colspan=3 align=center valign=bottom>
	    	<BR><INPUT type=checkbox <% if gfTrim(Request("AgreeChk")) = "1" then %>checked<% end if %> disabled>本画面の入力内容をゲートでの搬入票の代わりとして使用することに同意します。
	    	<BR>　※チェックを入れずに「登録」をした場合は、仮登録であり、予約受付は完了していません。
	    	<INPUT type=hidden name="AgreeChk" value="<%=gfTrim(Request("AgreeChk"))%>" >	    			<!-- 2016/08/08 H.Yoshikawa Add -->
		</TD>
	  </TR>
	<!-- 2016/08/08 H.Yoshikawa Add-E -->
	</TABLE>
	</TD>
  </TR>

    <TD colspan=4 align=center valign=bottom>
<% If Mord=1 AND Request("UpFlag")<>1 Then %>
    <DIV class=bgw>指示元への回答　　　Yes　　　　　</DIV><P>
<% ElseIf Mord =2 Then %>
    <DIV class=bgw>指示元への回答　　　No　　　　　</DIV><P>
    <DIV class=alert><B>＜注意＞</B>回答をNoで指定の場合は入力したデータは反映されません。</DIV><P>
<% End If %>
       <INPUT type=hidden name="SakuNo"   value="<%=Request("SakuNo")%>">
       <INPUT type=hidden name="UpFlag"   value="<%=Request("UpFlag")%>">
       <INPUT type=hidden name="compFlag" value="<%=Request("compFlag")%>">
       <INPUT type=hidden name="Mord"     value="<%=Mord%>"><%'CW-028 ADD%>
       <INPUT type=hidden name="WkCNo"    value="<%=Request("WkCNo")%>" >
       <INPUT type=hidden name="partFlg"  value="<%=Request("partFlg")%>" >
       <INPUT type=hidden name="TruckerFlag" value="<%=Request("TruckerFlag")%>" >
       <INPUT type=hidden name="kariflag" value="<%=Request("kariflag")%>">					<!-- 2016/10/12 H.Yoshikawa Add -->
<% If Not ret Then %>
       <P><DIV class=alert>
       	<!-- 2016/11/03 H.Yoshikawa Upd Start 
        指定された会社コードは存在しません。<BR>
       「戻る」ボタンを押下し、再入力してください。-->
        <%=ErrMsg%>
        <!-- 2016/11/03 H.Yoshikawa Upd End -->
       </DIV></P>
<% Else %>
       <INPUT type=button value="ＯＫ" onClick="GoEntry()">
<% End If %>
       <INPUT type=button value="戻る" onClick="GoBackT()">
      </TD></TR>
<%end if%>
<!--2017/02/06 T.Okui Del End-->
</TABLE>
</FORM>
<!-------------画面終わり--------------------------->
</BODY></HTML>
