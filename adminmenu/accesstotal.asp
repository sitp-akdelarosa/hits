<%@LANGUAGE="VBSCRIPT" CODEPAGE="932"%>
<%
'**********************************************
'  【プログラムＩＤ】　: 
'  【プログラム名称】　: 
'
'  （変更履歴）
'	2010/01/28	C.Pestano	画面にｺﾒﾝﾄ追加
'**********************************************
	
	Option Explicit
	Response.Expires = 0

	call CheckLoginH()
%>
<!--#include File="./Common/Common.inc"-->
<SCRIPT src="./Common/function.js" type=text/javascript></SCRIPT>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<TITLE>利用件数表示</TITLE>
<link href="./Common/style.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">

//**************************************************
// 概要　 : 文字列中に数字以外が無いかチェックする
// 引数　 : ストリング
// 戻り値 : true(ない)/false(ある)
//**************************************************
function CheckSu(str){
  checkstr="0123456789";
  for (i=0; i<str.length; i++){
    c = str.charAt(i);
    if (checkstr.indexOf(c,0) < 0){
      return false;
    }
  }
  return true;
}

//**************************************************
// 概要　 : 閏年チェック
// 引数　 : YYYY
// 戻り値 : true(閏年)/false(平年)
//**************************************************
function isURU(year){
  if((year % 4 == 0 && year % 100 != 0) || year % 400 == 0)
    return true;
  else
    return false;
}
//**************************************************
// 概要　 : 日付の正当性チェック
// 引数　 : targetYYYY,targetMM,targetDD,Mode(1:年月日 2:年月)
// 戻り値 : true(正当)/false(不当)
//**************************************************
function CheckDate(targetYYYY,targetMM,targetDD,Mode){
  YYYY=targetYYYY.value;
  MM=targetMM.value;
  if (Mode==1){
    DD=targetDD.value;
  }else{
    DD=1;
  }
  //Nullチェック
  if(YYYY==null || YYYY==""){
    alert("年は必須入力項目です。");
    targetYYYY.focus();
    return false;
  }else if(MM==null || MM==""){
    alert("月は必須入力項目です。");
    targetMM.focus();
    return false;
  }else if(DD==null || DD==""){
    if (Mode==1){
       alert("日は必須入力項目です。");
       targetDD.focus();
       return false;
    }
  }
  
  //文字チェック
  if(!CheckSu(YYYY)){
     alert("半角数字以外の文字を入力しないでください");
     targetYYYY.focus();
     return false;
  }else if(!CheckSu(MM)){
     alert("半角数字以外の文字を入力しないでください");
     targetMM.focus();
     return false;
  }
  if(!CheckSu(DD)){
    if (Mode==1){
        alert("半角数字以外の文字を入力しないでください");
        targetDD.focus();
        return false;
    }
  }
  //期間チェック
  //月
  if(MM<1 || MM>12){
     alert("月は1～12の数字を入力してください");
     targetMM.focus();
     return false;
  }
  //日
  if(targetMM.value==2){  //2月ならば閏年チェックを行う
    if(isURU(YYYY)){
      //閏年
      MaxDay=29;
    } else {
      //平年
      MaxDay=28;
    }
  } else if(MM==4 || MM==6 || MM==9 || MM==11){
      MaxDay=30;
  } else {
      MaxDay=31;
  }
  if(DD<1 || DD>MaxDay){
     alert(YYYY+"年"+MM+"月なので、日は1～"+ MaxDay +"の数字を入力してください");
     targetDD.focus();
     return false;
  }
  return true;
}

//期間検索のチェック
function CheckList(){
	var FromDate;
	var ToDate;
	//入力チェック
	if(document.frm.rdoDetail[0].checked){
		//日付のチェックFrom
		if(CheckDate(document.frm.txtSYearFrom,document.frm.txtSMonthFrom,document.frm.txtSDayFrom,1)==true){
			FromDate=document.frm.txtSYearFrom.value;
			if(document.frm.txtSMonthFrom.value.length==1){
				FromDate=FromDate+"/0"+document.frm.txtSMonthFrom.value;
			}else{
				FromDate=FromDate+"/"+document.frm.txtSMonthFrom.value;
			}
			if(document.frm.txtSDayFrom.value.length==1){
				FromDate=FromDate+"/0"+document.frm.txtSDayFrom.value;	
			}else{
				FromDate=FromDate+"/"+document.frm.txtSDayFrom.value;
			}
		}else{
			return false;
		}
		//日付のチェックTo
		if(CheckDate(document.frm.txtSYearTo,document.frm.txtSMonthTo,document.frm.txtSDayTo,1)==true){
			ToDate=document.frm.txtSYearTo.value;
			if(document.frm.txtSMonthTo.value.length==1){
				ToDate=ToDate+"/0"+document.frm.txtSMonthTo.value;
			}else{
				ToDate=ToDate+"/"+document.frm.txtSMonthTo.value;
			}
			if(document.frm.txtSDayTo.value.length==1){
				ToDate=ToDate+"/0"+document.frm.txtSDayTo.value;
			}else{
				ToDate=ToDate+"/"+document.frm.txtSDayTo.value;
			}
		}else{
			return false;
		}
	}else{
		//日付のチェックFrom
		if(CheckDate(document.frm.txtSYearFrom,document.frm.txtSMonthFrom,"",2)==true){
			FromDate=document.frm.txtSYearFrom.value;
			if(document.frm.txtSMonthFrom.value.length==1){
				FromDate=FromDate+"/0"+document.frm.txtSMonthFrom.value;
			}else{
				FromDate=FromDate+"/"+document.frm.txtSMonthFrom.value;
			}
			FromDate=FromDate+"/01"
		}else{
			return false;
		}
		//日付のチェックTo
		if(CheckDate(document.frm.txtSYearTo,document.frm.txtSMonthTo,"",2)==true){
			ToDate=document.frm.txtSYearTo.value;
			if(document.frm.txtSMonthTo.value.length==1){
				ToDate=ToDate+"/0"+document.frm.txtSMonthTo.value;
			}else{
				ToDate=ToDate+"/"+document.frm.txtSMonthTo.value;
			}
			ToDate=ToDate+"/01"
		}else{
			return false;
		}
	}
	//日付の前後チェック
	if(FromDate > ToDate){
		alert("日付の範囲に誤りがあります。");
		document.frm.txtSYearFrom.focus();
		return false;
	}
return true;
}

//累計表のチェック
function CheckTotal(){
	var FromDate;
	var ToDate;
	//入力チェック
	//日付のチェックFrom
	if(CheckDate(document.frm.txtRYearFrom,document.frm.txtRMonthFrom,"",2)==true){
		FromDate=document.frm.txtRYearFrom.value;
		if(document.frm.txtRMonthFrom.value.length==1){
			FromDate=FromDate+"/0"+document.frm.txtRMonthFrom.value;
		}else{
			FromDate=FromDate+"/"+document.frm.txtRMonthFrom.value;
		}
	}else{
		return false;
	}
	//日付のチェックTo
	if(CheckDate(document.frm.txtRYearTo,document.frm.txtRMonthTo,"",2)==true){
		ToDate=document.frm.txtRYearTo.value;
		if(document.frm.txtRMonthTo.value.length==1){
			ToDate=ToDate+"/0"+document.frm.txtRMonthTo.value;
		}else{
			ToDate=ToDate+"/"+document.frm.txtRMonthTo.value;
		}
	}else{
		return false;
	}
	//日付の前後チェック
	if(FromDate > ToDate){
		alert("日付の範囲に誤りがあります。");
		document.frm.txtRYearFrom.focus();
		return false;
	}
return true;
}
// 期間検索を押下時
function fListSearch(){
	var FromDate;
	var ToDate;
	var Mode;

	ret=CheckList();
	if (ret==true){	
		//From日付作成
		FromDate=document.frm.txtSYearFrom.value;
		if(document.frm.txtSMonthFrom.value.length==1){
			FromDate=FromDate+"/0"+document.frm.txtSMonthFrom.value;
		}else{
			FromDate=FromDate+"/"+document.frm.txtSMonthFrom.value;
		}
		//Mode設定が日別の場合
		if(document.frm.rdoDetail[0].checked){
			if(document.frm.txtSDayFrom.value.length==1){
				FromDate=FromDate+"/0"+document.frm.txtSDayFrom.value;	
			}else{
				FromDate=FromDate+"/"+document.frm.txtSDayFrom.value;
			}
		}else{
			FromDate=FromDate+"/01"
		}
		//To日付作成
		ToDate=document.frm.txtSYearTo.value;
		if(document.frm.txtSMonthTo.value.length==1){
			ToDate=ToDate+"/0"+document.frm.txtSMonthTo.value;
		}else{
			ToDate=ToDate+"/"+document.frm.txtSMonthTo.value;
		}
		//Mode設定が日別の場合
		if(document.frm.rdoDetail[0].checked){
			if(document.frm.txtSDayTo.value.length==1){
				ToDate=ToDate+"/0"+document.frm.txtSDayTo.value;
			}else{
				ToDate=ToDate+"/"+document.frm.txtSDayTo.value;
			}
		}else{
			ToDate=ToDate+"/01"
		}
		//Mode設定
		if(document.frm.rdoDetail[0].checked){
			Mode="D"
		}else{
			Mode="M"
		}
		document.frm.action="./logview.asp?fDate="+FromDate+"&tDate="+ToDate+"&Mode="+Mode;
		document.frm.submit();
		return true;
	}
	return ;

}

// 累計表を押下時
function fTotalSearch(){
	var FromDate;
	var ToDate;

	ret=CheckTotal();
	if (ret==true){	
		//From日付作成
		FromDate=document.frm.txtRYearFrom.value;
		if(document.frm.txtRMonthFrom.value.length==1){
			FromDate=FromDate+"/0"+document.frm.txtRMonthFrom.value;
		}else{
			FromDate=FromDate+"/"+document.frm.txtRMonthFrom.value;
		}
		//To日付作成
		ToDate=document.frm.txtRYearTo.value;
		if(document.frm.txtRMonthTo.value.length==1){
			ToDate=ToDate+"/0"+document.frm.txtRMonthTo.value;
		}else{
			ToDate=ToDate+"/"+document.frm.txtRMonthTo.value;
		}
		document.frm.action="./logListview.asp?fDate="+FromDate+"&tDate="+ToDate;
		document.frm.submit();
		return true;
	}
	return false;

}
</script>

</script>
</HEAD>
<body class="bckcolor" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" >
<form name="frm" action="accesstotal.asp" method="post" enctype="multipart/form-data">
<!-------------ここからログイン入力画面--------------------------->
<table class="main2" align="center" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td valign=top>
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
		<%
			DisplayHeader2("利用件数表示")
		  %>
		</table>
		<center>
		<table class="square" border="0" cellspacing="4" cellpadding="0" >
			<tr>
			<td>
			<table border="0" cellspacing="3" cellpadding="4">
				<tr>
				<td>
				<table width="100%" border="0" cellspacing="2" cellpadding="3">
					<tr> 
					<td colspan="3" align=left valign=center nowrap>１．期間検索</td>
					</tr>
					<tr>
					<td width="50"></td>
					<td>
					<P>
					<INPUT name="txtSYearFrom" Type="Text" size="3" maxlength="4" style="WIDTH: 35px; HEIGHT: 20px" >&nbsp; 年 &nbsp;
					<INPUT name="txtSMonthFrom" Type="Text" size="1" maxlength="2" style="LEFT: 75px; WIDTH: 23px; TOP: 2px; HEIGHT: 21px">&nbsp; 月&nbsp; 
					<INPUT name="txtSDayFrom" size="1" maxlength="2" style="LEFT: 75px; WIDTH: 23px; TOP: 2px; HEIGHT: 21px">&nbsp; 日&nbsp;&nbsp;から&nbsp; 
					<INPUT name="txtSYearTo" size="3" maxlength="4" style="WIDTH: 35px; HEIGHT: 20px">&nbsp; 年&nbsp;
					<INPUT name="txtSMonthTo" size="1" maxlength="2" style="LEFT: 75px; WIDTH: 23px; TOP: 2px; HEIGHT: 21px">&nbsp;&nbsp;&nbsp;月&nbsp;
					<INPUT name="txtSDayTo" size="1" maxlength="2" style="LEFT: 75px; WIDTH: 23px; TOP: 2px; HEIGHT: 21px">&nbsp; 日 まで
					</P>
					<P><INPUT type=radio name="rdoDetail" value="1" checked>&nbsp; 日別明細</P>
					<P><INPUT type=radio name="rdoDetail" value="2">&nbsp; 月別明細</P>
					<P>※過去３年分のみ検索可能です。<BR>
					※画面上では指定期間のTOTAL件数のみ表示されます。<BR>&nbsp;&nbsp;&nbsp; 
					CSV出力では日別、または、月別の明細が表示されます。<BR>
					<!-- 2010/01/28 Add-S C.Pestano -->
					※合計欄には福岡市統計データを反映させています。<BR>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;（2000年11月～2009年10月）
					<!-- 2010/01/28 Add-E C.Pestano -->
					</P> 
					</td>
					<td valign="top">&nbsp;</td>
					</tr>
					<tr> 
					<td colspan="3" align=middle valign=center nowrap>
					<P>
					<INPUT style="WIDTH: 112px; HEIGHT: 29px" type="button" size="37" value="期間検索" name="btnSearch" Onclick="fListSearch();"></P>
					</td>
					</tr>
					<TR>
					<td>
					<P>２．累計表</P>
					</td>
					<tr> 
					<td width="50"></td>
					<td colspan="2" nowrap>
					<P>
					<INPUT name="txtRYearFrom" maxLength="4" size="3" style="WIDTH: 35px; HEIGHT: 20px"> 年&nbsp;
					<INPUT name="txtRMonthFrom" maxLength="2" size="1" style="LEFT: 75px; WIDTH: 23px; TOP: 2px; HEIGHT: 21px">&nbsp;月&nbsp;&nbsp;&nbsp;から&nbsp; 
					<INPUT name="txtRYearTo" maxLength="4" size="3" style="WIDTH: 35px; HEIGHT: 20px">年&nbsp; 
					<INPUT name="txtRMonthTo" maxLength="2" size="1" style="LEFT: 75px; WIDTH: 23px; TOP: 2px; HEIGHT: 21px">&nbsp;月&nbsp; まで
					</P>
					<P>※３年以内は月単位、４年以上は年単位で表示されます。<BR>
					※HiTSは2000年11月からのスタートです。<BR>
					※2000年11月～2009年10月の期間は、福岡市統計データを反映させています。<!-- 2010/01/28 Add C.Pestano -->
					</P> 
					</td>
					<td></td>
					</tr>
					<TR>
					<td colspan="3" align=middle valign=center nowrap>
					<INPUT style="WIDTH: 112px; HEIGHT: 29px" type="button" size="37" value="累計表" name="btnRuikei" Onclick="fTotalSearch();">
					<P></P>
					</td>
					<td></td></TR>
					</table>
					<center>
					<A href="menu.asp">閉じる</A>
					</center></td>
					</tr>
					</table>
				</td>
				</tr>
			</table></center>
			</td>
			</tr>
		<%
			DisplayFooter
		%>
		</table>
		</td>
	</Tr>
</form>
</body>
</HTML>
