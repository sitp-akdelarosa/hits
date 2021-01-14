<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi220.asp				_/
'_/	Function	:事前空搬出入力画面			_/
'_/	Date		:2003/05/28				_/
'_/	Code By		:SEIKO Electric.Co 大重			_/
'_/	Modify		:C-002	2003/08/06	備考欄追加	_/
'_/	Modify		:3th	2003/01/31	3次全面改修	_/
'_/	Modify		:2017/02/22 ピック予定日時に前日以前を入力禁止に変更	_/
'_/	Modify		:2017/05/09 属性と本数の入力欄を１０行に増加、行削除追加　など	_/
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
'サーバ日付の取得
 dim DayTime
 getDayTime DayTime
'データ所得
  dim BookNo, COMPcd0, COMPcd1, Mord, TFlag
  dim Dflag,plintStr,i
  dim WkOutFlag, OutStyle				'2016/08/22 H.Yoshikawa Add
  dim Dflag2							'2016/08/22 H.Yoshikawa Add
  BookNo  = Request("BookNo")
  COMPcd0 = Request("COMPcd0")
  COMPcd1 = Request("COMPcd1")
  Mord    = Request("Mord")
  Dflag=""
  Dflag2=""								'2016/08/22 H.Yoshikawa Add
  plintStr=""

  Const RowNum = 10						'2017/05/09 H.Yoshikawa Add
  
  If Mord=0 Then '新規登録時
  
  Else          '更新時
    WriteLogH "b302", "空搬出事前情報入力","12",""
    TFlag   = Request("TFlag")
'Chang 20050303 STAT fro Fourth Recon By SEIKO N.Oosige
'    If COMPcd0 <> UCase(Session.Contents("userid")) OR TFlag = 1 Then
    If COMPcd0 <> UCase(Session.Contents("userid")) OR TFlag = "1" OR Request("compFlag")<>"0" Then
      Dflag="readOnly"
    End If
    plintStr="(更新モード)"
  End If

'2016/10/26 H.Yoshikawa Add Start
'DB接続
  dim ObjConn, ObjRS, StrSQL
  ConnDBH ObjConn, ObjRS

  dim OdrNum, OutNum, RsvNum
  dim DflagZokusei
'2016/10/26 H.Yoshikawa Add End

  '2017/08/04 H.Yoshikawa Add Start キャッシュ対策
  dim sysdate
  sysdate = Replace(Replace(Replace(Now(), "/", ""), ":", ""), " ", "")
  '2017/08/04 H.Yoshikawa Add End

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>空バンピック情報入力更新</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<!-- 2017/08/04 H.Yoshikawa Upd Start -->
<!-- <SCRIPT src="./JS/Common.js"></SCRIPT> -->
<!-- <SCRIPT src="./JS/CommonSub.js"></SCRIPT> -->
<SCRIPT src="./JS/Common.js?ver=<%=sysdate%>"></SCRIPT>
<SCRIPT src="./JS/CommonSub.js?ver=<%=sysdate%>"></SCRIPT>
<!-- 2017/08/04 H.Yoshikawa Upd End -->
<SCRIPT language=JavaScript>
<!--
function setParam(target){
//2016/08/22 H.Yoshikawa Upd Start
  //window.resizeTo(550,680);
  window.moveTo(120,20);
  window.resizeTo(1366,768);			// 2017/05/09 H.Yoshikawa Upd(770⇒820) // edited by AK.DELAROSA 2021-01-14
//2016/08/22 H.Yoshikawa Upd End

//2016/11/09 H.Yoshikawa Add Start
  target=document.dmi220F;
<%
'コンボボックスデータ取得
'コンテナタイプ取得＆表示
  StrSQL = "select * from mContType ORDER BY ContType ASC"
  ObjRS.Open StrSQL, ObjConn
  Response.Write "  list = new Array(''"
  Do Until ObjRS.EOF
    Response.Write ",'" & ObjRS("ContType") & "'"
    ObjRS.MoveNext
  Loop 
  Response.Write ");" & vbCrLf
  ObjRS.Close
  for i = 0 to RowNum-1					'2017/05/09 H.Yoshikawa Upd(4⇒RowNum-1)
    Response.Write "  setList(target.elements[""ContTypeSel" & i & """],list,'" & Request("ContType"&i) & "');" & vbCrLf
    Response.Write " UpdFlagChg(" & i & ");"
  next
%>

//2016/11/09 H.Yoshikawa Add End

  bgset(target);
}
//更新
function GoNext(){
  target=document.dmi220F;
  if(!check(target))
    return;
  chengeUpper(target);
  target.action="./dmi230.asp";
  target.submit();
}
//削除
function GoDell(){
//2017/05/10 H.Yoshikawa Add Start
  target=document.dmi220F;
  for(idx=0;idx<<%=RowNum%>;idx++){
    if(Number(target.elements["OutNum" + idx].value) > 0){
	    alert("搬出済みのコンテナが存在するため、削除できません。");
	    return false;
  	}
  }
//2017/05/10 H.Yoshikawa Add End

<%If TFlag<>"1" Then%>
  flag = confirm('削除しますか？');
<%Else%>
  flag = confirm('指示先が受諾回答済です。\n削除する前に指示先に確認してください。\n削除しますか？');
<%End If%>
  if(flag){
    target=document.dmi220F;
    target.action="./dmi290.asp";
    len = target.elements.length;
    for (i=0; i<len; i++) target.elements[i].disabled = false;
    target.submit();
  }
}
//保留
function Suspend(){
  target=document.dmi220F;
  if(target.way[1].checked){
    flag = confirm('回答をNoにしますか？');
    if(!flag) return false;
    target.Res.value=2;
  }
  target.action="./dmi230.asp";
  target.submit();
}
//ブッキング情報
function GoBookI(){
  target=document.dmi220F;
  BookInfo(target);
}

//入力情報チェック
function check(target){
  //2016/11/09 H.Yoshikawa Add Start
  for(idx=0;idx<<%=RowNum%>;idx++){			//2017/05/09 H.Yoshikawa Upd(5⇒<%=RowNum%>)
	if(target.elements["ContTypeSel" + idx].disabled == true){
		target.elements["ContType" + idx].value = target.elements["Bef_ContType" + idx].value;
	}else{
		target.elements["ContType" + idx].value = target.elements["ContTypeSel" + idx].options[target.elements["ContTypeSel" + idx].selectedIndex].value;
	}
  }
//2016/11/09 H.Yoshikawa Add End

  //2016.08.26 H.Yoshikawa Add Start
  //必須チェック
  if(target.shipName.value.length==0 || target.VoyCtrl.value.length==0 || target.VslCode.value.length==0 || target.ExVoyage.value.length==0){
    alert("船名、次航が正しくありません。検索画面よりセットしてください。");
    target.shipName.focus();
    return false;
  }

  strA    = new Array();
//2016.10.11 H.Yoshikawa Upd Start
  //strA[0] = target.COMPcd1;
  //strA[1] = target.TruckerSubName;
  //strA[2] = target.Tel;
  //strA[3] = target.Mail;
  //strM    = new Array("会社コード","登録担当者","電話番号","メールアドレス");
  strA[0] = target.TruckerSubName;
  strA[1] = target.Tel;
  strA[2] = target.Mail;
  strM    = new Array("登録担当者","電話番号","メールアドレス");
//2016.10.11 H.Yoshikawa Upd End
  for(k=0;k<strA.length;k++){
    Num=LTrim(strA[k].value);
    if(Num.length==0){
      alert(strM[k]+"を記入してください");
      strA[k].focus();
      return false;
    }
  }
  //2016.08.26 H.Yoshikawa Add End

  if(!CheckEisu2(target.COMPcd1.value)){
    alert("会社コードに半角英数字以外の文字を記入しないでください");
    target.COMPcd1.focus();
    return;
  }
  
  //2016.08.25 H.Yoshikawa Del Start
  //strA    = new Array();
  //strA[0] = target.ContSize0;
  //strA[1] = target.ContSize1;
  //strA[2] = target.ContSize2;
  //strA[3] = target.ContSize3;
  //strA[4] = target.ContSize4;
  //strA[5] = target.ContHeight0;
  //strA[6] = target.ContHeight1;
  //strA[7] = target.ContHeight2;
  //strA[8] = target.ContHeight3;
  //strA[9] = target.ContHeight4;
  //strA[10]= target.PickNum0;
  //strA[11]= target.PickNum1;
  //strA[12]= target.PickNum2;
  //strA[13]= target.PickNum3;
  //strA[14]= target.PickNum4;
  //strA[15]= target.vanMin;
  //for(k=0;k<16;k++){
  //  if(strA[k].value!="" && strA[k].value!=null){
  //    ret = CheckSu(strA[k].value); 
  //    if(ret==false){
  //      alert("数字以外を入力しないでください。");
  //      strA[k].focus();
  //      return false;
  //    }
  //  }
  //}
  //strA    = new Array();
  //strA[0] = target.ContType0;
  //strA[1] = target.ContType1;
  //strA[2] = target.ContType2;
  //strA[3] = target.ContType3;
  //strA[4] = target.ContType4;
  //strA[5] = target.Material0;
  //strA[6] = target.Material1;
  //strA[7] = target.Material2;
  //strA[8] = target.Material3;
  //strA[9] = target.Material4;
  //for(k=0;k<10;k++){
  //if(strA[k].value!="" && strA[k].value!=null){
  //    ret = CheckEisu2(strA[k].value); 
  //    if(ret==false){
  //      alert("半角英数字以外の文字を入力しないでください");
  //      strA[k].focus();
  //      return false;
  //    }
  //  }
  //}
  //2016.08.25 H.Yoshikawa Del End

  //2016.08.25 H.Yoshikawa Add Start
  //属性、本数のチェック（変更チェック有の場合のみ）
  var today = new Date();											//2016/12/07 H.Yoshikawa Add
  var RFtodayFlg = false;											//2016/12/07 H.Yoshikawa Add
  for(idx=0;idx<<%=RowNum%>;idx++){									//2017/05/09 H.Yoshikawa Upd(5⇒<%=RowNum%>)
	if(target.elements["UpdFlag" + idx].checked == true){
		//必須チェック
		strA    = new Array();
		strA[0] = target.elements["ContSize" + idx];
		strA[1] = target.elements["ContType" + idx];
		strA[2] = target.elements["ContHeight" + idx];
		strA[3] = target.elements["PickDate" + idx];
		strA[4] = target.elements["PickNum" + idx];
		strM    = new Array("サイズ","タイプ","高さ","ピック予定日", "本数");
		for(k=0;k<strA.length;k++){
			Num=LTrim(strA[k].value);
			if(Num.length==0){
			  alert(strM[k]+"を記入してください");
			  strA[k].focus();
			  return false;
			}
		}
		
		//2017/08/25 H.Yoshikawa Add Start (RFのときはPcool必須)
		if(target.elements["ContType" + idx].value == "RF"){
			if(target.elements["Pcool" + idx].value == "0" || target.elements["Pcool" + idx].value == ""){
			  alert("プレクールを選択してください");
			  target.elements["Pcool" + idx].focus();
			  return false;
			}
		}
		//2017/08/25 H.Yoshikawa Add End
		
		if(target.elements["Pcool" + idx].value == "1"){
			Num=LTrim(target.elements["PickHour" + idx].value);
			if(Num.length==0){
			  alert("ピック予定時を記入してください");
			  target.elements["PickHour" + idx].focus();
			  return false;
			}
			Num=LTrim(target.elements["PickMinute" + idx].value);
			if(Num.length==0){
			  alert("ピック予定分を記入してください");
			  target.elements["PickMinute" + idx].focus();
			  return false;
			}
		}
	
		//数値チェック
		strA    = new Array();
		strA[0] = target.elements["ContSize" + idx];
		strA[1] = target.elements["ContHeight" + idx];
		strA[2] = target.elements["Ventilation" + idx];
		strA[3] = target.elements["PickHour" + idx];
		strA[4] = target.elements["PickMinute" + idx];
		strA[5] = target.elements["PickNum" + idx];
		for(k=0;k<strA.length;k++){
			if(strA[k].value!="" && strA[k].value!=null){
			  ret = CheckSu(strA[k].value); 
			  if(ret==false){
			    alert("数字以外を入力しないでください。");
			    strA[k].focus();
			    return false;
			  }
			}
		}

		// 2016/11/15 H.Yoshikawa Add Start
		//本数チェック（１本以上）
		if(Number(target.elements["PickNum" + idx].value) <= 0){
			alert("本数は1本以上で指定してください。");
			target.elements["PickNum" + idx].focus();
			return false;
		}
		// 2016/11/15 H.Yoshikawa Add End

		//英数字チェック
		ret = CheckEisu2(target.elements["ContType" + idx].value); 
		if(ret==false){
			alert("英数字以外を入力しないでください。");
			target.elements["ContType" + idx].focus();
			return false;
		}

		//日付チェック
		ret = CheckYMD(target.elements["PickDate" + idx]); 
		if(ret==false){
			alert("日付が正しくありません。");
			target.elements["PickDate" + idx].focus();
			return false;
		}
		
		//20170222 T.Okui Add S
		//過去分の修正を行えないようにする
		
		//if((Rtrim(target.elements["ContSize" + idx].value, ' ') == "")||(Rtrim(target.elements["ContSize" + idx].value, ' ') != "" && target.elements["PickNum" + idx].value==target.elements["Bef_PickNum" + idx].value)){
		if(target.elements["PickDate" + idx].value!=target.elements["Bef_PickDate" + idx].value){			//2017/05/09 H.Yoshikawa Add
			var tmpDate = target.elements["PickDate" + idx].value;
			// 現在の日付＆時刻を取得
			var today = new Date();
			// 時間を0:00にする
			today.setHours(0);
			today.setMinutes(0);
			today.setSeconds(0);
			today.setMilliseconds(0);
			
			// 文字列から年月日を抜き出し、数値型に変換
			var vYear = parseInt( tmpDate.substr( 0, 4  ),10);
			var vMonth = parseInt( tmpDate.substr( 5, 2 ),10 ) -1;
			var vDay = parseInt( tmpDate.substr( 8, 2 ),10 );
			var adate = new Date( vYear, vMonth, vDay );

			if( adate.getTime() < today.getTime() ){
			//前日以前
				alert("ピック予定日は本日以降である必要があります。");
				target.elements["PickDate" + idx].focus();
				return false;
			}
		}		//2017/05/09 H.Yoshikawa Add
		//}
		//20170222 T.Okui Add S	
		
		//2017/08/25 H.Yoshikawa Add Start (日祝日は入力不可)
		if(ktHolidayName(target.elements["PickDate" + idx].value) != ""){
				alert("ピック予定日に日祝日は入力できません。");
				target.elements["PickDate" + idx].focus();
				return false;
		}
		//2017/08/25 H.Yoshikawa Add End

		//時間チェック
		if(target.elements["PickHour" + idx].value>23){
	      alert("時は0～23で入力してください");
	      target.elements["PickHour" + idx].focus();
	      return false;
	    }
		//分チェック
		if(target.elements["PickMinute" + idx].value>59){
	      alert("分は0～59で入力してください");
	      target.elements["PickMinute" + idx].focus();
	      return false;
	    }
	    
		//2016/12/07 H.Yoshikawa Add Start
		//RFで当日の予約があるか？
		var inputDay = new Date(target.elements["PickDate" + idx].value);
		if(target.elements["ContType" + idx].value == "RF" 
		&& inputDay.getFullYear() == today.getFullYear()
		&& inputDay.getMonth() == today.getMonth()
		&& inputDay.getDate() == today.getDate()){
			RFtodayFlg = true;
		}
		//2016/12/07 H.Yoshikawa Add End
		
		//2017/06/22 H.Yoshikawa Add Start
		//設定温度チェック
		str = target.elements["SetTemp" + idx].value;
		if(!CheckSu2(str, "+-.")){
			alert("設定温度は半角数字または、+、-、.のみで入力してください。");
			target.elements["SetTemp" + idx].focus();
			return false;
		}
		//2017/06/22 H.Yoshikawa Add End
 	}
 	//2017/05/10 H.Yoshikawa Add Start
	if(target.elements["DelFlag" + idx].checked == true){
	    if(Number(target.elements["OutNum" + idx].value) > 0){
		    alert("搬出済みのコンテナが存在するため、行削除できません。");
		    target.elements["DelFlag" + idx].focus();
		    return false;
	  	}
	}
 	//2017/05/10 H.Yoshikawa Add End
  }

  //バン詰め日時：分 数値チェック
  ret = CheckSu(target.vanMin.value); 
  if(ret==false){
    alert("数字以外を入力しないでください。");
    target.vanMin.focus();
    return false;
  }
  //2016.08.25 H.Yoshikawa Add End


//日付のチェック
  if(!CheckDate('<%=DayTime(0)%>','<%=DayTime(1)%>',target.vanMon,target.vanDay,target.vanHou)){
    return false;
  }else{
    if(target.vanHou.value=="")
      target.vanMin.value="";
    if(target.vanMin.value>59){
      alert("分は0～59で入力してください");
      target.vanMin.focus();
      return false;
    }
  }
  NumA    = new Array();
  //2016.08.26 H.Yoshikawa Del Start
  //strA[0] = target.PickPlace0;	NumA[0]=20;
  //strA[1] = target.PickPlace1;	NumA[1]=20;
  //strA[2] = target.PickPlace2;	NumA[2]=20;
  //strA[3] = target.PickPlace3;	NumA[3]=20;
  //strA[4] = target.PickPlace4;	NumA[4]=20;
  strA[5] = target.vanPlace1;	NumA[5]=70;
  strA[6] = target.vanPlace2;	NumA[6]=70;
  strA[7] = target.goodsName;	NumA[7]=20;
  strA[8] = target.Comment1;	NumA[8]=70;
  strA[9] = target.Comment2;	NumA[9]=70;
  strA[10] = target.TruckerSubName;	NumA[10]=16;
  //for(k=0;k<11;k++){
  for(k=5;k<11;k++){
  //2016.08.26 H.Yoshikawa Del End
    if(strA[k].value!="" && strA[k].value!=null){
      ret = CheckKin(strA[k].value); 
      if(ret==false){
        alert("「\"」や「\'」等の半角記号を入力しないでください。");
        strA[k].focus();
        return false;
      }
      retA=getByte(strA[k].value);
      if(retA[0]>NumA[k]){
        if(retA[2]>(NumA[k]/2)){
          alertStr="全角文字を"+(NumA[k]/2)+"文字以内で入力してください。";
        }else{
          alertStr="全角文字を"+Math.floor((NumA[k]-retA[1])/2)+"文字にするか\n";
          alertStr=alertStr+"半角文字を"+(NumA[k]-retA[2]*2)+"文字にしてください。";
        }
        alert(NumA[k]+"バイト以内で入力してください。\n"+NumA[k]+"バイト以内にするには"+alertStr);
        strA[k].focus();
        return false;
      }
    }
  }
  /* 2009/09/27 C.Pestano Del-S
   ret = CheckKana(target.TruckerSubName.value); 
   if(ret==false){
     alert("半角カナ文字は入力できません");
     target.TruckerSubName.focus();
     return false;
   }2009/09/27 C.Pestano Del-E
   */
   
   //2016.08.26 H.Yoshikawa Add Start
   ret = CheckMail(target.Mail.value); 
   if(ret==false){
     alert("メールアドレスが正しくありません。");
     target.Mail.focus();
     return false;
   }
   
   ret = CheckTel(target.Tel.value); 
   if(ret==false){
     alert("電話番号が正しくありません。");
     target.Tel.focus();
     return false;
   }
   //2016.08.26 H.Yoshikawa Add End
   
   //2016/12/07 H.Yoshikawa Add Start
   if(RFtodayFlg == true){
     retValue = showModalDialog("dmlModalRFToday.asp", window, "dialogWidth:500px; dialogHeight:200px; center:1; scroll: no; dialogTop:300px; ");
     if(retValue != true){
       return false;
     }
   }
   //2016/12/07 H.Yoshikawa Add Start

  return true;
}
//2008-01-31 Add-S M.Marquez
function finit(){
    document.dmi220F.shipName.focus();					//2016.08.22 H.Yoshikawa Upd (COMPcd1→shipName)
}
//2008-01-31 Add-E M.Marquez

function CheckKana(str){
  checkstr="｡｢｣､･ｦｧｨｩｪｫｬｭｮｯｰｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜﾝﾞﾟ";
   for (i=0; i<str.length; i++){
    c = str.charAt(i);
    if (checkstr.indexOf(c,0) >= 0){
      return false;
    }
  }
  return true;
}
//2009/07/27 Add-S C.Pestano
function CheckLen(obj,mesgon,focuson,mandatory) {
	var kanjicheck = gfStrLen(obj.value);
	
	if (kanjicheck == false){
		alert("半角文字を入力してください。");
		obj.focus();
		return false;
	}	
	
	if (mandatory && objlength==0)
		return false;	
	return true;
}

function gfStrLen(StrSrc) {
	var r = 0;
	for (var i = 0; i < StrSrc.length; i++) {
		var c = StrSrc.charCodeAt(i);
		// Shift_JIS: 0x0 ～ 0x80, 0xa0  , 0xa1   ～ 0xdf  , 0xfd   ～ 0xff
		// Unicode  : 0x0 ～ 0x80, 0xf8f0, 0xff61 ～ 0xff9f, 0xf8f1 ～ 0xf8f3
		if ( (c >= 0x0 && c < 0x81) || (c == 0xf8f0) || (c >= 0xff61 && c < 0xffa0) || (c >= 0xf8f1 && c < 0xf8f4)) {
			
		} else {			
			return false;		
		}
	}
	return true;
}
//2009/07/27 Add-E C.Pestano

//2016/08/23 H.Yoshikawa Add Start
//船名・次航の検索画面表示
function VslSelect(){
	var winname="searchVsl";
	var target=document.dmi220F;
	var vslnm = target.shipName.value;
  	var retValue = window.open("./dmlModalVslVoy.asp?tgt=dmi220F&fldvn=shipName&fldvc=VslCode&fldvy=VoyCtrl&flddspvy=ExVoyage&dspkbn=LD", winname, "width=600, height=700, menubar=no, toolbar=no, scrollbars=yes");
  	return true;
}

//2017/05/10 H.Yoshikawa Add Start
//属性行削除設定
function DelFlagChg(idx){
  	var target;
	target=document.dmi220F;
	if(target.elements["DelFlag" + idx].checked == true){
		target.elements["UpdFlag" + idx].checked = false;
		UpdFlagChg(idx);
	}
}
//2017/05/10 H.Yoshikawa Add End

//属性変更可否設定
function UpdFlagChg(idx){
  var target;
	target=document.dmi220F;
	
	if(target.elements["UpdFlag" + idx].checked == true){
		target.elements["DelFlag" + idx].checked = false;					//2017/05/10 H.Yoshikawa Add
		if(target.COMPcd1.readOnly == true && Rtrim(target.elements["Bef_ContSize" + idx].value, ' ') != ""){
			//属性、本数以外は変更不可
			//変更された値をもとに戻す
			target.elements["SetTemp" + idx].value = target.elements["Bef_SetTemp" + idx].value;
			target.elements["Pcool" + idx].value = target.elements["Bef_Pcool" + idx].value;
			target.elements["Ventilation" + idx].value = target.elements["Bef_Ventilation" + idx].value;
			//target.elements["PickDate" + idx].value = target.elements["Bef_PickDate" + idx].value;			// 2016/11/11 H.Yoshikawa Del
			//target.elements["PickHour" + idx].value = target.elements["Bef_PickHour" + idx].value;			// 2016/11/11 H.Yoshikawa Del
			//target.elements["PickMinute" + idx].value = target.elements["Bef_PickMinute" + idx].value;		// 2016/11/11 H.Yoshikawa Del
			target.elements["PickPlace" + idx].value = target.elements["Bef_PickPlace" + idx].value;
			target.elements["Terminal" + idx].value = target.elements["Bef_Terminal" + idx].value;
			target.elements["SetTemp" + idx].readOnly  = true;
			target.elements["Pcool" + idx].disabled  = true;
			target.elements["Ventilation" + idx].readOnly  = true;
			//target.elements["PickDate" + idx].readOnly  = true;			// 2016/11/11 H.Yoshikawa Del
			//target.elements["PickHour" + idx].readOnly  = true;			// 2016/11/11 H.Yoshikawa Del
			//target.elements["PickMinute" + idx].readOnly  = true;			// 2016/11/11 H.Yoshikawa Del
			
			//属性は同一属性の搬出済みが存在する場合は不可（ピック予定日時も…2017/06/20 H.Yoshikawa Upd）
			if(Number(target.elements["OutNum" + idx].value) > 0){
				//変更された値をもとに戻す
				target.elements["ContSize" + idx].value = target.elements["Bef_ContSize" + idx].value;
				//2016/11/09 Yoshikawa Upd Start
				//target.elements["ContType" + idx].value = target.elements["Bef_ContType" + idx].value;
				pulldown_option = target.elements["ContTypeSel" + idx].getElementsByTagName('option');
				for(i=0; i<pulldown_option.length;i++){
					if(pulldown_option[i].value == target.elements["Bef_ContType" + idx].value){
						pulldown_option[i].selected = true;
					break;
					}
				}
				target.elements["ContHeight" + idx].value = target.elements["Bef_ContHeight" + idx].value;
				target.elements["ContSize" + idx].readOnly  = true;
				//target.elements["ContType" + idx].readOnly  = true;			//2016/11/09 Del Yoshikawa
				target.elements["ContTypeSel" + idx].disabled  = true;				//2016/11/09 Add Yoshikawa
				target.elements["ContHeight" + idx].readOnly  = true;

				//2017/06/20 H.Yoshikawa Add Start
				target.elements["PickDate" + idx].value = target.elements["Bef_PickDate" + idx].value;
				target.elements["PickHour" + idx].value = target.elements["Bef_PickHour" + idx].value;
				target.elements["PickMinute" + idx].value = target.elements["Bef_PickMinute" + idx].value;
				target.elements["PickDate" + idx].readOnly  = true;
				target.elements["PickHour" + idx].readOnly  = true;
				target.elements["PickMinute" + idx].readOnly  = true;
				//2017/06/20 H.Yoshikawa Add End
				
			}else{
				target.elements["ContSize" + idx].readOnly  = false;
				//target.elements["ContType" + idx].readOnly  = false;			//2016/11/09 Del Yoshikawa
				target.elements["ContTypeSel" + idx].disabled  = false;			//2016/11/09 Add Yoshikawa
				target.elements["ContHeight" + idx].readOnly  = false;

				//2017/06/20 H.Yoshikawa Add Start
				target.elements["PickDate" + idx].readOnly  = false;
				target.elements["PickHour" + idx].readOnly  = false;
				target.elements["PickMinute" + idx].readOnly  = false;
				//2017/06/20 H.Yoshikawa Add End

			}
			
			//本数は変更可
			target.elements["PickNum" + idx].readOnly  =false;
			
			//2017/06/20 H.Yoshikawa Del Start
			//ピック予定日も変更可
			//target.elements["PickDate" + idx].readOnly  = false;				// 2016/11/11 H.Yoshikawa Add
			//target.elements["PickHour" + idx].readOnly  = false;				// 2016/11/11 H.Yoshikawa Add
			//target.elements["PickMinute" + idx].readOnly  = false;				// 2016/11/11 H.Yoshikawa Add
			//2017/06/20 H.Yoshikawa Del End
		}else{
			//すべて変更可
			target.elements["ContSize" + idx].readOnly  = false;
			//target.elements["ContType" + idx].readOnly  = false;			//2016/11/09 Del Yoshikawa
			target.elements["ContTypeSel" + idx].disabled  = false;			//2016/11/09 Add Yoshikawa
			target.elements["ContHeight" + idx].readOnly  = false;
			target.elements["SetTemp" + idx].readOnly  = false;
			target.elements["Pcool" + idx].disabled  = false;
			target.elements["Ventilation" + idx].readOnly  = false;
			target.elements["PickDate" + idx].readOnly  = false;
			target.elements["PickHour" + idx].readOnly  = false;
			target.elements["PickMinute" + idx].readOnly  = false;
			target.elements["PickNum" + idx].readOnly  =false;
		}
	}else{
		//変更された値をもとに戻す
		target.elements["ContSize" + idx].value = target.elements["Bef_ContSize" + idx].value;
		//2016/11/09 Yoshikawa Upd Start
		//target.elements["ContType" + idx].value = target.elements["Bef_ContType" + idx].value;
		pulldown_option = target.elements["ContTypeSel" + idx].getElementsByTagName('option');
		for(i=0; i<pulldown_option.length;i++){
			if(pulldown_option[i].value == target.elements["Bef_ContType" + idx].value){
				pulldown_option[i].selected = true;
			break;
			}
		}
		//2016/11/09 Yoshikawa Upd End
		target.elements["ContHeight" + idx].value = target.elements["Bef_ContHeight" + idx].value;
		target.elements["SetTemp" + idx].value = target.elements["Bef_SetTemp" + idx].value;
		target.elements["Pcool" + idx].value = target.elements["Bef_Pcool" + idx].value;
		target.elements["Ventilation" + idx].value = target.elements["Bef_Ventilation" + idx].value;
		target.elements["PickDate" + idx].value = target.elements["Bef_PickDate" + idx].value;
		target.elements["PickHour" + idx].value = target.elements["Bef_PickHour" + idx].value;
		target.elements["PickMinute" + idx].value = target.elements["Bef_PickMinute" + idx].value;
		target.elements["PickNum" + idx].value = target.elements["Bef_PickNum" + idx].value;
		target.elements["PickPlace" + idx].value = target.elements["Bef_PickPlace" + idx].value;
		target.elements["Terminal" + idx].value = target.elements["Bef_Terminal" + idx].value;
		//すべて変更不可
		target.elements["ContSize" + idx].readOnly  = true;
		//target.elements["ContType" + idx].readOnly  = true;				//2016/11/09 Yoshikawa Del
		target.elements["ContTypeSel" + idx].disabled  = true;					//2016/11/09 Yoshikawa Add
		target.elements["ContHeight" + idx].readOnly  = true;
		target.elements["SetTemp" + idx].readOnly  = true;
		target.elements["Pcool" + idx].disabled = true;
		target.elements["Ventilation" + idx].readOnly  = true;
		target.elements["PickDate" + idx].readOnly  = true;
		target.elements["PickHour" + idx].readOnly  = true;
		target.elements["PickMinute" + idx].readOnly  = true;
		target.elements["PickNum" + idx].readOnly  = true;
	}
	bgset(target);
}
//2016/08/23 H.Yoshikawa Add End

//2017/08/25 H.Yoshikawa Add Start（日祝日判定追加）
var SUNDAY = 0;
var MONDAY = 1;
var TUESDAY = 2;
var WEDNESDAY = 3;

// JavaScriptで扱える日付は1970/1/1～のみ
//var cstImplementTheLawOfHoliday = new Date("1948/7/20");  // 祝日法施行
//var cstAkihitoKekkon = new Date("1959/4/10");              // 明仁親王の結婚の儀
var cstShowaTaiso = new Date("1989/2/24");                // 昭和天皇大喪の礼
var cstNorihitoKekkon = new Date("1993/6/9");            // 徳仁親王の結婚の儀
var cstSokuireiseiden = new Date("1990/11/12");          // 即位礼正殿の儀
var cstImplementHoliday = new Date("1973/4/12");        // 振替休日施行

// [prmDate]には "yyyy/m/d"形式の日付文字列を渡す
function ktHolidayName(prmDate)
{
  var MyDate = new Date(prmDate);
  var HolidayName = prvHolidayChk(MyDate);
  var YesterDay;
  var HolidayName_ret;

  if (HolidayName == "") {
      if (MyDate.getDay() == MONDAY) {
          // 月曜以外は振替休日判定不要
          // 5/6(火,水)の判定はprvHolidayChkで処理済
          // 5/6(月)はここで判定する
          if (MyDate.getTime() >= cstImplementHoliday.getTime()) {
              YesterDay = new Date(MyDate.getFullYear(),
                                     MyDate.getMonth(),(MyDate.getDate()-1));
              HolidayName = prvHolidayChk(YesterDay);
              if (HolidayName != "") {
                  HolidayName_ret = "振替休日";
              } else {
                  HolidayName_ret = "";
              }
          } else {
              HolidayName_ret = "";
          }
      } else {
          if (MyDate.getDay() == SUNDAY) {
              HolidayName_ret = "日曜";
          }else{
              HolidayName_ret = "";
          }
      }
  } else {
      HolidayName_ret = HolidayName;
  }

  return HolidayName_ret;
}


//===============================================================

function prvHolidayChk(MyDate)
{
  var MyYear = MyDate.getFullYear();
  var MyMonth = MyDate.getMonth() + 1;    // MyMonth:1～12
  var MyDay = MyDate.getDate();
  var Result = "";
  var NumberOfWeek;
  var MyAutumnEquinox;

// JavaScriptで扱える日付は1970/1/1～のみで祝日法施行後なので下記は不要
// if (MyDate.getTime() < cstImplementTheLawOfHoliday.getTime()) {
// 　　return ""; // 祝日法施行(1948/7/20)以前
// } else;

  switch (MyMonth) {
// １月 //
  case 1:
      if (MyDay == 1) {
          Result = "元日";
      } else {
          if (MyYear >= 2000) {
              NumberOfWeek = Math.floor((MyDay - 1) / 7) + 1;
              if ((NumberOfWeek == 2) && (MyDate.getDay() == MONDAY)) {
                  Result = "成人の日";
              } else;
          } else {
              if (MyDay == 15) {
                  Result = "成人の日";
              } else;
          }
      }
      break;
// ２月 //
  case 2:
      if (MyDay == 11) {
          if (MyYear >= 1967) {
              Result = "建国記念の日";
          } else;
      //2019/03/25 Add-S Tanaka 天皇即位対応
      } else if(MyDay==23) {
          if (MyYear >= 2020) {
              Result = "天皇誕生日";
          } else;
      //2019/03/25 Add-E Tanaka
      } else {
          if (MyDate.getTime() == cstShowaTaiso.getTime()) {
              Result = "昭和天皇の大喪の礼";
          } else;
      }
      break;
// ３月 //
  case 3:
      if (MyDay == prvDayOfSpringEquinox(MyYear)) {  // 1948～2150以外は[99]
          Result = "春分の日";                       // が返るので､必ず≠になる
      } else;
      break;
// ４月 //
  case 4:
      if (MyDay == 29) {
          if (MyYear >= 2007) {
              Result = "昭和の日";
          } else {
              if (MyYear >= 1989) {
                  Result = "みどりの日";
              } else {
                Result = "天皇誕生日";
              }
          }
      } else {
          // JavaScriptで扱える日付は1970/1/1～のみなので下記は不要
          // if (MyDate.getTime() == cstAkihitoKekkon.getTime()) {
          // 　　Result = "皇太子明仁親王の結婚の儀";　　// (=1959/4/10)
          // } else;
          //2019/03/25 Add-S Tanaka 天皇即位対応
          if (MyYear == 2019 && MyDay==30) {
              Result = "平成天皇退位";
          } else;
          //2019/03/25 Add-E Tanaka
      }
      break;
// ５月 //
  case 5:
      switch ( MyDay ) {
        //2019/03/25 Add-S Tanaka 天皇即位対応
        case 1:  // ５月１日
          if (MyYear == 2019) {
              Result = "新天皇即位";
          } else;
          break;

        case 2:  // ５月２日
          if (MyYear == 2019) {
              Result = "新天皇即位翌日";
          } else;
          break;
          //2019/03/25 Add-E Tanaka

        case 3:  // ５月３日
          Result = "憲法記念日";
          break;
        case 4:  // ５月４日
          if (MyYear >= 2007) {
              Result = "みどりの日";
          } else {
              if (MyYear >= 1986) {
                  if (MyDate.getDay() > MONDAY) {
                  // 5/4が日曜日は『只の日曜』､月曜日は『憲法記念日の振替休日』(～2006年)
                      Result = "国民の休日";
                  } else;
              } else;
          }
          break;
        case 5:  // ５月５日
          Result = "こどもの日";
          break;
        case 6:  // ５月６日
          if (MyYear >= 2007) {
              if ((MyDate.getDay() == TUESDAY) || (MyDate.getDay() == WEDNESDAY)) {
                  Result = "振替休日";    // [5/3,5/4が日曜]ケースのみ、ここで判定
              } else;
          } else;
          break;
      }
      break;
// ６月 //
  case 6:
      if (MyDate.getTime() == cstNorihitoKekkon.getTime()) {
          Result = "皇太子徳仁親王の結婚の儀";
      } else;
      break;
// ７月 //
  case 7:
      //2019/03/25 Upd-S Tanaka 2020オリンピック特例
      //if (MyYear >= 2003) {
      if (MyYear == 2020) {
          if (MyDay == 23) {
              Result = "海の日";
          } else if(MyDay == 24){
              Result = "スポーツの日";
          } else;
      } else if (MyYear >= 2003) {
      //2019/03/25 Upd-E Tanaka
          NumberOfWeek = Math.floor((MyDay - 1) / 7) + 1;
          if ((NumberOfWeek == 3) && (MyDate.getDay() == MONDAY)) {
              Result = "海の日";
          } else;
      } else {
          if (MyYear >= 1996) {
              if (MyDay == 20) {
                  Result = "海の日";
              } else;
          } else;
      }
      break;

// 8月 //
  case 8:
      //2019/03/25 Upd-S Tanaka 2020オリンピック特例
      //if (MyYear >= 2016) {
          if (MyYear == 2020) {
          if (MyDay == 10) {
              Result = "山の日";
          }else;
      } else if (MyYear >= 2016) {
          //2019/03/25 Upd-E Tanaka
            if (MyDay == 11) {
                Result = "山の日";
            }
        }
        break;
// ９月 //
  case 9:
      //第３月曜日(15～21)と秋分日(22～24)が重なる事はない
      MyAutumnEquinox = prvDayOfAutumnEquinox(MyYear);
      if (MyDay == MyAutumnEquinox) {    // 1948～2150以外は[99]
          Result = "秋分の日";           // が返るので､必ず≠になる
      } else {
          if (MyYear >= 2003) {
              NumberOfWeek = Math.floor((MyDay - 1) / 7) + 1;
              if ((NumberOfWeek == 3) && (MyDate.getDay() == MONDAY)) {
                  Result = "敬老の日";
              } else {
                  if (MyDate.getDay() == TUESDAY) {
                      if (MyDay == (MyAutumnEquinox - 1)) {
                          Result = "国民の休日";
                      } else;
                  } else;
              }
          } else {
              if (MyYear >= 1966) {
                  if (MyDay == 15) {
                      Result = "敬老の日";
                  } else;
              } else;
          }
      }
      break;
// １０月 //
  case 10:
      //2019/03/25 Upd-S Tanaka 天皇即位対応
      //if (MyYear >= 2000) {
      if (MyYear == 2019 && MyDay == 22) {
          Result = "即位礼正殿の儀";

          //2019/03/25 2020オリンピック関連
      } else if (MyYear >= 2020) {
          NumberOfWeek = Math.floor(( MyDay - 1) / 7) + 1;
          if ((NumberOfWeek == 2) && (MyDate.getDay() == MONDAY)) {
              if (MyYear == 2020) {
                  //「スポーツの日」は特例として7/24になるよ
              }else{
                  Result = "スポーツの日";
              }
          }
      } else if (MyYear >= 2000) {
      //2019/03/25 Upd-E Tanaka
          NumberOfWeek = Math.floor(( MyDay - 1) / 7) + 1;
          if ((NumberOfWeek == 2) && (MyDate.getDay() == MONDAY)) {
              Result = "体育の日";
          } else;
      } else {
          if (MyYear >= 1966) {
              if (MyDay == 10) {
                  Result = "体育の日";
              } else;
          } else;
      }
      break;
// １１月 //
  case 11:
      if (MyDay == 3) {
          Result = "文化の日";
      } else {
          if (MyDay == 23) {
              Result = "勤労感謝の日";
          } else {
              if (MyDate.getTime() == cstSokuireiseiden.getTime()) {
                  Result = "即位礼正殿の儀";
              } else;
          }
      }
      break;
// １２月 //
  case 12:
      if (MyDay == 23) {
          //2019/03/25 Upd-S Tanaka 天皇即位対応
          //if (MyYear >= 1989) {
          if (MyYear >= 1989 && MyYear <= 2018 ) {
          //2019/03/25 Upd-E Tanaka
              Result = "天皇誕生日";
          } else;
      } else;
      break;
  }

  return Result;
}

//===================================================================
// 春分/秋分日の略算式は
// 『海上保安庁水路部 暦計算研究会編 新こよみ便利帳』
// で紹介されている式です。
function prvDayOfSpringEquinox(MyYear)
{
  var SpringEquinox_ret;

  if (MyYear <= 1947) {
      SpringEquinox_ret = 99;    //祝日法施行前
  } else {
      if (MyYear <= 1979) {
          // Math.floor 関数は[VBAのInt関数]に相当
          SpringEquinox_ret = Math.floor(20.8357 + 
            (0.242194 * (MyYear - 1980)) - Math.floor((MyYear - 1980) / 4));
      } else {
          if (MyYear <= 2099) {
              SpringEquinox_ret = Math.floor(20.8431 + 
                (0.242194 * (MyYear - 1980)) - Math.floor((MyYear - 1980) / 4));
          } else {
              if (MyYear <= 2150) {
                  SpringEquinox_ret = Math.floor(21.851 + 
                    (0.242194 * (MyYear - 1980)) - Math.floor((MyYear - 1980) / 4));
              } else {
                  SpringEquinox_ret = 99;    //2151年以降は略算式が無いので不明
              }
          }
      }
  }
  return SpringEquinox_ret;
}

//=====================================================================
function prvDayOfAutumnEquinox(MyYear)
{
  var AutumnEquinox_ret;

  if (MyYear <= 1947) {
      AutumnEquinox_ret = 99; //祝日法施行前
  } else {
      if (MyYear <= 1979) {
          // Math.floor 関数は[VBAのInt関数]に相当
          AutumnEquinox_ret = Math.floor(23.2588 + 
            (0.242194 * (MyYear - 1980)) - Math.floor((MyYear - 1980) / 4));
      } else {
          if (MyYear <= 2099) {
              AutumnEquinox_ret = Math.floor(23.2488 + 
                (0.242194 * (MyYear - 1980)) - Math.floor((MyYear - 1980) / 4));
          } else {
              if (MyYear <= 2150) {
                  AutumnEquinox_ret = Math.floor(24.2488 + 
                    (0.242194 * (MyYear - 1980)) - Math.floor((MyYear - 1980) / 4));
              } else {
                  AutumnEquinox_ret = 99;    //2151年以降は略算式が無いので不明
              }
          }
      }
  }
  return AutumnEquinox_ret;
}
//2017/08/25 H.Yoshikawa Add End

// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0"  onLoad="setParam(document.dmi220F);finit();">
<!-------------空搬出情報入力更新画面--------------------------->
<FORM name="dmi220F" method="POST">
<TABLE border=0 cellPadding=3 cellSpacing=0 width="100%">
  <TR>
    <TD colspan=2>
      <B>空バンピック情報入力<%=plintStr%></B></TD></TR>
  <TR>
    <TD><DIV class=bgb>ブッキングＮｏ．</DIV></TD>
    <TD><INPUT type=text name="BookNoM" value="<%=Request("BookNoM")%>" readOnly tabindex=-1 size=40>
        <INPUT type=hidden name="BookNo" value="<%=Request("BookNo")%>"></TD></TR>
  <TR>
    <TD><DIV class=bgb>船社</DIV></TD>
    <TD><INPUT type=text name="shipFact" value="<%=Request("shipFact")%>" readOnly tabindex=-1 size=40></TD></TR>
  <TR>
    <TD><DIV class=bgb>*船名</DIV></TD>
    <TD><INPUT type=text name="shipName" value="<%=Request("shipName")%>" readonly size=40>					<!-- 2016/08/22 H.Yoshikawa Upd (readonlyに変更) -->
    	<% if Dflag = "" then %>
    	<INPUT type=button value="検索" onClick="VslSelect()">
    	<% end if %>
    	<INPUT type=hidden name="VslCode" value="<%=Request("VslCode")%>">									<!-- 2016/08/22 H.Yoshikawa Add -->
    </TD></TR>
  <TR>
  	<!-- 2016/08/22 H.Yoshikawa Upd Start -->
    <TD><DIV class=bgb><!--仕向地-->*Voyage</DIV></TD>
    <TD><INPUT type=hidden name="delivTo" value="<%=Request("delivTo")%>">
    	<INPUT type=text name="ExVoyage" value="<%=Request("ExVoyage")%>" size=10 readonly maxlength=12>	<!-- 2016/10/17 H.Yoshikawa Add -->
    	<INPUT type=hidden name="VoyCtrl" value="<%=Request("VoyCtrl")%>" >									<!-- 2016/10/17 H.Yoshikawa Upd(text→hidden) -->
    </TD></TR>
  	<!-- 2016/08/22 H.Yoshikawa Upd End -->
  <TR>
    <TD><DIV class=bgb>会社コード(陸運)</DIV></TD>
    <TD><INPUT type=text name="COMPcd1" value="<%=Trim(COMPcd1)%>" size=5 <%=Dflag%> maxlength=2>
        <INPUT type=hidden name="oldCOMPcd1" value="<%=Request("oldCOMPcd1")%>"></TD></TR>
  <TR>
    <TD><DIV class=bgb>属性と本数</DIV></TD>
    <TD></TD></TR>
  <TR>
    <TD colspan=2>
    <TABLE border=0 cellPadding=0 cellSpacing=0 width="90%" align=center>				<!-- 2017/05/10 H.Yoshikawa Upd(width:920⇒980) -->
    <!-- 2016/08/16 H.Yoshikawa Upd Start -->
    <!-- <TR><TD></TD><TD>サイズ</TD><TD>タイプ</TD><TD>高さ</TD><TD>材質</TD><TD>ピック場所</TD><TD></TD><TD>本数</TD></TR> -->
    <TR>
    	<TD></TD>
    	<TD>*サイズ</TD>
    	<TD>*タイプ</TD>
    	<TD>*高さ</TD>
    	<TD>設定温度</TD>
    	<TD>プレクール</TD>
    	<TD>ベンチレーション</TD>
    	<TD>*ピック予定日時(時間はﾌﾟﾚｸｰﾙ時のみ必須)</TD>
    	<TD>　*本数</TD>
    	<TD>搬出可否</TD>
    	<TD>ピックアップ場所</TD>
    	<TD>変更</TD>
    	<TD>行削除</TD>									<!-- 2017/05/10 H.Yoshikwawa Add -->
    </TR>
    <!-- 2016/08/16 H.Yoshikawa Upd End -->
<% For i=0 To RowNum-1%>								<!-- 2017/05/09 H.Yoshikawa Upd(4⇒RowNum-1) -->
	<% '2016/10/26 H.Yoshikawa Add Start
		OutNum = 0
		
		if gfTrim(Request("ContSize" & i)) <> "" then
			if gfTrim(Request("BackFlag")) <> "1" then		'2017/08/25 H.Yoshikawa Add(確認画面からの戻りの場合は、新たに取得しない)
				'同一属性の搬出済み本数を取得
				StrSQL = "SELECT Count(Exc.ContNo) as NumCont FROM ExportCont Exc "
				'2017/08/24 H.Yoshikawa Upd Start
				''2017/05/10 H.Yoshikawa Upd Start
				''StrSQL = StrSQL & " INNER JOIN Container Con ON Con.VslCode = Exc.VslCode AND Con.VoyCtrl = Exc.VoyCtrl AND Con.ContNo = Exc.ContNo "
				'StrSQL = StrSQL & " INNER JOIN Pickup Con ON Con.VslCode = Exc.VslCode AND Con.VoyCtrl = Exc.VoyCtrl AND Con.BookNo = Exc.BookNo "
				''2017/05/10 H.Yoshikawa Upd End
				StrSQL = StrSQL & " INNER JOIN Container Con ON Con.VslCode = Exc.VslCode AND Con.VoyCtrl = Exc.VoyCtrl AND Con.ContNo = Exc.ContNo "
				'2017/08/24 H.Yoshikawa Upd End
				StrSQL = StrSQL & "WHERE Exc.VslCode    = '" & gfSQLEncode(Request("VslCode")) & "' "
				StrSQL = StrSQL & "  AND Exc.VoyCtrl    = '" & gfSQLEncode(Request("VoyCtrl")) & "' "
				StrSQL = StrSQL & "  AND Exc.BookNo     = '" & BookNo & "' "
				StrSQL = StrSQL & "  AND Con.ContSize   = '" & gfSQLEncode(Request("ContSize" & i)) & "'"
				StrSQL = StrSQL & "  AND Con.ContType   = '" & gfSQLEncode(Request("ContType" & i)) & "'"
				StrSQL = StrSQL & "  AND Con.ContHeight = '" & gfSQLEncode(Request("ContHeight" & i)) & "'"
				StrSQL = StrSQL & "  AND Exc.EmpDelTime IS NOT NULL"
				ObjRS.Open StrSQL, ObjConn
				if err <> 0 then
					DisConnDBH ObjConn, ObjRS
					jampErrerP "1","b303","01","空搬出：搬出済み本数取得","101","SQL:<BR>"&strSQL
				end if
				if not ObjRS.eof then
					OutNum=ObjRS("NumCont")
				end if
				ObjRS.close
			'2017/08/25 H.Yoshikawa Add Start (確認画面からの戻りの場合は、新たに取得しない)
			else
				OutNum  = gfTrim(Request("OutNum" & i))
			'2017/08/25 H.Yoshikawa Add End
			end if
		end if		
	   '2016/10/26 H.Yoshikawa Add End %>
	<% '2016/08/22 H.Yoshikawa Add Start %>
	<%if Dflag = "" then
		If Mord=0 Then '新規登録時
			if Request("UpdFlag"&i) = "1" then
				Dflag2 = ""
			else
				Dflag2 = "readOnly"
			end if
		elseif Mord = 1 then
			if Request("UpdFlag"&i) = "1" then
				Dflag2 = ""
			else
				Dflag2 = "readOnly"
			end if
		else
           Dflag2="readOnly"
		end if
		DflagZokusei = Dflag2
	  else
		if Request("UpdFlag"&i) = "1" then
			if RTrim(Request("Bef_ContSize"&i)) = "" then
				Dflag2=""
				DflagZokusei=""
			else
				Dflag2="readOnly"
				if OutNum > 0 then
					DflagZokusei = "readOnly"
				else
					DflagZokusei = ""
				end if
			end if
		else
			Dflag2="readOnly"
			DflagZokusei = "readOnly"
		end if
	  end if
	%>
	<% '2016/08/22 H.Yoshikawa Add End %>
      <TR><TD>(<%=i+1%>)</TD>
          <TD><INPUT type=text name="ContSize<%=i%>"       value="<%=Request("ContSize"&i)%>" size=4 <%=DflagZokusei%> maxlength=2>
              <INPUT type=hidden name="Bef_ContSize<%=i%>" value="<%=Request("Bef_ContSize"&i)%>">
          </TD>
          <TD><!-- 2016/11/09 H.Yoshikawa Upd Start -->
              <!-- <INPUT type=text name="ContType<%=i%>"       value="<%=Request("ContType"&i)%>" size=4 <%=DflagZokusei%> maxlength=2> -->
              <select name="ContTypeSel<%=i%>" ></select>
              <INPUT type=hidden name="ContType<%=i%>" value="<%=Request("ContType"&i)%>">
              <!-- 2016/11/09 H.Yoshikawa Upd End -->
              <INPUT type=hidden name="Bef_ContType<%=i%>" value="<%=Request("Bef_ContType"&i)%>">
          </TD>
          <TD><INPUT type=text name="ContHeight<%=i%>"       value="<%=Request("ContHeight"&i)%>" size=4 <%=DflagZokusei%> maxlength=2>
              <INPUT type=hidden name="Bef_ContHeight<%=i%>" value="<%=Request("Bef_ContHeight"&i)%>">
          </TD>
      <!-- 2016/08/22 H.Yoshikawa Upd Start -->
          <!--<TD><INPUT type=text name="Material<%=i%>"   value="<%=Request("Material"&i)%>" size=4 <%=Dflag%> maxlength=1></TD>
          <TD><INPUT type=text name="PickPlace<%=i%>"  value="<%=Request("PickPlace"&i)%>"  size=25 <%=Dflag%> maxlength=20></TD>
          <TD>・・・</TD>
          <TD><INPUT type=text name="PickNum<%=i%>" value="<%=Request("PickNum"&i)%>" size=4  <%=Dflag%> maxlength=3></TD></TR> -->
          <TD><INPUT type=text name="SetTemp<%=i%>"  value="<%=Request("SetTemp"&i)%>" size=8 <%=Dflag2%> maxlength=5>℃
              <INPUT type=hidden name="Bef_SetTemp<%=i%>" value="<%=Request("Bef_SetTemp"&i)%>">
          </TD>
          <TD><select name="Pcool<%=i%>" <% if Dflag2 <> "" then%>disabled<% end if %>>
				<option value="0"></option>
				<option value="1" <% if gfTrim(Request("Pcool"&i)) = "1" then %>selected<% end if %> >有</option>
				<option value="2" <% if gfTrim(Request("Pcool"&i)) = "2" then %>selected<% end if %> >無</option>	<!-- 2017/08/25 H.Yoshikawa Add -->
			  </select>
              <INPUT type=hidden name="Bef_Pcool<%=i%>" value="<%=Request("Bef_Pcool"&i)%>">
          </TD>
          <TD><INPUT type=text name="Ventilation<%=i%>"  value="<%=Request("Ventilation"&i)%>" size=5 <%=Dflag2%> maxlength=3>%（開口）
              <INPUT type=hidden name="Bef_Ventilation<%=i%>" value="<%=Request("Bef_Ventilation"&i)%>">
          </TD>
          <TD><INPUT type=text name="PickDate<%=i%>"  value="<%=Request("PickDate"&i)%>" size=15 <% if RTrim(Request("UpdFlag"&i)) <> "1" then%>readOnly<% end if %> maxlength=10>		<!-- 2016/11/11 H.Yoshikawa Upd (readOnlyの条件変更：Dflag2→変更チェックONなら常に編集可能に) -->
              <INPUT type=hidden name="Bef_PickDate<%=i%>" value="<%=Request("Bef_PickDate"&i)%>">
              <INPUT type=text name="PickHour<%=i%>"  value="<%=Request("PickHour"&i)%>" size=4 <% if RTrim(Request("UpdFlag"&i)) <> "1" then%>readOnly<% end if %> maxlength=2>時		<!-- 2016/11/11 H.Yoshikawa Upd (readOnlyの条件変更：Dflag2→変更チェックONなら常に編集可能に) -->
              <INPUT type=text name="PickMinute<%=i%>"  value="<%=Request("PickMinute"&i)%>" size=4 <% if RTrim(Request("UpdFlag"&i)) <> "1" then%>readOnly<% end if %> maxlength=2>分	<!-- 2016/11/11 H.Yoshikawa Upd (readOnlyの条件変更：Dflag2→変更チェックONなら常に編集可能に) -->
              <INPUT type=hidden name="Bef_PickHour<%=i%>" value="<%=Request("Bef_PickHour"&i)%>">
              <INPUT type=hidden name="Bef_PickMinute<%=i%>" value="<%=Request("Bef_PickMinute"&i)%>">
          </TD>
          <!--<TD>・・・</TD>-->
          <TD>…<INPUT type=text name="PickNum<%=i%>" value="<%=Request("PickNum"&i)%>" size=4 <% if RTrim(Request("UpdFlag"&i)) <> "1" then%>readOnly<% end if %> maxlength=3>
                <INPUT type=hidden name="Bef_PickNum<%=i%>" value="<%=Request("Bef_PickNum"&i)%>">
                <INPUT type=hidden name="OutNum<%=i%>" value="<%=OutNum%>">  <!-- 2016/10/26 H.Yoshikawa Add -->
          </TD>
          <% select case Trim(Request("OutFlag"&i))
               case "0"
                 WkOutFlag = "確認中"
                 OutStyle = ""
               case "1"
                 WkOutFlag = "可"
                 OutStyle = ""
               case "9"
                 WkOutFlag = "不可"
                 OutStyle = "color:red;"
               case else
                 WkOutFlag = ""
                 OutStyle = ""
             end select
          %>
          <TD style="<%=OutStyle%>"><INPUT type=hidden name="OutFlag<%=i%>"  value="<%=Request("OutFlag"&i)%>" ><%=WkOutFlag %></TD>
          <TD><INPUT type=hidden name="PickPlace<%=i%>"  value="<%=Request("PickPlace"&i)%>"><%=gfHTMLEncode(Request("PickPlace"&i))%>
              <INPUT type=hidden name="Terminal<%=i%>"  value="<%=Request("Terminal"&i)%>">
          </TD>
          <TD><INPUT type=checkbox name="UpdFlag<%=i%>"  value="1"  <% if RTrim(Request("UpdFlag"&i)) = "1" then%> checked <% end if %> onclick="UpdFlagChg(<%=i%>);"></TD>
          <TD><INPUT type=checkbox name="DelFlag<%=i%>"  value="1"  <% if RTrim(Request("DelFlag"&i)) = "1" then%> checked <% end if %> onclick="DelFlagChg(<%=i%>);"></TD>
	  </TR>
      <!-- 2016/08/22 H.Yoshikawa Upd Start -->
	<% '2016/10/28 H.Yoshikawa Upd End %>
	  <INPUT type=hidden name="Bef_OutFlag<%=i%>"     value="<%=Request("Bef_OutFlag"&i)%>">
	  <INPUT type=hidden name="Bef_PickPlace<%=i%>"   value="<%=Request("Bef_PickPlace"&i)%>">
	  <INPUT type=hidden name="Bef_Terminal<%=i%>"    value="<%=Request("Bef_Terminal"&i)%>">
	<% '2016/10/28 H.Yoshikawa Upd End %>
<% Next %>
    </TABLE>
    </TD></TR>
  <TR>
    <TD><DIV class=bgb>バン詰め日時</DIV></TD>
    <TD><INPUT type=text name="vanMon" value="<%=Request("vanMon")%>" size=3 <%=Dflag%> maxlength=2>月
        <INPUT type=text name="vanDay" value="<%=Request("vanDay")%>" size=3 <%=Dflag%> maxlength=2>日
        <INPUT type=text name="vanHou" value="<%=Request("vanHou")%>" size=3 <%=Dflag%> maxlength=2>時
        <INPUT type=text name="vanMin" value="<%=Request("vanMin")%>" size=3 <%=Dflag%> maxlength=2>分
        </TD></TR>
  <TR>
    <TD><DIV class=bgb>バン詰め場所１</DIV></TD>
    <TD><INPUT type=text name="vanPlace1" value="<%=Request("vanPlace1")%>" size=73 <%=Dflag%> maxlength=70></TD></TR>
  <TR>
    <TD><DIV class=bgb>バン詰め場所２</DIV></TD>
    <TD><INPUT type=text name="vanPlace2" value="<%=Request("vanPlace2")%>" size=73 <%=Dflag%> maxlength=70></TD></TR>
  <TR>
    <TD><DIV class=bgb>品名</DIV></TD>
    <TD><INPUT type=text name="goodsName" value="<%=Request("goodsName")%>" size=30 <%=Dflag%> maxlength=20></TD></TR>
  <TR>
    <TD><DIV class=bgb>搬入先ＣＹ．ＣＹカット日</DIV></TD>
    <TD><INPUT type=text name="Terminal" value="<%=Request("Terminal")%>" readOnly tabindex=-1>
        <INPUT type=text name="CYCut" value="<%=Request("CYCut")%>" readOnly tabindex=-1></TD></TR>
  <TR>
    <TD><DIV class=bgb>備考１</DIV></TD>
    <TD><INPUT type=text name="Comment1" value="<%=Request("Comment1")%>" size=73 <%=Dflag%> maxlength=70></TD></TR>
  <TR>
    <TD><DIV class=bgb>備考２</DIV></TD>
    <TD><INPUT type=text name="Comment2" value="<%=Request("Comment2")%>" size=73 <%=Dflag%> maxlength=70></TD></TR>

  <TR>
<!-- 2009/03/10 R.Shibuta Add-S -->
  	<TD><DIV class=bgy>*登録担当者</DIV></TD>
	<!-- 2009/07/25 Update C.Pestano -->
 	<TD><INPUT type=text name="TruckerSubName" value="<%=Request("TruckerSubName")%>" maxlength=16 ></TD></TR>
<!-- 2009/03/10 R.Shibuta Add-E -->
<!-- 2016/08/22 H.Yoshikawa Add Start -->
  <TR>
  	<TD><DIV class=bgy>*電話番号</DIV></TD>
 	<TD><INPUT type=text name="Tel" value="<%=Request("Tel")%>" maxlength=14 onBlur="CheckLen(this,true,true,false)"></TD></TR>
  <TR>
  	<TD><DIV class=bgy>*メールアドレス</DIV></TD>
 	<TD><INPUT type=text name="Mail" value="<%=Request("Mail")%>"  size=60 maxlength=100 onBlur="CheckLen(this,true,true,false)">
 		<INPUT type=checkbox name="MailFlag" value="1" <% if Request("MailFlag") = "1" then %>checked <% end if %>>
 		搬出可否状態変更時にメールを受け取る
 	</TD></TR>
<!-- 2016/08/22 H.Yoshikawa Add End -->
  <TR>
    <TD colspan=2 align=center>
<% If Request("ErrerM")<>"" Then %>
       <%= Request("ErrerM") %><BR>
<% Else %>
       <P><BR></P>
<% End If %>
       <INPUT type=hidden name=COMPcd0 value="<%=COMPcd0%>" >
<%'Add-s 2006/03/06 h.matsuda%>
       <INPUT type=hidden name=shipline value="<%=Request("shipline")%>" >
	   <INPUT type=hidden name="ShoriMode" value="EMoutInf">
<%'Add-e 2006/03/06 h.matsuda%>
<%'2016/08/30 H.Yoshikawa Add Start%>
       <INPUT type=hidden name=compFlag value="<%=Request("compFlag")%>" >
<%'2016/08/30 H.Yoshikawa Add End%>

<% If Mord=0 Then %>
       <INPUT type=hidden name=Mord value="0" >
       <INPUT type=button value="登録" onClick="GoNext()">
<% ElseIf COMPcd0 = UCase(Session.Contents("userid")) Then%>
       <INPUT type=hidden name=Mord value="1" >
  <%'If TFlag<>"1" AND Request("compFlag")="0" Then					2016/10/25 H.Yoshikawa Del %>
       <INPUT type=button value="更新" onClick="GoNext()">
  <% 'End If 														2016/10/25 H.Yoshikawa Del %>
       <INPUT type=button value="削除" onClick="GoDell()">
<% Else %>
       <INPUT type=hidden name=Mord value="2" >
       <DIV class=bgw>指示元へ回答　　　
       <INPUT type=radio name="way" checked>Yes　
       <INPUT type=radio name="way">No</DIV>
       <INPUT type=hidden name=Res value="1" >
    </TD></TR>
    <TR><TD colspan=2 align=center>
       <INPUT type=button value="更新" onClick="Suspend()">
<% End If %>
       <INPUT type=button value="キャンセル" onClick="window.close()">
       <P>
       <INPUT type=button value="ブッキング情報" onClick="GoBookI()">
    </TD></TR>


</TABLE>
</FORM>
<!-------------画面終わり--------------------------->
</BODY></HTML>
<% DisConnDBH ObjConn, ObjRS %>
