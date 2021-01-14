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
  dim Dflag2,Dflag3						'2016/08/22 H.Yoshikawa Add
  BookNo  = Request("BookNo")
  COMPcd0 = Request("COMPcd0")
  COMPcd1 = Request("COMPcd1")
  Mord    = Request("Mord")
  Dflag=""
  Dflag2=""								'2016/08/22 H.Yoshikawa Add
  Dflag3=""								'2016/08/22 H.Yoshikawa Add
  plintStr=""

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
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>空バンピック情報入力更新</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT src="./JS/CommonSub.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
function setParam(target){
//2016/08/22 H.Yoshikawa Upd Start
  //window.resizeTo(550,680);
  window.moveTo(120,20);
  window.resizeTo(1200,770);
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
  for i = 0 to 4
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
  for(idx=0;idx<5;idx++){
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
  for(idx=0;idx<5;idx++){
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
 	}
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
  	var retValue = window.open("./dmlModalVslVoy.asp?tgt=dmi220F&fldvn=shipName&fldvc=VslCode&fldvy=VoyCtrl&flddspvy=ExVoyage&dspkbn=LD", winname, "width=600, height=600, menubar=no, toolbar=no, scrollbars=yes");
  	return true;
}

//属性変更可否設定
function UpdFlagChg(idx){
  var target;
	target=document.dmi220F;
	
	if(target.elements["UpdFlag" + idx].checked == true){
		if(target.COMPcd1.readOnly == true && Rtrim(target.elements["ContSize" + idx].value, ' ') != ""){
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
			
			//属性は同一属性の搬出済みが存在する場合は不可
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
			}else{
				target.elements["ContSize" + idx].readOnly  = false;
				//target.elements["ContType" + idx].readOnly  = false;			//2016/11/09 Del Yoshikawa
				target.elements["ContTypeSel" + idx].disabled  = false;			//2016/11/09 Add Yoshikawa
				target.elements["ContHeight" + idx].readOnly  = false;
			}
			
			//本数は変更可
			target.elements["PickNum" + idx].readOnly  =false;
			//ピック予定日も変更可
			target.elements["PickDate" + idx].readOnly  = false;				// 2016/11/11 H.Yoshikawa Add
			target.elements["PickHour" + idx].readOnly  = false;				// 2016/11/11 H.Yoshikawa Add
			target.elements["PickMinute" + idx].readOnly  = false;				// 2016/11/11 H.Yoshikawa Add
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
    <TABLE border=0 cellPadding=0 cellSpacing=0 width=920 align=center>
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
    </TR>
    <!-- 2016/08/16 H.Yoshikawa Upd End -->
<% For i=0 To 4%>
	<% '2016/10/26 H.Yoshikawa Add Start
		OutNum = 0
		
		if gfTrim(Request("ContSize" & i)) <> "" then
			'同一属性の搬出済み本数を取得
			StrSQL = "SELECT Count(Exc.ContNo) as NumCont FROM ExportCont Exc "
			StrSQL = StrSQL & " INNER JOIN Container Con ON Con.VslCode = Exc.VslCode AND Con.VoyCtrl = Exc.VoyCtrl AND Con.ContNo = Exc.ContNo "
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
		end if		
	   '2016/10/26 H.Yoshikawa Add End %>
	<% '2016/08/22 H.Yoshikawa Add Start %>
	<% if Dflag = "" then
		If Mord=0 Then '新規登録時
			if Request("UpdFlag"&i) = "1" then
				Dflag2 = ""
				Dflag3 = ""
			else
				Dflag2 = "readOnly"
				Dflag3 = ""
			end if
		elseif Mord = 1 then
			if Request("UpdFlag"&i) = "1" then
				Dflag2 = ""
				Dflag3 = ""
			else
				Dflag2 = "readOnly"
				Dflag3 = ""
			end if
		else
           Dflag2="readOnly"
           Dflag3="disabled"
		end if
		DflagZokusei = Dflag2
	  else
		if Request("UpdFlag"&i) = "1" then
			if RTrim(Request("Bef_ContSize"&i)) = "" then
				Dflag2=""
				Dflag3=""
				DflagZokusei=""
			else
				Dflag2="readOnly"
				Dflag3="disabled"
				if OutNum > 0 then
					DflagZokusei = "readOnly"
				else
					DflagZokusei = ""
				end if
			end if
		else
			Dflag2="readOnly"
			Dflag3="disabled"
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
