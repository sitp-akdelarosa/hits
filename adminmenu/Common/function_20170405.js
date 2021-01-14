// JavaScript Document

function CheckEisuji(str){
  checkstr="0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
  for (i=0; i<str.length; i++){
    c = str.charAt(i);
    if (checkstr.indexOf(c,0) < 0){
      return false;
    }
  }
  return true;
}

//**************************************************
// 概要　 : 未入力チェック
// 引数　 : ls_str(項目名称）
// 戻り値 : true(正常)/false(異常)
//**************************************************
function gfCHKNull(ls_str)
{
    if (ls_str.value.length == 0) {
        window.alert("必須入力項目です。");
        return false;
    }
    return true;
}

//*********************************************************
//  関数名 ： gfCHKDate
//  概要   ： 入力日付データ型チェック
//  引数   ： ls_str    (画面の項目名）
//  戻り値 ： TURE(正常)/FALSE(異常)
//  作成日 ： 2000年04月20日
//*********************************************************
function gfCHKDate(ls_str)
{
    var p_val=ls_str.value;
    var v_yyyy;
    var v_mm;
    var v_dd;
    var v_kuguriY;
    var v_kuguriM;
	var errMsg = "実在する日付を入力して下さい。";
	
    if (p_val.length == 0){return(true);}
    
	if (p_val.length != 10){
        window.alert("10桁で入力して下さい。(YYYY/MM/DD)");
		return false;     // invalid length
    }
	
    if (gfCHKNumberD(ls_str) == false){return(false);}                   // not numeric
    //年
    var scode=p_val.substring(0, 4);
    for( var i=0; i < scode.length; i++ )   {
        if( "0123456789.".indexOf(scode.charAt(i)) == -1 )      {
            window.alert(errMsg);
            return false;
        }
    }
    v_yyyy = parseInt(p_val.substring(0, 4),10);
    //月
    var scode=p_val.substring(5, 7);
    for( var i=0; i < scode.length; i++ )   {
        if( "0123456789.".indexOf(scode.charAt(i)) == -1 )      {
            window.alert(errMsg);
            return false;
        }
    }
    v_mm = parseInt(p_val.substring(5, 7),10);
    //日
    var scode=p_val.substring(8, 10);
    for( var i=0; i < scode.length; i++ )   {
        if( "0123456789.".indexOf(scode.charAt(i)) == -1 )      {
            window.alert(errMsg);
            return false;
        }
    }
    v_dd = parseInt(p_val.substring(8, 10),10);
    v_kuguriY = p_val.substring(4, 5);
    v_kuguriM = p_val.substring(7, 8);
    if ((v_kuguriY != "/") || (v_kuguriM != "/")){
        window.alert(errMsg);return(false);
    }
    if (v_yyyy < 1900){
        window.alert(errMsg);return(false);
    }
    if ((v_mm < 1) || (v_mm > 12)){
        window.alert(errMsg);return(false);         // invalid month
    }
    if ((v_mm == 1) || (v_mm == 3) || (v_mm == 5) || (v_mm == 7) || (v_mm == 8) || (v_mm == 10) || (v_mm == 12)){
        if ((v_dd < 1) || (v_dd > 31)){
            window.alert(errMsg);return(false);     // invalid date
        }
    } else {
        if ((v_dd < 1) || (v_dd > 30)){
            window.alert(errMsg);return(false);     // invalid date
        }
    }
    if (v_mm == 2){                     // check leap year
        if ((v_yyyy % 400 == 0) || ((v_yyyy % 4 == 0) && (v_yyyy % 100 != 0))){
            if (v_dd > 29){
                window.alert(errMsg);return(false); // invalid date, not leap year
            }       // invalid date, leap year
        } else {
            if (v_dd > 28){
                window.alert(errMsg);return(false); // invalid date, not leap year
            }
        }
    }   
    return(true);
}

//**************************************************
// 概要　 : 入力値数値データ型チェック
// 引数　 : ls_str(項目名称）
// 戻り値 : true(正常)/false(異常)
//**************************************************
function gfCHKNumber(ls_str)
{
   var scode = ls_str.value;

    for (var i = 0; i < scode.length; i++)  {
        if ("0123456789".indexOf(scode.charAt(i)) == -1) {
            window.alert("入力値が正しくありません。");
            return false;
        }
    }
    return true;
}

//**************************************************
// 概要　 : 日付データ型チェック
// 引数　 : ls_str(項目名称）
// 戻り値 : true(正常)/false(異常)
//**************************************************
function gfCHKNumberD(ls_str)
{
    var scode = ls_str.value;

    for (var i = 0; i < scode.length; i++)  {
        if ("0123456789/".indexOf(scode.charAt(i)) == -1) {
            window.alert("日付で入力して下さい。(YYYY/MM/DD)");
            return false;
        }
    }
    return true;
}