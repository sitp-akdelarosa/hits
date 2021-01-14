
function Rtrim(str, chars) {
	chars = chars || "\\s";
	return str.replace(new RegExp("[" + chars + "]+$", "g"), "");
}

function SpaceDel(str) {
	retStr = str;
	
	//全角⇒半角変換
	retStr = toHalfWidth(retStr);
	
	//スペース削除
	while ( retStr.indexOf(" ",0) != -1 )
	{
		retStr = retStr.replace(" ", "");
	}

	return retStr;
}

/**
 * 全角から半角への変革関数
 * 入力値の英数記号を半角変換して返却
 * [引数]   strVal: 入力値
 * [返却値] String(): 半角変換された文字列
 */
function toHalfWidth(strVal){
  // 半角変換
  var halfVal = strVal.replace(/[！-～]/g,
    function( tmpStr ) {
      // 文字コードをシフト
      return String.fromCharCode( tmpStr.charCodeAt(0) - 0xFEE0 );
    }
  );

  // 文字コードシフトで対応できない文字の変換
  return halfVal.replace(/”/g, "\"")
    .replace(/’/g, "'")
    .replace(/‘/g, "`")
    .replace(/￥/g, "\\")
    .replace(/　/g, " ")
    .replace(/～/g, "~");
}

//メールアドレスの正当性チェック
function CheckMail(str) {
	a=str;
	if(a==""){
		return true;
	}
	var b=a.replace(/[a-zA-Z0-9_@\.\-?]/g,'');
	if(b.length!=0){
		return false;
	}
	var p1=a.indexOf("@");
	var p2=a.lastIndexOf("@");
	var p3=a.lastIndexOf(".");
	if(0<p1 && p1==p2 && p1<p3 && p3<a.length-1 ){
		return true;
	}
	else{
		return false;
	}
}

//電話番号の正当性チェック
function CheckTel(v){
	if(v.length==0){
		return true;
	}
	var w=v.replace(/[0-9\-]/g,'');
	if(w.length!=0){
		return false;
	}
	return true;
}

//y:year,m:month,d:day,
function CheckYMD(obj){
	var y, m, d;
	var str = obj.value;
	var a = Rtrim(str, " ");
    if(a.length==0){
      return true;
    }
	var w=a.replace(/[0-9\/]/g,'');
	if(w.length!=0){
		return false;
	}

    if(str.indexOf("/")>0){
    	var ymd = str.split("/");
    	if(ymd.length!=3){
    		return false;
    	}
    	y = ymd[0];
    	m = ymd[1];
    	d = ymd[2];
    }else{
    	if(str.length != 8){
    		return false;
    	}
    	y = str.substring(0, 4);
    	m = str.substring(4, 6);
    	d = str.substring(6);
    }
    
	if(!CheckSu(y) || !CheckSu(m) || !CheckSu(d)){
		return false;
	}
	if(eval(y)==0||eval(m)==0||eval(d)==0){
		return false;
	}
	if(eval(y)<2000){
		return(false);
	}
	
	dt=new Date(y,m-1,d);
    if(dt.getFullYear()!=y || dt.getMonth()!=m-1 || dt.getDate()!=d){
    	return false;
    }
    obj.value = y + "/" + ("00" + m).substr(("00" + m).length-2, 2) + "/" + ("00" + d).substr(("00" + d).length-2, 2);
	return true;
}

