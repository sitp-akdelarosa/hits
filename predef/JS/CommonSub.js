
function Rtrim(str, chars) {
	chars = chars || "\\s";
	return str.replace(new RegExp("[" + chars + "]+$", "g"), "");
}

function SpaceDel(str) {
	retStr = str;
	
	//�S�p�˔��p�ϊ�
	retStr = toHalfWidth(retStr);
	
	//�X�y�[�X�폜
	while ( retStr.indexOf(" ",0) != -1 )
	{
		retStr = retStr.replace(" ", "");
	}

	return retStr;
}

/**
 * �S�p���甼�p�ւ̕ϊv�֐�
 * ���͒l�̉p���L���𔼊p�ϊ����ĕԋp
 * [����]   strVal: ���͒l
 * [�ԋp�l] String(): ���p�ϊ����ꂽ������
 */
function toHalfWidth(strVal){
  // ���p�ϊ�
  var halfVal = strVal.replace(/[�I-�`]/g,
    function( tmpStr ) {
      // �����R�[�h���V�t�g
      return String.fromCharCode( tmpStr.charCodeAt(0) - 0xFEE0 );
    }
  );

  // �����R�[�h�V�t�g�őΉ��ł��Ȃ������̕ϊ�
  return halfVal.replace(/�h/g, "\"")
    .replace(/�f/g, "'")
    .replace(/�e/g, "`")
    .replace(/��/g, "\\")
    .replace(/�@/g, " ")
    .replace(/�`/g, "~");
}

//���[���A�h���X�̐������`�F�b�N
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

//�d�b�ԍ��̐������`�F�b�N
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

