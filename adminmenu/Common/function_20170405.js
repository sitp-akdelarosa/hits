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
// �T�v�@ : �����̓`�F�b�N
// �����@ : ls_str(���ږ��́j
// �߂�l : true(����)/false(�ُ�)
//**************************************************
function gfCHKNull(ls_str)
{
    if (ls_str.value.length == 0) {
        window.alert("�K�{���͍��ڂł��B");
        return false;
    }
    return true;
}

//*********************************************************
//  �֐��� �F gfCHKDate
//  �T�v   �F ���͓��t�f�[�^�^�`�F�b�N
//  ����   �F ls_str    (��ʂ̍��ږ��j
//  �߂�l �F TURE(����)/FALSE(�ُ�)
//  �쐬�� �F 2000�N04��20��
//*********************************************************
function gfCHKDate(ls_str)
{
    var p_val=ls_str.value;
    var v_yyyy;
    var v_mm;
    var v_dd;
    var v_kuguriY;
    var v_kuguriM;
	var errMsg = "���݂�����t����͂��ĉ������B";
	
    if (p_val.length == 0){return(true);}
    
	if (p_val.length != 10){
        window.alert("10���œ��͂��ĉ������B(YYYY/MM/DD)");
		return false;     // invalid length
    }
	
    if (gfCHKNumberD(ls_str) == false){return(false);}                   // not numeric
    //�N
    var scode=p_val.substring(0, 4);
    for( var i=0; i < scode.length; i++ )   {
        if( "0123456789.".indexOf(scode.charAt(i)) == -1 )      {
            window.alert(errMsg);
            return false;
        }
    }
    v_yyyy = parseInt(p_val.substring(0, 4),10);
    //��
    var scode=p_val.substring(5, 7);
    for( var i=0; i < scode.length; i++ )   {
        if( "0123456789.".indexOf(scode.charAt(i)) == -1 )      {
            window.alert(errMsg);
            return false;
        }
    }
    v_mm = parseInt(p_val.substring(5, 7),10);
    //��
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
// �T�v�@ : ���͒l���l�f�[�^�^�`�F�b�N
// �����@ : ls_str(���ږ��́j
// �߂�l : true(����)/false(�ُ�)
//**************************************************
function gfCHKNumber(ls_str)
{
   var scode = ls_str.value;

    for (var i = 0; i < scode.length; i++)  {
        if ("0123456789".indexOf(scode.charAt(i)) == -1) {
            window.alert("���͒l������������܂���B");
            return false;
        }
    }
    return true;
}

//**************************************************
// �T�v�@ : ���t�f�[�^�^�`�F�b�N
// �����@ : ls_str(���ږ��́j
// �߂�l : true(����)/false(�ُ�)
//**************************************************
function gfCHKNumberD(ls_str)
{
    var scode = ls_str.value;

    for (var i = 0; i < scode.length; i++)  {
        if ("0123456789/".indexOf(scode.charAt(i)) == -1) {
            window.alert("���t�œ��͂��ĉ������B(YYYY/MM/DD)");
            return false;
        }
    }
    return true;
}