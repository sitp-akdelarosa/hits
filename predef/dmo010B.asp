<%@LANGUAGE = VBScript%>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmo010B.asp				_/
'_/	Function	:実搬出情報一覧画面フッタ		_/
'_/	Date		:2003/05/26				_/
'_/	Code By		:SEIKO Electric.Co 大重			_/
'_/	Modify		:C-001 2003/07/29	CSV出力対応	_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTPコンテンツタイプ設定
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>実搬出情報一覧</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
//照会済
function GoSyokaizumi(){
  try{
    parent.DList.GoSyokaizumi();
  }catch(e){}
}
//CSV
function GoCSV(){
  try{
    parent.DList.GoCSV();
  }catch(e){}
}

<%'//印刷画面表示
'function GoPlint(){
'  Win = window.open('', 'PrintWin', 'width=1000,height=600,menubar=yes,scrollbars=yes');
'  document.next.target="PrintWin";
'  document.next.action="./dmo090.asp";
'  document.next.submit();
'}
%>
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!-------------実搬出情報一覧画面Bottom--------------------------->
<CENTER>
  <FORM name="next" action="">
    <TABLE border=0 cellPadding=0 cellSpacing=0 width="100%" height=35>
    <TR><TD>
        <A HREF='JavaScript:GoSyokaizumi()'>指示受諾</A>・・・表示されている全ての未回答データの回答を「Yes」にします。
        </TD>
        <TD>
        <A HREF="JavaScript:GoHelp(1)">ヘルプ</A>・・・画面内の機能の説明画面を表示します。
        </TD></TR>
    <TR><TD colspan=2>
<!--        <A HREF="JavaScript:GoPlint()">印刷画面表示</A>・・・表示内容を印刷に適した画面で表示します。ＢＬ番号指定や一覧から選択されたものは、当該するコンテナに展開して表示します。 -->
        <A HREF="JavaScript:GoCSV()">CSVファイル出力</A>・・・表示内容をCSVファイルに出力します。ＢＬ番号指定や一覧から選択されたものは、当該するコンテナに展開して表示します。
        <INPUT type=hidden name="SortFlag" value="">
        </TD></TR>
    </TABLE>
  </FORM>
</CENTER>
<!-------------画面終わり--------------------------->
</BODY></HTML>
