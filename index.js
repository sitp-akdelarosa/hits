function winOpen(winName,url,W,H){
  var Win1=window.open(url,winName,'scrollbars=yes,resizable=yes,width='+W+',height='+H+'');
  Win1.document.close();
}