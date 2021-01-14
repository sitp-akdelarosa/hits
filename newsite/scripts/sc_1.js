//window.onload=move_icon;
var i_tmp;
var spd = 1;

//座標テーブル
var img_x;
var img_width = 130;	//画像幅
var img_count = 8;	//画像個数
var disp_count = 8;	//画像表示幅個数　8枚
img_x = new Array( img_count );//画像個数

for( i_tmp = 1; i_tmp <= img_count ; i_tmp++ ){
	img_x[i_tmp] = img_width * ( disp_count - i_tmp );
}

function move_icon(){

	var dgiz=document.getElementById("zentai");
	var gz_array;
	var gbz_array;
	gz_array = new Array( img_count );
	gbz_array = new Array( img_count );

	var i;
	
	dgiz.style.width="980px";//全体のサイズ
	dgiz.style.height="70px";
	dgiz.style.border="1px solid silver";
	dgiz.style.position="relative";
	dgiz.style.top="0px";
	dgiz.style.left="0px";
	dgiz.style.overflow="hidden";

	for( i = 1 ; i <= img_count ; i++ ){
		gz_array[i] = document.getElementById("gazou"+i).style;
		gbz_array[i] = document.getElementById("b_gazou"+i).style;
		if( gz_array[i] ){
			gz_array[i].position="absolute";
			gbz_array[i].style="visibility:hidden;";
			gz_array[i].top="0px";
		}
		if( gbz_array[i] ){
			gbz_array[i].position="absolute";
			gbz_array[i].style="visibility:hidden;";
			gbz_array[i].top="0px";
		}
	}

	for( i = 1 ; i <=img_count ; i++ ){
		if( spd > 0 ){
			if ( img_x[i] < img_width * disp_count && img_x[i] > img_width * ( disp_count - 1 ) ){
				img_x[i]+= spd;
				if( gz_array[i] ){
					gz_array[i].right= img_x[i] +"px";
				}
				if( gbz_array[i] ){
					gbz_array[i].style="visibility:visible; ";
					gbz_array[i].right= img_x[i] - ( img_width * img_count ) +"px";
				}
			}
			else if ( img_x[i] <= img_width * ( disp_count - 1 ) ){
				img_x[i]+= spd;
				if( gz_array[i] ){
					gz_array[i].right= img_x[i] +"px";
				}
				if( gbz_array[i] ){
					gbz_array[i].right= img_width * -1 + "px";
					gbz_array[i].style="visibility:hidden; ";
				}
			}//全体の（左側）外まで移動した後
			else if( img_x[i] > 0 && img_x[i] >= img_width * disp_count ){
				img_x[i] = img_width * ( disp_count - img_count ) + spd;
				if( gz_array[i] ){
					gz_array[i].right= img_x[i] +"px";
				}
				if( gbz_array[i] ){
					gbz_array[i].right= img_width * -1 + "px";
					gbz_array[i].style="visibility:hidden; ";
				}
			}
			else{
				img_x[i] = img_width * ( disp_count - img_count );		//右側の外へ移動　right　0px　からマイナス１枚分の距離
				if( gbz_array[i] ){
					gbz_array[i].right= img_width * -1 + "px";
					gbz_array[i].style="visibility:hidden; ";
				}
			}
		}else{
			if ( img_x[i] < 0 && img_x[i] > img_width * -1 ){
				img_x[i]+= spd;
				if( gz_array[i] ){
					gz_array[i].right= img_x[i] +"px";
				}
				if( gbz_array[i] ){
					gbz_array[i].style="visibility:visible; ";
					gbz_array[i].right= ( img_width * img_count ) + img_x[i] +"px";
				}
			}
			else if ( img_x[i] > img_width * -1  ){
				img_x[i]+= spd;
				if( gz_array[i] ){
					gz_array[i].right= img_x[i] +"px";
				}
				if( gbz_array[i] ){
					gbz_array[i].right= img_width * img_count + "px";
					gbz_array[i].style="visibility:hidden; ";
				}
			}//全体の（左側）外まで移動した後
			else if( img_x[i] < 0 && img_x[i] <= img_width * -1 ){
				img_x[i] = img_width * ( img_count - 1 ) + spd ;		//右側の外へ移動　right　0px　からマイナス１枚分の距離
				if( gz_array[i] ){
					gz_array[i].right= img_x[i] +"px";
				}
				if( gbz_array[i] ){
					gbz_array[i].right= img_width * -1 + "px";
					gbz_array[i].style="visibility:hidden; ";
				}
			}
			else{
				img_x[i] = img_width * ( img_count - 1 );		//右側の外へ移動　right　0px　からマイナス１枚分の距離
				if( gbz_array[i] ){
					gbz_array[i].right= img_width * img_count + "px";
					gbz_array[i].style="visibility:hidden; ";
				}
			}

		}
	}
	
	//var str="";
	//for( i = 1 ; i <= img_count ; i++ ){
	//	str = str + "img_x[" + i + "]:" + img_x[i] + "\n";
	//}
	//if( spd == 1 ){
	//	alert( str );
	//}
	
	setTimeout("move_icon()",30);

}

