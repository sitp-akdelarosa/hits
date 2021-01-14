//Enterキーはタブ移動とする。
function gfkeyDown(){	
    if (document.activeElement.type != "button" && document.activeElement.type != "file" && document.activeElement.type != "submit" && document.activeElement.type != "reset" && document.activeElement.name !="TOPIC_CONTENT"){
        if(window.event.keyCode=="13"){			
			window.event.keyCode=9
		}
    }	
}
document.onkeydown = gfkeyDown;
