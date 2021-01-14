function openwin() {
    var w=900;
    var h=550;
    var l=0;
    var t=0;
    if(screen.width){
        l=(screen.width-w)/2;
    }
    if(screen.availWidth){
        l=(screen.availWidth-w)/2;
    }
    if(screen.height){
        t=(screen.height-h)/2;
    }
    if(screen.availHeight){
        t=(screen.availHeight-h)/2;
    }
    
    var win=window.open("../../download/download_list.asp","","status=no,width="+w+",height="+h+",top="+t+",left="+l);
}

var canvas, stage, exportRoot, anim_container, dom_overlay_container, fnStartAnimation;
function init() {
    try {
        canvas = document.getElementById("canvas");
        anim_container = document.getElementById("animation_container");
        dom_overlay_container = document.getElementById("dom_overlay_container");
        var comp=AdobeAn.getComposition("8462D70EDA407345A9D849B291B45E23");
        var lib=comp.getLibrary();

        var loader = new createjs.LoadQueue(false);
        loader.addEventListener("fileload", function(evt){handleFileLoad(evt,comp)});
        loader.addEventListener("complete", function(evt){handleComplete(evt,comp)});

        showHide('animation_container','block');

        var lib=comp.getLibrary();
        loader.loadManifest(lib.properties.manifest);
    }
    catch(err) {
        showHide('gif_alternative','block'); // this will show a gif animation alternative for the html5 animation
    }
}
function handleFileLoad(evt, comp) {
    var images=comp.getImages();	
    if (evt && (evt.item.type == "image")) { images[evt.item.id] = evt.result; }	
}
function handleComplete(evt,comp) {
    //This function is always called, irrespective of the content. You can use the variable "stage" after it is created in token create_stage.
    var lib=comp.getLibrary();
    var ss=comp.getSpriteSheet();
    var queue = evt.target;
    var ssMetadata = lib.ssMetadata;
    for(i=0; i<ssMetadata.length; i++) {
        ss[ssMetadata[i].name] = new createjs.SpriteSheet( {"images": [queue.getResult(ssMetadata[i].name)], "frames": ssMetadata[i].frames} )
    }

    try {
        exportRoot = new lib.top(); // this always error when in Edge
    }
    catch(err) {
        showHide('gif_alternative','block'); // this will show a gif animation alternative for the html5 animation
    }
    
    stage = new lib.Stage(canvas);	
    //Registers the "tick" event listener.
    fnStartAnimation = function() {
        stage.addChild(exportRoot);
        createjs.Ticker.framerate = lib.properties.fps;
        createjs.Ticker.addEventListener("tick", stage);
    }	    
    //Code to support hidpi screens and responsive scaling.
    AdobeAn.makeResponsive(false,'both',false,1,[canvas,anim_container,dom_overlay_container]);	
    AdobeAn.compositionLoaded(lib.properties.id);
    fnStartAnimation();
}

function showHide(elem_id, display_condition) {
    var x = document.getElementById(elem_id);
    x.style.display = display_condition;
}