
var x=y=o=null;


if (document.layers) {
    window.captureEvents(Event.MOUSEDOWN | Event.MOUSEMOVE | Event.MOUSEUP);
    window.onmousedown=down, window.onmouseup=up, window.onmousemove=move;
}
else if (document.all)
    document.onmousedown=down, document.onmouseup=up, document.onmousemove=move;

function down(e) {

try {

    if ( document.all && window.event.srcElement.parentElement.parentElement.parentElement.parentElement)
     o=window.event.srcElement.parentElement.parentElement.parentElement.parentElement, x=window.event.offsetX, y=window.event.offsetY;
    else {
        if (document.layers) {
            if (document.layers['o'+e.target.name]) {
                o=document.layers['o'+e.target.name], x=e.layerX, y=e.layerY;
            }
            else return true;
        }
   }

   return false;
}
catch(er) {

}
}

function move(e) {

    var xxleft = 0;
    var xxtop  = 0;

    if (document.all && o)
     {
        xxleft = window.event.clientX-x-5;
        xxtop  = window.event.clientY-y-140;

        if (xxleft<5) {xxleft=5;}
        if (xxleft>665) {xxleft=665;}
        if (xxtop<5) {xxtop=5;}
        if (xxtop>325) {xxtop=325;}

        o.style.posLeft=xxleft, o.style.posTop=xxtop;
     }
    else if (document.layers && o)
        o.left=e.pageX-x, o.top=e.pageY-y;
    return false;
}

function up() { o=null; }
