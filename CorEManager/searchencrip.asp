<%@ LANGUAGE = VBScript.Encode %>




<%#@~^4gAAAA==@#@&@#@&EZKwXMkT4Y,`^#,5C.DPnDzPdYN,+T!yOy!Tc@#@&EK4kk~wMWL.m:PCU9PrYk~m/dKmblYN,dW!D^+,mG[Pkk~Nb/YMr8ED+9P;x[nMPOt~YDhd,W0~O4+~Mgi,Mn	+MlV,n!8Vbm@#@&ESr^x/ @#@&B?n~Y4+,lOYC^4+[P6rVP/6h5qH!cYaY,WKD~sWM+Pbx6GDslOkKx@#@&xEYAAA==^#~@%>


<!-- #include file="configuration.asp" -->

<%#@~^zwAAAA==@#@&B./aWxk+c^tmDd+DPx~T{m4CDk+Y@#@&vZKwHDro4O,`^#,elMY~KDXPJO9P+!Z+  TZ&@#@&BzV^~Dbo4YkP.nk+D7nN@#@&B:4r/,/KE.mn~1W[+,r/,/;8N+mO~DW~Y4n,Vr1+	/k	o,^W	NrYbWUd,lY,4YDw)J&AhSRHl.Y ^K: l!&lMYr^^+/&ebNnDcCkw@#@&0EIAAA==^#~@%>


<html>



<%#@~^NwAAAA==@#@&DnkwKx/RS.kD+~W6\ oKD:_nl9+Dvdn/kkKxcJoGM:HJ*b@#@&OBEAAA==^#~@%>

<style type="text/css">

#dhtmltooltip{
position: absolute;
left: -300px;
width: 250px;
border: 1px solid black;
padding: 2px;
background-color: lightyellow;
font-family: arial;
font-size:9pt;
font-style:italic;
visibility: hidden;
z-index: 100;
/*Remove below line to remove shadow. Below line should always appear last within this CSS*/
filter: progid:DXImageTransform.Microsoft.Shadow(color=gray,direction=135);
}

#dhtmlpointer{
position:absolute;
left: -300px;
z-index: 101;
visibility: hidden;
}

</style>

<script type="text/javascript">

/***********************************************
* Cool DHTML tooltip script II- © Dynamic Drive DHTML code library (www.dynamicdrive.com)
* This notice MUST stay intact for legal use
* Visit Dynamic Drive at http://www.dynamicdrive.com/ for full source code
***********************************************/

var offsetfromcursorX=12 //Customize x offset of tooltip
var offsetfromcursorY=10 //Customize y offset of tooltip

var offsetdivfrompointerX=10 //Customize x offset of tooltip DIV relative to pointer image
var offsetdivfrompointerY=14 //Customize y offset of tooltip DIV relative to pointer image. Tip: Set it to (height_of_pointer_image-1).

document.write('<div id="dhtmltooltip"></div>') //write out tooltip DIV
document.write('<img id="dhtmlpointer" src="imagenes/arrow2.gif">') //write out pointer image

var ie=document.all
var ns6=document.getElementById && !document.all
var enabletip=false
if (ie||ns6)
var tipobj=document.all? document.all["dhtmltooltip"] : document.getElementById? document.getElementById("dhtmltooltip") : ""

var pointerobj=document.all? document.all["dhtmlpointer"] : document.getElementById? document.getElementById("dhtmlpointer") : ""

function ietruebody(){
return (document.compatMode && document.compatMode!="BackCompat")? document.documentElement : document.body
}

function ddrivetip(thetext, thewidth, thecolor){
if (ns6||ie){
if (typeof thewidth!="undefined") tipobj.style.width=thewidth+"px"
if (typeof thecolor!="undefined" && thecolor!="") tipobj.style.backgroundColor=thecolor
tipobj.innerHTML=thetext
enabletip=true
return false
}
}

function positiontip(e){
if (enabletip){
var nondefaultpos=false
var curX=(ns6)?e.pageX : event.clientX+ietruebody().scrollLeft;
var curY=(ns6)?e.pageY : event.clientY+ietruebody().scrollTop;
//Find out how close the mouse is to the corner of the window
var winwidth=ie&&!window.opera? ietruebody().clientWidth : window.innerWidth-20
var winheight=ie&&!window.opera? ietruebody().clientHeight : window.innerHeight-20

var rightedge=ie&&!window.opera? winwidth-event.clientX-offsetfromcursorX : winwidth-e.clientX-offsetfromcursorX
var bottomedge=ie&&!window.opera? winheight-event.clientY-offsetfromcursorY : winheight-e.clientY-offsetfromcursorY

var leftedge=(offsetfromcursorX<0)? offsetfromcursorX*(-1) : -1000

//if the horizontal distance isn't enough to accomodate the width of the context menu
if (rightedge<tipobj.offsetWidth){
//move the horizontal position of the menu to the left by it's width
tipobj.style.left=curX-tipobj.offsetWidth+"px"
nondefaultpos=true
}
else if (curX<leftedge)
tipobj.style.left="5px"
else{
//position the horizontal position of the menu where the mouse is positioned
tipobj.style.left=curX+offsetfromcursorX-offsetdivfrompointerX+"px"
pointerobj.style.left=curX+offsetfromcursorX+"px"
}

//same concept with the vertical position
if (bottomedge<tipobj.offsetHeight){
tipobj.style.top=curY-tipobj.offsetHeight-offsetfromcursorY+"px"
nondefaultpos=true
}
else{
tipobj.style.top=curY+offsetfromcursorY+offsetdivfrompointerY+"px"
pointerobj.style.top=curY+offsetfromcursorY+"px"
}
tipobj.style.visibility="visible"
if (!nondefaultpos)
pointerobj.style.visibility="visible"
else
pointerobj.style.visibility="hidden"
}
}

function hideddrivetip(){
if (ns6||ie){
enabletip=false
tipobj.style.visibility="hidden"
pointerobj.style.visibility="hidden"
tipobj.style.left="-1000px"
tipobj.style.backgroundColor=''
tipobj.style.width=''
}
}

document.onmousemove=positiontip

</script>
<!-- #include file="CYiderSearch.asp" -->
<!-- #include file="search_button_input.asp" -->

<!-- #include file="search_include.asp" -->

<%#@~^YgAAAA==@#@&@#@&M+kwW	/ hMkO+,WW-c4WDGxDlAjj(vJZJSJESrJSJrSJr~EEBJJSEr~EJBE	WEBbTD+TlMSA!/^lM~^n^Nl*@#@&axkAAA==^#~@%>


</body>
</html>

<%#@~^dAIAAA==@#@&@#@&b0,o{k+D~@!@*PTPmx[~&x?D.`T{EMsmYK{kwrNn.BPEFy{RZRT 8J#~@!@*PTPD4x@#@&P,D+kwKU/R	DbYn~74mMs0,[Pr@!dmMkaY~VCUTECoxJr9C-m?m.raYEJ@*E@#@&~,D/wKxknRSDrYPEC^+DDcB:t+,erND,tC/~WKEUN,Cx,+..KDR~L|/nY,rkPE,[,o{k+D~[,Jv#pJ@#@&,PDdwKx/ ADbYP-4^.^0~[,E@!J/^.bwY@*E@#@&~PUn	N3MDKDIwK.Y,JL{k+OE@#@&+	[Pb0@#@&@#@&k6PT{GwnU,@!@*PZ~l	N~(	?Y.cT{;D^mDWmkwbN+M~,EFyG !c! qr#P@!@*PZPY4nU@#@&P,Dn/2G	/nRq.kD+~-(mDsW,[~J@!d1DraY,Vl	o!Co'EJxl-CUmDb2YrJ@*r@#@&P,D/2WUdRADbO+,JCsDYcv:tnPIr9+.,tm/P6W!UN,lUPD.GMRPTmWa+x,rdPrPLPL{G2x~[,EB*iE@#@&PP.nkwGxknch.bYP\(mMs0,[~J@!zd^MkwD@*J@#@&P,jnx92MDGD]naW.Y,Eo|W2n	J@#@&n	N~k6@#@&@#@&kLQAAA==^#~@%>