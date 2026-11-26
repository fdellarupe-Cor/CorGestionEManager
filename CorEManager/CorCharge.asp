<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

%>
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<META HTTP-EQUIV="expires" content='-1'>
<style>
v\:*     { behavior: url(#default#VML) }
o\:*     { behavior: url(#default#VML) }
.shape    { behavior: url(#default#VML) }
</style>

</head>
<body>


<%
  if Session("FULLGUI") = "N" then

     shtml = ""

     shtml = shtml & "<center><div ><br><br><br><br><br><br><H2><font face='tahoma' color='" & Session("CMFMENU") & "' >Cargando el documento,<br> aguarde por favor.....</H3></div></center>"

     response.write shtml
  else

%>



<table style='display:none;' >
<tr>
<td id='y1' style='background-color:#eeeeee;' >aaaa</td>
<td id='y2' style='background-color:#eeeeee;' >aaaa</td>
<td id='y3' style='background-color:#eeeeee;' >aaaa</td>
</tr>
</table>

<script type="text/javascript">

/***********************************************
* 3D Spinning Message Script- By Copyright (c) 2003 Peter Gehrig
* Website: http://www.24fun.com
* Script available at/modified by Dynamic Drive (http://www.dynamicdrive.com)
* This notice must stay intact for use
***********************************************/
 var CorMcol1;
 var CorMcol2;
 var CorMcol3;

 CorMcol1   = '<% response.write Session("APLSHW") %>';
 CorMcol2   = '<% response.write Session("CMFMENU") %>';
 CorMcol3   = '<% response.write Session("CMFPNBOT") %>';

document.getElementById('y1').style.backgroundColor=CorMcol1;
document.getElementById('y2').style.backgroundColor=CorMcol2;
document.getElementById('y3').style.backgroundColor=CorMcol3;

document.getElementById('y1').innerHTML=CorMcol1;
document.getElementById('y2').innerHTML=CorMcol2;
document.getElementById('y3').innerHTML=CorMcol3;

// Add as many messages as you like
var message=new Array("E-MANAGER", "Cargando Documento ...", "COR E-MANAGER ...", "Cargando Documento ...", "Cargando Documento ...")


// Set the outline-color. Add as many colors as you like
//var outlinecolor=new Array("#99ccee", "#99ccee")
var outlinecolor=new Array(CorMcol1, CorMcol1)

// Set fillcolors 1. Add as many colors as you like
//var fillcolor1=new Array("#6699cc", "#6699cc", "#6699cc", "#6699cc")
var fillcolor1=new Array(CorMcol2, CorMcol2, CorMcol2, CorMcol2)

// Set fillcolors 2. Add as many colors as you like
//var fillcolor2=new Array("#336699", "#336699", "#336699", "#336699")
var fillcolor2=new Array(CorMcol3, CorMcol3, CorMcol3, CorMcol3)

// Set the letter marking the circle
var circlemark=new Array("-")

// Set the width of the outline
var strkweight=1

// Set the waiting time between the messages (seconds)
var pause=2

// Set the strength of the opacity (transparency of letters)
var strengthopacity="50%"

// Set the size of the circle (values range from 0.1 to 1)
var circlesize=0.5

// Always keep messages in view even if page is scrolled? (DD added option)
var keepinview="yes"

// Do not edit below this line

mytruebody=(!window.opera && document.compatMode && document.compatMode!="BackCompat")? document.documentElement : document.body //Dynamicdrive added object

var outerwidth=mytruebody.clientWidth
var outerheight=mytruebody.clientHeight

var innerwidth=Math.floor(circlesize*outerwidth)
var innerheight=Math.floor(circlesize*outerheight)

var posleft=(outerwidth-innerwidth)/2
var postop=(outerheight-innerheight)/2

var path=new Array()
var i_message=0
var i_outlinecolor=0
var i_fillcolor1=0
var i_fillcolor2=0
var i_messagelength=0
var longestmessage=0
pause*=1000

var ie=document.getElementById&&document.all?1:0

for (i=0;i<=message.length-1;i++) {
	if (message[i].length>longestmessage) {
		longestmessage=message[i].length
	}
	longestmessage+=4
}

for (i=0;i<=message.length-1;i++) {

	var emptyspace=""
	var i_emptyspace=(longestmessage-message[i].length)/2
	for (ii=0;ii<=i_emptyspace;ii++) {
		emptyspace+=circlemark
	}
	message[i]=emptyspace+" "+message[i]+" "+emptyspace
}

function changeform() {
	if (keepinview=="yes") //DD added
	document.getElementById("roofid").style.top=mytruebody.scrollTop //DD added
	if (i_outlinecolor >= outlinecolor.length) {i_outlinecolor=0}
	if (i_fillcolor1 >= fillcolor1.length) {i_fillcolor1=0}
	if (i_fillcolor2 >= fillcolor2.length) {i_fillcolor2=0}
	document.getElementById('strokeid').color=outlinecolor[i_outlinecolor]
	document.getElementById('fillid').color=fillcolor1[i_fillcolor1]
	document.getElementById('fillid').color2=fillcolor2[i_fillcolor2]
	if (i_message < message.length) {tick()}
	else {document.getElementById('textpathid').string=""
	document.getElementById("roofid").style.display="none" //DD added
	}
}

function tick() {
	if (i_messagelength <= message[i_message].length) {
		var messagestringend=""
		var messagestring=message[i_message].substring(0, i_messagelength)+messagestringend
		document.getElementById('textpathid').string=messagestring
		var timer=setTimeout("tick()",50)
		i_messagelength++
	}
	else {
		clearTimeout(timer)
		i_messagelength=0
		i_message++
		i_outlinecolor++
		i_fillcolor1++
		i_fillcolor2++
		var timer=setTimeout("changeform()",pause)
	}

}

if (ie) {
	document.write('<div id="roofid" style="position:absolute;left:0px;top:0px;width:'+outerwidth+'px;height:'+outerheight+'px;overflow:hidden;">')
	document.write('<v:oval id="tc" style="position:absolute;top:'+postop+'px;left:'+posleft+'px;width:'+innerwidth+'px;height:'+innerheight+'px">')
	document.write('<v:shadow on="t" opacity="'+strengthopacity+'" weight="5px" />')
	document.write('<v:stroke id="strokeid" weight="'+strkweight+'pt" color="blue"/>')
	document.write('<v:fill id="fillid" on="True" color="'+fillcolor1[0]+'" color2="'+fillcolor2[0]+'" opacity="'+strengthopacity+'" opacity2="'+strengthopacity+'" type="gradient"/>')
	document.write('<v:path textpathok="t"/>')
	document.write('<v:textpath id="textpathid" on="t" id="mytp" style="font-family:\'verdana\';font-weight:bold; " fitpath="t" string=""/>')
	document.write('</v:oval></div>')
	if (window.attachEvent) //DD added code
	window.attachEvent("onload", changeform) //DD added code
	else
	window.onload=changeform
}
</script>

<%

 end if

%>



</body>

</html>

