<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
   ' cc=ConUsrset()
    op=request.querystring("op")
    empre=request.querystring("empre")
    shtml = "<html> " & CorStyler2()
   ' shtml = CorStyler2() & ofv.FormHeader("")
    Whtml=""
    If (SESSION("UN") & "") = "" Then


			xxxx = NEWEMPRE("NO")


			SEMPRE = ATREMP(0, 0)
			SESSION("UN") = SEMPRE & ""

			If (SESSION("UN") & "") = "" Then SESSION("UN") = 1


	End If

    Whtml = Whtml & " <link rel=""stylesheet"" type=""text/css"" href=""../v48/line-awesome/css/MENU.asp"">"
    Whtml = Whtml & " <link rel=""stylesheet"" type=""text/css"" href=""../v48/line-awesome/css/line-awesome.css"">"
	Whtml = Whtml & " <script src='../v48/line-awesome/jquery.min.js'></script>"
    Whtml = Whtml & " <script src='../v48/line-awesome/jquery.balloon.js'></script>"
    Whtml = Whtml & " <script src='../v48/line-awesome/jqex.js'></script>"
    Whtml = Whtml & " <script src='../v48/line-awesome/jqestilo.js'></script>"



     sHTML=sHTML & Whtml
    sHTML = sHTML & "<body id=bodyM name=bodyM class=xx bgcolor=" & session("cmfmenu") &"><div id=conteM>" '& corpaneltop()

   sHTML = sHTML '& "<div align=right xstyle='position:absolute;top:0px;width:100%;' xonclick=cerrar()><i class='la la-bars'></i></div>"
   select case session("menuex")
    case "S"
     sHTML = sHTML  & ARMAM2()
    case else
     sHTML = sHTML  & ARMAM()
    end select

    	WSPFRMN="frm"
    	WSPFORME3=""
    	wjs=""
    if op="ex" then
      WSPFORME3 = "<FORM ID='" & WSPFRMN & "3' NAME='" & WSPFRMN & "3' ACTION='../coremanager/novedmenu.asp?EMPRE=" & empre & "' TARGET='info' style='display:none;' METHOD='POST' ></FORM>"
      wjs= "document." & WSPFRMN & "3.submit();"
      sHTML=shtml & "<script>"
	  sHTML =shtml & "window.addEventListener('load', function(){"

	  sHTML =shtml &  wjs

	   sHTML =shtml & "});</script>"

    end if

    sHTML = sHTML & WSPFORME3& "</div></body></html>"

end if

response.write sHTML

%>
<script>
//alert(window.top.CorMoc);
var iframeBody = window.top.document;
var iframe = document.getElementById('bodyM');
//iframe.addEventListener("mouseover", func, false);
//iframe.addEventListener("mouseleave", func2, false);
function func2()
{
if (window.top.CorMoc=='O1'){
 window.top.CorMoc='C';
 document.getElementById('conteM').style.visibility = 'hidden';
  var iframeBody = window.top.document;
   var iframe = iframeBody.getElementById('fxt');
   iframe.cols='50,*,0,0,0,0,0,0';
 }
}
function func()
{

   var iframeBody = window.top.document;
   var iframe = iframeBody.getElementById('fxt');
   if (window.top.CorMoc=='C'){
     iframe.cols='250,*,0,0,0,0,0,0';
     window.top.CorMoc='O1';
     document.getElementById('conteM').style.visibility = '';

   }

  // alert(window.top.CorMoc);
}
window.addEventListener("load", function(){

    <%=wjsaa%>

});
 function ie_ver(){
  msie = parseInt((/msie (\d+)/.exec(navigator.userAgent.toLowerCase()) || [])[1]);
 if (isNaN(msie)) {
 msie = parseInt((/trident\/.*; rv:(\d+)/.exec(navigator.userAgent.toLowerCase()) || [])[1]);
 if (isNaN(msie)) {msie=-1;}
   }else {msie=-1;
} return msie;}

if (ie_ver() > 18 )
{
//var XCPT000O = document.getElementById('XCPT000O').innerHTML;
//var XCPT001O = document.getElementById('XCPT001O').innerHTML;
//var XCPT001= document.getElementById('XCPT001').innerHTML;
//var XCPT007O= document.getElementById('XCPT007O').innerHTML;
//var XCPT003O= document.getElementById('XCPT003O').innerHTML;



var z=document.getElementById('estil').innerHTML;
var storage = window.localStorage;
storage.clear();
storage.setItem("estil", z);
//storage.setItem("XCPT000O", XCPT000O);
//storage.setItem("XCPT001O", XCPT001O);
//storage.setItem("XCPT001", XCPT001);
//storage.setItem("XCPT007O", XCPT007O);
//storage.setItem("XCPT003O", XCPT003O);









}

</script>