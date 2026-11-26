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

    Whtml = Whtml & " <link rel=""stylesheet"" type=""text/css"" href=""../v48/line-awesome/css/MENU.asp"">"
    Whtml = Whtml & " <link rel=""stylesheet"" type=""text/css"" href=""../v48/line-awesome/css/line-awesome.css"">"
	Whtml = Whtml & " <script src='../v48/line-awesome/jquery.min.js'></script>"
    Whtml = Whtml & " <script src='../v48/line-awesome/jquery.balloon.js'></script>"
    Whtml = Whtml & " <script src='../v48/line-awesome/jqex.js'></script>"
    Whtml = Whtml & " <script src='../v48/line-awesome/jqestilo.js'></script>"



     sHTML=sHTML & Whtml
    sHTML = sHTML & "<body class=xx xcolor=white>" '& corpaneltop()

   sHTML = sHTML '& "<div align=right xstyle='position:absolute;top:0px;width:100%;' xonclick=cerrar()><i class='la la-bars'></i></div>"
   select case session("menuex")
    case "S"
     sHTML = sHTML  & ARMAsolapa(8623,"COM","")
    case else
     sHTML = sHTML  & ARMAM()
    end select

    	WSPFRMN="frm"
    	WSPFORME3=""
    	wjs=""
    if op="ex" then
    WSPFORME3 = "<FORM ID='" & WSPFRMN & "3' NAME='" & WSPFRMN & "3' ACTION='../coremanager/novedmenu.asp?EMPRE=" & empre & "' TARGET='info' style='display:none;' METHOD='POST' ></FORM>"
    wjs= "document." & WSPFRMN & "3.submit();"
    end if

    sHTML = sHTML & WSPFORME3& "</body></html>"

end if

response.write sHTML

%>
<script>

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