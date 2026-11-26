
function AbrirVentanaEx(codigo,tipo) {
var n = location.href.toUpperCase().indexOf('COREMANAGER') ;
var Datosx ='Consultar.asp?DOCMODAL=CL&Codigo=' + codigo +  '&tipo='  + tipo +  '&MAINAPL=EM&Datos=*smenu=PORESTADO*lista=3';


Datosx=location.href.substring(0, n) + 'coremanager/' + Datosx;
//alert(Datosx);
var is_chrome = navigator.userAgent.toLowerCase().indexOf('chrome') > -1;
var ww='_self';
	if(is_chrome)
	{
	ww='edi';
	}


 var h=570;w=790 ;if (window.screen) {w = window.screen.availWidth;h = window.screen.availHeight-60;}  myWin= open(Datosx,ww,  'status=yes,toolbar=no,width='+w+',height='+h+',top=0,left=0,resizable=yes' );

 }