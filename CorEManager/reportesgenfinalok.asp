<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/reportes.inc" -->
<!-- #INCLUDE FILE="../coress20/IncFile/CorStyles.inc" -->
<%


 Response.CacheControl = "no-cache"
 Response.AddHeader "Pragma", "no-cache"
 Response.Expires = -1


ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
  Response.Write sHTML
else

call ofv.browser()



set oPrtG = new ClsReportes

Set oPrtG.Ofv = Ofv



sbloque1 =    oPrtG.Generar()



'response.write request.form("formato")
IF cstr(request.form("formato")&"") ="HTML" THEN

oPrtG.shead=replace(oPrtG.shead,"&&","")


%>
<html><Title>COR EManager - Reporte de Esquemas Normativos</title>
<style>P {margin-top :0px;margin-bottom:0 px;};</style>

<%

response.write corstyler()

%>
<body onload='renderPag();'>
<div id="divhead"  style='visibility:hidden; '>
<%=oPrtG.shead%>
</div>

<div id="divheader"  style='visibility:hidden; '>
  <%=oPrtG.srephead%>
</div>
<div id="divbody" name="divbody"   style='visibility:hidden; '>
<%=sbloque1%>

</div>

<form id=MSForm  style='display:none; ' target="_new"  method="post" action="vert.asp" onsubmit='asignar();'>
<textarea id="txt1" name="txt1" style='display:none; '>
</textarea>
<input type=submit>
</form>
</body>

<script languaje=javascript>

var Obody  =document.getElementById('divbody');
var OHeader=document.getElementById('divheader')
var OHead=document.getElementById('divhead')


//var AltoPagina = 900 - (OHeader.childNodes(0).clientHeight + OHead.childNodes(0).clientHeight);
var AltoPagina = 848 - (OHeader.childNodes(0).clientHeight + OHead.childNodes(0).clientHeight);

//AltoPagina=AltoPagina + 15;
AltoPagina=AltoPagina + 18;
var restoimg=0;
//var celpad=2;
var sdiv = "<div style='page-break-after:always;'>&nbsp;</div>";
var sdivprn='';//'<div class=prt1 width=100%>';
var sRepHead=OHeader.innerHTML;
var sHead=OHead.innerHTML;
var Pagina = 1;
var OHead2=clonarTag2(OHead);
OHead2=sRepHead + OHead2.replace('XX',Pagina);

var sHTML= OHead2;




var nSubtotal=0;

 var flag =0;

 function asignar()
{
 // document.getElementById("txt1").value=document.getElementById(editor).innerHTML;
   BreakItUp();

  return false;
  }
 function BreakItUp()
 {
   //Set the limit for field size.
   var FormLimit = 102399;

   //Get the value of the large input object.
   var TempVar = new String;
   TempVar = document.getElementById("txt1").value;

   //If the length of the object is greater than the limit, break it
   //into multiple objects.
   if (TempVar.length > FormLimit)
   {
     document.getElementById("txt1").value = TempVar.substr(0, FormLimit);
     TempVar = TempVar.substr(FormLimit);

     while (TempVar.length > 0)
     {
       var objTEXTAREA = document.createElement("TEXTAREA");
       objTEXTAREA.name = "txt1";
       objTEXTAREA.value = TempVar.substr(0, FormLimit);
       objTEXTAREA.style.display = 'none';
       document.MSForm.appendChild(objTEXTAREA);

       TempVar = TempVar.substr(FormLimit);
     }
   }
 }



 function calcPorcentaje(alto,porc)
 {
  var result;
  result=((porc*alto)/100);
  return result
 // alert(result);
  }


function Recorrer2(Ocol)
{
   var oElemento0=Ocol.childNodes;
   var nHeight=0;
   var MAXHEIGHT = AltoPagina;
   var celpad=1;


for(f=0;f< oElemento0.length;f++)

{

 Otag=oElemento0(f);

 var flag =0;
 switch (Otag.tagName)


    {
     case "TABLE":
                 var cellpadding=Otag.cellPadding;
                  if (cellpadding==undefined)
                  {
                  cellpadding=1;
                  }
                  //alert(Otag.clientHeight);
                  parseTabla(Otag,parseFloat(cellpadding));
             break;

     case "IMG":
             if (Otag.clientHeight > AltoPagina - nSubtotal)

	        {
	          var porc=calcPorcentaje( Otag.height,80)
	       //   alert(porc);
	          if (porc <= AltoPagina - nSubtotal)
	          {
                  //Otag.height=AltoPagina - nSubtotal  - restoimg ;
                  Otag.height=porc  - restoimg ;
                  }
                  else
                  {
                    nSubtotal=0;
		    Pagina++;
		    var OHead2=clonarTag2(OHead);
		    OHead2=OHead2.replace('XX',Pagina);
		    //sHTML = sHTML + '</table>' + sdiv + sRepHead +  OHead2  + clonarTag(oTabla);
		    if (typeof (oTabla) != "undefined")
			    {

			    sHTML = sHTML + '</table>' + sdiv + sRepHead +  OHead2  + clonarTag(oTabla);
			    }
			     else
			    {
			  //   alert('1');
	               sHTML = sHTML  + sdiv + sRepHead +  OHead2  ;
		    		     }
		      if (Otag.height > AltoPagina - restoimg - (celpad * 2))
		         {
		   	  Otag.height=AltoPagina - restoimg - (celpad * 2) ;
		         }
                  }
                }

			 nSubtotal=  nSubtotal  + Otag.clientHeight  ;
			 if (nSubtotal    < AltoPagina )
				 {
				   sHTML = sHTML  + Otag.outerHTML;
				 }
		     else
				 {
				  nSubtotal=0;
				  Pagina++;
				  var OHead2=clonarTag2(OHead);
				  OHead2=OHead2.replace('XX',Pagina);
				  sHTML = sHTML  + sdiv  + sRepHead +  OHead2  + Otag.outerHTML;
	         	  	 }

            break;

      case "DIV":
			   nSubtotal=0;
			   Pagina++;
			   var OHead2=clonarTag2(OHead);
			   OHead2=OHead2.replace('XX',Pagina);
			   sHTML = sHTML  + sdiv + sRepHead +  OHead2 ;
	   break;

    }




} // fin for f=0



}


function parseTabla(Otag,celpad)
{
  var oTabla =Otag;
  sHTML = sHTML   + clonarTag(oTabla);

  if (Otag.clientHeight!=undefined)
   {

        for(var irow=0;irow < oTabla.rows.length;irow++)
	 {
            var alto=false;
            var oFila=oTabla.rows(irow);

             for (var cols=0;cols<oFila.childNodes.length;cols++)
             {
              var Ocell=oFila.childNodes(cols)
              if (Ocell.childNodes.length >= 1 )
               {
               if (Ocell.childNodes(0).tagName=='IMG')
                 {
                  if (Ocell.childNodes(0).id=='IMD')
                  {

                      if (Ocell.childNodes(0).width > 500)
					   {Ocell.childNodes(0).width=500}
					  else
					   {Ocell.childNodes(0).width=Ocell.childNodes(0).width}


                     if (Ocell.childNodes(0).height > AltoPagina - nSubtotal)

                        {

                           var porc;
                           porc=calcPorcentaje( Ocell.childNodes(0).height,80)
			 //  alert(porc);
			   if (porc <= AltoPagina - nSubtotal)
			     {
			         Ocell.childNodes(0).height=porc  - restoimg ;
			    //     alert('1');
                             }
                             else
			       {
			    //    alert('2');
				nSubtotal=0;
			    	Pagina++;
			    	var OHead2=clonarTag2(OHead);
			    	OHead2=OHead2.replace('XX',Pagina);
			     	sHTML = sHTML + '</table>' + sdiv + sRepHead +  OHead2  + clonarTag(oTabla);
			     	if (Ocell.childNodes(0).height > AltoPagina - restoimg - (celpad * 2))
			     	 {
				 Ocell.childNodes(0).height=AltoPagina - restoimg - (celpad * 2) ;
				 }
                  		}


                           //Ocell.childNodes(0).height=AltoPagina - nSubtotal - restoimg - (celpad * 2) ;


                        }

                  }
                  }
                 }
             }

               nSubtotal= celpad + nSubtotal  + oFila.clientHeight ;


              if (nSubtotal     < AltoPagina )
                {
                  sHTML = sHTML  + oFila.outerHTML;

                   if  (irow == oTabla.rows.length-1)
                        {

                         sHTML = sHTML  + '</table>';

                        }

                }
              else
                {
                 nSubtotal=0;
                 Pagina++;

                 var OHead2=clonarTag2(OHead);
                  OHead2=OHead2.replace('XX',Pagina);
                  sHTML = sHTML + '</table>' + sdiv + sRepHead +  OHead2  + clonarTag(oTabla) +  oFila.outerHTML;


                }

              }


        }

}








function clonarTag(Otag)
{

 var txt = Otag.cloneNode().outerHTML;
  return txt.replace('</TABLE>','');


}



function clonarTag2(Otag)
{

 var txt = Otag.cloneNode(true).innerHTML;

 return txt;


}



// main

function renderPag()
{
Recorrer2(Obody);
//document.body.innerHTML='';
for(var p=1;p<=Pagina;p++)
{
 sHTML=sHTML.replace("XY",Pagina);
}
document.body.innerHTML=sHTML;
//document.getElementById('txt1').value=sHTML ;

//document.getElementById('MSForm').submit();
//	document.write(sHTML);
}
</script>


<%



set oprt=nothing
ELSE
 Response.addHeader  "Content-Type", "application/msword"
 %>



<%
logo=replace(Session("clilogop"),"../","HTTP://SERVER-20/CORV34DES/")
oPrtG.RCliNomb=REPLACE(oPrtG.RCliNomb," ","_")
oPrtG.titdoc=REPLACE(oPrtG.titdoc," ","_")

%>


<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:w="urn:schemas-microsoft-com:office:word"
xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<meta name=ProgId content=Word.Document>
<meta name=Generator content="Microsoft Word 11">
<meta name=Originator content="Microsoft Word 11">
<link rel=File-List href="reportesgen_archivos/filelist.xml">
<!--[if gte mso 9]><xml>
 <o:DocumentProperties>
  <o:Author>WinXP</o:Author>
  <o:LastAuthor>WinXP</o:LastAuthor>
  <o:Revision>1</o:Revision>
  <o:TotalTime>0</o:TotalTime>
  <o:Created>2008-05-19T15:41:00Z</o:Created>
  <o:LastSaved>2008-05-19T15:41:00Z</o:LastSaved>
  <o:Pages>1</o:Pages>
  <o:Company>Microsoft</o:Company>
  <o:Lines>1</o:Lines>
  <o:Paragraphs>1</o:Paragraphs>
  <o:Version>11.5606</o:Version>
 </o:DocumentProperties>
</xml><![endif]--><!--[if gte mso 9]><xml>
 <w:WordDocument>
  <w:View>Print</w:View>
  <w:HyphenationZone>21</w:HyphenationZone>
  <w:PunctuationKerning/>
  <w:ValidateAgainstSchemas/>
  <w:SaveIfXMLInvalid>false</w:SaveIfXMLInvalid>
  <w:IgnoreMixedContent>false</w:IgnoreMixedContent>
  <w:AlwaysShowPlaceholderText>false</w:AlwaysShowPlaceholderText>
  <w:Compatibility>
   <w:BreakWrappedTables/>
   <w:SnapToGridInCell/>
   <w:WrapTextWithPunct/>
   <w:UseAsianBreakRules/>
   <w:DontGrowAutofit/>
  </w:Compatibility>
  <w:BrowserLevel>MicrosoftInternetExplorer4</w:BrowserLevel>
 </w:WordDocument>
</xml><![endif]--><!--[if gte mso 9]><xml>
 <w:LatentStyles DefLockedState="false" LatentStyleCount="156">
 </w:LatentStyles>
</xml><![endif]-->
<style>
<!--
 /* Style Definitions */
 p.MsoNormal, li.MsoNormal, div.MsoNormal
	{mso-style-parent:"";
	margin:0cm;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:12.0pt;
	font-family:"Times New Roman";
	mso-fareast-font-family:"Times New Roman";}
p.MsoHeader, li.MsoHeader, div.MsoHeader
	{margin:0cm;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	tab-stops:center 212.6pt right 425.2pt;
	font-size:12.0pt;
	font-family:"Times New Roman";
	mso-fareast-font-family:"Times New Roman";}
p.MsoFooter, li.MsoFooter, div.MsoFooter
	{margin:0cm;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	tab-stops:center 212.6pt right 425.2pt;
	font-size:12.0pt;
	font-family:"Times New Roman";
	mso-fareast-font-family:"Times New Roman";}
 /* Page Definitions */

@page
	{mso-footnote-separator:url("reportesgen_archivos/footer.asp?logo=<%=oPrtG.RCliLogo%>&RCliNomb=<%=oPrtG.RCliNomb%>&STITULO=<%=oPrtG.titdoc%>") fs;
	mso-footnote-continuation-separator:url("reportesgen_archivos/footer.asp?logo=<%=oPrtG.RCliLogo%>&RCliNomb=<%=oPrtG.RCliNomb%>&STITULO=<%=oPrtG.titdoc%>") fcs;
	mso-endnote-separator:url("reportesgen_archivos/footer.asp") es;
	mso-endnote-continuation-separator:url("reportesgen_archivos/footer.asp?logo=<%=oPrtG.RCliLogo%>&STITULO=<%=oPrtG.titdoc%>") ecs;}
@page Section1

	{size:595.3pt 841.9pt;
	margin:1.0cm 1.0cm 1.0cm 1.0cm;
	mso-header-margin:35.4pt;
	mso-footer-margin:35.4pt;
	mso-even-header:url("reportesgen_archivos/footer.asp?logo=<%=oPrtG.RCliLogo%>&RCliNomb=<%=oPrtG.RCliNomb%>&STITULO=<%=oPrtG.titdoc%>") eh1;
	mso-header:url("reportesgen_archivos/footer.asp?logo=<%=oPrtG.RCliLogo%>&RCliNomb=<%=oPrtG.RCliNomb%>&STITULO=<%=oPrtG.titdoc%>") h1;
	mso-even-footer:url("reportesgen_archivos/footer.asp?logo=<%=oPrtG.RCliLogo%>&RCliNomb=<%=oPrtG.RCliNomb%>&STITULO=<%=oPrtG.titdoc%>") ef1;
	mso-footer:url("reportesgen_archivos/footer.asp?logo=<%=oPrtG.RCliLogo%>&RCliNomb=<%=oPrtG.RCliNomb%>&STITULO=<%=oPrtG.titdoc%>") f1;
	mso-first-header:url("reportesgen_archivos/footer.asp?logo=<%=oPrtG.RCliLogo%>&RCliNomb=<%=oPrtG.RCliNomb%>&STITULO=<%=oPrtG.titdoc%>") fh1;
	mso-first-footer:url("reportesgen_archivos/footer.asp?logo=<%=oPrtG.RCliLogo%>&RCliNomb=<%=oPrtG.RCliNomb%>&STITULO=<%=oPrtG.titdoc%>") ff1;
	mso-paper-source:0;}
div.Section1
	{page:Section1;}

	</style>
	<!--[if gte mso 10]>
	<style>
	 /* Style Definitions */
	 table.MsoNormalTable
		{mso-style-name:"Tabla normal";
		mso-tstyle-rowband-size:0;
		mso-tstyle-colband-size:0;
		mso-style-noshow:yes;
		mso-style-parent:"";
		mso-padding-alt:0cm 5.4pt 0cm 5.4pt;
		mso-para-margin:0cm;
		mso-para-margin-bottom:.0001pt;
		mso-pagination:widow-orphan;
		font-size:10.0pt;
		font-family:"Times New Roman";
		mso-ansi-language:#0400;
		mso-fareast-language:#0400;
		mso-bidi-language:#0400;}
	</style>
	<![endif]-->

	</head>


<body lang=ES style='tab-interval:35.4pt'>

<div class=Section1>


<%


   response.write CorStyler2()
   response.write  sbloque1

%>
</div>

</body>

</html>



<%
end if

END IF

%>