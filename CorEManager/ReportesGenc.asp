<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/reportes.inc" -->
<!-- #INCLUDE FILE="../coress20/IncFile/CorStyles.inc" -->
<%



ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
  Response.Write sHTML
else

call ofv.browser()


  ' Conset(session("CONSET"))
  ' response.write CorStyler()
 
''' set oPrtG = server.createobject("CorGeneratorXP.Reportes")
''' set oPrtG = server.createobject("CorGeneratorEX.Reportes")

set oPrtG = new ClsReportes ' server.createobject("CorGeneratorEXLR3.Reportes")

'Set oPrtG.Request = Request
'Set oPrtG.Server = Server
'Set oPrtG.Response = Response
'Set oPrtG.Session = Session
Set oPrtG.Ofv = Ofv



sbloque1 =  oPrtG.Generar() 
'sbloque1=replace(sbloque1,"<P>",chr(13) & "<td >")
'sbloque1=replace(sbloque1,"</P>",chr(13) & "</TD>")
'sbloque1=replace(sbloque1,"</TR>",chr(13) & "</TR>")
'sbloque1=replace(sbloque1,"</TABLE>",chr(13) & "</TABLE>")



IF request.form("formato")="HTML" or request.form("formato")="" THEN

oPrtG.shead=replace(oPrtG.shead,"&&","")


%>
<html><Title>COR EManager - Reporte de Esquemas Normativos</title><meta HTTP-EQUIV="expires" CONTENT="60" >
<style>P {margin-top :0px;margin-bottom:0 px;};</style>
<%
 
response.write corstyler()

%>

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
		    sHTML = sHTML + '</table>' + sdiv + sRepHead +  OHead2  + clonarTag(oTabla);
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
</script>


<%



set oprt=nothing
ELSE
strBody ="<html " 
strBody =strBody & "xmlns:o='urn:schemas-microsoft-com:office:office' " 

strBody =strBody &        "xmlns:w='urn:schemas-microsoft-com:office:word' "
'strBody =strBody &        "xmlns:w='http://schemas.microsoft.com/office/word/2003/wordml' "

strBody =strBody &       " xmlns='http://www.w3.org/TR/REC-html40'>"

strBody =strBody &       "<head><title></title>"




strBody =strBody & " <!--[if gte mso 9]>" 

strBody =strBody &     "<xml>"

strBody =strBody &                         "<w:WordDocument>" 

strBody =strBody &                        "<w:View>Print</w:View>" 
strBody =strBody &                        "<w:Zoom>90</w:Zoom>" 

strBody =strBody &                       "<w:DoNotOptimizeForBrowser/>" 
'strBody =strBody & "<w:AttachedTemplate href="&chr(34) &"c:\pruebaprn\NAME.dot" & chr(34) &"/></w:AttachedTemplate>"

strBody =strBody &                      "</w:WordDocument>" 

strBody =strBody &                     "</xml>" 

strBody =strBody &                     "<![endif]-->"



'we can tweak the MsoFooter class that is referenced by the footer, as required

strBody =strBody &"<style>" 

strBody =strBody &                       "<!-- /* Style Definitions */" 

strBody =strBody &                     "p.MsoFooter, li.MsoFooter, div.MsoFooter" 

strBody =strBody &                    "{margin:0in;" 

strBody =strBody &                    "margin-bottom:.0001pt;" 

strBody =strBody &                   "mso-pagination:widow-orphan;" 

strBody =strBody &                   "tab-stops:center 3.0in right 6.0in;" 

strBody =strBody &                  "font-size:12.0pt;}"




strBody =strBody &"@page Section1" 

strBody =strBody &                    "   {size:8.5in 11.0in; "

strBody =strBody &                "   margin:1.0in 1.0in 1.0in 1.0in ; " 

   strBody =strBody &                "   mso-header-margin:.5in; " 

strBody =strBody &                "   mso-footer: f1;" 

strBody =strBody &               "   mso-footer-margin:.1in; mso-paper-source:0;}" 

strBody =strBody &               " div.Section1" 

strBody =strBody &               "   {page:Section1;}" 

strBody =strBody &            "-->" 

strBody =strBody &         "</style></head>"



strBody =strBody & "<body style='tab-interval:.5in'>"

strBody =strBody &                       "<div class=Section1>"

'  strBody =strBody &                      "<h1>A dynamically generated document with Footer</h1>" 
' strBody =strBody &                      "<p style='color:red'><I> This doc was generated on "
'  strBody =strBody &                     Now & "</I></p><hr>"



'We are building up a big string here so that the generated doc runs into multiple pages.

' For counter = 1 To 50



'     strBody =strBody &"Lorem ipsum dolor sit amet, consectetur adipisicing elit, " 

'      strBody =strBody &              "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua."

' Next
sbloque1=replace(sbloque1,"SRC="," SRC='HTTP://SERVER-41/CoRESV3/COREMANAGER/")
sbloque1=replace(sbloque1,"src="," SRC='HTTP://SERVER-41/CORESV3/COREMANAGER/")
sbloque1=replace(sbloque1,"/COREMANAGER/'" ,"/COREMANAGER/")
sbloque1=replace(sbloque1,"/COREMANAGER/"& CHR(34) ,"/COREMANAGER/")
sbloque1=replace(sbloque1,"JPG" & chr(34),"JPG'")
strBody =strBody  &sbloque1



strBody =strBody &"</div>"



'Word marks and stores information for simple fields by means of the Span element with the 

'mso-field-code style. Refer to the topic "Fields" in the Office XML & HTML Reference

strBody =strBody &"<div style='mso-element:footer' id=f1>" 

 'strBody =strBody &                     " <p class=MsoFooter><span>paginaaaaaaaaa</span></p>" 

 strBody =strBody & " pagina &nbsp;&nbsp; <span style='mso-field-code:"" PAGE ""'></span>&nbsp;&nbsp;" 
 strBody =strBody & "de &nbsp;&nbsp;<span style='mso-field-code:"" NUMPAGES ""'></span>"
 strBody =strBody &  " </div>"






'strBody =strBody &                  "</body></html>"





'Force this content to be downloaded 

'as a Word document with the name of your choice

Response.addHeader  "Content-Type", "application/msword"

'Response.addHeader "Content-disposition",  "attachment; filename=bbb.doc"

Response.Write strBody


END IF

END IF

%>