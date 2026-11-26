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


'response.write request.form("formato")
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
 Response.addHeader  "Content-Type", "application/msword"
 %>
 
 
'Response.addHeader "Content-disposition ",  " attachment;filename=bbb.doc"
'Session("cliente")
'strBody ="<html " 
'strBody =strBody & "xmlns:o='urn:schemas-microsoft-com:office:office' " 

'strBody =strBody &        "xmlns:w='urn:schemas-microsoft-com:office:word' "
'strBody =strBody &        "html xmlns:v='urn:schemas-microsoft-com:vml'"

'strBody =strBody &       " xmlns='http://www.w3.org/TR/REC-html40'>"

'strBody =strBody &       "<head><meta http-equiv=Content-Type content='text/html; charset=windows-1252'><link rel=File-List href='Doc1_archivos/filelist.xml'>"



'strBody =strBody & " <!--[if gte mso 9]>" 

'strBody =strBody &     "<xml>"
'
'strBody =strBody &                         "<w:WordDocument>" 

'strBody =strBody &                        "<w:View>Print</w:View>" 
'strBody =strBody &                        "<w:Zoom>100</w:Zoom>  <w:ValidateAgainstSchemas/>" 
'
'strBody =strBody &                       "<w:DoNotOptimizeForBrowser/> " 
'strBody =strBody & "<w:AttachedTemplate href="&chr(34) &"http://server-41/prueba/header.asp" & chr(34) &"/></w:AttachedTemplate>"

'strBody =strBody & "<w:evenAndOddHeaders/> "
'strBody =strBody &                      "</w:WordDocument>" 

'strBody =strBody &                     "</xml>" 

'strBody =strBody &                     "<![endif]-->"

<%
logo=replace(Session("clilogop"),"../","HTTP://SERVER-41/CORESV3/")

'we can tweak the MsoFooter class that is referenced by the footer, as required
%>

'strBody =strBody &"<style> " 

'strBody =strBody &                       "<!-- /* Style Definitions */" 

'strBody =strBody &                     " p.MsoFooter, li.MsoFooter, div.MsoFooter " 

'strBody =strBody &                    "{margin:0in; " 

'strBody =strBody &                    "margin-bottom:.0001pt;" 

'strBody =strBody &                   "mso-pagination:widow-orphan;" 

'strBody =strBody &                   "tab-stops:center 3.0in right 6.0in;" 

'strBody =strBody &                  "font-size:12.0pt;} "

'%>
<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:w="urn:schemas-microsoft-com:office:word"
xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<meta name=ProgId content=Word.Document>
<meta name=Generator content="Microsoft Word 11">
<meta name=Originator content="Microsoft Word 11">
<link rel=File-List href="Doc1_archivos/filelist.xml">
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
	{mso-footnote-separator:url("Doc1_archivos/header.htm") fs;
	mso-footnote-continuation-separator:url("Doc1_archivos/header.htm") fcs;
	mso-endnote-separator:url("Doc1_archivos/header.htm") es;
	mso-endnote-continuation-separator:url("Doc1_archivos/header.htm") ecs;}
@page Section1
	{size:595.3pt 841.9pt;
	margin:70.85pt 3.0cm 70.85pt 3.0cm;
	mso-header-margin:35.4pt;
	mso-footer-margin:35.4pt;
	mso-even-header:url("Doc1_archivos/header.htm") eh1;
	mso-header:url("Doc1_archivos/header.htm") h1;
	mso-even-footer:url("Doc1_archivos/header.htm") ef1;
	mso-footer:url("Doc1_archivos/header.htm") f1;
	mso-first-header:url("Doc1_archivos/header.htm") fh1;
	mso-first-footer:url("Doc1_archivos/header.htm") ff1;
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

	

 
'strBody =strBody &" @page Section1 " 

'strBody =strBody &                    "   {size:8.5in 11.0in; "

'strBody =strBody &                "   margin:0.5in 0.5in 0.5in 0.5in ; " 

'   strBody =strBody &                "   mso-header-margin:.5in; " 

'strBody =strBody &                "   mso-footer: f1 ;" 
' strBody =strBody &                "   mso-even-footer:url('Doc1_archivos/header.htm') ef1; " 
' strBody =strBody &                "   mso-footer:url('Doc1_archivos/header.htm') f1;  "
' strBody =strBody &                "   mso-header: f1;"
'strBody =strBody &                "  mso-even-footer: f1;"
'strBody =strBody & " 		       mso-footer-data :'&L Proprietary & Confidential &R Page &P of &N';  "
'strBody =strBody &               "   mso-footer-margin:.1in; mso-paper-source:0;}" 

'strBody =strBody &               " div.Section1 " 

'strBody =strBody &               "   {page:Section1;} " 


'strBody =strBody &" @page Section1" 

'strBody =strBody &                    "   {size:8.5in 11.0in; "

'strBody =strBody &                "   margin:1.0in 1.0in 1.0in 1.0in  " 
'strBody =strBody &                "   margin:1.0in 1.0in 1.0in 1.0in ; mso-page-orientation:landscape; " 
'  strBody =strBody &                "   mso-header-margin:.5in; " 

'strBody =strBody &                "   mso-footer:url('http://server-41/prueba/header.asp') f1;" 
' strBody =strBody &                "   mso-footer: f1;"
' strBody =strBody &                "   mso-header: h1;"
' strBody =strBody &                "  mso-even-footer: f1;"
'strBody =strBody &               "   mso-footer-margin:.1in; mso-paper-source:0;}" 

'strBody =strBody &               " div.Section3 " 

'strBody =strBody &               "   {page:Section4;} " 


'strBody =strBody &            "-->" 




'strBody =strBody &         "</style></head>"





 
 'header = "<div style='mso-element:footer;' id=f1 ><p class=MsoFooter> "
 'header =header & "<span style='mso-tab-count:2;'> pagina &nbsp;&nbsp;</span> <span style='mso-field-code:""PAGE ""'></span>&nbsp;&nbsp;" 
 'header =header & "de &nbsp;&nbsp;<span style='mso-field-code:NUMPAGES'></span></p>"
 'header =header &  "<span style='mso-element:field-begin'><span style='mso-no-proof:yes'>SS</span></span><span style='mso-element:field-end'></span> </div>"
 
 
 
 'header=""
' header =header &  "<div style='mso-element:footer' id=f1>"
 
' header =header &  "<div style='mso-element:frame;mso-element-wrap:around;mso-element-anchor-vertical:"
' header =header &  "paragraph;mso-element-anchor-horizontal:margin;mso-element-left:right;"
' header =header &  "mso-element-top:.05pt;mso-height-rule:exactly'>"
 
' header =header &  "<table cellspacing=0 cellpadding=0 hspace=0 vspace=0 align=right>"
'  header =header &  "<tr>"
'  header =header &  " <td valign=top align=left style='padding-top:0cm;padding-right:0cm;"
' header =header &  "  padding-bottom:0cm;padding-left:0cm'>"
'  header =header &  " <p class=MsoFooter style='mso-element:frame;mso-element-wrap:around;"
'  header =header &  " mso-element-anchor-vertical:paragraph;mso-element-anchor-horizontal:margin;"
' header =header &  "  mso-element-left:right;mso-element-top:.05pt;mso-height-rule:exactly'><!--[if supportFields]><span "
'  header =header &  " class=MsoPageNumber><span style='mso-element:field-begin'></span>PAGE<span "
' header =header &  "  style='mso-spacerun:yes'>  </span><span style='mso-element:field-separator'></span></span><![endif]--><span "
' header =header &  "  class=MsoPageNumber><span style='mso-no-proof:yes'>1</span></span><!--[if supportFields]><span "
' header =header &  "  class=MsoPageNumber><span style='mso-element:field-end'></span></span><![endif]--><span "
'  header =header &  " class=MsoPageNumber><o:p></o:p></span></p>"
'  header =header &  " </td>"
' header =header &  " </tr>"
' header =header &  "</table>"
' 
' header =header &  "</div>"
' 
'' header =header &  "<p class=MsoFooter style='margin-right:18.0pt'>Pagina</p>"
' 
' header =header &  "</div>"
 
 
 
 
 


 
strBody =strBody  &"<body >" 

strBody =strBody &                       "<div class=Section1>"   


sbloque1=replace(sbloque1,"SRC="," SRC='HTTP://SERVER-41/CoRESV3/COREMANAGER/")
sbloque1=replace(sbloque1,"src="," SRC='HTTP://SERVER-41/CORESV3/COREMANAGER/")
sbloque1=replace(sbloque1,"/COREMANAGER/'" ,"/COREMANAGER/")
sbloque1=replace(sbloque1,"/COREMANAGER/"& CHR(34) ,"/COREMANAGER/")
sbloque1=replace(sbloque1,"JPG" & chr(34),"JPG'")
strBody =strBody  &sbloque1







'Word marks and stores information for simple fields by means of the Span element with the 

'mso-field-code style. Refer to the topic "Fields" in the Office XML & HTML Reference
strBody =strBody&"</div>" 

'strBody =strBody &"<div style='mso-element:header' id=h1 >" 



 'strBody =strBody & "<span style='mso-tab-count:2'> pagina &nbsp;&nbsp; </span> <span style='mso-field-code:"" PAGE ""'></span>&nbsp;&nbsp;" 
 'strBody =strBody & "de &nbsp;&nbsp;<span style='mso-field-code:"" NUMPAGES ""'></span> <span style='mso-field-code:"" meta15 ""'>hhhh</span>"
 'strBody =strBody &  " </div>"
'
'strBody =strBody &                       "<div class=Section3>"

'strBody =strBody &                  "</body></html>"







' strBody =strBody &"<div style='mso-element:footer' id=f1><p class=MsoFooter> "
   '<p class=""MsoFooter"">" 


' strBody =strBody & "<span style='mso-tab-count:2'> pagina &nbsp;&nbsp; </span> <span style='mso-field-code:"" PAGE ""'></span>&nbsp;&nbsp;" 
' strBody =strBody & "de &nbsp;&nbsp;<span style='mso-field-code:NUMPAGES'></span><span style='mso-no-proof:yes'>1</span></p>"
' strBody =strBody &  " </div>"
'strBody =strBody &"<div style='mso-element:footer' id=f1 > <p class=MsoFooter>" 



 'strBody =strBody & " <span style='mso-tab-count:1'></span> <span style='mso-field-code:PAGE' ></span>" 
 'strBody =strBody & "&nbsp;&nbsp;<span style='mso-field-code:"" NUMPAGES ""'></span>"
 'strBody =strBody &  " </p></div>"

'strBody =strBody &                  "</div>"

'strBody =strBody & "<script>"
'strBody =strBody & "var y = document.createElement('div');"
'strBody =strBody & "y.id = 'f1'; y.style='mso-element:footer';"
'strBody =strBody & "document.write('1111111111111111');"


 
 strBody =strBody &                  " </body>"  
 strBody =strBody & "</html>"
response.write strBody

'Force this content to be downloaded 

'as a Word document with the name of your choice




'http://server-41/prueba/header.asp




END IF

END IF

%>