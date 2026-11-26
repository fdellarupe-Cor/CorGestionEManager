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


 Sbase =Left(request.servervariables("SCRIPT_NAME"), InstrRev(request.servervariables("SCRIPT_NAME"),"/"))

'  End If


   If session("URLPORT") = "N" Then
        Sbase = TipoSite("")  & request.servervariables("SERVER_NAME") & Sbase
   Else
        Sbase = TipoSite("")  & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT") & Sbase
   End If



IF REQUEST.FORM("formato") = "PDF" THEN
SHTML="<h3>GENERANDO PDF AGUARDE POR FAVOR....</H3><FORM style='visibility:hidden' ACTION=HTTP://SERVER-40:8055/ALepdf.ASPX METHOD=POST ID=PEPE>"

id=Request.querystring("ID")


sPerfil = Request.querystring("Perfil")
tipo = Request.querystring("tipobj")
archivo="pdf" & id & "_" & tipo & ".pdf"


 Set cn = OFV.conectar(OFV.strconn2)
strsql = "SELECT OBJETOS.codiden, OBJETOS.descri, TIPOBJETO.descri, OBJPRTY.estado, "
strsql = strsql & "OBJPRTY.fecvigen, OBJPRTY.fecbaja, OBJPRTY.fecemis, OBJPRTY.fecvto, "
strsql = strsql & "OBJPRTY.autor, estados.estado, objetos.objUN,textoimpresion, tipobjeto.perfil "
strsql = strsql & "FROM OBJETOS, TIPOBJETO, OBJPRTY, estados "
strsql = strsql & "where OBJETOS.tipobjeto = TIPOBJETO.tipobjeto and "
strsql = strsql & " OBJETOS.codobjeto = OBJPRTY.codobjeto and "
strsql = strsql & " OBJPRTY.estado = estados.objid "
strsql = strsql & " and OBJETOS.codobjeto = " & id
Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")
  sEsquema = rs(2)
  sEsquemaEx = rs(0) & " - " & rs(1)
   RCliLogo = Session("clilogop")
   RCliNomb = Session("cliente")
  If rs(10) > 1 Then
          strsql = "Select * from clienteplan where id = " & rs(10)
          Set rsc = OFV.crearconsultaEx(strsql, cn, 1, "")
          If Not rsc.EOF Then
              RCliLogo = "../COREss20/Empresa/" & rsc(5)
              if (rsc(11) & "" ) <> "" then RCliLogo = "../COREss20/Empresa/" & rsc(11)
              RCliNomb = rsc(1)
          End If
          Call OFV.cerrarconsulta(rsc)

     End If

 Call OFV.cerrarconsulta(rs)
 Call OFV.cerrarconn(cn)
sRepHead = RepHeader2pdf(sEsquema & " " & sEsquemaEx, "600")
 ruta=server.mappath(request.servervariables("script_name"))

 pos=instrrev(ruta,"\")
 ruta=mid(ruta,1,pos)


SHTML=SHTML &  "<INPUT SIZE=220 TYPE=text NAME=url value='"& Sbase &"/ReportesGenex2.asp?id="&id &"&tipobj="&tipo&"&perfil="&sperfil&"'>"

SHTML =SHTML & "<INPUT SIZE=220 TYPE=text NAME=archivo  value='" & archivo &"'>"

SHTML =SHTML &"<INPUT SIZE=220 TYPE=text NAME=destinourl value='"&Sbase&"/pnlprint/' >"
SHTML =SHTML & "<INPUT SIZE=220 TYPE=text NAME=path value='"& ruta &"\PnlPrint\'>"
SHTML =SHTML & "<INPUT SIZE=220 TYPE=text NAME=header value=" &  chr(34)  & sRepHead & chr(34) & ">"

SHTML =SHTML & "<INPUT SIZE=220 TYPE=text NAME=footer value=""<div style='text-align:right;font-family:Tahoma; font-size:9px;'>Página <span class=page></span> de <span class=topage></span></div>"">"

SHTML =SHTML & "</FORM> <SCRIPT>document.getElementById('PEPE').submit();</SCRIPT>"

response.write shtml
else

call ofv.browser()
dim jsHashPefiles
dim jsHashsectores
dim sperfiles
dim ssectores
aplcheck()
'response.write session("APLMDCLI")

set oPrtG = new ClsReportes

Set oPrtG.Ofv = Ofv


if  session("APLMDCLI") = "BGAL" or session("APLMDCLI") = "BANELCO" or session("APLMDCLI") = "BCR"  then
sbloque1 =     oPrtG.Generar() & "<script type='text/javascript' src='hashtable.js'></script><script>"& jsHashPefiles & ";" & jsHashsectores & ";change_span('',hashperfiles,'spanpf');change_span('',hashsectores,'spansec');</script>"
else
sbloque1 =     oPrtG.Generar() '& "<script type='text/javascript' src='hashtable.js'></script><script>"& jsHashPefiles & ";" & jsHashsectores & ";change_span('',hashperfiles,'spanpf');change_span('',hashsectores,'spansec');</script>"
end if


'If session("URLTYP") = "V" Then



IF request.form("formato")="HTML" or request.form("formato")="" THEN

oPrtG.shead=replace(oPrtG.shead,"&&","")


%>
<html><Title>COR EManager - Reporte de Esquemas Normativos</title><meta HTTP-EQUIV="expires" CONTENT="60" >
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


//var AltoPagina = 900 - (OHeader.childNodes[0].clientHeight + OHead.childNodes(0).clientHeight);
var AltoPagina = 848 - (OHeader.clientHeight + OHead.clientHeight);

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

 Otag=oElemento0[f];

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
            var oFila=oTabla.rows[irow];

             for (var cols=0;cols<oFila.childNodes.length;cols++)
             {
              var Ocell=oFila.childNodes[cols]
              if (Ocell.childNodes.length >= 1 )
               {
               if (Ocell.childNodes[0].tagName=='IMG')
                 {
                  if (Ocell.childNodes[0].id=='IMD')
                  {

                      if (Ocell.childNodes[0].width > 500)
					   {Ocell.childNodes[0].width=500}
					  else
					   {Ocell.childNodes[0].width=Ocell.childNodes[0].width}


                     if (Ocell.childNodes[0].height > AltoPagina - nSubtotal)

                        {

                           var porc;
                           porc=calcPorcentaje( Ocell.childNodes[0].height,80)

			   if (porc <= AltoPagina - nSubtotal)
			     {
			         Ocell.childNodes[0].height=porc  - restoimg ;
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
			     	if (Ocell.childNodes[0].height > AltoPagina - restoimg - (celpad * 2))
			     	 {
				 Ocell.childNodes[0].height=AltoPagina - restoimg - (celpad * 2) ;
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
  txt=txt.replace('</table>','')
  txt=txt.replace('</TABLE>','');
  return txt;

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
 Response.AddHeader "content-disposition", "attachment; filename=salida.doc"
 %>



<%
'logo=replace(Session("clilogop"),"../","HTTP://SERVER-20/CORV34DES/")
oPrtG.RCliNomb=REPLACE(oPrtG.RCliNomb," ","_")
oPrtG.titdoc=REPLACE(oPrtG.titdoc," ","_")
if session("APLMDCLI")="BC" THEN
%>
<html xmlns:v="urn:schemas-microsoft-com:vml"
xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:w="urn:schemas-microsoft-com:office:word"
xmlns:m="http://schemas.microsoft.com/office/2004/12/omml"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<base href="<%=sbase%>" />
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<meta name=ProgId content=Word.Document>
<meta name=Generator content="Microsoft Word 12">
<meta name=Originator content="Microsoft Word 12">
<link rel=File-List href="reportesgen_archivoss/filelist.xml">
<!--[if gte mso 9]><xml>
 <o:DocumentProperties>
  <o:Author>WinXP</o:Author>
  <o:LastAuthor>dm011099</o:LastAuthor>
  <o:Revision>2</o:Revision>
  <o:TotalTime>1</o:TotalTime>
  <o:Created>2012-05-14T16:52:00Z</o:Created>
  <o:LastSaved>2012-05-14T16:52:00Z</o:LastSaved>
  <o:Pages>1</o:Pages>
  <o:Words>358</o:Words>
  <o:Characters>1973</o:Characters>
  <o:Company>Microsoft</o:Company>
  <o:Lines>16</o:Lines>
  <o:Paragraphs>4</o:Paragraphs>
  <o:CharactersWithSpaces>2327</o:CharactersWithSpaces>
  <o:Version>12.00</o:Version>
 </o:DocumentProperties>
</xml><![endif]-->
<link rel=themeData href="reportesgen_archivos/themedata.thmx">
<link rel=colorSchemeMapping href="reportesgen_archivos/colorschememapping.xml">
<!--[if gte mso 9]><xml>
 <w:WordDocument>
  <w:View>Print</w:View>
  <w:TrackMoves>false</w:TrackMoves>
  <w:TrackFormatting/>
  <w:HyphenationZone>21</w:HyphenationZone>
  <w:PunctuationKerning/>
  <w:ValidateAgainstSchemas/>
  <w:SaveIfXMLInvalid>false</w:SaveIfXMLInvalid>
  <w:IgnoreMixedContent>false</w:IgnoreMixedContent>
  <w:AlwaysShowPlaceholderText>false</w:AlwaysShowPlaceholderText>
  <w:DoNotPromoteQF/>
  <w:LidThemeOther>ES-AR</w:LidThemeOther>
  <w:LidThemeAsian>X-NONE</w:LidThemeAsian>
  <w:LidThemeComplexScript>X-NONE</w:LidThemeComplexScript>
  <w:Compatibility>
   <w:BreakWrappedTables/>
   <w:SnapToGridInCell/>
   <w:WrapTextWithPunct/>
   <w:UseAsianBreakRules/>
   <w:DontGrowAutofit/>
   <w:SplitPgBreakAndParaMark/>
   <w:DontVertAlignCellWithSp/>
   <w:DontBreakConstrainedForcedTables/>
   <w:DontVertAlignInTxbx/>
   <w:Word11KerningPairs/>
   <w:CachedColBalance/>
  </w:Compatibility>
  <w:BrowserLevel>MicrosoftInternetExplorer4</w:BrowserLevel>
  <m:mathPr>
   <m:mathFont m:val="Cambria Math"/>
   <m:brkBin m:val="before"/>
   <m:brkBinSub m:val="&#45;-"/>
   <m:smallFrac m:val="off"/>
   <m:dispDef/>
   <m:lMargin m:val="0"/>
   <m:rMargin m:val="0"/>
   <m:defJc m:val="centerGroup"/>
   <m:wrapIndent m:val="1440"/>
   <m:intLim m:val="subSup"/>
   <m:naryLim m:val="undOvr"/>
  </m:mathPr></w:WordDocument>
</xml><![endif]--><!--[if gte mso 9]><xml>
 <w:LatentStyles DefLockedState="false" DefUnhideWhenUsed="false"
  DefSemiHidden="false" DefQFormat="false" LatentStyleCount="267">
  <w:LsdException Locked="false" QFormat="true" Name="Normal"/>
  <w:LsdException Locked="false" QFormat="true" Name="heading 1"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   QFormat="true" Name="heading 2"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   QFormat="true" Name="heading 3"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   QFormat="true" Name="heading 4"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   QFormat="true" Name="heading 5"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   QFormat="true" Name="heading 6"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   QFormat="true" Name="heading 7"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   QFormat="true" Name="heading 8"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   QFormat="true" Name="heading 9"/>
  <w:LsdException Locked="false" Priority="99" Name="footer"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   QFormat="true" Name="caption"/>
  <w:LsdException Locked="false" QFormat="true" Name="Title"/>
  <w:LsdException Locked="false" Priority="1" Name="Default Paragraph Font"/>
  <w:LsdException Locked="false" QFormat="true" Name="Subtitle"/>
  <w:LsdException Locked="false" QFormat="true" Name="Strong"/>
  <w:LsdException Locked="false" QFormat="true" Name="Emphasis"/>
  <w:LsdException Locked="false" Priority="99" Name="No List"/>
  <w:LsdException Locked="false" Priority="99" SemiHidden="true"
   Name="Placeholder Text"/>
  <w:LsdException Locked="false" Priority="1" QFormat="true" Name="No Spacing"/>
  <w:LsdException Locked="false" Priority="60" Name="Light Shading"/>
  <w:LsdException Locked="false" Priority="61" Name="Light List"/>
  <w:LsdException Locked="false" Priority="62" Name="Light Grid"/>
  <w:LsdException Locked="false" Priority="63" Name="Medium Shading 1"/>
  <w:LsdException Locked="false" Priority="64" Name="Medium Shading 2"/>
  <w:LsdException Locked="false" Priority="65" Name="Medium List 1"/>
  <w:LsdException Locked="false" Priority="66" Name="Medium List 2"/>
  <w:LsdException Locked="false" Priority="67" Name="Medium Grid 1"/>
  <w:LsdException Locked="false" Priority="68" Name="Medium Grid 2"/>
  <w:LsdException Locked="false" Priority="69" Name="Medium Grid 3"/>
  <w:LsdException Locked="false" Priority="70" Name="Dark List"/>
  <w:LsdException Locked="false" Priority="71" Name="Colorful Shading"/>
  <w:LsdException Locked="false" Priority="72" Name="Colorful List"/>
  <w:LsdException Locked="false" Priority="73" Name="Colorful Grid"/>
  <w:LsdException Locked="false" Priority="60" Name="Light Shading Accent 1"/>
  <w:LsdException Locked="false" Priority="61" Name="Light List Accent 1"/>
  <w:LsdException Locked="false" Priority="62" Name="Light Grid Accent 1"/>
  <w:LsdException Locked="false" Priority="63" Name="Medium Shading 1 Accent 1"/>
  <w:LsdException Locked="false" Priority="64" Name="Medium Shading 2 Accent 1"/>
  <w:LsdException Locked="false" Priority="65" Name="Medium List 1 Accent 1"/>
  <w:LsdException Locked="false" Priority="99" SemiHidden="true" Name="Revision"/>
  <w:LsdException Locked="false" Priority="34" QFormat="true"
   Name="List Paragraph"/>
  <w:LsdException Locked="false" Priority="29" QFormat="true" Name="Quote"/>
  <w:LsdException Locked="false" Priority="30" QFormat="true"
   Name="Intense Quote"/>
  <w:LsdException Locked="false" Priority="66" Name="Medium List 2 Accent 1"/>
  <w:LsdException Locked="false" Priority="67" Name="Medium Grid 1 Accent 1"/>
  <w:LsdException Locked="false" Priority="68" Name="Medium Grid 2 Accent 1"/>
  <w:LsdException Locked="false" Priority="69" Name="Medium Grid 3 Accent 1"/>
  <w:LsdException Locked="false" Priority="70" Name="Dark List Accent 1"/>
  <w:LsdException Locked="false" Priority="71" Name="Colorful Shading Accent 1"/>
  <w:LsdException Locked="false" Priority="72" Name="Colorful List Accent 1"/>
  <w:LsdException Locked="false" Priority="73" Name="Colorful Grid Accent 1"/>
  <w:LsdException Locked="false" Priority="60" Name="Light Shading Accent 2"/>
  <w:LsdException Locked="false" Priority="61" Name="Light List Accent 2"/>
  <w:LsdException Locked="false" Priority="62" Name="Light Grid Accent 2"/>
  <w:LsdException Locked="false" Priority="63" Name="Medium Shading 1 Accent 2"/>
  <w:LsdException Locked="false" Priority="64" Name="Medium Shading 2 Accent 2"/>
  <w:LsdException Locked="false" Priority="65" Name="Medium List 1 Accent 2"/>
  <w:LsdException Locked="false" Priority="66" Name="Medium List 2 Accent 2"/>
  <w:LsdException Locked="false" Priority="67" Name="Medium Grid 1 Accent 2"/>
  <w:LsdException Locked="false" Priority="68" Name="Medium Grid 2 Accent 2"/>
  <w:LsdException Locked="false" Priority="69" Name="Medium Grid 3 Accent 2"/>
  <w:LsdException Locked="false" Priority="70" Name="Dark List Accent 2"/>
  <w:LsdException Locked="false" Priority="71" Name="Colorful Shading Accent 2"/>
  <w:LsdException Locked="false" Priority="72" Name="Colorful List Accent 2"/>
  <w:LsdException Locked="false" Priority="73" Name="Colorful Grid Accent 2"/>
  <w:LsdException Locked="false" Priority="60" Name="Light Shading Accent 3"/>
  <w:LsdException Locked="false" Priority="61" Name="Light List Accent 3"/>
  <w:LsdException Locked="false" Priority="62" Name="Light Grid Accent 3"/>
  <w:LsdException Locked="false" Priority="63" Name="Medium Shading 1 Accent 3"/>
  <w:LsdException Locked="false" Priority="64" Name="Medium Shading 2 Accent 3"/>
  <w:LsdException Locked="false" Priority="65" Name="Medium List 1 Accent 3"/>
  <w:LsdException Locked="false" Priority="66" Name="Medium List 2 Accent 3"/>
  <w:LsdException Locked="false" Priority="67" Name="Medium Grid 1 Accent 3"/>
  <w:LsdException Locked="false" Priority="68" Name="Medium Grid 2 Accent 3"/>
  <w:LsdException Locked="false" Priority="69" Name="Medium Grid 3 Accent 3"/>
  <w:LsdException Locked="false" Priority="70" Name="Dark List Accent 3"/>
  <w:LsdException Locked="false" Priority="71" Name="Colorful Shading Accent 3"/>
  <w:LsdException Locked="false" Priority="72" Name="Colorful List Accent 3"/>
  <w:LsdException Locked="false" Priority="73" Name="Colorful Grid Accent 3"/>
  <w:LsdException Locked="false" Priority="60" Name="Light Shading Accent 4"/>
  <w:LsdException Locked="false" Priority="61" Name="Light List Accent 4"/>
  <w:LsdException Locked="false" Priority="62" Name="Light Grid Accent 4"/>
  <w:LsdException Locked="false" Priority="63" Name="Medium Shading 1 Accent 4"/>
  <w:LsdException Locked="false" Priority="64" Name="Medium Shading 2 Accent 4"/>
  <w:LsdException Locked="false" Priority="65" Name="Medium List 1 Accent 4"/>
  <w:LsdException Locked="false" Priority="66" Name="Medium List 2 Accent 4"/>
  <w:LsdException Locked="false" Priority="67" Name="Medium Grid 1 Accent 4"/>
  <w:LsdException Locked="false" Priority="68" Name="Medium Grid 2 Accent 4"/>
  <w:LsdException Locked="false" Priority="69" Name="Medium Grid 3 Accent 4"/>
  <w:LsdException Locked="false" Priority="70" Name="Dark List Accent 4"/>
  <w:LsdException Locked="false" Priority="71" Name="Colorful Shading Accent 4"/>
  <w:LsdException Locked="false" Priority="72" Name="Colorful List Accent 4"/>
  <w:LsdException Locked="false" Priority="73" Name="Colorful Grid Accent 4"/>
  <w:LsdException Locked="false" Priority="60" Name="Light Shading Accent 5"/>
  <w:LsdException Locked="false" Priority="61" Name="Light List Accent 5"/>
  <w:LsdException Locked="false" Priority="62" Name="Light Grid Accent 5"/>
  <w:LsdException Locked="false" Priority="63" Name="Medium Shading 1 Accent 5"/>
  <w:LsdException Locked="false" Priority="64" Name="Medium Shading 2 Accent 5"/>
  <w:LsdException Locked="false" Priority="65" Name="Medium List 1 Accent 5"/>
  <w:LsdException Locked="false" Priority="66" Name="Medium List 2 Accent 5"/>
  <w:LsdException Locked="false" Priority="67" Name="Medium Grid 1 Accent 5"/>
  <w:LsdException Locked="false" Priority="68" Name="Medium Grid 2 Accent 5"/>
  <w:LsdException Locked="false" Priority="69" Name="Medium Grid 3 Accent 5"/>
  <w:LsdException Locked="false" Priority="70" Name="Dark List Accent 5"/>
  <w:LsdException Locked="false" Priority="71" Name="Colorful Shading Accent 5"/>
  <w:LsdException Locked="false" Priority="72" Name="Colorful List Accent 5"/>
  <w:LsdException Locked="false" Priority="73" Name="Colorful Grid Accent 5"/>
  <w:LsdException Locked="false" Priority="60" Name="Light Shading Accent 6"/>
  <w:LsdException Locked="false" Priority="61" Name="Light List Accent 6"/>
  <w:LsdException Locked="false" Priority="62" Name="Light Grid Accent 6"/>
  <w:LsdException Locked="false" Priority="63" Name="Medium Shading 1 Accent 6"/>
  <w:LsdException Locked="false" Priority="64" Name="Medium Shading 2 Accent 6"/>
  <w:LsdException Locked="false" Priority="65" Name="Medium List 1 Accent 6"/>
  <w:LsdException Locked="false" Priority="66" Name="Medium List 2 Accent 6"/>
  <w:LsdException Locked="false" Priority="67" Name="Medium Grid 1 Accent 6"/>
  <w:LsdException Locked="false" Priority="68" Name="Medium Grid 2 Accent 6"/>
  <w:LsdException Locked="false" Priority="69" Name="Medium Grid 3 Accent 6"/>
  <w:LsdException Locked="false" Priority="70" Name="Dark List Accent 6"/>
  <w:LsdException Locked="false" Priority="71" Name="Colorful Shading Accent 6"/>
  <w:LsdException Locked="false" Priority="72" Name="Colorful List Accent 6"/>
  <w:LsdException Locked="false" Priority="73" Name="Colorful Grid Accent 6"/>
  <w:LsdException Locked="false" Priority="19" QFormat="true"
   Name="Subtle Emphasis"/>
  <w:LsdException Locked="false" Priority="21" QFormat="true"
   Name="Intense Emphasis"/>
  <w:LsdException Locked="false" Priority="31" QFormat="true"
   Name="Subtle Reference"/>
  <w:LsdException Locked="false" Priority="32" QFormat="true"
   Name="Intense Reference"/>
  <w:LsdException Locked="false" Priority="33" QFormat="true" Name="Book Title"/>
  <w:LsdException Locked="false" Priority="37" SemiHidden="true"
   UnhideWhenUsed="true" Name="Bibliography"/>
  <w:LsdException Locked="false" Priority="39" SemiHidden="true"
   UnhideWhenUsed="true" QFormat="true" Name="TOC Heading"/>
 </w:LatentStyles>
</xml><![endif]-->
<style id=estil>
<!--body
	{SCROLLBAR-FACE-COLOR:#eaeaea;
	SCROLLBAR-HIGHLIGHT-COLOR: #fdfdfd;
	SCROLLBAR-SHADOW-COLOR: #eeeeee;
	SCROLLBAR-3DLIGHT-COLOR: #ffffff;
	SCROLLBAR-ARROW-COLOR:#6699cc;
	SCROLLBAR-TRACK-COLOR: #fafafa;
	SCROLLBAR-DARKSHADOW-COLOR: #e2e2e2;}
a.ISTT
	{00;}
a.NN
	{text-transform:capitalize;}
a.NTSLP
	{xtext-decoration:none;}
input.BT
	{DXImageTransform.Microsoft.Gradient(startColorStr=#dddddd, endColorStr=#C8DBFF, gradientType=0);
	cursor:pointer;}
input.BTE
	{DXImageTransform.Microsoft.Gradient(startColorStr=#990000, endColorStr=#dddddd, gradientType=0);
	cursor:pointer;}
input.BTG
	{DXImageTransform.Microsoft.Gradient(startColorStr=#eeeeee, endColorStr=#dddddd, gradientType=0);
	cursor:pointer;}
input.BTR
	{DXImageTransform.Microsoft.Gradient(startColorStr=#eeeeee, endColorStr=#550000, gradientType=0);
	cursor:pointer;}
input.DTN
	{text-align=right;}
input.DTNS
	{text-align=right;}
td.AA
	{text-transform:capitalize;}
td.TDCM1
	{cursor:pointer;}
td.TDCM2
	{cursor:pointer;}
td.BTBS
	{cursor:pointer;}
td.BTBN
	{cursor:pointer;}
td.BTBA
	{cursor:pointer;}
input.BIETX
	{0;}
td.IETT
	{00;}
input.IETX
	{000;}
a.IETX
	{000;}
td.IETX
	{000;}
input.IETXG
	{000;}
td.ISTT
	{00;}
td.ISTXZ
	{000;}
td.IBTT
	{0;}
td.ICTT
	{000;}
td.ICTX
	{000;}
td.ICTR
	{0font-style:italic;
	0;}
td.ICXR
	{00;}
span.SEARCHWORD
	{cursor:pointer;}
input.BTSB
	{cursor:pointer;}
input.BTSBG
	{cursor:pointer;}
td.TS
	{filter:"alpha\(opacity=100\,finishOpacity=60\,STYLE=1\)";}
td.MBT
	{cursor:pointer;
	filter:"dropShadow\(Color=\#f9f9f9\,OFFX=1\,OFFY=1\)";}
td.FT
	{filter:"dropShadow\(Color=\#cccccc\,OFFX=2\,OFFY=3\)";}
td.FTI
	{filter:"dropShadow\(Color=\#cccccc\,OFFX=2\,OFFY=3\)";}
td.UTS
	{filter:"shadow\(color=\#dddddd\,direction=125\)";}
td.UTIS
	{filter:"shadow\(color=\#dddddd\,direction=125\)";}
td.UTGS
	{filter:"shadow\(color=\#dddddd\,direction=125\)";}
td.UTGIS
	{filter:"shadow\(color=\#dddddd\,direction=125\)";}
td.UT2
	{font-weight:none;}
table.TBB
	{visibility:hidden;}
table.TBH
	{filter:"alpha\(opacity=0\,finishOpacity=100\,STYLE=1\)";}
table.TBP
	{DXImageTransform.Microsoft.Gradient(startColorStr=#FFFFFF, endColorStr=#CCCCCC, gradientType=0);}
table.TDCM
	{DXImageTransform.Microsoft.Gradient(startColorStr=#FFFFFF, endColorStr=#CCCCCC, gradientType=0);}
table.TWIND
	{DXImageTransform.Microsoft.Gradient(startColorStr=#FFFFFF, endColorStr=#AAAAAA, gradientType=0);}

 /* Font Definitions */
 @font-face
	{font-family:"Cambria Math";
	panose-1:2 4 5 3 5 4 6 3 2 4;
	mso-font-charset:0;
	mso-generic-font-family:roman;
	mso-font-pitch:variable;
	mso-font-signature:-1610611985 1107304683 0 0 159 0;}
@font-face
	{font-family:Tahoma;
	panose-1:2 11 6 4 3 5 4 4 2 4;
	mso-font-charset:0;
	mso-generic-font-family:swiss;
	mso-font-pitch:variable;
	mso-font-signature:1627400839 -2147483648 8 0 66047 0;}
@font-face
	{font-family:"Trebuchet MS";
	panose-1:2 11 6 3 2 2 2 2 2 4;
	mso-font-charset:0;
	mso-generic-font-family:swiss;
	mso-font-pitch:variable;
	mso-font-signature:647 0 0 0 159 0;}
@font-face
	{font-family:Verdana;
	panose-1:2 11 6 4 3 5 4 4 2 4;
	mso-font-charset:0;
	mso-generic-font-family:swiss;
	mso-font-pitch:variable;
	mso-font-signature:536871559 0 0 0 415 0;}
 /* Style Definitions */
 p.MsoNormal, li.MsoNormal, div.MsoNormal
	{mso-style-unhide:no;
	mso-style-qformat:yes;
	mso-style-parent:"";
	margin:0cm;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:12.0pt;
	font-family:"Times New Roman","serif";
	mso-fareast-font-family:"Times New Roman";
	color:windowtext;}
p.MsoHeader, li.MsoHeader, div.MsoHeader
	{mso-style-unhide:no;
	mso-style-link:"Encabezado Car";
	margin:0cm;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	tab-stops:center 212.6pt right 425.2pt;
	font-size:12.0pt;
	font-family:"Times New Roman","serif";
	mso-fareast-font-family:"Times New Roman";
	color:#888888;}
p.MsoFooter, li.MsoFooter, div.MsoFooter
	{mso-style-priority:99;
	mso-style-unhide:no;
	mso-style-link:"Pie de página Car";
	margin:0cm;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	tab-stops:center 212.6pt right 425.2pt;
	font-size:12.0pt;
	font-family:"Times New Roman","serif";
	mso-fareast-font-family:"Times New Roman";
	color:#888888;}
a:link, span.MsoHyperlink
	{mso-style-unhide:no;
	color:blue;
	mso-text-animation:none;
	text-decoration:none;
	text-underline:none;
	text-decoration:none;
	text-line-through:none;}
a:visited, span.MsoHyperlinkFollowed
	{mso-style-unhide:no;
	color:purple;
	mso-text-animation:none;
	text-decoration:none;
	text-underline:none;
	text-decoration:none;
	text-line-through:none;}
p
	{mso-style-unhide:no;
	margin:0cm;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:9.0pt;
	font-family:"Trebuchet MS","sans-serif";
	mso-fareast-font-family:"Times New Roman";
	mso-fareast-theme-font:minor-fareast;
	mso-bidi-font-family:"Times New Roman";
	color:#888888;}
p.MsoAcetate, li.MsoAcetate, div.MsoAcetate
	{mso-style-unhide:no;
	mso-style-link:"Texto de globo Car";
	margin:0cm;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:8.0pt;
	font-family:"Tahoma","sans-serif";
	mso-fareast-font-family:"Times New Roman";
	color:windowtext;}
span.EncabezadoCar
	{mso-style-name:"Encabezado Car";
	mso-style-unhide:no;
	mso-style-locked:yes;
	mso-style-link:Encabezado;
	mso-ansi-font-size:12.0pt;
	mso-bidi-font-size:12.0pt;}
span.PiedepginaCar
	{mso-style-name:"Pie de página Car";
	mso-style-priority:99;
	mso-style-unhide:no;
	mso-style-locked:yes;
	mso-style-link:"Pie de página";
	mso-ansi-font-size:12.0pt;
	mso-bidi-font-size:12.0pt;}
span.at
	{mso-style-name:at;
	mso-style-unhide:no;
	mso-ansi-font-size:9.0pt;
	mso-bidi-font-size:9.0pt;
	font-family:"Trebuchet MS","sans-serif";
	mso-ascii-font-family:"Trebuchet MS";
	mso-hansi-font-family:"Trebuchet MS";
	color:#888888;
	background:#14468F;}
span.xmn
	{mso-style-name:xmn;
	mso-style-unhide:no;
	mso-ansi-font-size:10.0pt;
	mso-bidi-font-size:10.0pt;
	font-family:"Trebuchet MS","sans-serif";
	mso-ascii-font-family:"Trebuchet MS";
	mso-hansi-font-family:"Trebuchet MS";
	color:black;
	border:solid black 1.0pt;
	padding:0cm;
	background:#C8DBFF;}
span.xsl
	{mso-style-name:xsl;
	mso-style-unhide:no;
	mso-ansi-font-size:11.0pt;
	mso-bidi-font-size:11.0pt;
	font-family:"Trebuchet MS","sans-serif";
	mso-ascii-font-family:"Trebuchet MS";
	mso-hansi-font-family:"Trebuchet MS";
	color:#14468F;
	border:solid white 1.0pt;
	padding:0cm;
	background:#C8DBFF;}
span.ntslp
	{mso-style-name:ntslp;
	mso-style-unhide:no;
	mso-ansi-font-size:11.0pt;
	mso-bidi-font-size:11.0pt;
	font-family:"Trebuchet MS","sans-serif";
	mso-ascii-font-family:"Trebuchet MS";
	mso-hansi-font-family:"Trebuchet MS";
	color:#FFFDFF;
	font-weight:normal;}
span.ntmnn
	{mso-style-name:ntmnn;
	mso-style-unhide:no;
	mso-ansi-font-size:10.0pt;
	mso-bidi-font-size:10.0pt;
	font-family:"Trebuchet MS","sans-serif";
	mso-ascii-font-family:"Trebuchet MS";
	mso-hansi-font-family:"Trebuchet MS";
	color:white;
	font-weight:normal;}
span.mt
	{mso-style-name:mt;
	mso-style-unhide:no;
	mso-ansi-font-size:9.0pt;
	mso-bidi-font-size:9.0pt;
	font-family:"Trebuchet MS","sans-serif";
	mso-ascii-font-family:"Trebuchet MS";
	mso-hansi-font-family:"Trebuchet MS";
	color:#888888;
	text-transform:uppercase;
	border:outset 1.0pt;
	padding:0cm;
	background:#14468F;
	font-weight:bold;}
span.searchword
	{mso-style-name:searchword;
	mso-style-unhide:no;
	background:yellow;}
span.TextodegloboCar
	{mso-style-name:"Texto de globo Car";
	mso-style-unhide:no;
	mso-style-locked:yes;
	mso-style-link:"Texto de globo";
	mso-ansi-font-size:8.0pt;
	mso-bidi-font-size:8.0pt;
	font-family:"Tahoma","sans-serif";
	mso-ascii-font-family:Tahoma;
	mso-hansi-font-family:Tahoma;
	mso-bidi-font-family:Tahoma;}
.MsoChpDefault
	{mso-style-type:export-only;
	mso-default-props:yes;
	font-size:10.0pt;
	mso-ansi-font-size:10.0pt;
	mso-bidi-font-size:10.0pt;}
 /* Page Definitions */
 @page

	{mso-footnote-separator:url("reportesgen_archivos/footer.asp?logo=<%=oPrtG.RCliLogo%>&RCliNomb=<%=oPrtG.RCliNomb%>&STITULO=<%=oPrtG.titdoc%>") fs;
	mso-footnote-continuation-separator:url("reportesgen_archivos/footer.asp?logo=<%=oPrtG.RCliLogo%>&RCliNomb=<%=oPrtG.RCliNomb%>&STITULO=<%=oPrtG.titdoc%>") fcs;
	mso-endnote-separator:url("reportesgen_archivos/footer.asp") es;
	mso-endnote-continuation-separator:url("reportesgen_archivos/footer.asp?logo=<%=oPrtG.RCliLogo%>&STITULO=<%=oPrtG.titdoc%>") ecs;}
@page WordSection1
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
div.WordSection1
	{page:WordSection1;}
-->
</style>
<!--[if gte mso 10]>
<style>
 /* Style Definitions */
 table.MsoNormalTable
	{mso-style-name:"Tabla normal";
	mso-tstyle-rowband-size:0;
	mso-tstyle-colband-size:0;
	mso-style-noshow:yes;
	mso-style-priority:99;
	mso-style-qformat:yes;
	mso-style-parent:"";
	mso-padding-alt:0cm 5.4pt 0cm 5.4pt;
	mso-para-margin:0cm;
	mso-para-margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:10.0pt;
	font-family:"Times New Roman","serif";}
</style>
<![endif]--><!--[if gte mso 9]><xml>
 <o:shapedefaults v:ext="edit" spidmax="2050"/>
</xml><![endif]--><!--[if gte mso 9]><xml>
 <o:shapelayout v:ext="edit">
  <o:idmap v:ext="edit" data="1"/>
 </o:shapelayout></xml><![endif]-->

</head>

<body  lang=ES-AR link=blue vlink=purple style='tab-interval:35.4pt'>

<div class=WordSection1>

<%
ELSE
%>

<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:w="urn:schemas-microsoft-com:office:word"
xmlns="http://www.w3.org/TR/REC-html40">
<head>
<base href="<%=sbase%>" />
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

END IF

%>
<%


   response.write CorStyler2()
'sbloque1=replace(sbloque1,"src=" & chr(34) ,"src=" & chr(34) & "http://bcbat212/normastest/coremanager/")
'sbloque1=replace(sbloque1,"src='" ,"src='http://bcbat212/normastest/coremanager/")
'sbloque1=replace(sbloque1,"SRC=" & chr(34) ,"src=" & chr(34) & "http://bcbat212/normastest/coremanager/")
arperfiles=split(sperfiles,"###")
for i=0 to ubound(arperfiles)
 xdatof=split(arperfiles(i),"#")

 sbloque1=   replace(sbloque1,"<SPAN id="& trim(xdatof(0)) &" class=spanpf name=""idsec" & xdatof(2) &""">","<SPAN id="& xdatof(0)& " class=spanpf>" & xdatof(1))
next

arsectores=split(ssectores,"###")
for i=0 to ubound(arsectores)
 xdatof=split(arsectores(i),"#")
 sbloque1=  replace(sbloque1,"<SPAN id="& xdatof(0)&" class=spansec name=""" & xdatof(0) &""">","<SPAN id="& xdatof(0)& " class=spanpf>" & xdatof(1))
next


  response.write   sbloque1

%>
 </body></html>



<%
end if

END IF
end if
%>

<%

Public Function RepHeader2pdf(STitulo, nAncho)

'''        Public RCliLogo
'''        Public RCliNomb
    Dim s, Fecha

    nAncho = "100%"

    s = "<table width=100%  marginwidth=0 cellspacing=0 cellpadding=0>"
    s = s & "<tr align=top>"
    If RCliLogo <> "" Then s = s & "<td width=4% ><img src='" & Sbase & RCliLogo & "' border=0 ></td>"
    s = s & "<td class=ietxprn align=left >&nbsp;" & RCliNomb & "</td></tr></table>"
    s = s & "<br><table width=" & nAncho & " marginwidth=0 cellspacing=0 cellpadding=0>"
    s = s & "<tr align=top><td class=iettprn align=left width=70% >" & STitulo & "</td>"
    Fecha = OFV.fechador()
    s = s & "<td class=ietxprn align=right>" & Fecha & "</td></tr></table>"
    s = s & "<hr color=#000088 size=1 width=100% align=left><br><br><br>"
    s=replace(s,chr(34),"'")
    RepHeader2pdf = s

End Function

%>