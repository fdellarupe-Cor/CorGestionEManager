<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->



<% 

ok=ofv.CheckUsuario()
 
if ok <> "Y" then     
  sHTML = ok
else  
  response.expires=0   
    
  sHTML = ofv.formHeader("")
  session("svuelveAlm")="LISTAPROYALL.ASP"    
   
           
  sHTML = sHTML  & "<script language=javascript> "
 ' sHTML = sHTML  & "        document.getElementById('MDPTOP1').style.display='none';  document.getElementById('tbhMAX').style.display='none';        hdemenu();     "
  sHTML = sHTML  & "function AbrirCons(Datos) { "
  sHTML = sHTML  & " myWin= open(Datos," & chr(34) & "Cons" & chr(34) & "," & chr(34) & "status=yes,resizable=yes,menubar=no,toolbar=no,height=570,width=790" & chr(34) & " ); "
   
  sHTML = sHTML  & " } "
  sHTML = sHTML  & "function asignardato(obj1,obj,valor) { "
  sHTML = sHTML  &  "obj.value=valor; document.getElementById('ppp').style.display='none'; "
  sHTML = sHTML  &  " document.getElementById('ppp1').innerHTML=obj1.title + document.getElementById('ppp1').innerHTML; "
  
  sHTML = sHTML  & "document.getElementById('ppp1').style.display='block'; }"
 
     
 
 
   sHTML = sHTML  & "function asignardato2(obj1,obj,valor) { " 
   sHTML = sHTML  &  "obj.value=valor; }"
 
    
  sHTML = sHTML  & "function VerOnOff(obj) { "
  sHTML = sHTML  &"  if ( document.getElementById(obj).style.display =='none'){"
  sHTML = sHTML  &"  document.getElementById(obj).style.display='block'; }"
  sHTML = sHTML  &" else {"
  sHTML = sHTML  &"document.getElementById(obj).style.display='none';}}"
  
   
  
  sHTML = sHTML  & "function cambiarHref(obj,frm,xcampo){"
  sHTML = sHTML  & "var x=document.getElementById(frm);"
  sHTML = sHTML  & " obj.href='../coress20/almanaque.asp?xcampo='+ xcampo + '&';"
  sHTML = sHTML  & "for (var i=0;i<x.length;i++){"
  sHTML = sHTML  & "  if (x.elements[i].name !=''){"
  sHTML = sHTML  & "  obj.href=obj.href +  x.elements[i].name + '=' + x.elements[i].value + '&';"
  
  sHTML = sHTML  & "  }  }"
    
  sHTML = sHTML  & " }" 

                       
  sHTML = sHTML  & "</script>"

  CantRegMostrarx = 6
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

 

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

    
 
   set cn = ofv.conectar(ofv.strconn6)
   set cn2 = ofv.conectar(ofv.strconn0)
   
   
   xestado=request.querystring("estado")
   xfdesde=clng("0" & request.querystring("fdesde") )
   xfhasta=clng("0" & request.querystring("fhasta"))
   xfdesde1=clng("0" & request.querystring("fdesde1") )
   xfhasta1=clng("0" & request.querystring("fhasta1"))
   xgersector=clng("0" & request.querystring("xgersector"))
   Xanalista= request.querystring("Xanalista")
   Xoriginante= request.querystring("Xoriginante")
   Xxls=""   
      
   if xfdesde="0" then xfdesde=clng(date())
   if xfhasta="0" then xfhasta=clng(date())
      
   if xfdesde1="0" then xfdesde1=clng(date())
   if xfhasta1="0" then xfhasta1=clng(date())
   

   


    sImagenO  = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=3>&#0128;</FONT></B>"
    sImagenI  = "<B><FONT FACE=WEBDINGS COLOR=#AA0000 SIZE=3>&#0128;</FONT></B>"
    sImagenE  = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0240;</FONT></B>"
    sImagenT  = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0174;</FONT></B>"
    sImagenC  = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=3>R</FONT></B>"
    sImagenAlm  = "<B><FONT FACE=WEBDINGS COLOR=#f7aa22 SIZE=3>&#0165;</FONT></B>"
    sImagenX  = "<B><FONT FACE=WEBDINGS COLOR=#f7aa22 SIZE=3>4</FONT></B>"
    sImagenIX = "<B><FONT FACE=WEBDINGS COLOR=#aa0000 SIZE=3>2</FONT></B>"

    sImagenpry = "<B><FONT FACE=WEBDINGS COLOR=#aa0000 SIZE=3>&#0143;</FONT></B>"
    sImagenreq = "<B><FONT FACE=WEBDINGS COLOR=#006600 SIZE=3>&#0153;</FONT></B>"
    sImagenXLS = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=2>&#048;</FONT></B>"


   celdas = ""    

   sinput =  "  <a class=cc   style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
   sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
   sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
   sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
   sinput = sinput  & "none" & chr(34) & ";' " 
   sinput = sinput  & " href='#' onclick=""asignardato2(this,document.all.datosx.Xxls,'XLS');asignardato2(this,document.all.datosx.estado,1);enviar('estado');""  >"& sImagenXLS& "</a>"  
  
  sinput = sinput & sImagent  & "  <a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
  sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
  sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
  sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
  sinput = sinput  & "none" & chr(34) & ";' " 
  sinput = sinput  & " href='#' onclick=""asignardato2(this,document.all.datosx.Xxls,'NO');asignardato2(this,document.all.datosx.estado,1);enviar('estado');"">Proyectos nuevos</a>"  
  Celda = ofv.GenCelda("","vt ","left valign=top","","","","",sinput,"si")
 
 
  Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")
  
    sinput =  "  <a class=cc   style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
    sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
    sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
    sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
    sinput = sinput  & "none" & chr(34) & ";' " 
    sinput = sinput  & " href='#' onclick=""asignardato2(this,document.all.datosx.Xxls,'XLS');asignardato2(this,document.all.datosx.estado,2);enviar('estado');""  >"& sImagenXLS& "</a>"  
   
    
    sinput = sinput & sImagent  & "  <a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
    sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
    sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
    sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
    sinput = sinput  & "none" & chr(34) & ";' " 
    sinput = sinput  & " href='#' onclick=""asignardato2(this,document.all.datosx.Xxls,'NO');asignardato2(this,document.all.datosx.estado,2);enviar('estado');"">Proyecto en curso completo</a>"  
    Celda = ofv.GenCelda("","vt ","left valign=top","","","","",sinput,"si")
   
   
  Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")
  
  
      
   
  
    sinput = sImagent  & "  <a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
    sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
    sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
    sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
    sinput = sinput  & "none" & chr(34) & ";' " 
    sinput = sinput  & " href='#' >Proyectos en curso</a>"  
    
        sinputd =  "  <a class=cc   style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
        sinputd = sinputd  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
        sinputd = sinputd  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
        sinputd = sinputd  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
        sinputd = sinputd  & "none" & chr(34) & ";' " 
        sinputd = sinputd  & " href='#' onclick=""asignardato2(this,document.all.datosx.Xxls,'XLS');asignardato2(this,document.all.datosx.estado,12);enviar('estado');""  >"& sImagenXLS& "</a>"  
   
    
    
	sinput = sinput & "<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & sinputd & sImagent  & "  <a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
	sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
	sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
	sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
	sinput = sinput  & "none" & chr(34) & ";' " 
	sinput = sinput  & " href='#' onclick=""asignardato2(this,document.all.datosx.Xxls,'NO');asignardato2(this,document.all.datosx.estado,12);enviar('estado');"">Proyecto en curso normal</a>"  


        sinputd =  "  <a class=cc   style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
        sinputd = sinputd  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
        sinputd = sinputd  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
        sinputd = sinputd  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
        sinputd = sinputd  & "none" & chr(34) & ";' " 
        sinputd = sinputd  & " href='#' onclick=""asignardato2(this,document.all.datosx.Xxls,'XLS');asignardato2(this,document.all.datosx.estado,1);enviar('estado');""  >"& sImagenXLS& "</a>"  


        
        sinput = sinput & "<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & sinputd & sImagent  & "  <a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
	sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
	sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
	sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
	sinput = sinput  & "none" & chr(34) & ";' " 
        sinput = sinput  & " href='#' onclick=""asignardato2(this,document.all.datosx.Xxls,'NO');asignardato2(this,document.all.datosx.estado,1);enviar('estado');"">Proyecto en curso con desvio</a>"  
        
        sinputd =  "  <a class=cc   style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
        sinputd = sinputd  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
        sinputd = sinputd  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
        sinputd = sinputd  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
        sinputd = sinputd  & "none" & chr(34) & ";' " 
        sinputd = sinputd  & " href='#' onclick=""asignardato2(this,document.all.datosx.Xxls,'XLS');asignardato2(this,document.all.datosx.estado,1);enviar('estado');""  >"& sImagenXLS& "</a>"  
        
        
        sinput = sinput & "<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & sinputd & sImagent  & "  <a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
	sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
	sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
	sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
	sinput = sinput  & "none" & chr(34) & ";' " 
	sinput = sinput  & " href='#' onclick=""asignardato2(this,document.all.datosx.Xxls,'NO');asignardato2(this,document.all.datosx.estado,1);enviar('estado');"">Proyecto en curso altamente desviado</a>"  

        
        
        Celda = ofv.GenCelda("","vt ","left valign=top","","","","",sinput,"si")
    
    
    
    
  
   
   
  Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")
  
  
  
  
        sinputd =  "  <a class=cc   style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
        sinputd = sinputd  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
        sinputd = sinputd  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
        sinputd = sinputd  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
        sinputd = sinputd  & "none" & chr(34) & ";' " 
        sinputd = sinputd  & " href='#' onclick=""asignardato2(this,document.all.datosx.Xxls,'XLS');asignardato2(this,document.all.datosx.estado,1);enviar('estado');""  >"& sImagenXLS& "</a>"  
  
  
  
  
        sinput = sinputd & sImagent  & "  <a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
        sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
        sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
        sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
        sinput = sinput  & "none" & chr(34) & ";' " 
        sinput = sinput  & " href='#' onclick=""asignardato2(this,document.all.datosx.Xxls,'NO');asignardato2(this,document.all.datosx.estado,1);enviar('estado');"">Proyectos detenidos</a>"  
        Celda = ofv.GenCelda("","vt ","left valign=top","","","","",sinput,"si")
   
   
  Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")
  
  
  
        sinputd =  "  <a class=cc   style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
        sinputd = sinputd  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
        sinputd = sinputd  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
        sinputd = sinputd  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
        sinputd = sinputd  & "none" & chr(34) & ";' " 
        sinputd = sinputd  & " href='#' onclick=""asignardato2(this,document.all.datosx.Xxls,'XLS');asignardato2(this,document.all.datosx.estado,13);enviar('estado');""  >"& sImagenXLS& "</a>"  
  
        sinput = sinputd & sImagent  & "  <a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
        sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
        sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
        sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
        sinput = sinput  & "none" & chr(34) & ";' " 
        sinput = sinput  & " href='#' onclick=""asignardato2(this,document.all.datosx.Xxls,'NO');asignardato2(this,document.all.datosx.estado,13);enviar('estado');"">Proyectos finalizados</a>"  
      Celda = ofv.GenCelda("","vt ","left valign=top","","","","",sinput,"si")
     
     
       Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")
  
  
        sinputd =  "  <a class=cc   style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
        sinputd = sinputd  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
        sinputd = sinputd  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
        sinputd = sinputd  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
        sinputd = sinputd  & "none" & chr(34) & ";' " 
        sinputd = sinputd  & " href='#' onclick=""asignardato2(this,document.all.datosx.Xxls,'XLS');enviar('RANGO');""  >"& sImagenXLS& "</a>"  
 
 
        sinputc = sinputd & sImagent  & "  <a class=cc style='cursor:hand;text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
	sinputc = sinputc  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
	sinputc = sinputc  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
	sinputc = sinputc  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
	sinputc = sinputc  & "none" & chr(34) & ";' " 
	sinputc = sinputc  & "   onclick=""asignardato2(this,document.all.datosx.Xxls,'NO');enviar('RANGO');"" >buscar</a>" 



 
  
 
 
   celda=ofv.GenCelda("","vt ","left","","","","","&nbsp;","si")
   Celda = celda & ofv.GenCelda("","vt ","left","","","","","&nbsp;","si")
   Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")
 
 
   sinput = sImagent & "LISTADO GENERAL DE PROYECTOS POR :" 
   Celda = ofv.GenCelda("","vt ","left ","","","","",sinput,"si")
   'sinput = sImagenX & "Desde :" & sImagenAlm & "----" & "Hasta:" & sImagenAlm
   
   
     
     '----------------------------
     
     
       
       sinput =   "  <a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
       sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
       sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
       sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
       sinput = sinput  & "none" & chr(34) & ";' "
       sinput = sinput  & "  href='javascript:void(0);' onclick="&chr(34)&"cambiarHref(this,'datosx','fdesde1');" & chr(34)
       sinput = sinput & " > "& sImagenAlm & "Desde : "&cdate(xfdesde1)&" </a>"
       
       
       sinput2 =   "  <a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
       sinput2 = sinput2  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
       sinput2 = sinput2  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
       sinput2 = sinput2  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
       sinput2 = sinput2  & "none" & chr(34) & ";' "
       sinput2 = sinput2  & "  href='javascript:void(0);' onclick="&chr(34)&"cambiarHref(this,'datosx','fhasta1');" & chr(34)
       sinput2 = sinput2 & " > "& sImagenAlm & "hasta : "&cdate(xfhasta1)&" </a>"
      
          
  
       
       
       Celda = celda & ofv.GenCelda("","vt ","left","","","","",sinput & sinput2   & sinputc,"si")
       Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")
  
     '----------------------------
     
     
     
     
     
     
     
     
     
   celda=ofv.GenCelda("","vt ","left","","","","","&nbsp;","si")
   Celda = celda & ofv.GenCelda("","vt ","left","","","","","&nbsp;","si")
   Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")
   
     
   sinput = "&nbsp;&nbsp;&nbsp;&nbspGERENCIA / SECTOR"'sImagent & "LISTADO GENERAL DE PROYECTOS POR :" 
   Celda = ofv.GenCelda("","vt ","left valign='top'","","","","",sinput,"si")
     
     
     'cargo gerencia sector
     strsql="Select DISTINCT  GERENCIA, SECTOR From PROYRESP a where   USRAFC <> 'NO' AND TIPO = 'AN'  ORDER BY GERENCIA, SECTOR"
     set rsGerSector = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      
     
     sInput = "<font class=at><select name=Sector  onchange='document.all.datosx.Xgersector.value=this.options[this.selectedIndex].value;'>"
     
    
     gerenciaN=""
     do while not rsGerSector.eof
     
        gerencia=rsGerSector(0)
       strsql="select   a.descri + ' / ' + b.descri,a.descri  from eco a,eco b where a.eco=0"& rsGerSector(0) &" and b.eco=0" & rsGerSector(1) &""        
       set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
       if gerencia <> gerenciaN then
        sInput = sInput & "<option value=" & rsGerSector(0)
        sInput = sInput & " >" & rsniv(1) 
        gerenciaN=gerencia
       end if
       sInput = sInput & "<option value=" & rsGerSector(1)
       sInput = sInput & " >" & rsniv(0) 
       call ofv.cerrarconsulta(rsniv)  
       rsGerSector.movenext      
     loop 
        
         sInput = sInput & "</select></font>&nbsp;&nbsp;&nbsp;&nbsp;"
        'sInput = sInput & ofv.GenerarInput("","buscar","submit","","12","dt","") 
          
          
         call ofv.cerrarconsulta(rsGerSector) 
            
        sinputd =  "  <a class=cc   style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
        sinputd = sinputd  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
        sinputd = sinputd  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
        sinputd = sinputd  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
        sinputd = sinputd  & "none" & chr(34) & ";' " 
        sinputd = sinputd  & " href='#' onclick=""asignardato2(this,document.all.datosx.Xxls,'XLS');enviar('GerSector');""  >"& sImagenXLS& "</a>"  

            
           
	   sinputc = sImagent  & "  <a class=cc style='cursor:hand;text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
	   sinputc = sinputc  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
	   sinputc = sinputc  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
	   sinputc = sinputc  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
	   sinputc = sinputc  & "none" & chr(34) & ";' " 
	   sinputc = sinputc  & "   onclick=""asignardato2(this,document.all.datosx.Xxls,'NO');enviar('GerSector');"" >buscar</a>"  
	   sInput = sInput & sinputd & sinputc
         
    sAccion="LstReq.asp"    
    sinput = ofv.GenForm("GErSeca","",sAccion,"","post","",sInput,"si") 
   Celda = celda & ofv.GenCelda("","vt ","left  ","","","","",sinput,"si")
   Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")
   
   
   sinput = "&nbsp;&nbsp;&nbsp;&nbspANALISTA"'sImagent & "" 
   Celda = ofv.GenCelda("","vt ","left ","","","","",sinput,"si")
   
   
         
   'por  Analista
        strsql="SELECT distinct asignado  from requerimiento"
        set rsreq = ofv.crearconsultaEx(StrSql,cn,1,parametros)
         
        
        sInput = "<br><font class=at><select name=analistas onchange='document.all.datosx.Xanalista.value=this.options[this.selectedIndex].value;'>"
        
       
        gerenciaN=""
        inx=" in("
        do while not rsreq.eof
         inx=inx &"'" &rsreq(0) & "',"
          rsreq.movenext      
        loop 
         
         inx=mid(inx,1,len(inx)-1)
         inx=inx &")"
         
         
            strsql="select  descripcion,id  from usuarios where id " & inx  & " order by descripcion"
            'response.write strsql
	    set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
	    do while not rsniv.eof
	    
	              sInput = sInput & "<option value=" & rsniv(1)
	              sInput = sInput & " >" & rsniv(0) 
	    rsniv.movenext          
	    loop
          call ofv.cerrarconsulta(rsniv)  
          
           sInput = sInput & "</select></font>&nbsp;&nbsp;&nbsp;&nbsp;"
          ' sInput = sInput & ofv.GenerarInput("","buscar","submit","","12","dt","") 
            call ofv.cerrarconsulta(rsreq)  
            
        sinputd =  "  <a class=cc   style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
        sinputd = sinputd  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
        sinputd = sinputd  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
        sinputd = sinputd  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
        sinputd = sinputd  & "none" & chr(34) & ";' " 
        sinputd = sinputd  & " href='#' onclick=""asignardato2(this,document.all.datosx.Xxls,'XLS');enviar('analista');""  >"& sImagenXLS& "</a>"  
          
       sinputc = sImagent  & "  <a class=cc style='cursor:hand;text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
       sinputc = sinputc  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
       sinputc = sinputc  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
       sinputc = sinputc  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
       sinputc = sinputc  & "none" & chr(34) & ";' " 
       sinputc = sinputc  & "   onclick=""asignardato2(this,document.all.datosx.Xxls,'NO');enviar('analista');"" >buscar</a>"  
       sInput = sInput & sinputd & sinputc
	   
      sinput = ofv.GenForm("GErSec","",sAccion,"","post","",sInput,"si") 
      Celda = celda & ofv.GenCelda("","vt ","left","","","","",sinput,"si")
      Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")
      
      
      
   sinput = "&nbsp;&nbsp;&nbsp;&nbspORIGINANTE"'sImagent & "" 
   Celda = ofv.GenCelda("","vt ","left ","","","","",sinput,"si")
   
   
   
   'por  originante
        strsql="SELECT distinct originante  from requerimiento"
        set rsreq = ofv.crearconsultaEx(StrSql,cn,1,parametros)
         
        
        sInput = "<br><font class=at><select name=originante onchange='document.all.datosx.Xoriginante.value=this.options[this.selectedIndex].value;'>"
        
       
        gerenciaN=""
        inx=" in("
        do while not rsreq.eof
         inx=inx &"'" &rsreq(0) & "',"
          rsreq.movenext      
        loop 
         
         inx=mid(inx,1,len(inx)-1)
         inx=inx &")"
         
         
            strsql="select  descripcion,id  from usuarios where id " & inx  & " order by descripcion"
            'response.write strsql
	    set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
	    do while not rsniv.eof
	    
	              sInput = sInput & "<option value=" & rsniv(1)
	              sInput = sInput & " >" & rsniv(0) 
	    rsniv.movenext          
	    loop
          call ofv.cerrarconsulta(rsniv)  
          
           sInput = sInput & "</select></font>&nbsp;&nbsp;&nbsp;&nbsp;"
          ' sInput = sInput & ofv.GenerarInput("","buscar","submit","","12","dt","") 
            call ofv.cerrarconsulta(rsreq)  
            
        sinputd =  "  <a class=cc   style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
        sinputd = sinputd  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
        sinputd = sinputd  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
        sinputd = sinputd  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
        sinputd = sinputd  & "none" & chr(34) & ";' " 
        sinputd = sinputd  & " href='#' onclick=""asignardato2(this,document.all.datosx.Xxls,'XLS');enviar('originante');""  >"& sImagenXLS& "</a>"  
           
           
           sinputc = sImagent  & "  <a class=cc style='cursor:hand;text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
	   sinputc = sinputc  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
	   sinputc = sinputc  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
	   sinputc = sinputc  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
	   sinputc = sinputc  & "none" & chr(34) & ";' " 
	   sinputc = sinputc  & "   onclick=""asignardato2(this,document.all.datosx.Xxls,'NO');enviar('originante');"" >buscar</a>"  
           SInput = sInput & sinputd & sinputc
       
       sinput = ofv.GenForm("GErSec","",sAccion,"","post","",sInput,"si") 
      Celda = celda & ofv.GenCelda("","vt ","left","","","","",sinput,"si")
      Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")
      
   
  
 ''' form datos
    sHTML = sHTML  & ofv.GenTabla("","aaa  ","","98%","","0","","-1","-1","","",Fila,"si")
   
   
   
   t1 = ofv.GenerarInput("estado ",xestado,"textsb","datosx","12","dt","") & "<br>"
   t1 = t1 & ofv.GenerarInput("fdesde",xfdesde,"textsb","datosx","12","dt","")& "<br>"
   t1 = t1 & ofv.GenerarInput("fhasta",xfhasta,"textsb","datosx","12","dt","")& "<br>"
   t1 = t1 & ofv.GenerarInput("fdesde1",xfdesde1,"textsb","datosx","12","dt","")& "<br>"
   t1 =  t1 & ofv.GenerarInput("fhasta1",xfhasta1,"textsb","datosx","12","dt","")& "<br>"
   t1 = t1 & ofv.GenerarInput("Xgersector",xgersector,"textsb","datosx","12","dt","")& "<br>"
   t1 = t1 & ofv.GenerarInput("Xanalista",Xanalista,"textsb","datosx","12","dt","")& "<br>"
    t1 = t1 & ofv.GenerarInput("Xoriginante",Xoriginante,"textsb","datosx","12","dt","")& "<br>"
   t1 = t1 & ofv.GenerarInput("Xreqaa5","","textsb","datosx","12","dt","")& "<br>"
    t1 = t1 & ofv.GenerarInput("Xxls","","textsb","datosx","12","dt","")& "<br>"
   
   t1 = t1 & ofv.GenerarInput("Xreqaa1",cdate( xfdesde),"textsb","datosx","12","dt","")& "<br>"
   t1 = t1 & ofv.GenerarInput("Xreqaa2",cdate( xfhasta),"textsb","datosx","12","dt","")& "<br>"
   t1 = t1 & ofv.GenerarInput("Xreqaa3",cdate( xfdesde1),"textsb","datosx","12","dt","")& "<br>"
   t1 = t1 & ofv.GenerarInput("Xreqaa4",cdate( xfhasta1),"textsb","datosx","12","dt","")& "<br>"
   t1 = t1 & "<input id='envxxx' value='aaaaaaaaaaaaa' type=submit class='bt' ><br>"

   
   
    sAccion="LstPROY.asp"    
    
    sHTML = sHTML  &  "<div style='{sdisplay=none}'>" &ofv.GenForm("'datosx'","",saccion ,"","post ","",t1,"si") & "</div>"
           
           
  
  sHTML = sHTML  & "<script language=javascript >"
  sHTML = sHTML  & "function enviar(p1) { " 
   sHTML = sHTML  &  " if (document.all.datosx.action.indexOf('busca=' + p1)==-1){"
  sHTML = sHTML  &  "      document.all.datosx.action=document.all.datosx.action + '?busca=' + p1;} "
  sHTML = sHTML & " var x=document.all.datosx; "
  sHTML = sHTML & " if(x.Xxls.value=='NO') {x.target='_self';}else{x.target='blank';}"
  sHTML = sHTML  &  "document.all.datosx.submit(); "
  sHTML = sHTML  &"  }"
  sHTML = sHTML  &"  if (document.all.datosx.Xanalista.value==''){document.all.datosx.Xanalista.value=document.all.analistas.options[0].value;};"
  sHTML = sHTML  &"  if (document.all.datosx.Xgersector.value=='0'){document.all.datosx.Xgersector.value=document.all.Sector.options[0].value;};"
  sHTML = sHTML  &"  if (document.all.datosx.Xoriginante.value==''){document.all.datosx.Xoriginante.value=document.all.originante.options[0].value;};"
  sHTML = sHTML  & "</script>"
  
  Agregar = ""
  Buscar = ""   
  
  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,"")

  sHTML = sHTML  & "</body></html>"
end if


response.write sHTML



%>                