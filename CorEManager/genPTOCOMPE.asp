<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<% 
   
ok=ofv.CheckUsuario() 
if ok <> "Y" then
  uv = ofv.uv 
  if uv = "NO" then ok = ofv.ValidarUsuario 
  sHTML = ok    
else     
  response.expires=0         
  'response.write session("codSol")          
  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar    
  if session("Nomodificar") = "S" then bModificar = False   

  sHTML = sHTML & "<HTML>"
  sHTML = sHTML & "<BODY onload='window.top.ldfobj();' >"
  

  SCOMPE = request.querystring("COMPE")    
  sDATO   = request.querystring("DATO")  
  sperfil = request.querystring("perfil")

  
  vuelta = "../genObsdoc.asp?id=" & ID  


    

   
    
  set cn = ofv.conectar(ofv.strconn2)





       
  'response.write "<center>"

    Fila = ""

    wkproccod = ""
    wkproccodA = ""

    sImage = "<font face='corflaticon' size=2 color='#f0f0f0'  >&#0120;</font>"
    sImagea = "<font face='corflaticon' size=2 color='#f0f0f0'  >&#0121;</font>"

    sImaged = "<font face='corflaticon' size=2 color='#f0f0f0' style='background-color:#993333;border-style:OUTSET;border-width:2;' >&#0214;</font>"

    sImageSH = "<font face='corflaticon' size=2 color='#f0f0f0'  >&#0244;</font>"
    sImageSW = "<font face='corflaticon' size=2 color='#f0f0f0'  >&#0243;</font>"
    sImageSS = "<font face='corflaticon' size=2 color='#f0f0f0'  >&#0113;</font>"


    wkproccod = "<a class=nt  href='javascript:void(0);' >"
    wkproccod = wkproccod & sImage
    wkproccod = wkproccod & "</a>"

    wkproccodA = "<a class=nt target='WK4' href='genObsdoc.asp?id=" & ID  & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&OP=INS' >"
    wkproccodA = wkproccodA & sImagea
    wkproccodA = wkproccodA & "</a>"

    wkproccodSH = "<a class=nt  href='javascript:void(0);' >"
    wkproccodSH = wkproccodSH & sImageSH
    wkproccodSH = wkproccodSH & "</a>"


    wkproccodSw = "<a class=nt  href='javascript:void(0);' >"
    wkproccodSw = wkproccodSw & sImageSw
    wkproccodSw = wkproccodSw & "</a>"

    wkproccodSS = "<a class=nt  href='javascript:void(0);' >"
    wkproccodSS = wkproccodSS & sImageSS
    wkproccodSS = wkproccodSS & "</a>"



    CELDAD = ""
    CeldaD = CeldaD & ofv.GenCelda("","utI  ","","95%","","","","LISTA DE COMPETENCIAS","si")
    CeldaD = CeldaD & ofv.GenCelda("","ut title='Ventana normal.'  onclick='GTCOBSSS();' style='background-color:#6699cc;border-style:OUTSET;border-width:1;' ","center","2%","","","",wkproccodSS,"si")
    CeldaD = CeldaD & ofv.GenCelda("","ut title='Ampliar alto.'  onclick='GTCOBSHM();' style='background-color:#f7aa22;border-style:OUTSET;border-width:1;' ","center","2%","","","",wkproccodSH,"si")
    CeldaD = CeldaD & ofv.GenCelda("","ut title='Ampliar ancho.'  onclick='GTCOBSWM();' style='background-color:#f7aa22;border-style:OUTSET;border-width:1;' ","center","2%","","","",wkproccodSw,"si")
    CeldaD = CeldaD & ofv.GenCelda("","ut title='Cerrar.' style='background-color:#bb0000;border-style:OUTSET;border-width:1;' onclick="" document.getElementById('DX4').style.display='none'; "" ","center","2%","","","",wkproccod,"si")
    FilaD     = ofv.genrow("","","","","","","",CeldaD,"si")
    sFORMhed = ofv.GenTabla("","TWIND","","100%","","0","","0","1","","",FilaD,"si")


   shtml = shtml & sFORMhed


     fila=""


           
 



  strsql = "Select * From TIPOCOMPEDET WHERE COMPETENCIA = 0" & SCOMPE & " order by SECUEN,descriPCION"
'' response.write strsql  
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   
    
  
   if  rs.eof then


    Celda = ofv.GenCelda("","ut colspan=2 ","","","","","","&nbsp;","si")
    Fila = fila & ofv.genrow("","","","","20","","",Celda,"si")

    Celda = ofv.GenCelda("","uti colspan=2 style='border-style:solid;border-color:#888888;border-width:1;'  ","center","","","","","<big>NO HAY COMPETENCIAS DEFINIDAS.</big>","si")
    Fila = fila & ofv.genrow("","","","","50","","",Celda,"si")





   end if
  
  iform=0
  
  do while  not rs.eof  

        wkproccodA = "<a class=nt  href='SECTORPERFILESFUN.ASP?SOLAPA=COM&PERFIL=" & SPERFIL & "&OPCION=UPD&DATO=" & SDATO & "&COMPE=" & RS(0) & "' ><b>"
        wkproccodA = wkproccodA & RS(2)
        wkproccodA = wkproccodA & "</b></a>"
	Celda = ofv.GenCelda("","utI  title='SELECCIONAR COMPETENCIA.' ","","","","","",wkproccodA,"si")
    Fila = fila & ofv.genrow("","","","","20","","",Celda,"si")

      
       

   

rs.movenext
iform=iform+1
loop  


 shtml =shtml & ofv.GenTabla("","","","80%   ","","0","","0","0","","",fila,"si")



  Call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML & "</BODY>"

  sHTML = sHTML & "</HTML>"
end if

 response.write shtml



%>    