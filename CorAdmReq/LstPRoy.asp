<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario() 

if ok <> "Y" then   
  sHTML = ok                  
  
         
else 
  response.expires=0   
    
  session("ModMain")="CK"
  sHTML =   ofv.formHeader("")   

  sHTML = sHTML  & "<script language=javascript> "
  sHTML = sHTML  & "function AbrirCons(Datos) { "
  sHTML = sHTML  & " myWin= open(Datos," & chr(34) & "Cons" & chr(34) & "," & chr(34) & "status=yes,resizable=yes,menubar=no,toolbar=no,height=570,width=790" & chr(34) & " ); "
  sHTML = sHTML  & " } "
  sHTML = sHTML  & "</script>"                 
        
  CantRegMostrarx = 6      
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  set cn = ofv.conectar(ofv.strconn6) 
  set cn2 = ofv.conectar(ofv.strconn0)   
  set cn3 = ofv.conectar(ofv.strconn2)  
                  
  ofv.ObtenerAtributos session("Form"),""     
  bAgregar = ofv.AAgregar           
  bModificar = ofv.AModificar 
 
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar  
 
  sopcion = request.querystring("opcion")   
  sopc    = request.querystring("opc")   
  sTIPO   = "PRY" 'request.querystring("TIPO")
  sRESP   = request.querystring("RESP") 
  sSEL    = request.querystring("SEL") 
  XOPT    = request.querystring("XOPT")
  XRESP   = request.querystring("RESP")
  Xestado   = request.form("estado")
  Xfdesde   = request.form("fdesde1") 
  Xfhasta   = request.form("fhasta1")
  XORIGINANTE=request.form("XORIGINANTE")               
  XANALISTA=request.form("XANALISTA")
  xgersector=request.form("xgersector")
  xBusca=  ucase(request.querystring("busca")&"")
  xxls=request.form("xxls")
  'response.write Xestado
  select case cstr(Xestado&"")
  
  case "1" 
     descriestado="NUEVO"
  case "12" 
     descriestado="EN CURSO"
     
   case "13"   
    
    descriestado="FINALIZADO"
  end select 
 
             LISTA = " LISTA DE  PROYECTOS"

 
    sImagenO  = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=3>&#0128;</FONT></B>"
    sImagenI  = "<B><FONT FACE=WEBDINGS COLOR=#AA0000 SIZE=3>&#0128;</FONT></B>"
    sImagenE  = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0240;</FONT></B>"
    sImagenC  = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=3>R</FONT></B>"

    sImagenX  = "<B><FONT FACE=WEBDINGS COLOR=#f7aa22 SIZE=3>4</FONT></B>"
    sImagenIX = "<B><FONT FACE=WEBDINGS COLOR=#aa0000 SIZE=3>2</FONT></B>"

    sImagenpry = "<B><FONT FACE=WEBDINGS COLOR=#aa0000 SIZE=3>&#0143;</FONT></B>"
    sImagenreq = "<B><FONT FACE=WEBDINGS COLOR=#006600 SIZE=3>&#0153;</FONT></B>"
 



 DIM ATABLA(50,2) 

 ITAB=0
 TTAB=0 
 ITABB=0



    celdas = "" 

  
  
  'Encabezados
  
      sinput = sImagent  & "  <a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
      sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
      sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
      sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
      sinput = sinput  & "none" & chr(34) & ";' " 
      sinput = sinput  & " href="& session("svuelveAlm")&">Volver</a>"  
  
  if xxls="XLS" then
       Response.ContentType = "application/vnd.ms-excel"
       sinput=""
    end if
 
 
  formulario = formulario & ofv.GenCelda("","ttT colspan=7","","","","","",LISTA & DESTINO & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & sinput,"si")


  select case xBusca
   case "ESTADO"
        strsql="Select * From PROYECTO WHERE ESTADO =0" & Xestado
   
   case "RANGO"     
       strsql="Select * From PROYECTO WHERE FECHAINI BETWEEN " & XFDESDE & " AND " & XFHASTA
       
       
   case "GERSECTOR"     
       strsql="Select a.* From proyecto a ,PROYRESP b " 
       strsql=strsql & " where a.codigo=b.codigo and (b.gerencia=0"& xgersector &" or b.sector=0"&xgersector&")"  
       
   case "ANALISTA"     
       strsql="Select * From PROYECTO WHERE ASIGNADO='" & XANALISTA &"'"
   
   case "ORIGINANTE"     
       strsql="Select * From PROYECTO WHERE ORIGINANTE='" & XORIGINANTE & "'"
  end select
    
    'response.write strsql & "<br>"
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros) 
     if not rs.eof then
        'rs.move(CantRegMoverx)
        posi = 0
        formulario = formulario & GENLISTA(RS)

      celdas = ""   

      sInput = "&nbsp;"
      celdas = celdas & ofv.GenCelda("","vt colspan=3","","","","","",sInput,"si")

      sInput = "<b>" & sImagenC & "</b>"
      celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")

        
      sInput = "<b>&nbsp;Fecha Creacion</b>"
      celdas = celdas & ofv.GenCelda("","vt" ,"","","","","",sInput,"si")

      sInput = "<b>" & sImagenE & "</b>"
      celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")

      if stipo = "PRY" then
         sInput = "<b>&nbsp;Fecha Inicio</b>" 
      else
         sInput = "<b>&nbsp;Fecha Solicitud</b>"
      end if    
      celdas = celdas & ofv.GenCelda("","vt" ,"","","","","",sInput,"si")

     ' formulario = formulario & ofv.GenRow("","aaa height=30  ","","","","","",celdas,"si")

      celdas = ""

      sInput = "&nbsp;"
      celdas = celdas & ofv.GenCelda("","vt colspan=3","","","","","",sInput,"si")

      sInput = "<b>" & sImagenO & "</b>"
      celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")

      sInput = "<b>&nbsp;Usuario Originante</b>" 
      celdas = celdas & ofv.GenCelda("","vt align=center" ,"","","","","",sInput,"si")

      sInput = "<b>" & sImagenI & "</b>"
      celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")

      sInput = "<b>&nbsp;Usuario Interviniente</b>"
      celdas = celdas & ofv.GenCelda("","vt align=center" ,"","","","","",sInput,"si")

      'formulario = formulario & ofv.GenRow("","aaa  ","","","","","",celdas,"si")

     ELSE

        celdas = ofv.GenCelda("","VT colspan=7","center","","","","","<BR><BR><BR>NO HAY " & STABLA & "S PARA LA SELECCION<BR><BR><BR>","si")
        formulario = formulario & ofv.GenRow("","aaa  ","","","","","",celdas,"si")

     END IF


 ' END IF

  tabla1 = ofv.GenTabla("","","","98% BORDER","","0","","-1","-1","","",formulario,"si")
  sHTML = sHTML  & ofv.GenTabla("","","","98% ","","0","","-1","-1","","",formulario,"si")


  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "PROYLISTDSP.asp?Opcion=SI&OPC=1&TIPO=" & STIPO & "&RESP=" & SRESP & "'  "
    Agregar = REPLACE(ofv.BotonAgregar(sAccion),"Agregar","NUEVO " & STABLA)
  end if

  Buscar = ""
   Agregar = ""

  
  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

if xxls="XLS" then
  response.write tabla1
else
 response.write sHTML
end if




FUNCTION GENLISTA(RSX)



     
   celdas = ofv.GenCelda("","ttr ","CENTER","","","","","PROYECTO","si")
   celdas = celdas & ofv.GenCelda("","ttr","CENTER","","","","","Nombre","si")
   celdas = celdas & ofv.GenCelda("","ttr ","CENTER","","","","","fecha","si")
   celdas = celdas & ofv.GenCelda("","ttr","CENTER","","","","","usuario","si")
   celdas = celdas & ofv.GenCelda("","ttr","CENTER","","","","","sector","si")
   celdas = celdas & ofv.GenCelda("","ttr","CENTER","","","","","estado","si")
   celdas = celdas & ofv.GenCelda("","ttr ","CENTER","","","","","responsable","si")
   celdas = celdas & ofv.GenCelda("","ttr ","CENTER","","","","","dependencia","si")

           
   filaX = ofv.GenRow("","","","","20","","",celdas,"si")
      

   xTrColor = "#FFFFFF"

  posi = 0
  prow = "N" 
  Do While (Not rsX.EOF) 'and (posi < CantRegMostrarx)
  filas=""
  celdas=""

    sform = "REQ" & rsX(0)

    if prow = "S" then
       prow = "N"
       sr   = " bgcolor=#eeeeee "
    else 
       prow = "S"
       sr   = ""
    end if

       sr   = ""
       sc = "ut"

    


       CeldaS=""
       celdas = celdas & ofv.GenCelda("","ut","","5%","","","","&nbsp;","si")
       if xxls<>"XLS" THEN
       xlink = "<A  href='&nbsp;' class=nt "
       xlink2 =  replace(rsX(2)&""," " ," " )
  
       xlink = xlink & " onclick=" & chr(34) & "AbrirCons('editarPRY.asp?codigo=" & rsx(1) & "&tipo=PRY'); return false;" & chr(34)
       xlink = xlink & " >" & sImagenpry & xlink2 & "</A>"
       ELSE
        xlink =  rsX(2)
       END IF
       


      sInput = xlink

      celdas = celdas & ofv.GenCelda("","ut","","40%","","","",sInput,"si")



      'sInput = xlink2

      'celdas = celdas & ofv.GenCelda("","ut","","30%","","","",sInput,"si")



    if clng(rsx(3)) = 0 then
      dFecha = "---"
    else
      dFecha = cdate(rsx(3) + clng(cdate("01-01-1900")))
    end if

      sInput = dFecha
      celdas = celdas & ofv.GenCelda("","ut ","","","","","",sInput,"si")

      sInput = ""'"<b>" & sImagenE & "</b>"
      'celdas = celdas & ofv.GenCelda("","ut","","1%","","","",sInput,"si")

    if clng(rsx(4)) = 0 then
      dFecha = "---"
    else
      dFecha = cdate(rsx(4) + clng(cdate("01-01-1900")))
    end if

      sInput = dFecha
 
   
      sniv = "---"

            'strsql = "Select * From USUARIOS where ID = '" & RSX(13) & "' "
             strsql = "Select A.DESCRIPCION,B.DESCRI From USUARIOS A,ECO B where A.SECTOR=B.ECO AND A.ID = '" & RSX(13) & "' "
            ' set rsorig = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
		   
	        set rsorig = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
               ' if not rsniv.eof then sniv = rsniv(1)
                call ofv.cerrarconsulta(rsniv)	   
		    if not rsorig.eof then
		       XORIGINANTE = rsorig(0)
		       XORIGINANTESECTOR = rsorig(1)   
		   ELSE
		       XORIGINANTE = "NO DETERMINADO"
		       XORIGINANTESECTOR = rsorig(1)
		   END IF
              call ofv.cerrarconsulta(rsorig)
              
               

           
      celdas = celdas & ofv.GenCelda("","ut" ,"","37% align=center","","","",XORIGINANTE,"si")

      sInput = replace(XORIGINANTESECTOR&""," " ," ") 
      celdas = celdas & ofv.GenCelda("","ut align=center ","","1%","","","",sInput,"si")
      celdas = celdas & ofv.GenCelda("","vt","","1%","","","",descriestado,"si")

      sniv = "---"

                strsql = "Select * From USUARIOS where ID = '" & RSX(14) & "' "
                set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
                if not rsniv.eof then sniv = rsniv(1)
                call ofv.cerrarconsulta(rsniv)

                        
              
      celdas = celdas & ofv.GenCelda("","ut" ,"","37% align=center","","","", replace(sniv&""," " ," ") ,"si")
      
      
      'DEPENDENCIAS
      
          
            XPRYNOM="&nbsp;"
   '       
   '         StrSql = "Select * From PROYtareasLST where codigo = '" & rs1(5) & "' and TAREA = '" & rs1(20) & "' order by orden "
   '         set rs1Dep = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   '        if rs1Dep.eof then
   '       
   '             XPRYNOM = "&nbsp;"
   '     
   '       
   '         ELSE    
   '           do while not rs1Dep.eof
   '             StrSql = "Select * From PROYECTO WHERE CODIGO = '" & rs1Dep(4) & "' "
   '             set RS2Dep = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   '               IF NOT RS2Dep.EOF THEN
   '                XPRYNOM = XPRYNOM & RS2Dep(2) & "<br>"
   '               END IF
   '             call ofv.cerrarconsulta(RS2Dep)
   '           rs1Dep.movenext
   '           loop
   '        
   '         end if
   '       
   '   call ofv.cerrarconsulta(rs1Dep)
      

     celdas = celdas & ofv.GenCelda("","ut" ,"","37%","","","",XPRYNOM,"si")

    filas = filas & ofv.GenRow("","aaa  BGCOLOR='" & xTrColor & "' ","","","","","",celdas,"si")




    formu = formu & filas

      celdas = ofv.GenCelda("","AAA COLSPAN=7 style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#CCCCCC;FONT-SIZE:5pt' ","","1%","","","","&nbsp;","si")

    'formu = formu & ofv.GenRow("","aaa  BGCOLOR='" & xTrColor & "' ","","","","","",celdas,"si")


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9" 
                         ELSE
                            xTrColor = "#FFFFFF" 
                         END IF

    posi = posi + 1
    rsX.Movenext
  loop



  GENLISTA = FILAX & FORMU

END FUNCTION

%>