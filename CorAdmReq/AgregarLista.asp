<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
  ncaller = request.querystring("caller")
  nCodigo = request.querystring("Cod")
  sTipo = request.querystring("Tipobjeto")
  sCod = request.querystring("CodOBJETO")
  Volver = request.querystring("Volver")
  CantRegMover= REQUEST.QUERYSTRING("CantRegMover")
  
  Stabla=request.querystring("tabla")
  SCampo=request.querystring("Campo")
  if SCampo<>"" then
   session("SCampo")=SCampo
  else
   SCampo=session("SCampo")
  end if
  
  if sCod<>"" then
     session("sCod")=sCod
    else
     sCod=session("sCod")
  end if
  
  sxopc1 = ""


'''''''''''''''''''''''''''''''''''''''''''''''''''

     sHTML = ofv.FormHeader("Agregar Usuario")

     SHTML = SHTML & GenBuscaUsr("N","N","N","S",CHR(9),CHR(9),CHR(9),CHR(9))

     IF SESSION("FBUSR") <> "NO" THEN

        set cn = ofv.conectar(ofv.strconn6)

        
    
      select case Stabla
      
      CASE "GPREQXUSR"
        strsql = "Select * From " & Stabla
	Volver = Volver & "?Cod=" & sCod       
        set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
        f1a = "asp/AgregarEN.asp?fname=" & ofv.convertircar(sCod," ","*") & "&volver=../"&Volver &"&Tabla=" &stabla
        f1a = f1a 

        sHTML = sHTML & ofv.GenForm("AgreArchb","",f1a,"","post","","","no")
        
        sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='"&SESSION("FBUSR") &"' NAME='" & rs(cint("0"& scampo)).name & "' >"
        sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='U' NAME='TIPOUSR' >"
        sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='"&sCod&"' NAME='TIPOREQCOD' >"
        call ofv.cerrarconsulta(rs)
        
      CASE "GPREQRESPONSABLES"
      
         f1a = "asp/Actualizar.asp?fname=" & ofv.convertircar(sCod," ","*") & "&volver=../"&Volver &"&Tabla=" &stabla
	 f1a = f1a
	  
	  sHTML = sHTML & ofv.GenForm("AgreArchb","",f1a,"","post","","","no")
	  sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='"&SESSION("FBUSR")&"' NAME='USUARIOSECTOR' >"

      CASE "GPREQRESPONSABLESJER"
      
         f1a = "asp/Actualizar.asp?fname=" & ofv.convertircar(sCod," ","*") & "&volver=../"&Volver &"&Tabla=" &stabla
	 f1a = f1a
	  
	  sHTML = sHTML & ofv.GenForm("AgreArchb","",f1a,"","post","","","no")


	  sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='"&SESSION("FBUSR")&"' NAME='"

          IF ncaller = "N1" THEN sHTML = sHTML & "USUARION1"
          IF ncaller = "N2" THEN sHTML = sHTML & "USUARION2"

	  sHTML = sHTML & "' >"

      END SELECT    
        
        call ofv.cerrarconn(cn)



      END IF

''''''''''''''''''''''''''''''''''''''''''''''''''''




  'Botones Aceptar y Cancelar
  'if ncaller = "AUT" then 
  'sAccion = "AutDefaultdet.asp"
  'else
  sAccion = Volver '"gpresponsables.asp"
  sHTML = sHTML & ofv.AgregarBoton("AgreArchb",sAccion,"_self")


  Agregar = ""
  Buscar  = ""
  Celda   = ""


  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>