<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  sTipo = request.querystring("Tipo")
  sNombre = request.querystring("Nombre")
  sNombre2 = ofv.convertircar(sNombre,"_"," ")
  sElemento = request.querystring("Elemento")
  n = instr(sElemento,"-")
  nID = mid(sElemento,n+1)




  sTitulo = "Agregar " & sNombre2
  sTitulo = UCase(sTitulo)


  sHTML = ofv.FormHeader(sTitulo)
  volver= ofv.vueltaasp2("../")


n = instr(sElemento,"-")
if n > 0 then
  nCodMaster = left(sElemento,n-1)
  sTipoMaster = mid(sElemento,n+1)
  n = instr(sTipoMaster,"-")
  sTipoMaster = left(sTipoMaster,n-2)
end if

SPTIPO = request.querystring("SPTIPO")

WKSPAPLICSET = "RG"

if SPTIPO = "" then SPTIPO = "SP"

    WKSPTIPO    = SPTIPO
    WKSPNOMBRE  = ""
    WKSPCODIDEN = ""


 if SPTIPO = "CR" then 
    WKSPTIPO    = "CROBJ"
    WKSPNOMBRE  = nCodMaster & "-" & sTipoMaster & "- VINCULACION EXPLICITA CR" 
    WKSPCODIDEN = nCodMaster & "-" & sTipoMaster & "- CR" 
 END IF

 if SPTIPO = "AI" then 
    WKSPTIPO    = "AIOBJ"
    WKSPNOMBRE  = nCodMaster & "-" & sTipoMaster & "- VINCULACION ACTIVOS INFORMATICOS AI" 
    WKSPCODIDEN = nCodMaster & "-" & sTipoMaster & "- AI" 
 END IF


 if SPTIPO = "PL" then 
    WKSPTIPO    = "PLOBJ"
    WKSPNOMBRE  = nCodMaster & "-" & sTipoMaster & "- SOPORTE-CIRCUITO-CONTINGENCIAS" 
    WKSPCODIDEN = nCodMaster & "-" & sTipoMaster & "- PL" 
    WKSPAPLICSET = "PL"
 END IF



    f1a = "asp/AgregarEsquema.asp?Tipo=" & sTipo & "&SPTIPO=" &  SPTIPO
    f1a = f1a & "&Nombre=" & sNombre & "&Elemento=" & sElemento
  sHTML = sHTML & ofv.GenForm("Agregarb","",f1a,"","post","","","no")

  sHTML = sHTML & ofv.GenTabla("panel","","","80%","",0,0,0,0,"","","","no")

  'Encabezado Descripcion
   t1 = ofv.GenCelda("","tt","","2%","","","","Descripcion","si")
  'Texto Descripcion
   t1 = t1 & ofv.GenCelda("","","","","","","",ofv.GenerarInput("Descri",WKSPNOMBRE,"agregar","","60","dt",""),"si")
  sHTML = sHTML & ofv.GenRow("","","","","","","",T1,"si")

   SINPUTD1 = ofv.GenerarInput("OBJUN",SESSION("UN"),"HIDDEN","","60","dt","")
   SINPUTD1 = SINPUTD1 &  ofv.GenerarInput("aplic",WKSPAPLICSET,"HIDDEN","","1","dt","")




   SINPUTD1 = SINPUTD1 &  ofv.GenerarInput("codvinOBJ",WKSPTIPO,"HIDDEN","","1","dt","")

   SINPUTD1 = SINPUTD1 &  ofv.GenerarInput("RELLINK","0","HIDDEN","","1","dt","")
   SINPUTD1 = SINPUTD1 &  ofv.GenerarInput("prioridad","1","HIDDEN","","1","dt","")

   t2 = ofv.GenCelda("","","","","","","","&nbsp;" &  SINPUTD1,"si")
   t2 = t2 & ofv.GenCelda("","","","","","","","&nbsp;","si")
  sHTML = sHTML & ofv.GenRow("","","","","","","",T2,"si")

  'Encabezado Identificacion
   t3 = ofv.GenCelda("","tt","","2%","","","","Identificacion","si")
  'Texto Identificacion
   t3 = t3 & ofv.GenCelda("","","","","","","",ofv.GenerarInput("Codiden",WKSPCODIDEN,"agregar","","40","dt",""),"si")
  sHTML = sHTML & ofv.GenRow("","","","","","","",T3,"si")

  sHTML = sHTML & ofv.GenFTabla()

  'Botones Aceptar y Cancelar
  sAccion = "RGDETAPROC.asp?Nombre=" & sNombre & "&codmaster=" & nCodMaster & "&tipmaster=" & sTipoMaster
  sHTML = sHTML & ofv.AgregarBoton("Agregarb",sAccion,"no")


 if SPTIPO = "CR" OR SPTIPO = "AI" OR SPTIPO = "PL" then 
          sHTML = sHTML & "<script language='javascript' >"
          sHTML = sHTML & " document.Agregarb.submit(); "
          sHTML = sHTML & "</script>"
 END IF


  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>