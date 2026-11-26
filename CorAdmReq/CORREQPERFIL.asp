<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))



  CantRegMostrarx = 12
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  set cn = ofv.conectar(ofv.strconn6)
''  set cn2 = ofv.conectar(ofv.strconn0)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  sopcion = request.querystring("opcion")


 if sopcion = "SI" then

    snv = request.querystring("version")

    if snv = "1" then
       GENPERFIL()
    else
       GENPERFILNV()
    end if

 end if


    sImagenC = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=4>&#0232;</FONT></B>"


    celdas = ""






    celdas = ""



  'Encabezados
  formulario = ofv.GenCelda("","ttt colspan=4","","","","","","PERFILES DE REQUERIMIENTOS","si")
  formulario = formulario & ofv.Encabezados("CODIGO","DESCRIPCION","ACTIVO","PROPIEDADES")


  StrSql = "Select * From TIPOREQNV ORDER BY TIPOREQCOD"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)

    sform = "REQNV" & rs(0)


       sc = "ut"

    'Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=TIPOREQNV&volver=" & vuelta & "&Eliminar=NSNC"
      Celdas = ofv.BotonEliminar(rs(2),sHref)
    else
      Celdas = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

 '   celdas = ofv.GenCelda("","","","","","","","&nbsp;","si")



       sInput = ofv.GenerarInput("",rs(1),"readonly","","5","dtn","")
       celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")




       if bModificar then
          sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sform,"90","dt","")
       else
          sInput = ofv.GenerarInput("",rs(2),"readonly","","90","dt","")
       end if
       celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")


         scomK = "N" & chr(9) & "S"
         scomD = "NO" & chr(9) & "SI"

         if bModificar then
             sInput = ofv.gencomboFX(sForm,rs(14).name,rs(14),scomK,scomD," ")
         else
             sInput = ofv.GenerarInput("",rs(14),"readonly","","70","dt","")
         end if
      CeldaS = celdaS & ofv.GenCelda("","uti","","","","","",sInput,"si")




       sinput = "<a  href=CORREQPERFILPROPNV.asp?PERFIL=" & RS(0) & "&CantRegMover=" & CantRegMoverx
       sinput = sinput  & " ><b>PROPIEDADES</b></a>" 

    '''   celdas = celdas & ofv.GenCelda("","TT","center","","","","",sInput,"si")


       sinput2 = "<a  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
       sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
       sinput2 = sinput2  & "none" & chr(34) & ";'  href='CORREQPERFILPROPNV.asp?PERFIL=" & RS(0) & "&CantRegMover=" & CantRegMoverx
       sinput2 = sinput2  & "' TITLE='Propiedades DE " & rs(2) & "'  ><b>" & sImagenC & "</b></a>" 

       celdas = celdas & ofv.GenCelda("","ut","CENTER","","","","",sInput2,"si")



    filas = ofv.GenRow("","","","","","","",celdas,"si")

    if bModificar then
      sAccion = "asp/Actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*") & "&volver=" & vuelta & "&Tabla=TIPOREQNV"
      formulario = formulario & ofv.GenForm(sform,"",sAccion,"","post","",filas,"si")
    else
      formulario = formulario & ofv.GenForm("","","","","post","",filas,"si")
    end if

    posi = posi + 1
    rs.Movenext
  loop

  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",formulario,"si")


  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "CORREQPERFIL.asp?Opcion=SI&VERSION=2"
    Agregar = ofv.BotonAgregar(sAccion)
  end if


  VOLVER = ""


  Buscar = ""
  

  
  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,VOLVER)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if


response.write sHTML

FUNCTION GENPERFIL()

  StrSql = "Select * From TIPOREQ "
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  DATA = ""   
  DATA2 = ""   
  TREQ = 100

  IF NOT RS1.EOF THEN

     StrSql = "Select MAX(TIPOREQ) From TIPOREQ "
     set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

     TREQ = CLNG(RS2(0)) + 10


  END IF

  FOR II=1 TO RS1.FIELDS.COUNT - 1

      IF II > 1 THEN
         DATA  = DATA & ","
         DATA2 = DATA2 & ","
      END IF

      DATA = DATA & RS1(II).NAME

      IF II=1 THEN 
         DATA2 = DATA2 & "'" & TREQ & "'"  
      ELSEIF II=2 THEN 
         DATA2 = DATA2 & "'NUEVO PERFIL'"  
      ELSE
         IF II > 2 AND II < 18 THEN
            IF RS1(II) THEN
               DATA2 = DATA2 & "'1'" 
            ELSE
               DATA2 = DATA2 & "'0'" 
            END IF 

         ELSE
           DATA2 = DATA2 & "'" & RS1(II) & "'" 
         END IF
      END IF
  NEXT

 ' RESPONSE.WRITE DATA



    strsql = "insert into TIPOREQ (" & DATA & ") values (" & DATA2 & ")"

'    strsql = "insert into TIPOREQ (" & DATA & ") values ('" & TREQ & "','NUEVO PERFIL',"
'    strsql = strsql & "'1','0','0','0','0','0','0','0','0','0','0','0','0','0','0'"
'    strsql = strsql & ",'COMERCIAL','OPERATIVA','SEGURIDAD','IMPOSITIVA','CONTABLE',"
'    strsql = strsql & "'LEGALES','EQUIPAMIENTO','LOGISTICA','ANTECEDENTES','NO','NO','NO','NO','NO','NO',0,0)"

    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

'' RESPONSE.WRITE strsql

END FUNCTION

FUNCTION GENPERFILNV()

  StrSql = "Select * From TIPOREQNV "
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  DATA = ""   
  DATA2 = ""   
  TREQ = 500

  IF NOT RS1.EOF THEN

     StrSql = "Select MAX(TIPOREQCOD) From TIPOREQNV "
     set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

     TREQ = CLNG(RS2(0)) + 10


  END IF


    strsql = "insert into TIPOREQNV (TIPOREQCOD,DESCRIPCION,CRPTPRY,CTADEF,SUPNIVEL,DESTSECTOR,DESTNIVEL,PERMISO,ACTIVO)"
    strsql = strsql & " values ("
    strsql = strsql & "'" & TREQ & "','NUEVO PERFIL','S','1','8','0','8','N','N'"
    strsql = strsql & ")"

    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

'' RESPONSE.WRITE strsql

END FUNCTION



%>