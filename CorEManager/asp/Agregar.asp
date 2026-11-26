<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%
ok=ofv.CheckUsuario()
response.buffer = true

response.expires=0

nCampo = request.querystring("Campo")
if nCampo = "" then nCampo = 0
nCampo = cdbl(nCampo)
Tabla = request.querystring("Tabla")
ID = request.querystring("ID")
Volver = request.querystring("Volver")
Volver = ofv.convertircar(Volver,"_","&")
Regisvuelta = cdbl(request.querystring("CantRegMover"))

on error resume next

sDupKey = 0

set cn = ofv.conectar(ofv.strconn2)

strsql = "Select * From " & tabla

set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

wkeyname = rs(0).name

s = ofv.QueEsx(rs(0).type)
if rs(0).type = 200 then s = "'"

wformname = ""
wmensaje = ""
For Each x In Request.form
    if x = wkeyname then
       wformname = x
       exit for
    end if
next

if wkeyname = wformname then
   sDupKey = 1
   strsql = "select * from " & tabla & " where " & wkeyname & " =" & s & request.form(1) & s

   set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
   if rs.eof then sDupKey = 0
end if

if sDupKey <> 1 and (Ucase(tabla) = "SUBTEMA" or Ucase(tabla) = "TIPOBJETO") then
   sDupKey = 1
   select case Ucase(tabla)
          case "SUBTEMA"
               tabla2 = "TIPOBJETO"
               wmensaje = " como Tipo de Esquema Normativo"
          case "TIPOBJETO"
               tabla2 = "SUBTEMA"
               wmensaje = " como Tipo de Sub Tema"
   end select

   strsql = "select * from " & tabla2 & " where " & wkeyname & " =" & s & request.form(1)  & s
   set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
   if rs.eof then sDupKey = 0
end if


if sDupKey <> 1 then
i = 0
for each fld in rs.fields
  redim preserve vDesc(i)
  vDesc(i) = fld.name
  i = i + 1
next
sCampo = rs(nCampo).name

s = ofv.QueEsx(rs(nCampo).type)
if rs(nCampo).type = 200 then s = "'"
sValor = s & UCase(request.form(1)) & s
if request.form(1) = "" and s <> "'" Then sValor = 0

strsql = "Insert Into " & tabla & " ("
For Each x In Request.form
  strsql = strsql & x & ","

Next
strsql = mid(strsql,1,len(strsql)-1)
strsql = strsql & ") Values ("

ValorNulo = False
TipoValido = True
For Each x In Request.form


    if  (x="SHWWORD" and ucase(tabla) = "GLOSARIO")   OR   (UCASE(x) = "TIPOMODO" and ucase(tabla) = "MODO") then
      vNuevoValor = UCase(CStr(ofv.convertircar(Request.Form(x),"_"," ")))
		 if  vNuevoValor ="NO" THEN
		  vNuevoValor = UCase(CStr(ofv.convertircar(Request.Form("KEYWORD"),"_"," ")))
	    end if
	ELSEIF  (UCASE(x) = "TIPOBJETO" and ucase(tabla) = "TIPOBJETO") OR   (UCASE(x) = "TIPOBJETO" and ucase(tabla) = "SUBTEMA") OR   (UCASE(x) = "TIPOPERFIL" and ucase(tabla) = "OBJPERFIL") OR   (UCASE(x) = "TIPOPERFIL" and ucase(tabla) = "COMPERFIL") OR   (UCASE(x) = "TIPO" and ucase(tabla) = "TIPOCOMP")   OR   (UCASE(x) = "TABPERFIL" and ucase(tabla) = "TABPERFIL") OR   (UCASE(x) = "TIPOPERFIL" and ucase(tabla) = "TIPOPERFIL") THEN
	    vNuevoValor=REPLACE(Request.Form(x)," ","")
	    vNuevoValor=REPLACE(vNuevoValor,"_","")
	    vNuevoValor=REPLACE(vNuevoValor,"-","")
	     vNuevoValor=REPLACE(vNuevoValor,"*","")
	     vNuevoValor=REPLACE(vNuevoValor,"/","")
         vNuevoValor=REPLACE(vNuevoValor,"\","")
         vNuevoValor=REPLACE(vNuevoValor,"'","")
         vNuevoValor=REPLACE(vNuevoValor,"´","")
         vNuevoValor=REPLACE(vNuevoValor,"&","")
         vNuevoValor=REPLACE(vNuevoValor,".","")


    elseif x <> "KEYWORD" and ucase(tabla) = "GLOSARIO" then

       vNuevoValor = Request.Form(x)
    elseif x = wkeyname and ucase(tabla) = "TABPERFIL" then
       vNuevoValor = UCase(CStr(ofv.convertircar(Request.Form(x)," ","1")))
    elseif ucase(tabla) = "SUBTEMASKR" then
     vNuevoValor = UCase(CSTR(Request.Form(x)))
    else

'' 07112014 - anulada por mayuscula minuscula
''   vNuevoValor = UCase(CStr(ofv.convertircar(Request.Form(x),"_"," ")))
'' 07112014 - anulada por mayuscula minuscula

     vNuevoValor = CStr(ofv.convertircar(Request.Form(x),"_"," "))
    end if
  nx = 0
  for i = 0 to UBound(vDesc)
    if x = vDesc(i) then
      nx = i
      exit for
    end if
  next
  s = ofv.QueEsx(rs(nx).type)
  if rs(nx).type = 200 then s = "'"
  ValorNulo = ofv.EsNulo(rs(nx),vNuevoValor)
  if ValorNulo then exit for
  TipoValido = ofv.ValidarTipo(rs(nx),vNuevoValor)
  if not TipoValido then exit for
  if vNuevoValor = "TRUE" then
     vNuevoValor = 1
  elseif vNuevoValor = "FALSE" then
     vNuevoValor = 0
  end if
  strsql = strsql & s & vNuevoValor & s & ","
 ' xvalact=vNuevoValor

Next

if not TipoValido or ValorNulo then
  on error resume next
  response.write ofv.ErrorTipo(volver,Request.Form(x),rs(nx).name)
  if err.number <> 0 then
     response.write ofv.ErrorTipo(volver,"","")
     response.flush
  end if

else
  strsql = mid(strsql,1,len(strsql)-1) & ")"

 ' strsqlw="select * from " & tabla & " where " &  wformname  &"=" & request.form(1) & "'"
  ' response.write strsql
   STRSQL="PRC_EXECSQL '" & REPLACE(STRSQL,"'","''") & "'"
' STRSQL="SET NOCOUNT ON;"& STRSQL & " ;SELECT @@IDENTITY AS ID;SET NOCOUNT OFF;"

 SET VV= ofv.crearconsultaEx(StrSql,cn,1,volver)


  set errores = cn.errors
  if cn.errors.count = 0 then
     IF NOT ISNULL(VV(0)) THEN
       IDINSERT=VV(0)
     ELSE
       IDINSERT=request.form(wkeyname)
     END IF
   ' RESPONSE.WRITE VV(0) & "222"

     xxxr=GENLOGEXTDATA("INSERT",TABLA,wkeyname,IDINSERT,"PAR")
     IF ucase(tabla) = "GLOSARIO" THEN
        vNuevoValor = UCase(CStr(ofv.convertircar(Request.Form("KEYWORD"),"_"," ")))
        strsql = "select OBJID,SHWWORD from GLOSARIO where KEYWORD = '" & vNuevoValor & "' "
        set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
        if NOT rs.eof then

           GLOSCOD = RS(0)
           GLOSWRD = RS(1)
           GLOSFEC = CLNG(DATE()) - CLNG(CDATE("01-01-1900"))
              xxxx = GETFECH()
           GLOSHOR = XLHORDAT

           strsql = "INSERT INTO GLOSDSC (CODWORD,FECVIG,HORVIG,DSCWORD) "
           strsql = strsql & " VALUES ('" & GLOSCOD & "','" & GLOSFEC & "','" & GLOSHOR & "','" & GLOSWRD & "') "
           call ofv.crearconsultaEx(StrSql,cn,1,volver)


           strsql = "select OBJID from GLOSDSC where CODWORD = " & GLOSCOD & " AND FECVIG = " & GLOSFEC & " AND HORVIG = " & GLOSHOR
           set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
           if NOT rs.eof then

              GLOSDSC = RS(0)
              GLOSUSR = SESSION("USUARIO")


              strsql = "INSERT INTO GLOSUSR (CODWORD,CODDSC,USRWORD,TIPMOD,FECMOD,HORMOD) "
              strsql = strsql & " VALUES ('" & GLOSCOD & "','" & GLOSDSC & "','" & GLOSUSR & "','A','" & GLOSFEC & "','" & GLOSHOR & "') "
              call ofv.crearconsultaEx(StrSql,cn,1,volver)




           END IF


        END IF


     END IF

    response.clear
  'RESPONSE.WRITE STRSQL
   response.redirect volver
  else
    response.write ofv.ErrorDB(cn,volver)
    response.flush
  end if
end if

else

  sHTML = ofv.MenuHeader("#eeeeee","")

  sHTML = sHTML & "<br><br><br><br><center>"
  'Jerarquia
  Celda = ofv.gencelda("","","","","","","","&nbsp;","si")

  Fila = ofv.genrow("","","","","","","",Celda,"si")

  Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

  Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

  Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

  Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")


  sInput = sInput & "<font face=tahoma size=2 color=#000080> "
  sInput = sInput & "Error al agregar: el Codigo <font face=tahoma size=2 color=#993333><b> "
  sInput = sInput & s & request.form(1) & s & "</b></font> ya existe " & wmensaje & "</font>"
  Celda = ofv.gencelda("","ut","center","100%","","","",sInput,"si")


  Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

  Celda = ofv.gencelda("","","","","","","","&nbsp;","si")

  Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

  Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

  Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

  Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

  h1 = ofv.GenCelda("","","center","80%","20","","","<input class=it type=submit size=" & ofv.wsize & " name=boton value=Volver>","si")
  Fila = Fila & ofv.GenRow("","","","","","","",h1,"si")

  sAccion = volver
  Formulario = ofv.GenForm("vuelta","",sAccion,"","post","",Fila,"no")

  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")

  sHTML = sHTML & "</center></body></html>"

  response.write sHTML

end if

call ofv.cerrarconsulta(rs)
call ofv.cerrarconn(cn)


%>