<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

response.expires=0
dim SvalidaBcra
SvalidaBcra="NO"
id = request.querystring("fname")
idsv = request.querystring("fname")
id = ofv.convertircar(id,"*"," ")
tabla = request.querystring("tabla")
volver = request.querystring("volver")
volver = ofv.convertircar(volver,"_","&")
nEditar=request.querystring("Editar")







on error resume next

set cn = ofv.conectar(ofv.strconn2)
nFecini = 0

if nFecini < clng(cdate("01-01-1900")) then
   nFecini = clng(cdate("01-01-1900"))
else
   nFecini = 0
end if

strsql = "Select * From " & tabla
set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
IDROWNAME=RS(0).NAME

CALL OFV.CERRARCONSULTA(RS)

strsql = "Select * From " & tabla & " WHERE " & IDROWNAME & "='" & id &"'"
'response.write strsql & "<br>"
set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)


strsql = "Update " & tabla & " Set "
XENTRO=FALSE
XENTRO2=FALSE
For Each x In Request.form

  ''  if ( ucase(tabla) = "TIPOBJETO" and x = "TEXTOIMPRESION" ) OR ( ucase(tabla) = "ECOPERFILDSC" and x = "PTODESCRI" ) or ( ucase(tabla) = "TEMA" and x = "PSURL" ) or ( ucase(tabla) = "PRODSERV" and x = "PSURL" ) or ( ucase(tabla) = "NOVEGENPARIMG"  )  or ( ucase(tabla) = "NOVEGENPAR"  ) then

  ''      vNuevoValor = Request.Form(x)

 ''   else

   ''     vNuevoValor = UCase(CStr(ofv.convertircar(Request.Form(x),"_"," ")))


  ''  end if

vNuevoValor = Request.Form(x) & ""



    if ( ucase(tabla) = "ORVARIABLES" and x = "VALOR" )  then

        vNuevoValor = REPLACE(CSTR(vNuevoValor & ""),",",".")


    end if

'response.write  ucase(tabla) & "  "  & ucase(x) & "   " & session("APLMDCLI")
	if ( ucase(tabla) = "TIPOPERFIL" and UCASE(x) = "ESTADO" ) and session("APLMDCLI") = "BNX" then
    'porque tiene * en el id
	        ID = idsv


    end if



   'response.write x & ": " & vNuevoValor & "<br>"

  if UCase(vNuevoValor) = "FALSE" then vNuevoValor = "0"
  if UCase(vNuevoValor) = "TRUE" then vNuevoValor = "-1"

  if UCase(vNuevoValor) = "FALSO" then vNuevoValor = "0"
  if UCase(vNuevoValor) = "VERDADERO" then vNuevoValor = "-1"

  nx = ucase(x)


'session("APLMDLOR")="BE"
 ' IF  session("APLMDLOR")="BE" THEN

  SELECT CASE nx



   CASE "CODBCRA"

     Str_a=cstr(vNuevoValor&"")
     Str_b=cstr(rs(x)&"")
	  if Str_a <> Str_b then

	    select case ucase(tabla)

	    case  "ORIMPRIESGO"


	      XAGRUPA=RS("AGRUPA")
	     strsqlbcra="select * from " & tabla & " where AGRUPA='-1'"

	    case "ORLINNEGO2"
	     XAGRUPA=RS("nivel1")
	     strsqlbcra="select * from " & tabla & " where nivel1='"& XAGRUPA& "' and codbcra='" & vNuevoValor & "'"

        case "ORLINNEGOCIOS"
           strsqlbcra="select * from " & tabla & " where nivel1=-1"

	    case "ORTIPOCATEGO"
	      XAGRUPA=RS("tipoevento")
	      strsqlbcra="select * from " & tabla & " where tipoevento='"& XAGRUPA& "' and codbcra='" & vNuevoValor & "'"

	   case else
	      strsqlbcra="select * from " & tabla & " where codbcra='" & vNuevoValor & "'"
	    end select

	    'if (XAGRUPA= 3 OR XAGRUPA= 1) and (ucase(tabla)="ORIMPRIESGO" and ucase(tabla)="ORLINNEGOCIOS") then

	   'else
	    'if ucase(tabla)="ORIMPRIESGO"   then
	    '  XAGRUPA=RS("AGRUPA")
	    '  strsqlbcra="select * from " & tabla & " where AGRUPA='"&XAGRUPA&"' AND codbcra='" & vNuevoValor & "'"

	   ' ELSE
	     'if ucase(tabla)="ORLINNEGO2" then
	     ' XAGRUPA=RS("nivel1")
	     ' strsqlbcra="select * from " & tabla & " where nivel1='"& XAGRUPA& "' and codbcra='" & vNuevoValor & "'"
	     'else
	   '  if ucase(tabla)="ORTIPOCATEGO" then
	   '   XAGRUPA=RS("tipoevento")
	   '   strsqlbcra="select * from " & tabla & " where tipoevento='"& XAGRUPA& "' and codbcra='" & vNuevoValor & "'"
	    ' else
        ' strsqlbcra="select * from " & tabla & " where codbcra='" & vNuevoValor & "'"
       '  end if
       '  end if
       ' END IF


   	    set rsbcra = ofv.crearconsultaEx(strsqlbcra,cn,1,volver)

	    if not rsbcra.eof  then SvalidaBcra="SI"

        call ofv.cerrarconsulta(rsbcra)
     ' end if
      end if

  CASE "CODINTER"


      Str_a=cstr(vNuevoValor&"")
	  Str_b=cstr(rs(x)&"")
	  if Str_a <> Str_b then
	      strsqlbcra="select * from " & tabla & " where CODINTER='" & vNuevoValor & "'"
    	  set rsbcra = ofv.crearconsultaEx(strsqlbcra,cn,1,volver)

   	      if not rsbcra.eof then SvalidaBcra="SI"

          call ofv.cerrarconsulta(rsbcra)
      END IF
  END SELECT

' END IF

  if (nx = "FECEMIS" or nx = "FECBAJA" or nx = "FECVIGEN" or nx = "FECVTO" or nx = "FECVIG") then



     if vNuevoValor = "" then
        vNuevoValor = clng(cdate("01-01-1900")) - nFecini
    else
       ' vNuevoValor = clng(cdate(vNuevoValor)) - nFecini
         if  nx = "FECVIG" then
	           vNuevoValor = clng(cdate(vNuevoValor))
	           else
	           vNuevoValor = clng(cdate(vNuevoValor)) - nFecini
        end if

     end if
  end if
  if not rs.eof then

      Str_a=cstr(vNuevoValor&"")
      Str_b=cstr(rs(x)&"")
     if Str_a <> Str_b then
        'response.write vNuevoValor &"="& rs(x) & "<br>"
        XVALANT	=Str_b
        XVALACT	=Str_a
        xxxr=GENLOGEXTDATA("UPDATE",TABLA,x,id,"PAR")
     end if
  end if

 '' s = ofv.QueEsx(rs(x).type)
  s = "'"
  if rs(x).type = 200 then s = "'"

 ' IF  X = "TRCPROCESO" AND UCASE(TABLA)="ORTRACER" THEN
 '         RESPONSE.WRITE "1"
 '        Str_a=cstr(vNuevoValor&"")
 '        Str_b=cstr(rs(x)&"")
 '     if Str_a <> Str_b then
  '       RESPONSE.WRITE "2"

  IF  (UCASE(X) = "TRCPROCESO" OR UCASE(X)="TRCEVENTORIESGO1" )AND UCASE(TABLA)="ORTRACER" THEN
   Str_a=cstr(vNuevoValor&"")
   Str_b=cstr(rs(x)&"")

   XENTRO =Str_a <> Str_B
  ' RESPONSE. WRITE XENTRO

  END IF

         IF (UCASE(X)="TRCPRODSER" OR UCASE(X)="TRCIDENTACT" OR UCASE(X)="TRCEVENTORIESGO2" ) AND  XENTRO THEN

         vNuevoValor=0
         END IF

    '   END IF
  'END IF

  if not XENTRO then
   IF  (UCASE(X) = "TRCPRODSER" OR UCASE(X)="TRCEVENTORIESGO2") AND UCASE(TABLA)="ORTRACER" THEN
     Str_a=cstr(vNuevoValor&"")
     Str_b=cstr(rs(x)&"")

     XENTRO2 =Str_a <> Str_B
    ' RESPONSE. WRITE XENTRO

  END IF
   IF (UCASE(X)="TRCIDENTACT") AND  XENTRO2 THEN

           vNuevoValor=0
   END IF
 end if

  strsql = strsql & x  & " = " & s & vNuevoValor & s & ","
Next




s = ofv.QueEsx(rs(0).type)
  if rs(0).type = 200 then s = "'"
strsql = mid(strsql,1,len(strsql)-1) & " Where " & rs(0).name & "="
strsql = strsql & s & id & s

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''CONTROL PARA QUE EL DESTINO A MODIFICAR NO ESTE SIENDO USADO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

IF UCASE(TABLA) = "DESTINOSURL" THEN

	estado = request.form("estado")

	IF CSTR(estado) = "4" THEN
		controlDestinos = request.querystring("usado")
	ELSE
		controlDestinos = "NO"
	END IF

ELSE

	controlDestinos = "NO"

END IF

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

 if SvalidaBcra="NO" AND controlDestinos = "NO" THEN
   call ofv.crearconsultaEx(StrSql,cn,1,volver)
 END IF

if cn.errors.count = 0 then
'response.write xvalant & "<br>" & xvalact & "<br>"


  if instr(volver,"?") =0 then volver=volver & "?"
  if volver <> "" then response.redirect  volver & "&Svalida=" & SvalidaBcra & "&control=" & controlDestinos
else
  response.write ofv.ErrorDB(cn,volver)
end if

call ofv.cerrarconsulta(rs)
call ofv.cerrarconn(cn)

%>