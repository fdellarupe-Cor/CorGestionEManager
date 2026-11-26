	<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else




    sOCesque = Request.querystring("Codobjeto")
    nCodigo = Request.querystring("Codproced")
    sTipo = Request.querystring("Tipobjeto")
    vuelta = Request.querystring("Volver")
    idsector=0
    Set cn = OFV.conectar(OFV.strconn2)


    StrSql = "Select codobjeto,tipobjeto From OBJPROCED Where codproced = '" & nCodigo & "'"
    Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
    If Not rs.EOF Then
       sOCesque = rs(0)
       sOTesque = rs(1)
    End If
    Call OFV.cerrarconsulta(rs)

  StrSqlx = "Select Version From Objprty Where codobjeto = " & sOCesque
 '' response.write strsqlx
  Set rs = OFV.crearconsultaEx(StrSqlx, cn, 1, "")
  bVersion = "NO"
  If Not rs.EOF Then
    If rs(0) = "V" Then bVersion = "SI"
  End If
  Call OFV.cerrarconsulta(rs)

    StrSql = "Select descri From tipobjeto Where tipobjeto = '" & sOTesque & "'"
    Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
    If Not rs.EOF Then sTesque = rs(0)
    Call OFV.cerrarconsulta(rs)

    StrSql = "Select codiden,descri From objetos Where codobjeto = " & sOCesque
    Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
    If Not rs.EOF Then
       sCesque = rs(0)
       sDesque = rs(1)
    End If
    Call OFV.cerrarconsulta(rs)

    StrSql = "Select Archivocur From Modo Where tipomodo = 'I'"
    Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
    If Not rs.EOF Then sInicio = rs(0)
    Call OFV.cerrarconsulta(rs)

    StrSql = "Select Archivocur From Modo Where tipomodo = 'F'"
    Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
    If Not rs.EOF Then sFin = rs(0)
    Call OFV.cerrarconsulta(rs)


    '' response.write "aa"

      sCaption = sTesque & "&nbsp;" & sCesque & " - " & sDesque



  WSINIC = 0

   WSFINC = 1

   ReDim asecto(0)

   asecto(0) = 0
   isec = 0
   isecb = 0

      Strsql = "DELETE OBJPROCURSODET Where codobjeto = " & sOCesque
      call ofv.crearconsultaEx(StrSql,cn,1,volver)


'''' loop sectores

    StrSql = "SELECT * from cursmaster "
    StrSql = StrSql & " Where codobjeto = " & sOCesque
    StrSql = StrSql & " order by secuen, secuencia "
    Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
    Do Until rs.EOF
       If isec = 0 Then
          isec = isec + 1
          asecto(0) = asecto(0) + 1
          ReDim Preserve asecto(isec)
          asecto(isec) = rs(1)
       Else
          For isecb = 1 To isec
              If asecto(isecb) = rs(1) Then
                 isecb = 9999
                 Exit For
              End If
          Next
          If isecb <> 9999 Then
             isec = isec + 1
             asecto(0) = asecto(0) + 1
             ReDim Preserve asecto(isec)
             asecto(isec) = rs(1)
          End If
       End If
    rs.movenext
    Loop
    Call OFV.cerrarconsulta(rs)



   ReDim alnkv(0)
   ReDim alnkl(0)
   ReDim alnkvD(0)
   ReDim alnklD(0)


   ilnkv = 0
   ilnkvb = 0
   ilnkl = 0
   ilnklb = 0

    StrSql = "SELECT * from cursmasterlnk "
    StrSql = StrSql & " Where codobjeto = " & sOCesque & " and ltipo = 'R' "
    StrSql = StrSql & " order by sece, secp, seca"
  '  response.write strsql
    Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
    Do Until rs.EOF
          ilnkv = ilnkv + 1
          ReDim Preserve alnkv(ilnkv)
          ReDim Preserve alnkvD(ilnkv)
          alnkv(ilnkv) = rs(1)
          daccD = "&nbsp;"
          daccP = " "
          daccE = "&nbsp;"

          StrSql = "Select codproced,secuen,descri From Cursograma Where codaccion = '" & rs(2) & "'"
          Set rs2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
          If Not rs2.EOF Then
             daccD = rs2(1) & "." & rs2(2)
             daccP = rs2(0)
          End If
          Call OFV.cerrarconsulta(rs2)


          StrSql = "SELECT  nombre "
          StrSql = StrSql & " from CSCEventoSectorPerfil Where codproced = '" & daccP & "' "
          Set rs2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
          If Not rs2.EOF Then
             daccE = rs2(0)
             daccE = Replace(daccE, " ", "&nbsp;")
          End If
          Call OFV.cerrarconsulta(rs2)

          alnkvD(ilnkv) = "&nbsp;A:&nbsp;" & daccE & Chr(13) & daccD

          ilnkl = ilnkl + 1
          ReDim Preserve alnkl(ilnkl)
          ReDim Preserve alnklD(ilnkl)
          alnkl(ilnkl) = rs(2)

          daccD = "&nbsp;"
          daccP = " "
          daccE = "&nbsp;"

          StrSql = "Select codproced,secuen,descri From Cursograma Where codaccion = '" & rs(1) & "'"
          Set rs2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
          If Not rs2.EOF Then
             daccD = rs2(1) & "." & rs2(2)
             daccP = rs2(0)
          End If
          Call OFV.cerrarconsulta(rs2)


          StrSql = "SELECT  nombre "
          StrSql = StrSql & " from CSCEventoSectorPerfil Where codproced = '" & daccP & "' "
          Set rs2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
          If Not rs2.EOF Then
             daccE = rs2(0)
             daccE = Replace(daccE, " ", "&nbsp;")
          End If
          Call OFV.cerrarconsulta(rs2)

          alnklD(ilnkl) = "&nbsp;DE:&nbsp;" & daccE & Chr(13) & daccD

    rs.movenext
    Loop
    Call OFV.cerrarconsulta(rs)





      Fila     = ""
      Tabla    = ""


' ''''''''''''''

  asectn = "no"
  aeven  = "NO"

    Set cnn = OFV.conectar(OFV.strconn2)
     Set cnx = OFV.conectar(OFV.strconn2)





        StrSql = "SELECT CH.*, CM.descri as evento, CM.ecodescri, CM.tipodescri, CL.iden3L,CL.iden4,CM.ECO,tt.descriacc from cursmaster CM"
        StrSql = StrSql & " inner join Cursogramanew CH on (CM.codproced = CH.codproced and CM.codobjeto = CH.codobjeto)"
        StrSql = StrSql & "  left join cursmasterlnk CL on (CH.codaccion = CL.codaccion and CH.codobjeto = CL.codobjeto and CL.ltipo = 'R')"
        StrSql = StrSql & " left join objaccion tt on(tt.codobjeto=0" & sOCesque & " and tt.codproced=ch.codproced and tt.codaccion=ch.codaccion and tt.tipoaccion=301)"
        StrSql = StrSql & " where CM.codobjeto = " & sOCesque & "  and  (CH.modi = 0 or ((CH.modi = -1 or CH.modi = 1) and CH.tipmodi <> 'E')) " & " and CM.codproced <> 'x243412548045730440553801' "
        StrSql = StrSql & " order by CM.secuen, CM.secuencia, CM.codproced, CH.secuen, CH.codaccion"
        Set rs = OFV.crearconsultaEx(StrSql, cnn, 1, "")
   '  response.write strsql & "<br>"

      'Encabezado


      Tabla = ""
    If Not rs.EOF Then



      nCont = 0
      If Not rs.EOF Then
        StrSql = "Select Continuador From AccionesAlt "
        StrSql = StrSql & "Where codproced = '" & rs(2) & "' and Continuador <> '0'"
      '  response.write strsql
        Set rsAlt = OFV.crearconsultaEx(StrSql, cnx, 1, "")
        z = 1
        ReDim Continuadores(z)
        Do Until rsAlt.EOF
          ReDim Preserve Continuadores(z)
          Continuadores(z) = rsAlt(0)
          z = z + 1
          rsAlt.movenext
        Loop
        Call OFV.cerrarconsulta(rsAlt)

      End If

      Do Until rs.EOF


    If Not IsNull(rs("evento"))     Then sEvento = rs("evento")
    If Not IsNull(rs("ecodescri"))  Then sSector = rs("ecodescri")
    If Not IsNull(rs("tipodescri")) Then sPerfil = rs("tipodescri")
    If Not IsNull(rs(2))            Then scodpro = rs(2)
    If Not IsNull(rs("ECO"))            Then idsector = rs("ECO")

 '   Call OFV.cerrarconsulta(rs)


    If aeven <> sEvento Then
       aeven = sEvento

    End If

    If asectn <> sSector Then
       asectn = sSector




          isecc = 1

          For isecb = 1 To isec
              If asecto(isecb) = rs("ecodescri") Then
                 isecc = isecb
                 isecb = 9999
                 Exit For
              End If
          Next





    End If






  IF    WSINIC <> 1 THEN

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	TWactividad   = "INICIO"
	TWsector      = sSector
	TWtipo        = "INICIO"
	TWsi          = 0
	TWno          = 0
	TWidsec       = isecc
	TWposx        = 0
	TWposy        = 0

	TWcodobjeto   = sOCesque
	TWtipobjeto   = sOTesque
	TWcodproced   = rs(2)
	TWcodaccion   = "1"
	TWcodaccionSI = rs(12)
    Tmodo="I"
	TWIDCOD       = 0


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

      Strsql = "insert into OBJPROCURSODET  (actividad,sector,tipo,si,NO,idsec,posx,posy,codobjeto,tipobjeto,codproced,codaccion,codaccionSI,codaccionNO,IDCOD,modo,idsector) values ("
      Strsql = Strsql & "'" & TWactividad & "','" & TWsector & "','" & TWtipo & "','" & TWsi & "','" & TWno & "','" & TWidsec & "','" & TWposx & "','"
      Strsql = Strsql & TWposy & "','" & TWcodobjeto & "','" & TWtipobjeto & "','" & TWcodproced  & "','" & TWcodaccion & "','" & TWcodaccionSI & "','" & TWcodaccionNO & "','" & TWIDCOD &  "','" & Tmodo & "','"& idsector &")"

      call ofv.crearconsultaEx(StrSql,cn,1,volver)

      WSINIC = 1

 END IF




     slnkv = ""
     slnkl = ""
     ilnkvt = 0
     ilnklt = 0

     WsCACC = rs(12)
     WsSecu = rs(3)




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    if rs("tipoaccionalt")=301 then
     if rs("descriacc") & "" <> "" then
      TWactividad   = rs("descriacc")
     else
      TWactividad   = RS(4)
     end if
     'TWactividad   = rs("descriacc")
    else
	TWactividad   = RS(4)
	end if
	TWsector      = sSector
	TWtipo        = "PROCESO"
	TWsi          = 0
	TWno          = 0
	TWidsec       = isecc
	TWposx        = 0
	TWposy        = 0

	TWcodobjeto   = sOCesque
	TWtipobjeto   = sOTesque
	TWcodproced   = rs(2)
	TWcodaccion   = rs(12)
	TWcodaccionSI = ""
	TWcodaccionNO = ""
	TWIDCOD       = 0
    Tmodo=rs("modo")



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''








        'Icono de la Accion


        'Secuencia
        sinputs = rs(3) & "&nbsp;" & slnkv & slnkl


       'Referencia
       sVerbo = rs(4)
        sAcciont = rs(4)
        i = InStr(sVerbo, " ")
        If  i > 0 Then
            sVerbo = Left(rs(4), i - 1)
            sAcciont = Mid(rs(4), i + 1)
        End If




        If  rs(16) > "0" Then
        ' Sector Continuador
            StrSql = "select * from eco where eco.eco = " & rs(16)
            Set rsc = OFV.crearconsultaEx(StrSql, cnn, 1, "")
            If Not rsc.EOF Then sconti = rsc(1)
        End If
        If  rs(17) > "0" Then
        ' Perfil Continuador
        StrSql = "select * from tipoperfil where tipoperfil = '" & rs(17) & "'"
            Set rsd = OFV.crearconsultaEx(StrSql, cnn, 1, "")
            If Not rsd.EOF Then sperfi = rsd(1)
        End If
        If  rs(16) > "0" Or rs(17) > "0" Then
            sArma = "Continuador:"
            sArma = sArma & sconti & " - " & sperfi
        Else
            sArma = "&nbsp;"
        End If

        'Call OFV.cerrarconn(cnn)



     wwwmark = rs("iden3L") & ""
     wwwmarkSN = rs("iden4") & ""

   ' if RS(12)="304313222140446356146693" then wwwmarkSN="N"


      if RS("control") then
	 TWtipo        = "DECISIÓN"
      else
	TWtipo        = "PROCESO"
      end if

		wTWtipo = TWtipo
        nSec          = ""




        rs.movenext


      TWcodaccioncn=""

'     WsCACC = rs(12)
 '    WsSecu = rs(3)

       IF wwwmark <> "" or (NOT RS.EOF) THEN
           WSNEXT = "FIN"


          if NOT RS.EOF THEN WSNEXT = RS(12)

    ''' Aqui conectores - conector '''
          if wwwmark <> "" then
            ' WSNEXT = wwwmark
             TWcodaccioncn=wwwmark
             if wwwmarkSN="N" then
               WSNEXT="NO"
             end if

             TWtipo = TWtipo  & "R"
          end if


       ELSE

          WSNEXT = "FIN"


       END IF

'''''''''''''''''
         wcontinuator = ""

        'Set cnn = OFV.conectar(OFV.strconn2)
        StrSql = "Select * From AccionesAlt Where codaccion = '" & WsCACC & "'"
        StrSql = StrSql & " and (modi = 0 or ((modi = -1 or modi = 1) and tipmodi <> 'E')) "
           StrSql = StrSql & " Order By secuencia "
'response.write strsql
        Set rsAlt = OFV.crearconsultaEx(StrSql, cn, 1, "")

           nAlt = ""




      IF NOT RSALT.EOF and wTWtipo = "DECISIÓN" THEN

'''	TWtipo        = "DECISIÓN"
	TWcodaccionNO = "ALT" & rsalt(10)
        nSec          = ""

      ELSE
''''	TWtipo        = "PROCESO"

	if TWtipo        = "DECISIÓN" then
           nSec          = "FIN"
	   TWcodaccionNO = "ALTFIN" & WSFINC
           WSFINC        =  WSFINC + 1
        else
           nSec          = ""
	   TWcodaccionNO = ""
        end if

      END IF

      IF WSNEXT = "FIN" THEN
	 TWcodaccionSI = WSNEXT & WSFINC
         WSFINC =    WSFINC + 1

      ELSE
	 TWcodaccionSI = WSNEXT
      END IF



      Strsql = "insert into OBJPROCURSODET  (actividad,sector,tipo,si,NO,idsec,posx,posy,codobjeto,tipobjeto,codproced,codaccion,codaccionSI,codaccionNO,IDCOD,modo,codaccioncn,idsector) values ("
      Strsql = Strsql & "'" & TWactividad & "','" & TWsector & "','" & TWtipo & "','" & TWsi & "','" & TWno & "','" & TWidsec & "','" & TWposx & "','"
      Strsql = Strsql & TWposy & "','" & TWcodobjeto & "','" & TWtipobjeto & "','" & TWcodproced  & "','" & TWcodaccion & "','" & TWcodaccionSI & "','" & TWcodaccionNO & "','" & TWIDCOD &  "','" & Tmodo & "','" & TWcodaccioncn & "',"& idsector &")"
      call ofv.crearconsultaEx(StrSql,cn,1,volver)

    '  response.write strsql & "<br"

     IF WSNEXT = "FIN" THEN
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	TWactividad   = "FIN"
	TWsector      = sSector
	TWtipo        = "FIN"
	TWsi          = 0
	TWno          = 0
	TWidsec       = isecc
	TWposx        = 0
	TWposy        = 0

	TWcodobjeto   = sOCesque
	TWtipobjeto   = sOTesque
	TWcodproced   = scodpro
	TWcodaccion   = TWcodaccionSI
        TWcodaccionSI = ""
        TWcodaccionNO = ""
	TWIDCOD       = 0
    tmodo="F"

      Strsql = "insert into OBJPROCURSODET  (actividad,sector,tipo,si,NO,idsec,posx,posy,codobjeto,tipobjeto,codproced,codaccion,codaccionSI,codaccionNO,IDCOD,modo,idsector) values ("
      Strsql = Strsql & "'" & TWactividad & "','" & TWsector & "','" & TWtipo & "','" & TWsi & "','" & TWno & "','" & TWidsec & "','" & TWposx & "','"
      Strsql = Strsql & TWposy & "','" & TWcodobjeto & "','" & TWtipobjeto & "','" & TWcodproced  & "','" & TWcodaccion & "','" & TWcodaccionSI & "','" & TWcodaccionNO & "','" & TWIDCOD &  "','" & Tmodo & "'," & idsector & ")"
      call ofv.crearconsultaEx(StrSql,cn,1,volver)



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   END IF


     IF nSec = "FIN" THEN
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	TWactividad   = "FIN"
	TWsector      = sSector
	TWtipo        = "FIN"
	TWsi          = 0
	TWno          = 0
	TWidsec       = isecc
	TWposx        = 0
	TWposy        = 0

	TWcodobjeto   = sOCesque
	TWtipobjeto   = sOTesque
	TWcodproced   = scodpro
	TWcodaccion   = TWcodaccionNO
        TWcodaccionSI = ""
        TWcodaccionNO = ""
	TWIDCOD       = 0
    tmodo="F"

      Strsql = "insert into OBJPROCURSODET  (actividad,sector,tipo,si,NO,idsec,posx,posy,codobjeto,tipobjeto,codproced,codaccion,codaccionSI,codaccionNO,IDCOD,modo,idsector) values ("
      Strsql = Strsql & "'" & TWactividad & "','" & TWsector & "','" & TWtipo & "','" & TWsi & "','" & TWno & "','" & TWidsec & "','" & TWposx & "','"
      Strsql = Strsql & TWposy & "','" & TWcodobjeto & "','" & TWtipobjeto & "','" & TWcodproced  & "','" & TWcodaccion & "','" & TWcodaccionSI & "','" & TWcodaccionNO & "','" & TWIDCOD &  "','" & Tmodo & "'" & idsector &")"
      call ofv.crearconsultaEx(StrSql,cn,1,volver)

        nSec          = ""

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     END IF


        wcontinuator = ""

if wTWtipo = "DECISIÓN" then
        Do Until rsAlt.EOF

          'Icono de la Accion Alt
        '  sImagen = OFV.GenImage("", "", "", "", "Iconos/" & rsAlt(7), "0", "")




          If nAlt = "" Or nAlt = "0" Then
                If Not IsNull(rsAlt(5)) and rsAlt(5) & "" <> "nsnc" Then
                   nAlt = rsAlt(5)
                   wcontinuator = rsAlt(5)
                end if
          End If



          naccion = rsAlt(10)



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	TWactividad   = rsAlt(0)
	TWsector      = sSector
	TWtipo        = "PROCESO"
	TWsi          = 0
	TWno          = 0
	TWidsec       = isecc
	TWposx        = 0
	TWposy        = 0

	TWcodobjeto   = sOCesque
	TWtipobjeto   = sOTesque
	TWcodproced   = scodpro
	TWcodaccion   = "ALT" & rsalt(10)
    'tmodo=rs("modo")
    tmodo=rsAlt("modo")
	TWIDCOD       = 0

  '  response.write tmodo & "<br>"


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


          rsAlt.movenext

          'Lineas
          If Not rsAlt.EOF Then

            If nAlt = "" Or nAlt = "0"  Then
                  If Not IsNull(rsAlt(5)) and rsAlt(5) & "" <> "nsnc" Then nAlt = rsAlt(5)
            End If
            sArma = "&nbsp;"
            nSec = "NO"
	    TWcodaccionSI = "ALT" & rsalt(10)
            TWcodaccionNO = ""
            If Not IsNull(rsAlt(5)) and rsAlt(5) & "" <> "nsnc"  Then wcontinuator = rsAlt(5)
          Else

''            StrSql = "Select secuen From Cursograma Where codaccion = '" & nAlt & "'"
''            StrSql = StrSql & " And codproced = '" & nCodigo & "'"

''            Set rsSec = OFV.crearconsultaEx(StrSql, cn, 1, "")
            nSec = "FIN"
''            If Not rsSec.EOF Then
''              If Not IsNull(rsSec(0)) Then nSec = rsSec(0)
''            End If
''            Call OFV.cerrarconsulta(rsSec)

            'Conector
''            sArma = nSec

       ''''     IF nSec = "FIN" THEN
           if wcontinuator = "" then
	       TWcodaccionSI = "ALT" & nSec & WSFINC
               WSFINC =    WSFINC + 1
            ELSE
               nSec = wcontinuator
	     '''  TWcodaccionSI = NALT
	       TWcodaccionSI = wcontinuator
            END IF
            TWcodaccionNO = ""

          End If


      Strsql = "insert into OBJPROCURSODET  (actividad,sector,tipo,si,NO,idsec,posx,posy,codobjeto,tipobjeto,codproced,codaccion,codaccionSI,codaccionNO,IDCOD,modo,idsector) values ("
      Strsql = Strsql & "'" & TWactividad & "','" & TWsector & "','" & TWtipo & "','" & TWsi & "','" & TWno & "','" & TWidsec & "','" & TWposx & "','"
      Strsql = Strsql & TWposy & "','" & TWcodobjeto & "','" & TWtipobjeto & "','" & TWcodproced  & "','" & TWcodaccion & "','" & TWcodaccionSI & "','" & TWcodaccionNO & "','" & TWIDCOD &  "','" & Tmodo & "'," & idsector & ")"
      call ofv.crearconsultaEx(StrSql,cn,1,volver)
    '  response.write strsql & "<br>"


     IF nSec = "FIN" THEN
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	TWactividad   = "FIN"
	TWsector      = sSector
	TWtipo        = "FIN"
	TWsi          = 0
	TWno          = 0
	TWidsec       = isecc
	TWposx        = 0
	TWposy        = 0

	TWcodobjeto   = sOCesque
	TWtipobjeto   = sOTesque
	TWcodproced   = scodpro
	TWcodaccion   = TWcodaccionSI
        TWcodaccionSI = ""
        TWcodaccionNO = ""
	TWIDCOD       = 0


      Strsql = "insert into OBJPROCURSODET  (actividad,sector,tipo,si,NO,idsec,posx,posy,codobjeto,tipobjeto,codproced,codaccion,codaccionSI,codaccionNO,IDCOD,idsector) values ("
      Strsql = Strsql & "'" & TWactividad & "','" & TWsector & "','" & TWtipo & "','" & TWsi & "','" & TWno & "','" & TWidsec & "','" & TWposx & "','"
      Strsql = Strsql & TWposy & "','" & TWcodobjeto & "','" & TWtipobjeto & "','" & TWcodproced  & "','" & TWcodaccion & "','" & TWcodaccionSI & "','" & TWcodaccionNO & "','" & TWIDCOD & "'," & idector & ")"
      call ofv.crearconsultaEx(StrSql,cn,1,volver)



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     END IF






        Loop

        end if


        Call OFV.cerrarconsulta(rsAlt)




        Loop



    end if


        Call OFV.cerrarconn(cnn)
        Call OFV.cerrarconn(cnx)


      Strsql = "DELETE FROM OBJPROCURSOCMP where codobjeto = " & sOCesque
      call ofv.crearconsultaEx(StrSql,cn,1,volver)


  response.redirect "CursogramaAdmin.asp?Codproced=" & nID & "&Codobjeto=" & sOCesque & "&Tipobjeto=" & sTipo




       end if

''response.write sHTML

%>