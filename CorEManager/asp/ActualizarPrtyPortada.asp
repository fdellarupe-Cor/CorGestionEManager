<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

response.expires=0

nCodigo = request.querystring("Codobjeto")
sTipo = request.querystring("Tipobjeto")
sEst = request.querystring("Tipo")
detcon = request.querystring("detcon")
ID = request.querystring("ID")
ID2 = request.querystring("ID2")
volver = "../edportada.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo

set cn = ofv.conectar(ofv.strconn2)


    sDettipo = "G"


dim aFName(18)
dim aFValue(18)

aFname(0) = "detfont"
aFname(1) = "detfsize"
aFname(2) = "detfcolor"
aFname(3) = "detbgcolor"
aFname(4) = "detborde"
aFname(5) = "detftipo"
aFname(6) = "detsang"
aFname(7) = "detindent1"
aFname(8) = "detindent2"
aFname(9) = "detvin"
aFname(10) = "detvinfont"
aFname(11) = "detvintipo"
aFname(12) = "detvinindent"
aFname(13) = "detnumer"
aFname(14) = "detnumertipo"
aFname(15) = "detnumerclase"
aFname(16) = "detalign"
aFname(17) = "dettextalign"
aFname(18) = "detdefault"
'response.write request.form



for each yy in request.form
   x = lcase(yy)
  if left(x,len(x)-1) = "detftipo" then
     
    sValue = request.form(x)
    aFValue(5) = aFValue(5) + sValue
  else
    if request.form(x) = "nsnc" then
      sValue = "''"
    else
      sValue = "'" & replace(Request.Form(x),"_"," ") & "'"
    end if
    for i = 0 to 18
      if aFName(i) = x then
        aFValue(i) = sValue
        exit for
      end if
    next
  end if
next

'''if aFValue(5) <> "" then aFValue(5) = "'" & aFValue(5) & "'"

'''if aFValue(5) <> "" then aFValue(5) = "'" & aFValue(5) & "'"
 aFValue(5) = "'" & aFValue(5) & "'"

scoma = ""
strsql = "Update Objportada Set "
for i = 0 to 18
  if aFValue(i) <> "" then
     strsql = strsql  & scoma & aFname(i) & "= " & aFvalue(i)
     scoma = ","
  end if
next

strsql = strsql & " Where Objid = " & ID

''' response.write strsql
call ofv.crearconsultaEx(StrSql,cn,1,volver)
if cn.errors.count = 0 then
'response.write volver
  response.redirect volver
else
  response.write ofv.ErrorDB(cn,volver)
end if 

call ofv.cerrarconn(cn)

%>