
<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%


ok = ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	shtml = ok
Else
	response.expires = 0

	If (SESSION("UN") & "") = "" Then


		xxxx = NEWEMPRE("NO")


		SEMPRE = ATREMP(0, 0)
		SESSION("UN") = SEMPRE & ""

		If (SESSION("UN") & "") = "" Then SESSION("UN") = 1


	End If
	Set cn = ofv.conectar(ofv.strconn2)
	Set cn0 = ofv.conectar(ofv.strconn0)
	Set cn5 = ofv.conectar(ofv.strconn5)

	'response.write session("sql")

	shtml = ofv.FormHeader(SESSION("FormN"))

   sHTML = sHTML  & armafiltros("")

    sHTML = sHTML & "<div class='chart-container'>"
    sHTML = sHTML & "      <canvas id=""miGrafico""></canvas>"
    sHTML = sHTML & "</div>"
	sHTML = sHTML  & "<iframe WIDTH=100% HEIGHT=400PX  style='display:none' id=enviaf name=enviaf></iframe></body></html>"

end if

function armafiltros(t)
dim str
str= " <form id='miFormulario' name='miFormulario'	 method='POST' target='enviaf' action='verdatos.asp'>"
 str=str & "<div class='form-group'>"

 str=str & "<br><label for='miSelect'>Selecciona que graficar </label><BR><hr>"
 str=str & "           <select id='miSelect' name='miSelect'>"

'  str=str & "<option value=''>-- Elige una opción --</option>"
  str=str & "     <option value='DOCMASLEIDOS'>Tipos de documentos mas leidos</option>"
  str=str & "     <option value='DOCPORESTADOS'>cantidad de documentos por estados</option>"
  str=str & "     <option value='DOCPORAUTOR'>cantidad de documentos por autor</option>"
  str=str & "     <option value='DOCPORSUBTEMA'>cantidad de documentos por SUBTEMAS</option>"

  str=str & "</select>"
  str=str & "       </div><br>"
  str=str & "<div class='form-group'>"
  str=str & "<br><label for='miSelect'>Selecciona tipo de grafico </label><BR><hr>"
   str=str & "           <select id='TIPOGRAF' name='TIPOGRAF'>"

  '  str=str & "<option value=''>-- Elige una opción --</option>"
    str=str & "     <option value='pie'>TORTA</option>"
    str=str & "     <option value='bar'>BARRAS</option>"
    str=str & "     <option value='line'>LINEA</option>"
    str=str & "     <option value='doughnut'>ROSQUILLA</option>"


    str=str & "</select>"
  str=str & "       </div><br>"



str=str & "       <div class='form-group'>"
str=str & "           <label for='fechaDesde'>Fecha Desde:</label>"
str=str & "           <input type='text' id='fechaDesde' name='fechaDesde' placeholder='Selecciona una fecha'>"
str=str & "       </div>"

str=str & "       <div class='form-group'>"
str=str & "           <label for='fechaHasta'>Fecha Hasta:</label>"
str=str & "          <input type='text' id='fechaHasta' name='fechaHasta' placeholder='Selecciona una fecha'>"
str=str 	& "       </div>"
 str=str & "  <br><button type='submit' id=graficar>GRAFICAR</button></FORM>"
armafiltros=str
end function

Function fechaArg(f)
	dia = Day(f)
	mes = Month(f)
	ano = Year(f)
	fechaArg = dia & "/" & mes & "/" & ano
End Function
response.write shtml
%>



