
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

 '   sHTML = sHTML & "<div class='chart-container'>"
  '  sHTML = sHTML & "      <canvas id=""miGrafico""></canvas>"
  '  sHTML = sHTML & "</div>"
	sHTML = sHTML  & "<iframe WIDTH=100% HEIGHT=400PX  style='display:none' id=enviaf name=enviaf></iframe></body></html>"

end if

function armafiltros(t)
dim str
str= "  <BR><div class='form-container norma-card2'>"
str= STR & "         <h2>Configuración del Gráfico</h2>"
str= STR & "         <p>Selecciona las opciones a continuación y presiona 'Graficar' para generar el reporte.</p>"
str= STR & "  <form id='miFormulario' name='miFormulario'	 method='POST' target='enviaf' action='verdatos.asp'>"
str= STR & " <div class='form-group'>"
str= STR & "                 <label for='miSelect'><strong>1. Selecciona qué graficar</strong></label>"
str= STR & "                 <select id='miSelect' name='miSelect'>"
str= STR & "                     <option value='DOCMASLEIDOS'>Tipos de documentos más leídos</option>"
str= STR & "                     <option value='DOCPORESTADOS'>Cantidad de documentos por estado</option>"
str= STR & "                     <option value='DOCPORAUTOR'>Cantidad de documentos por autor</option>"
str= STR & "                     <option value='DOCPORSUBTEMA'>Cantidad de documentos por subtema</option>"
str= STR & "                     <option value='WKF'>WKF POR ESTADO</option>"
str= STR & "                     <option value='WKFPASO30'>WKF EN PASO APROBADO Y RECHAZADO</option>"
str= STR & "                 </select>"
str= STR & "             </div>"

str= STR & "             <div class='form-group'>"
str= STR & "                 <label for='tipoGrafico'><strong>2. Selecciona el tipo de gráfico</strong></label>"
str= STR & "                 <select id='TIPOGRAF' name='TIPOGRAF'>"
str= STR & "                     <option value='pie'>Torta (Pie)</option>"
str= STR & "                     <option value='bar' selected>Barras (Bar)</option>"
str= STR & "                     <option value='line'>Línea (Line)</option>"
str= STR & "                     <option value='doughnut'>Rosquilla (Donut)</option>" '''Doughnut
str= STR & "                 </select>"
str= STR & "             </div>"

str= STR & "             <div class='date-row'>"
str= STR & "                 <div class='form-group'>"
str= STR & "                     <label for='fechaDesde'><strong>3. Fecha Desde:</strong></label>"
str= STR & "                     <input type='text' id='fechaDesde' name='fechaDesde'>"
str= STR & "                 </div>"
str= STR & "                 <div class='form-group'>"
str= STR & "                     <label for='fechaHasta'><strong>4. Fecha Hasta:</strong></label>"
str= STR & "                     <input type='text' id='fechaHasta' name='fechaHasta'>"
str= STR & "                 </div>"
str= STR & "            </div>"

 str=str & "<button type='submit' id=graficar>GRAFICAR</button></FORM>"

'str= STR & "  <hr>"

str= STR & "         <div class='chart-container'>"
str= STR & "              <canvas id='miGrafico'></canvas>"
str= STR & "         </div>"
str= STR & "    </div>"
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



