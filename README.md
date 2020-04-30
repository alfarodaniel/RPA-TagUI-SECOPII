<h2>
Automatización de SECOP II por medio de RPA (Robotic Process Automation) con TagUI</h2>
A partir de una herramienta <b>Open Source</b> (Código Abierto) de&nbsp;<b>RPA</b>&nbsp;(Automatización Robótica de Procesos) denominada&nbsp;<b>TagUI</b>&nbsp;(<a href="https://github.com/kelaberetiv/TagUI" target="_blank">https://github.com/kelaberetiv/TagUI</a>) se desarrollaron robots para automatizar los diferentes procesos de cargue manual de información en la plataforma <b>SECOP II</b> (<a href="https://www.colombiacompra.gov.co/secop-ii">https://www.colombiacompra.gov.co/secop-ii</a>) de <b>Colombia Compra Eficiente</b>.<br />
<h3>
Logros</h3>
<div>
<ul>
<li>Cumplimiento de las Directrices Legales que en materia de contratación rigen y obligan a la empresa en el uso de la plataforma del SECOP II a partir del 1 de enero de 2020</li>
<li>Reducción de los tiempos para la elaboracion y publicacion del contrato (de 30 minutos a &lt; 3 minutos) y sus modificaciones contractuales (de 15 minutos a &lt; 2.5. minutos).&nbsp;</li>
<li>Minimiza los errores en la digitación de la información, mejorando la calidad de la misma.</li>
<li>Reducción del espacio en la gestión documental (archivo digital en SECOP II)</li>
<li>Mejora los procesos de las auditorías&nbsp;</li>
<li>Disminuye los costos financieros y de personal en $60.000.000 anuales&nbsp;</li>
<li>Mejora el acceso a la información por parte de contratista.</li>
<li>Aprovechamiento de las 24 horas del día para el cargue de información en SECOP II.</li>
</ul>
<div>
<h3 style="background-color: white; font-family: Merriweather, Georgia, serif;">
Requisitos</h3>
<span style="background-color: white; font-family: &quot;merriweather&quot; , &quot;georgia&quot; , serif; font-size: 16px;">Para la ejecución del robot, se necesita:</span><br />
<ul style="background-color: white;">
<li style="font-family: merriweather, georgia, serif; font-size: 16px;">Tener instalado&nbsp;<a href="https://github.com/kelaberetiv/TagUI#set-up" style="background: transparent; color: #729c0b; text-decoration-line: none;" target="_blank">TagUI</a></li>
<li style="font-family: merriweather, georgia, serif; font-size: 16px;">Crear la carpeta&nbsp;<b>C:\secop2</b></li>
<li style="font-family: merriweather, georgia, serif; font-size: 16px;">Crear la carpeta&nbsp;<b>C:\secop2\documentos</b>&nbsp;con los archivos de cada contrato a cargar</li>
<li style="font-family: merriweather, georgia, serif; font-size: 16px;">Descargar los siguientes archivos:</li>
<ul style="font-family: merriweather, georgia, serif; font-size: 16px;">
<li>Código fuente de cada robot</li>
<li><a href="https://github.com/alfarodaniel/RPA-TagUI-SECOPII/blob/master/Base_de_datos_Contratacion.csv" style="background: transparent; color: #729c0b; text-decoration-line: none;" target="_blank">Base_de_datos_Contratacion.csv</a>&nbsp;- listado de información de contratos a cargar</li>
</ul>
<li><span style="font-family: &quot;merriweather&quot; , &quot;georgia&quot; , serif;">En el código de cada robot&nbsp;en la linea "select seldpCompany as #" colocar el número de identificación de su empresa</span></li>
</ul>
</div>
<h3>
Robot Precontractual</h3>
</div>
<div>
Crear la parte precontractual del contrato partiendo del requerimiento del perfil.</div>
<div class="separator" style="clear: both; text-align: center;">
<iframe allowfullscreen="" class="YOUTUBE-iframe-video" data-thumbnail-src="https://i.ytimg.com/vi/lkpWoyAoh0o/0.jpg" frameborder="0" height="266" src="https://www.youtube.com/embed/lkpWoyAoh0o?feature=player_embedded" width="320"></iframe></div>
<div>
Código fuente del robot: <a href="https://github.com/alfarodaniel/RPA-TagUI-SECOPII/blob/master/secop2precontractual" target="_blank">secop2precontractual</a><br />
La estructura del nombre de los documentos a cargar es&nbsp;req-" + numproceso + "-2020.pdf<br />
<h3>
Robot Aprobación</h3>
Aprobación de la etapa precontractual por la Dirección de contratación.<br />
<div class="separator" style="clear: both; text-align: center;">
<iframe allowfullscreen="" class="YOUTUBE-iframe-video" data-thumbnail-src="https://i.ytimg.com/vi/nvFwcx4E1EA/0.jpg" frameborder="0" height="266" src="https://www.youtube.com/embed/nvFwcx4E1EA?feature=player_embedded" width="320"></iframe></div>
Código fuente del robot:&nbsp;<a href="https://github.com/alfarodaniel/RPA-TagUI-SECOPII/blob/master/secop2aprobacion" target="_blank">secop2aprobacion</a><br />
<h3>
Robot Contrato</h3>
<div>
Inclusión del CDP, documentos de la hoja de vida del contratista y creación del contrato para la adjudicación y aprobación del contratista.</div>
<div class="separator" style="clear: both; text-align: center;">
<iframe allowfullscreen="" class="YOUTUBE-iframe-video" data-thumbnail-src="https://i.ytimg.com/vi/iOZcHAE3oRI/0.jpg" frameborder="0" height="266" src="https://www.youtube.com/embed/iOZcHAE3oRI?feature=player_embedded" width="320"></iframe></div>
Código fuente del robot:&nbsp;<a href="https://github.com/alfarodaniel/RPA-TagUI-SECOPII/blob/master/secop2contrato" target="_blank">secop2contrato</a><br />
La estructura del nombre de los documentos a cargar es&nbsp;CPS-" + numproceso + "-2020.pdf<br />
La estructura del nombre de los documentos anexos a cargar es " + documento + ".zip<br />
Por cada contrato cargado, el robot escribe una linea con el enlace al contrato en el archivo&nbsp;<a href="https://github.com/alfarodaniel/RPA-TagUI-SECOPII/blob/master/enlace.csv" target="_blank">enlace.csv</a><br />
<h3>
Robot Modificación</h3>
<div>
Subir y adjudicar toda modificación contractual.</div>
<div class="separator" style="clear: both; text-align: center;">
<iframe allowfullscreen="" class="YOUTUBE-iframe-video" data-thumbnail-src="https://i.ytimg.com/vi/mZyPGMEaY5M/0.jpg" frameborder="0" height="266" src="https://www.youtube.com/embed/mZyPGMEaY5M?feature=player_embedded" width="320"></iframe></div>
<div>
Código fuente del robot:&nbsp;<a href="https://github.com/alfarodaniel/RPA-TagUI-SECOPII/blob/master/secop2modificacion" target="_blank">secop2modificacion</a><br />
La estructura del nombre de los documentos a cargar es&nbsp;OTROSI-" + clasen + "-CPS-" + numproceso + "-2019.pdf</div>
<div>
<br /></div>
</div>
