<h1>Documentación sobre la Aplicación de Lectura del Sistema</h1>

<p>Esta aplicación permite obtener información detallada sobre el hardware y sistema operativo de un equipo Windows. Utiliza las librerías <strong>WMI</strong>, <strong>psutil</strong>, <strong>platform</strong>, <strong>pandas</strong>, y <strong>customtkinter</strong> para recolectar y mostrar los datos, así como para exportarlos a un archivo Excel.</p>

<h2>Características</h2>
<ul>
    <li>Recopila información sobre el procesador, la memoria RAM, el disco duro, la placa base, las tarjetas gráficas y los adaptadores de red.</li>
    <li>Muestra los datos en una interfaz gráfica creada con <strong>CustomTkinter</strong>.</li>
    <li>Permite guardar toda la información en un archivo Excel para futuras referencias.</li>
</ul>

<h2>Requisitos</h2>
<p>Para ejecutar esta aplicación, necesitas tener instaladas las siguientes librerías:</p>
<ul>
    <li><code>wmi</code></li>
    <li><code>psutil</code></li>
    <li><code>platform</code></li>
    <li><code>pandas</code></li>
    <li><code>customtkinter</code></li>
</ul>

<h2>Uso</h2>
<ol>
    <li>Al iniciar la aplicación, verás un botón llamado "Get System Info". Haz clic en él para recolectar la información del sistema.</li>
    <li>La información recolectada será mostrada en el cuadro de texto de la interfaz.</li>
    <li>Si deseas guardar la información en un archivo Excel, haz clic en el botón "Save to Excel". Se abrirá un cuadro de diálogo para que elijas la ubicación y el nombre del archivo.</li>
</ol>

<h2>Estructura del Proyecto</h2>
<p>El proyecto está estructurado de la siguiente manera:</p>
<ul>
    <li><strong>get_full_system_info()</strong>: Función que recopila la información del sistema y la devuelve en un diccionario.</li>
    <li><strong>show_system_info()</strong>: Muestra la información recopilada en el cuadro de texto de la interfaz.</li>
    <li><strong>save_to_excel()</strong>: Guarda la información recopilada en un archivo Excel usando la librería <code>pandas</code>.</li>
</ul>

<h2>Interfaz Gráfica</h2>
<p>La interfaz gráfica está construida con <strong>CustomTkinter</strong>. Incluye un cuadro de texto para mostrar la información del sistema, un botón para iniciar el proceso de recolección de datos, y otro botón para guardar la información en un archivo Excel.</p>

<h2>Instalación de Dependencias</h2>
<p>Para instalar las dependencias necesarias, puedes ejecutar el siguiente comando:</p>
<pre><code>pip install wmi psutil pandas customtkinter</code></pre>

<h2>Captura de Pantalla</h2>
<p>A continuación se muestra una captura de pantalla de la interfaz de la aplicación:</p>


<h2>Contacto</h2>
<p>Si tienes alguna duda o sugerencia sobre la aplicación, no dudes en contactarme a través de <strong>correo@ejemplo.com</strong>.</p>
