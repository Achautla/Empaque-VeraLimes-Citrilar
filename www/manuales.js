/* v8.32.1 · Manuales de usuario por puesto (Menú › Sesión › Mi cuenta).
   Un manual por rol operativo: Jefe de empaque, Jefe de inocuidad, Asistencia en tarimas y Conteo de segundas.
   Administración y superadministración no tienen manual propio, pero pueden consultar los cuatro.
   Solo texto: este archivo no guarda ni cambia datos. */
(function(){
'use strict';
var VER='1.2',VIG='30 de octubre de 2026',EMPRESAS='Exportadora de Cítricos Vera Limes, S.A. de C.V. y Citrilar';

/* ---------- capítulos comunes ---------- */
function acceso(puesto,extra){
  return '<p>Tu cuenta es personal. Todo lo que guardas queda registrado con tu nombre, la fecha y la hora, así que nadie más debe usarla.</p>'+
  '<h4>Entrar</h4><ol>'+
  '<li>Abre <b>Empaque Vera Limes</b> (APK en Android, o la versión web instalada en iPhone o computadora).</li>'+
  '<li>Escribe el correo y la contraseña que te dio administración. El botón del ojo muestra lo que escribes.</li>'+
  '<li>Marca que leíste el aviso de privacidad y los términos de uso, y toca <b>Entrar</b>. Cuando cambia la versión de la información legal, se vuelve a pedir.</li>'+
  '<li>La primera vez necesitas internet. Después la app abre aunque no haya señal.</li></ol>'+
  '<h4>Huella (solo APK de Android)</h4><p>En <b>Mi cuenta</b> puedes vincular tu huella para entrar más rápido. La huella solo abre tu propia cuenta en tu propio teléfono.</p>'+
  '<h4>Sesión que se cierra sola</h4><ul>'+
  '<li>En el teléfono, si sales de la app sin cerrar sesión y regresas después de <b>30 minutos</b>, la app vuelve a pedir contraseña o huella. Lo que capturaste se conserva.</li>'+
  '<li>Un gesto de regresar retrocede una pantalla. Dos gestos rápidos salen de la app sin cerrar sesión.</li></ul>'+
  '<h4>Contraseña</h4><ul><li>Las contraseñas las administra la empresa. Si la olvidaste o crees que alguien la conoce, avisa de inmediato a administración para que la cambie.</li>'+
  '<li>Nunca la mandes por WhatsApp ni la escribas en papel junto al equipo.</li></ul>'+
  '<h4>Cerrar sesión</h4><p>Al terminar tu turno en un equipo compartido, toca <b>Sincronizar</b> y después <b>Cerrar sesión</b>. Si cambian tu puesto, el nuevo rol aplica la siguiente vez que entres.</p>'+
  (extra||'');
}
function sincro(){
  return '<p>La app guarda primero en tu dispositivo y después envía a la base de datos de la empresa. Por eso puedes trabajar en piso sin señal.</p>'+
  '<ul><li>La <b>S</b> de arriba es Sincronizar. Muestra cuántos cambios faltan por enviar y dice <b>Al día</b> o <b>en vivo</b> cuando todo está enviado.</li>'+
  '<li>Con internet, lo que guarda una persona lo ven las demás casi al instante. Sin internet se envía en cuanto vuelve la red.</li>'+
  '<li>Si dos personas cambian <b>exactamente el mismo dato</b>, se queda el de quien sincronizó al último. Cada quien debe capturar lo que le toca.</li>'+
  '<li>Antes de apagar el teléfono, cambiar de equipo o cerrar sesión, sincroniza y espera a ver <b>Al día</b>.</li></ul>';
}
function legal(puesto){
  return '<p>La app es de uso interno de '+EMPRESAS+'. Solo personal autorizado puede usarla, y solo para su trabajo.</p>'+
  '<h4>Lo que te toca cuidar</h4><ul>'+
  '<li><b>Veracidad:</b> captura lo que de verdad pasó. Con lo que capturas se paga a productores, se emiten documentos de exportación y se responde a auditorías de SENASICA y PrimusGFS.</li>'+
  '<li><b>Registro de cambios:</b> cada alta, cambio, anulación o borrado queda con tu nombre, fecha y hora. Corregir un error está bien; esconderlo no.</li>'+
  '<li><b>Datos personales:</b> la app tiene nombres, teléfonos y datos de productores, conductores y personal. No los compartas fuera de la empresa ni los uses para otra cosa. El detalle está en Menú › Uso interno · Legal (aviso de privacidad, aviso laboral y términos de uso).</li>'+
  '<li><b>Documentos:</b> manifiestos, packing lists, registros y etiquetas se envían solo a quien corresponde por su trabajo.</li>'+
  '<li><b>Equipo:</b> mantén el teléfono con bloqueo de pantalla. Si lo pierdes o te lo roban, avisa a administración el mismo día para bloquear tu cuenta.</li>'+
  '<li><b>No intentes</b> entrar con la cuenta de otra persona, saltarte un candado o modificar la app. Además de ser una falta, la base de datos lo rechaza y queda registrado.</li></ul>'+
  '<blockquote>Si algo no cuadra o la app no te deja hacer algo que necesitas para tu trabajo, no busques un atajo: avisa a tu jefe inmediato o a administración.</blockquote>';
}
function tabla(rows){return '<table class="mtbl"><thead><tr><th>Módulo</th><th>Qué puedes hacer</th></tr></thead><tbody>'+rows.map(function(r){return '<tr><th scope="row">'+r[0]+'</th><td>'+r[1]+'</td></tr>'}).join('')+'</tbody></table>'}

var M={};

/* ======================= JEFE DE EMPAQUE ======================= */
M.jefe={puesto:'Jefe de empaque',caps:[
{t:'Tu puesto en la app',h:
'<p>El jefe de empaque es la fuente de casi todos los datos de la empresa. Lo que registras en cada viaje decide:</p><ul>'+
'<li><b>Cuánto se paga a cada productor:</b> kilos de báscula, cajas empacadas, rendimiento, segundas, cera, Torreón, desecho y Juguera.</li>'+
'<li><b>Qué dicen los documentos de exportación:</b> manifiesto, packing list, código de trazabilidad y etiquetas con QR.</li>'+
'<li><b>Qué hay en almacén</b> y qué se manda a Juguera al cerrar la semana.</li>'+
'<li><b>Las cifras que ve dirección</b> en Estadísticas, Finanzas e Inteligencia, aunque tú no veas esas pantallas.</li></ul>'+
'<p>Un error en un calibre, un productor o una cantidad no se queda en el empaque: llega al pago, al cliente en Estados Unidos y a la auditoría. Por eso la app avisa con colores cuando algo no cuadra.</p>'+
'<h4>Tu jornada en cinco pasos</h4><ol><li>Revisar Entrada de Fruta y abrir el pedido del viaje.</li><li>Agregar a los productores en el orden en que entran a la línea y poner en turno al primero para asistencia en tarimas.</li><li>Capturar el conteo de cada productor (o validarlo desde las tarimas) y sus otras cajas.</li><li>Armar y revisar tarimas en Código de Trazabilidad.</li><li>Completar registro, autorizar el packing list, cerrar el manifiesto y notificar «empaque procesado».</li></ol>'},
{t:'Alcance: qué ves y qué puedes cambiar',h:
tabla([
['Entrada de Fruta','Ver y capturar entradas, pesos, inventario, procesado y la pestaña Segundas (Juguera y desecho).'],
['Almacén','Ver lo que no va en un thermo y registrar movimientos con justificación.'],
['Registro de Empaque','Crear pedidos, agregar productores, capturar conteos, hoja 1.5, sobrantes, reempaque, manifiesto y completar registro.'],
['Código de Trazabilidad','Encabezado, tarimas, turno de asistencia, autorizar packing list, cerrar y reabrir manifiesto.'],
['Documentos de viaje','Ver, descargar y enviar todos los documentos del viaje; agregar manifiestos por venta.'],
['Productores','Solo consultar la lista. Altas y cambios los hace inocuidad o administración.'],
['Peso por caja','Solo consultar. Los pesos los cambia administración.']])+
'<h4>Lo que no ves</h4><ul><li>Estadísticas, Finanzas, Inteligencia, precios y pagos.</li><li>Control Logístico (los catálogos los mantiene inocuidad o administración; tú los eliges de la lista).</li><li>Administración: pendientes, Registro de cambios completo, usuarios y respaldos.</li><li>Periodos: abrir y cerrar semanas.</li></ul>'},
{t:'Acceso, sesión y Mi cuenta',h:acceso('Jefe de empaque')},
{t:'Trabajo sin señal y sincronización',h:sincro()},
{t:'Recepción: Entrada de Fruta',h:
'<p>Cada entrega de fruta se registra una vez y queda ligada al empaque donde se corre. De ahí salen los kilos con que se paga.</p>'+
'<h4>Registrar una entrada</h4><ol><li>Menú › Recepción › <b>Entrada de Fruta</b> › agregar.</li><li>Fecha y productor (al escribir el código o nombre de la lista se completa el otro).</li><li>Cajas grandes, cajas chicas y <b>peso de báscula</b>. Si el peso llega después, se captura luego aquí o en la ficha del productor en el empaque: es el mismo dato.</li><li>Lo que ya llega catalogado: Segunda, Torreón y Juguera.</li><li>Zona o huerta de origen y muestreo de recepción (% primera, segunda y desecho), si se hizo.</li><li>«+ Agregar pesos» para pesos de muestra por caja; la app estima el peso del lote.</li></ol>'+
'<blockquote>Si los kilos entre cajas salen fuera de 12 a 25 kg por caja de campo, la app avisa en naranja. Revisa que no falte un cero o que no se hayan cruzado cajas grandes y chicas.</blockquote>'+
'<h4>Pestañas</h4><ul><li><b>Entradas:</b> las que todavía tienen fruta por empacar. Cada una dice por qué sigue ahí: sin cantidades, sin ligar a un empaque, o en empaque con fruta pendiente.</li><li><b>Inventario:</b> lo disponible por productor y por semana, con sus segundas.</li><li><b>Procesado:</b> entradas ya empacadas. La etiqueta naranja «Sin asignar» indica que los kilos están en un empaque pero la clasificación todavía no cuadra.</li><li><b>Segundas:</b> ver el capítulo «Hoja 1.5, segundas, Juguera y desecho».</li></ul>'+
'<h4>Corregir</h4><ul><li><b>Ligar con VL-…:</b> liga la entrada con el renglón del productor sin cambiar las cajas capturadas.</li><li><b>Usar el peso de báscula en …:</b> cuando el renglón tiene kilos escritos a mano distintos a la báscula.</li><li>Una semana cerrada no deja cambiar el peso, porque los pagos ya se congelaron.</li></ul>'+
'<h4>Hoja semanal</h4><p><b>Hoja semanal · Excel o PDF</b> saca el reporte de entrada de fruta de la semana con el formato de la empresa, para enviar por WhatsApp o guardar.</p>'},
{t:'Registro de Empaque: crear el pedido',h:
'<ol><li>Menú › Empaque › <b>Registro de Empaque</b>. Elige la semana arriba.</li><li><b>Nuevo pedido</b>: elige por qué empresa sale (Vera Limes o Citrilar), el cliente de la lista y revisa la factura sugerida.</li><li>La app pone el consecutivo del año (por ejemplo VL-2026-189) y el código de embarque (año, semana y número de empaque de la semana).</li><li>Al elegir un cliente del catálogo se llenan destino, INCOTERM y días de pago si estaban vacíos.</li></ol>'+
'<h4>Encabezado</h4><p>En «Editar encabezado» van destino, INCOTERM, cajas rechazadas y su motivo, naviera o transportista, contenedor, fechas, personas en el turno y minutos de paro. Se capturan una vez y valen para las hojas 1 y 1.5.</p>'+
'<h4>Cambiar la empresa del pedido</h4><p>Con el interruptor «Sale por» un pedido a medio registro puede pasar de Vera Limes a Citrilar o al revés. Cambian juntas sus hojas, el código, las marcas de tarimas y los membretes. Hazlo solo si de verdad cambió la venta.</p>'+
'<h4>Peso por caja</h4><p>El engrane muestra los pesos con que se calculan los kilos empacados (RPC 17.8, CTN 16, CTN 10 lb 4.53 y PLT 17.8 kg, salvo que administración los cambie). Tú los consultas; si ves uno raro, avisa.</p>'},
{t:'Registro de Empaque: productores y conteo',h:
'<h4>Agregar productores</h4><ul><li>Agrégalos en el orden en que entran a la línea. Ese orden se usa para ordenar las tarimas.</li><li><b>Turno automático:</b> si nadie está en turno, el productor que agregas queda en turno para asistencia en tarimas. El tipo de empaque que elijas en su ficha pasa al turno.</li><li><b>Poner en turno:</b> en la ficha de cualquier productor, dentro de «Tarimas (Código de Trazabilidad)», pone a ese productor en turno. La ficha del que está en turno lo dice en verde.</li><li>Al agregarlo se carga su entrega más antigua pendiente de Entrada de Fruta. Si trae otra entrega para la misma corrida, usa <b>+ Agregar otra entrega</b>; se suman cajas y pesadas.</li><li>Si el productor ya estaba capturado sin ligar, usa <b>Ligar con la entrega del …</b> (no sobrescribe lo capturado).</li></ul>'+
'<h4>Dos formas de contar</h4><p>Todo viaje nuevo nace como <b>Viaje contado en tarimas</b>. Si ese día cuentas con el contador de la línea, apágalo al inicio de Captura.</p><ol><li><b>Conteo acumulado:</b> capturas el acumulado por calibre. Las casillas aceptan sumas (escribe +50+50) y guardan el total al salir.</li><li><b>Viaje contado en tarimas</b> (al inicio de Captura): las cajas de cada productor salen de las tarimas que arma asistencia. En cada productor aparece <b>Cajas en tarimas</b> con su rendimiento; <b>Validar y pasar al siguiente</b> llena el registro y pone en turno al siguiente productor. Si después cambian las tarimas, la app avisa y ofrece <b>Usar lo de tarimas</b>.</li></ol>'+
'<h4>Otras cajas</h4><p>Segundas, Cera, Torreón, Desecho y Juguera. Si el equipo de conteo de segundas ya contó, verás el recuadro <b>Conteo de segundas</b> con su total; <b>Usar conteo</b> llena tus casillas y después puedes corregirlas. <b>Tu casilla es la que cuenta.</b></p>'+
'<h4>La diferencia «Dif»</h4><p>Dif = cajas campo − kilos empacados ÷ 17.8 − segundas − cera − Torreón − desecho − Juguera. Queda en verde con hasta 3 cajas o 3 %. En rojo significa que falta o sobra fruta por clasificar: revisa antes de completar.</p>'+
'<h4>Botones de la ficha del productor</h4><ul><li><b>Se llevó sus segundas:</b> registra cuántas segundas se llevó el productor; ya no entran a inventario.</li><li><b>Mover a otro viaje:</b> pasa el renglón completo a otro viaje como si ahí hubiera sido su primer empaque. La ventana explica antes qué pasa.</li><li>El bloque <b>Tarimas (Código de Trazabilidad)</b> compara lo registrado con lo que hay en tarimas: ✓ cuadra, ▲ te pasaste, • falta registrar.</li></ul>'},
{t:'Hoja 1.5, segundas, Juguera y desecho',h:
'<p>La hoja 1.5 es la corrida de segunda del mismo pedido. «1.5» es la fruta que va en tarima a Estados Unidos; «segundas» es la de venta nacional. Las cajas de la hoja 1.5 se reparten a cada productor en proporción a las segundas que aportó.</p>'+
'<h4>Armar la hoja 1.5</h4><ol><li>Crea la hoja 1.5 del pedido. Se abre <b>Segundas disponibles</b> por productor y entrega, incluidas las de viajes anteriores.</li><li>Escribe cuántas van, o usa «Todo» por productor o «Seleccionar todo».</li><li><b>Agregar a CITRILAR SEGUNDAS</b> las suma en ese renglón; su «Resumen» dice quién aportó qué.</li><li>Lo que no eliges se queda disponible. «+ Segundas» vuelve a abrir la ventana.</li></ol>'+
'<p>Lo que sobra en las casillas Segundas y Con cera del renglón CITRILAR SEGUNDAS vuelve a estar disponible como «Remanente de la hoja 1.5». Si aparecen aportes de un pedido eliminado, usa <b>Quitar los N aportes</b> y corrige tu casilla con lo que de verdad quedó.</p>'+
'<h4>Juguera</h4><ul><li><b>Enviar a Juguera</b> (en Segundas disponibles o en Entrada de Fruta › Segundas) manda segundas por productor y entrega en la semana elegida.</li><li>Mientras la semana esté abierta, un envío se anula tocando dos veces; queda tachado con quién lo anuló.</li><li>Al cerrar la semana, administración manda solas a Juguera las segundas que sigan disponibles.</li></ul>'+
'<h4>Desecho</h4><p>En Recepción › Segundas, <b>Vaciar desecho</b> registra lo que se vacía por viaje, con kilos y nota opcional. La pestaña muestra % de desecho y % a Juguera por productor; con 8 % o más sale el aviso de desecho alto.</p>'},
{t:'Sobrantes, próximo viaje y reempaque',h:
'<h4>Tarima al próximo viaje</h4><ol><li>En Código de Trazabilidad abre la tarima y toca <b>Próximo viaje →</b>.</li><li>Elige el viaje que la recibe (o «Pedido nuevo») y cuántas cajas de cada productor van. Si va una parte, se crea una tarima nueva con esas cajas.</li><li>Confirma. Si ya hay manifiesto, la app avisa si el viaje quedaría corto.</li></ol>'+
'<p>En el viaje que recibe, las cajas entran solas al renglón <b>SOBRANTES</b> con su productor de origen. No llevan kilos ni pago: se pagaron donde se recibieron. La tarima recibida tiene <b>↩ Regresar a su viaje</b> y <b>Enviar a otro viaje →</b>.</p>'+
'<h4>Sobrantes en piso</h4><p><b>Mover sobrantes en piso a otro viaje →</b> (debajo de Existencia en empaque) manda cajas sin entarimar. La app no deja mover más de lo que el viaje tiene de más fuera de tarimas. El destino puede ser de la otra empresa; no se ofrecen viajes completos, cerrados ni con manifiesto cerrado.</p>'+
'<h4>Reempaque</h4><ol><li>En la tarjeta del pedido toca <b>+ Reempaque</b> (aparece cuando ya hay cajas registradas).</li><li>Elige renglones por productor, presentación y calibre, en piso o de una tarima, y cuántas cajas.</li><li>Elige la presentación nueva y el viaje que recibe (el mismo en otra tarima, otro en proceso o uno nuevo).</li><li>La app propone las cajas por peso o por calibre (por ejemplo 250 → 63 al pasar a 10 lb); corrige con las reales y la diferencia queda como merma.</li></ol>'+
'<p>Al productor no le cambia nada: su pago, kilos y rendimiento se quedan en su primer empaque. «↩ Regresar a …» deshace el reempaque mientras el viaje que recibe no esté completo y la semana siga abierta. Con el manifiesto cerrado ya no se reempaca de sus tarimas.</p>'},
{t:'Resumen, manifiesto y completar registro',h:
'<ol><li>En <b>Resumen</b> captura el manifiesto por calidad (Primera y 1-1/2) y tipo de empaque (RPC, CTN, CTN 10 lb, PLT). La tabla «Manifiesto contra empacado» muestra la diferencia.</li><li>Decide en «Tarimas incompletas» si se completan con segunda de la misma presentación.</li><li>Con manifiesto capturado aparece <b>Completar registro</b>. Al confirmar, el pedido queda completo y guarda los pesos por caja con que se completó.</li><li>Cuando el empaque ya terminó, <b>Notificar: empaque procesado</b> manda el aviso «Empaque VL-… procesado» a dirección, administración e inocuidad. Solo en la semana en curso y pide confirmar.</li></ol>'+
'<h4>Descargas del registro</h4><ul><li>Pestaña <b>Excel</b>: vista previa en vivo y «Descargar o enviar» en Excel, PDF, PowerPoint o imagen JPG.</li><li><b>Resumen para enviar</b> de cada productor: imagen con calidad, tipos de empaque y calibres, para WhatsApp.</li><li><b>Descargar Excel de viajes de la semana</b>: todos los registros de la semana en un archivo.</li></ul>'+
'<blockquote>El Excel se puede descargar en cualquier momento. Mientras falte algo, los documentos salen con la leyenda BORRADOR.</blockquote>'},
{t:'Código de Trazabilidad y tarimas',h:
'<h4>Encabezado</h4><p>Transportista, conductor, trailer, termo, vendedor y cliente se eligen del catálogo; teléfonos, placas, CAAT y ALPHA se llenan solos. Si falta alguno en la lista, pídelo a inocuidad o a administración.</p>'+
'<h4>Turno para asistencia en tarimas</h4><p>En <b>En turno para asistencia en tarimas</b> indica productor, calidad y presentación. Asistencia en tarimas solo suma cajas por calibre al productor en turno; si no lo indicas, les sale «Falta el productor en turno».</p>'+
'<h4>Armar tarimas</h4><ul><li>Cada tarima tiene una sola presentación; dentro se mezclan calidades, calibres y productores.</li><li>Cajas estándar: RPC 50, CTN 60, CTN 10 lb 210, PLT 50. En RPC se aceptan 54, 55 y 60 como excepción.</li><li>El punto verde dice que la tarima cuadra; naranja, que faltan calibres por registrar; rojo, que falta calibre, cajas o productor, o que algo no cuadra con el registro completo.</li><li><b>Acomodar tarimas como el empaque</b> pone en tarimas las cajas de primera ya registradas que faltan y ordena todo como entraron los productores. No mete cajas no registradas ni quita las que armaste. La 1.5 se completa con «Completar con segunda».</li><li>Ordena con el asa ☰, con ↑ ↓ o con «Lugar». El orden es el del thermo y lo siguen el packing list y el dibujo de la caja.</li><li><b>Buscar productor en las tarimas</b> marca dónde está su fruta.</li></ul>'+
'<h4>Revisar lo que hizo asistencia</h4><p><b>Movimientos de asistencia en tarimas</b> lista lo que sumaron; cada movimiento se puede <b>Deshacer</b>. En viajes antiguos con conteos de la versión 8.31 se autorizan, rechazan o corrigen uno por uno.</p>'+
'<h4>Supervisar al equipo de piso</h4><p>En el Código de Trazabilidad de cada viaje abierto está el panel <b>Equipo de piso</b>, que ven solo el jefe de empaque y administración:</p><ul><li><b>Tarimas del viaje:</b> abre la misma pantalla de Asistencia en tarimas con ese viaje elegido. Puedes sumar, quitar o mover cajas y cambiar el lugar de una tarima.</li><li><b>Conteo de segundas:</b> abre la pantalla del equipo de conteo con ese viaje. Ves el total por productor y la lista <b>Conteos del viaje (supervisión)</b>, donde puedes anular cualquier conteo (dos toques). También puedes capturar un conteo.</li><li>Arriba sale el aviso <b>Supervisión</b>; lo que guardes queda con tu nombre. <b>Volver a Código de Trazabilidad</b> (o el gesto de regresar) te devuelve al viaje.</li><li>Con el manifiesto o la semana cerrados, los botones quedan apagados.</li></ul>'+
'<h4>Existencia por productor</h4><p>Registradas = reempacadas + próximo viaje + bajas en almacén + en tarima + en piso. Úsala para encontrar cajas perdidas antes de cerrar.</p>'+
'<h4>Cerrar el viaje</h4><ol><li>Revisa que no haya puntos rojos ni avisos en «Contra Registro de Empaque» y «Contra el manifiesto».</li><li>Toca <b>Autorizar packing list y cerrar manifiesto</b>. Si el código sigue en borrador, pide confirmar dos veces.</li><li>Las tarimas quedan congeladas: ni tú ni asistencia las pueden cambiar, y sus cajas ya no se reempacan.</li><li>Para corregir, <b>Reabrir manifiesto</b> pide escribir el motivo, que queda registrado.</li></ol>'},
{t:'Documentos de viaje y etiquetas',h:
'<p>Menú › <b>Documentos de viaje</b> junta todo lo de un viaje: Manifiesto, Manifiesto interno, Registro de Empaque, Packing List, Código de Trazabilidad y Etiquetas con QR. Cada uno tiene vista previa, descarga y un botón para ir a donde se corrige.</p>'+
'<h4>Varios manifiestos por viaje</h4><p>En «Editar datos», <b>+ Agregar manifiesto</b> crea uno por venta, con su cliente, factura y P.O. Asigna tarimas tocándolas o por rango. Cada venta saca su packing list con sus tarimas sombreadas. Con el manifiesto cerrado ya no se cambian las asignaciones.</p>'+
'<h4>Etiquetas con QR</h4><p>Una etiqueta por productor, calidad y presentación. El QR lleva el lote: embarque, empresa y código del productor, calidad y fecha (por ejemplo JC263301-EC003-1.5-20260812). Los productores sin código no se etiquetan: pide que se les asigne uno.</p>'+
'<blockquote>Antes de mandar un documento al cliente, revisa que no diga BORRADOR y que el número de cajas sea el del thermo.</blockquote>'},
{t:'Almacén',h:
'<p>Menú › Recepción › <b>Almacén</b> junta todo lo que no va en un thermo, de las dos empresas: cajas en piso de viajes que ya salieron o siguen en proceso, cajas en tarimas sin salir y segundas sin empacar. Cada caja conserva productor, calidad, presentación, calibre, viaje de origen y días desde su empaque. Solo se ven las semanas abiertas.</p>'+
'<h4>Mover</h4><ul><li><b>Pasar a segunda:</b> la caja deja de ser de exportación y queda como lote de segunda de su productor.</li><li><b>Merma:</b> se echó a perder; pide los kilos de báscula.</li><li><b>Ajuste por conteo:</b> no está físicamente (por ejemplo, sí salió en el thermo).</li></ul>'+
'<p>Siempre pide justificación y guarda quién y cuándo. Solo salen cajas en piso; las tarimas no se tocan. <b>Anular</b> (dos toques) regresa las cajas y deja el movimiento tachado. Ningún movimiento de almacén cambia el pago al productor.</p>'},
{t:'Restricciones y candados',h:
'<ul><li><b>Semana cerrada:</b> sus pedidos quedan en solo lectura y no se crean pedidos nuevos en ella. Solo administración la reabre.</li><li><b>Manifiesto cerrado:</b> tarimas, asignaciones y reempaques de esas tarimas quedan congelados hasta reabrir con motivo.</li><li><b>Registro completo:</b> se puede ajustar el peso de báscula (queda registrado), pero revisa antes de completar.</li><li><b>Renglones informativos:</b> CITRILAR SEGUNDAS, REEMPAQUE y SOBRANTES no son productores; su fruta se atribuye a los productores que traen dentro.</li><li><b>Un productor por código.</b> No captures a dos personas con el mismo código.</li><li>No puedes cambiar precios, pagos, pesos por caja, catálogos de logística, productores ni usuarios.</li><li>La base de datos revisa tu rol en cada guardado: aunque una pantalla se viera editable por error, el cambio se rechaza.</li></ul>'},
{t:'Implicaciones y errores comunes',h:
'<table class="mtbl"><thead><tr><th>Error</th><th>A quién afecta</th><th>Cómo se evita o corrige</th></tr></thead><tbody>'+
'<tr><td>Productor equivocado en un renglón o tarima</td><td>Pago, trazabilidad y etiqueta QR</td><td>Busca por código, no por nombre. Corrige antes de cerrar el manifiesto.</td></tr>'+
'<tr><td>Calibre o presentación equivocada</td><td>Packing list, precio de venta y cliente</td><td>Revisa los puntos de color y «Contra Registro de Empaque».</td></tr>'+
'<tr><td>Peso de báscula faltante o mal escrito</td><td>Rendimiento y pago</td><td>Atiende el aviso «Falta el peso de báscula» y el de kg por caja fuera de rango.</td></tr>'+
'<tr><td>Dif en rojo al completar</td><td>Pago y cuadre de segundas</td><td>Revisa segundas, cera, Torreón, desecho y Juguera antes de completar.</td></tr>'+
'<tr><td>Cerrar el manifiesto con el código en BORRADOR</td><td>Cliente y aduana</td><td>Lee el motivo del borrador debajo del encabezado y complétalo.</td></tr>'+
'<tr><td>Dos personas capturando el mismo pedido</td><td>Datos que se pisan</td><td>Un pedido lo captura una sola persona a la vez.</td></tr>'+
'<tr><td>Olvidar sincronizar</td><td>Dirección no ve el avance; avisos que no llegan</td><td>Revisa que la S diga Al día antes de irte.</td></tr></tbody></table>'},
{t:'Responsabilidad y uso legal',h:legal('Jefe de empaque')}
]};

/* ======================= JEFE DE INOCUIDAD ======================= */
M.inocuidad={puesto:'Jefe de inocuidad',caps:[
{t:'Tu puesto en la app',h:
'<p>El jefe de inocuidad cuida que cada caja se pueda rastrear de la huerta al cliente, y que los catálogos con que se arman los documentos sean correctos. Eso es lo que piden SENASICA y PrimusGFS, y lo que permite retirar un lote si hace falta.</p><ul>'+
'<li><b>Hacia atrás:</b> de una etiqueta QR al embarque, al productor, a su entrega (fecha, huerta, báscula y muestreo).</li>'+
'<li><b>Hacia adelante:</b> de una entrega a los viajes, tarimas, thermo, transportista y cliente donde terminó.</li>'+
'<li><b>Datos maestros:</b> productores con su código y clasificación, y los catálogos de Control Logístico. Si un catálogo está mal, sale mal en todos los documentos.</li></ul>'+
'<p>Además puedes trabajar el Registro de Empaque, el Almacén y los documentos igual que el jefe de empaque. Lo único que no puedes es borrar pedidos.</p>'},
{t:'Alcance: qué ves y qué puedes cambiar',h:
tabla([
['Entrada de Fruta','Ver y capturar entradas, pesos, zona o huerta, muestreo; pestaña Segundas: Juguera, desecho e índices.'],
['Código de Trazabilidad','Encabezado, datos generales y tarimas; cerrar y reabrir manifiesto con motivo.'],
['Productores','Agregar, corregir o liberar códigos; nombre y clasificación (Chicos, Grandes, Internos). El acuerdo de compra lo ve y cambia solo administración.'],
['Control Logístico','Trailers, termos, conductores, transportistas, vendedores y clientes.'],
['Registro de Empaque','Capturar y corregir igual que el jefe de empaque; no borra pedidos.'],
['Almacén','Acomodar, enviar al almacén y cargar tarimas al viaje.'],
['Documentos de viaje','Editar, generar y descargar.']])+
'<h4>Lo que no ves</h4><ul><li>Estadísticas, Finanzas, precios y pagos.</li><li>Administración (pendientes, Registro de cambios completo, usuarios, respaldo) y Periodos.</li></ul>'+
'<p>Recibes el aviso <b>Empaque procesado</b> cuando el jefe de empaque termina un viaje.</p>'},
{t:'Acceso, sesión y Mi cuenta',h:acceso('Jefe de inocuidad')},
{t:'Trabajo sin señal y sincronización',h:sincro()},
{t:'Recepción: Entrada de Fruta',h:
'<p>La entrada es el primer eslabón de la trazabilidad. Lo que captures aquí es lo que se podrá responder en una auditoría.</p>'+
'<h4>Registrar una entrada</h4><ol><li>Menú › Recepción › <b>Entrada de Fruta</b> › agregar.</li><li>Fecha, productor por código, cajas grandes y chicas, y peso de báscula.</li><li><b>Zona o huerta de origen.</b> Sin huerta, la trazabilidad hacia atrás termina en el productor.</li><li><b>Muestreo de recepción:</b> % de primera, segunda y desecho. Se compara después contra el desecho real.</li><li>Lo que llega catalogado (Segunda, Torreón, Juguera) y notas.</li></ol>'+
'<p>La app avisa si los kilos por caja salen de 12 a 25 kg. Revisa en el momento: después es difícil reconstruirlo.</p>'+
'<h4>Revisar el día</h4><ul><li><b>Entradas</b> muestra por qué cada una sigue pendiente: sin cantidades, sin ligar a un empaque o con fruta por pasar.</li><li><b>Procesado</b> marca «Sin asignar» cuando los kilos ya están en un empaque pero la clasificación no cuadra.</li><li><b>Hoja semanal · Excel o PDF</b> sirve como evidencia de recepción de la semana.</li></ul>'},
{t:'Segundas, desecho e índices',h:
'<p>En Recepción › <b>Segundas</b> eliges la semana y ves disponibles, con cera, procesadas en hoja 1.5 y enviadas a Juguera, con la revisión «salieron = disponibles + procesadas + Juguera».</p>'+
'<h4>Desecho</h4><ol><li><b>Vaciar desecho</b> muestra lo que falta por vaciar por viaje.</li><li>Registra todo o una parte, con kilos de báscula y nota.</li><li>Cada vaciado queda con quién y cuándo; <b>Anular</b> con dos toques.</li><li>Al cerrar la semana lo que falte se vacía solo.</li></ol>'+
'<h4>Índices</h4><ul><li>% de desecho y % a Juguera de la semana y por productor.</li><li>Aviso de <b>desecho alto</b> con 8 % o más.</li><li>Columna <b>Muestreo</b>: el % de desecho del muestreo de recepción contra el real. Una diferencia grande indica que el muestreo no representa la fruta o que algo pasó en la línea.</li><li><b>Últimas semanas</b> marca en rojo una semana pasada que quedó sin vaciar todo.</li></ul>'+
'<p><b>Enviar a Juguera</b> desde esta pestaña manda segundas por productor y entrega en la semana elegida; mientras la semana esté abierta se puede anular.</p>'},
{t:'Productores',h:
'<p>Regla de la empresa: <b>un productor por código</b>. El código es lo que viaja en tarimas, documentos y etiquetas QR.</p>'+
'<h4>Qué puedes hacer</h4><ul><li>Agregar un código nuevo (la lista muestra los espacios vacíos), corregir el nombre o liberar un código.</li><li>Asignar la clasificación: Chicos, Grandes o Internos.</li><li>Buscar por código o nombre en todas las listas.</li></ul>'+
'<h4>Cuidados</h4><ul><li>No reutilices el código de un productor que ya tuvo fruta en la temporada.</li><li>Si un nombre junta a dos personas, sepáralas en dos códigos.</li><li>Los productores sin código no se etiquetan con QR.</li><li>Los cambios se guardan en la base al momento; quedan en el Registro de cambios.</li></ul>'},
{t:'Control Logístico',h:
'<p>Seis catálogos que se capturan una sola vez y se eligen en Registro de Empaque, Código de Trazabilidad, Manifiesto y Packing List:</p>'+
'<table class="mtbl"><thead><tr><th>Catálogo</th><th>Datos</th></tr></thead><tbody>'+
'<tr><td>Trailers</td><td>Número económico, placas del tractor, transportista</td></tr>'+
'<tr><td>Termos</td><td>Número de caja, placas del termo, transportista</td></tr>'+
'<tr><td>Conductores</td><td>Nombre, teléfono, placas del tractor, transportista</td></tr>'+
'<tr><td>Transportistas</td><td>Nombre, teléfono, código CAAT, código ALPHA</td></tr>'+
'<tr><td>Vendedores</td><td>Nombre, teléfono</td></tr>'+
'<tr><td>Clientes</td><td>Nombre, país y puerto de destino, INCOTERM, días de pago</td></tr></tbody></table>'+
'<ul><li>Al elegir un registro en un documento se completan teléfonos, placas y códigos. Un dato mal escrito aquí sale mal en todos los viajes.</li><li>Revisa placas y CAAT/ALPHA contra los documentos del transportista.</li><li>Son datos personales y de terceros: no los compartas fuera de la empresa.</li></ul>'},
{t:'Código de Trazabilidad',h:
'<h4>Encabezado y datos generales</h4><p>Elige transportista, conductor, trailer, termo, vendedor y cliente del catálogo. Si falta el transportista se toma del trailer, termo o conductor elegido.</p>'+
'<h4>Tarimas</h4><ul><li>Una presentación por tarima; dentro pueden ir varias calidades, calibres y productores.</li><li>Cada línea debe tener calidad, calibre, cajas y productor. El punto rojo te dice qué falta.</li><li>El orden de las tarimas es el orden real en el thermo. Así se puede ubicar una tarima en una inspección.</li><li><b>Buscar productor en las tarimas</b> marca en qué tarimas va la fruta de un productor y cuántas cajas.</li></ul>'+
'<h4>Cierre del manifiesto</h4><p>Puedes autorizar el packing list y cerrar el manifiesto, o reabrirlo escribiendo el motivo. El cierre congela tarimas y lo que dicen los documentos.</p>'+
'<blockquote>Antes de cerrar, confirma que ninguna línea esté sin productor: una caja sin productor es una caja sin trazabilidad.</blockquote>'},
{t:'Documentos de viaje y etiquetas QR',h:
'<p>En <b>Documentos de viaje</b> consultas y descargas Manifiesto, Manifiesto interno, Registro de Empaque, Packing List, Código de Trazabilidad y Etiquetas con QR de cualquier viaje, por empresa y semana.</p>'+
'<h4>El lote del QR</h4><p>Formato: <b>EMBARQUE-EMPRESA+CÓDIGO-CALIDAD-AAAAMMDD</b>, por ejemplo JC263301-EC003-1.5-20260812. Con el lote sabes embarque, empresa, productor, calidad y fecha de empaque.</p>'+
'<h4>Para una auditoría</h4><ul><li>Descarga en PDF el Código de Trazabilidad, el Packing List y la hoja semanal de entrada de fruta.</li><li>Guárdalos en la carpeta que indique la empresa; no los mandes a personas ajenas.</li></ul>'},
{t:'Rastreo hacia atrás y hacia adelante',h:
'<h4>De una caja a su origen</h4><ol><li>Lee el lote de la etiqueta QR.</li><li>En Documentos de viaje busca el embarque y abre el Código de Trazabilidad: ahí ves la tarima y la línea del productor.</li><li>En Registro de Empaque (consulta) abre la ficha del productor: sus entregas en esa corrida con fecha y peso.</li><li>En Entrada de Fruta abre la entrega: huerta, muestreo, báscula y notas.</li></ol>'+
'<h4>De una entrega a sus destinos</h4><ol><li>Abre la entrega en Entrada de Fruta: «Corridas de esta entrega» dice en qué pedidos se empacó.</li><li>En cada pedido, Código de Trazabilidad › buscar el productor: tarimas y cajas.</li><li>Documentos de viaje: cliente, transportista, conductor, placas y termo.</li><li>Si hubo sobrantes o reempaque, sigue la leyenda «Sobrante de …» o el renglón REEMPAQUE al viaje que la recibió.</li></ol>'+
'<p>Haz este ejercicio con un viaje real al menos una vez al mes, como simulacro.</p>'},
{t:'Consulta de Registro de Empaque y Almacén',h:
'<p>Ves estas pantallas en solo lectura. Sirven para verificar, no para corregir:</p><ul>'+
'<li>Revisa en cada productor la «Dif» y el bloque de tarimas: ▲ te pasaste o • falta registrar indican diferencias entre lo registrado y lo entarimado.</li>'+
'<li>En Almacén revisa la caja más vieja y los días desde su empaque: fruta con muchos días en piso es un riesgo de calidad.</li>'+
'<li>Si encuentras un error, avísale al jefe de empaque o a administración; ellos corrigen y queda registrado.</li></ul>'},
{t:'Restricciones y candados',h:
'<ul><li>No puedes cambiar Registro de Empaque ni Almacén; la base de datos tampoco lo permite.</li><li>No puedes borrar pedidos.</li><li>No ves ni cambias precios, pagos, acuerdos de compra, pesos por caja ni periodos.</li><li>Un manifiesto cerrado congela tarimas hasta reabrirlo con motivo.</li><li>Una semana cerrada deja sus pedidos en solo lectura.</li><li>Cada cambio de catálogo, productor o tarima queda en el Registro de cambios con tu nombre.</li></ul>'},
{t:'Implicaciones y errores comunes',h:
'<table class="mtbl"><thead><tr><th>Error</th><th>Consecuencia</th><th>Cómo se evita</th></tr></thead><tbody>'+
'<tr><td>Entrada sin huerta o sin muestreo</td><td>La trazabilidad termina en el productor</td><td>Captura zona y muestreo en el momento de la recepción.</td></tr>'+
'<tr><td>Dos productores con el mismo código</td><td>Pago y lote QR mezclados</td><td>Un productor por código; separa nombres combinados.</td></tr>'+
'<tr><td>Placas o CAAT/ALPHA mal escritos</td><td>Documentos de transporte con datos falsos</td><td>Compara con la tarjeta de circulación y los documentos del transportista.</td></tr>'+
'<tr><td>Liberar un código con historia</td><td>Lotes antiguos apuntan a otra persona</td><td>No reutilices códigos de la temporada.</td></tr>'+
'<tr><td>Desecho sin vaciar</td><td>Índices falsos y cierre de semana con faltantes</td><td>Registra el vaciado del día.</td></tr></tbody></table>'},
{t:'Responsabilidad y uso legal',h:legal('Jefe de inocuidad')+'<p>El aviso de privacidad laboral contempla datos de salud del personal solo con consentimiento por escrito. Si por inocuidad manejas ese tipo de información, no la captures en la app.</p>'}
]};

/* ======================= ASISTENCIA EN TARIMAS ======================= */
M.tarima={puesto:'Asistencia en tarimas',caps:[
{t:'Tu puesto en la app',h:
'<p>Asistencia en tarimas registra, caja por caja, qué fruta sube a cada tarima. Lo que capturas es lo que el cliente recibe en el thermo:</p><ul>'+
'<li>Con tus tarimas se arma el <b>packing list</b> y el <b>código de trazabilidad</b>.</li>'+
'<li>En los viajes contados en tarimas, tus cajas son las que el jefe de empaque valida para el <b>pago de cada productor</b>.</li>'+
'<li>Cada caja lleva el código de su productor: si alguien pregunta de dónde vino una caja, la respuesta sale de tu captura.</li></ul>'},
{t:'Alcance: qué ves y qué puedes cambiar',h:
tabla([
['Tarimas del viaje','Sumar y corregir cajas, abrir tarimas, mover cajas y cambiar el lugar de la tarima en el thermo, solo en viajes con manifiesto abierto.'],
['Sincronizar (S de arriba)','Enviar y recibir cambios.'],
['Mi cuenta y Cerrar sesión','Tu cuenta, huella y este manual.']])+
'<p>No ves Registro de Empaque, Entrada de Fruta, documentos, finanzas ni administración. La base de datos solo te deja cambiar las tarimas de un viaje abierto; nada más del pedido.</p>'},
{t:'Acceso, sesión y Mi cuenta',h:acceso('Asistencia en tarimas','<p>Al entrar ves un menú con la tarjeta <b>Tarimas del viaje</b> y el panel de tu sesión. Desde la pantalla de tarimas, la flecha de arriba regresa al menú.</p>')},
{t:'Trabajo sin señal y sincronización',h:sincro()+'<p>Cada asistente suma en su propio renglón, así que dos personas tocando la misma tarima no se borran cajas al sincronizar.</p>'},
{t:'La pantalla Tarimas del viaje',h:
'<ol><li><b>Viaje:</b> salen los viajes abiertos más recientes. Si dice «No hay viajes abiertos», espera a que el jefe de empaque abra uno.</li><li><b>En turno:</b> el productor, la calidad y la presentación que puso el jefe de empaque. Si dice «Falta el productor en turno», avísale antes de contar.</li><li><b>Tarimas dibujadas:</b> en el orden en que suben al thermo, de dos en dos como en el almacén.</li></ol>'+
'<h4>Cómo leer el dibujo</h4><ul><li>Cada tarima se dibuja por camas: RPC 5×10, CTN 6×10, CTN 10 lb 14×15 y PLT 5×10.</li><li>Cada caja tiene el color de su calibre; la 1.5 se ve más clara.</li><li>Lo que pasa del estándar se marca en naranja. Se permite, pero confirma que sea correcto.</li><li>Arriba de cada tarima: su número, presentación y cajas contra el estándar (por ejemplo 48/50).</li></ul><p>Las tarimas que armas aquí son las mismas del Código de Trazabilidad y del packing list: lo que sumas aparece ahí al momento.</p>'},
{t:'Sumar cajas',h:
'<ol><li>Toca la tarima donde vas a poner las cajas.</li><li>Elige el calibre de las cajas.</li><li>Suma: <b>+1</b> por caja, <b>+ cama</b> por una cama completa, o escribe la cantidad y toca <b>Sumar</b>.</li><li>Si te equivocaste, <b>−1</b> quita una caja.</li></ol>'+
'<p>Las cajas se suman al productor en turno. Si la fruta es de otro productor del viaje o de otra calidad, elígelo en la tarima antes de sumar.</p>'+
'<blockquote>Cuenta lo que físicamente subió. Si dudas del calibre, pregunta antes de sumar: corregir después cuesta más que preguntar.</blockquote>'},
{t:'Tarimas nuevas, mover cajas y lugar en el thermo',h:
'<h4>Tarima nueva</h4><ol><li>Elige la presentación de la tarima nueva.</li><li>Toca <b>+ Tarima</b>: se abre vacía.</li><li>Una tarima lleva <b>una sola presentación</b>; dentro sí se mezclan calidades, calibres y productores.</li><li>Una tarima que quedó vacía se quita con <b>Quitar tarima vacía</b>.</li></ol>'+
'<h4>Mover cajas</h4><p><b>Mover cajas</b> pasa cajas a otra tarima de la misma presentación.</p>'+
'<h4>Lugar en el thermo</h4><p>Con ↑, ↓ o «Lugar» cambias la posición de la tarima. Antes de cambiar el orden la app pregunta «¿Estás seguro?». El orden es el del thermo y lo usa el packing list, así que cámbialo solo cuando de verdad cambió el acomodo.</p>'},
{t:'Tus movimientos y la revisión del jefe',h:
'<ul><li><b>Tus últimos movimientos</b> muestra lo que sumaste, quitaste o moviste.</li><li>El jefe de empaque revisa tus movimientos y puede deshacer uno; en tu lista aparece «Deshecho por el jefe».</li><li>El jefe de empaque y administración pueden abrir tu misma pantalla desde el Código de Trazabilidad para supervisar o corregir. Sus cambios quedan con su nombre.</li><li>En viajes contados en tarimas, el jefe valida productor por productor con lo que capturaste.</li></ul>'},
{t:'Cuando se cierra el viaje',h:
'<p>Al <b>autorizar el packing list y cerrar el manifiesto</b> (o al cerrar la semana) todo queda congelado y verás «Manifiesto cerrado. Las tarimas ya no se pueden cambiar.». Si encuentras un error después, avisa al jefe de empaque: solo él o administración pueden reabrir, escribiendo el motivo.</p>'},
{t:'Restricciones',h:
'<ul><li>Solo trabajas en viajes con manifiesto abierto.</li><li>No cambias productores, registro, encabezados ni documentos.</li><li>No mezclas presentaciones en una tarima.</li><li>No puedes deshacer lo que el jefe ya validó o cerró.</li><li>Todo lo que haces queda registrado con tu nombre y la hora.</li></ul>'},
{t:'Implicaciones y errores comunes',h:
'<table class="mtbl"><thead><tr><th>Error</th><th>Consecuencia</th><th>Cómo se evita</th></tr></thead><tbody>'+
'<tr><td>Sumar al productor equivocado</td><td>Se paga mal y el lote del QR apunta a otra persona</td><td>Revisa «En turno» antes de cada productor.</td></tr>'+
'<tr><td>Calibre equivocado</td><td>Packing list y precio de venta incorrectos</td><td>Confirma el calibre antes de tocar +1.</td></tr>'+
'<tr><td>Cambiar el lugar sin querer</td><td>El packing list no coincide con el thermo</td><td>Responde con cuidado a «¿Estás seguro?».</td></tr>'+
'<tr><td>Contar sin señal y no sincronizar</td><td>El jefe no ve tus cajas</td><td>Sincroniza en cuanto haya red.</td></tr></tbody></table>'},
{t:'Responsabilidad y uso legal',h:legal('Asistencia en tarimas')}
]};

/* ======================= CONTEO DE SEGUNDAS ======================= */
M.segundas={puesto:'Conteo de segundas',caps:[
{t:'Tu puesto en la app',h:
'<p>Conteo de segundas registra por productor las cajas que no salen como primera. Tu conteo sirve para:</p><ul>'+
'<li><b>Pagar al productor</b> sus segundas y las de cera, que se pagan como segunda.</li>'+
'<li><b>Medir el desecho</b> (terceras) de cada productor y de la semana.</li>'+
'<li><b>Comparar</b> con el conteo del jefe de empaque: dos conteos independientes detectan errores a tiempo.</li></ul>'},
{t:'Alcance: qué ves y qué puedes cambiar',h:
tabla([
['Conteo de segundas','Elegir un viaje abierto y un productor, contar segundas, con cera y terceras, guardar y anular tus propios conteos.'],
['Sincronizar','Enviar y recibir conteos.'],
['Mi cuenta y Cerrar sesión','Tu cuenta, huella y este manual.']])+
'<p>No ves ni cambias el Registro de Empaque, tarimas, documentos, finanzas ni administración.</p>'},
{t:'Acceso, sesión y Mi cuenta',h:acceso('Conteo de segundas','<p>Al entrar se abre directo <b>Conteo de segundas</b>. Abajo están <b>Sincronizar</b>, <b>Mi cuenta</b> y <b>Cerrar sesión</b>.</p>')},
{t:'Trabajo sin señal y sincronización',h:sincro()},
{t:'Contar paso a paso',h:
'<ol><li><b>Viaje:</b> elige el viaje. Solo salen los abiertos; si no hay ninguno, avisa al jefe de empaque.</li><li><b>Productor:</b> escribe el código (VL-156) o el nombre y elígelo de la lista.</li><li><b>Qué cuentas:</b> Segundas, Con cera o Terceras (desecho). Cada botón muestra lo que llevas.</li><li>Escribe las cajas. Puedes sumar por partes: 20+15+12. Toca <b>Sumar</b> (o Enter).</li><li>Cada número queda a la vista; tócalo para quitarlo si te equivocaste.</li><li>Repite con los otros dos tipos si aplica.</li><li>Toca <b>Guardar conteo</b>.</li></ol>'+
'<p><b>Empezar de nuevo</b> borra lo que llevas sin guardar.</p>'},
{t:'Totales, tus conteos y anular',h:
'<ul><li><b>Total contado en el viaje:</b> tabla por productor con segundas, con cera, terceras y número de conteos.</li><li><b>Mis conteos:</b> tus últimos 15, con viaje, productor, hora y cantidades.</li><li><b>Anular:</b> toca dos veces. El conteo queda marcado como anulado; no se borra.</li></ul>'+
'<blockquote>Si un conteo quedó mal, anúlalo y vuelve a contar. No guardes un conteo «de ajuste» con números negativos.</blockquote>'},
{t:'Qué pasa con tu conteo',h:
'<ol><li>El jefe de empaque ve en la ficha del productor el recuadro <b>Conteo de segundas</b> con tu total y si coincide con lo suyo.</li><li>Con <b>Usar conteo</b> pasa tus números a sus casillas de Segundas, Cera y Desecho, y las puede corregir.</li><li><b>La casilla del jefe es la que cuenta</b> para el pago y los reportes.</li><li>El jefe de empaque y administración pueden abrir esta misma pantalla desde el Código de Trazabilidad, ver todos los conteos del viaje y anular uno si quedó mal. La anulación queda con su nombre.</li><li>Las segundas de cada productor quedan disponibles para la hoja 1.5 o para Juguera.</li></ol>'},
{t:'Restricciones',h:
'<ul><li>Solo cuentas en viajes abiertos; un viaje cerrado responde «El viaje … ya está cerrado».</li><li>El total no puede quedar abajo de cero.</li><li>Solo anulas tus propios conteos.</li><li>No cambias el registro del jefe ni las tarimas.</li><li>Todo queda registrado con tu nombre y la hora.</li></ul>'},
{t:'Implicaciones y errores comunes',h:
'<table class="mtbl"><thead><tr><th>Error</th><th>Consecuencia</th><th>Cómo se evita</th></tr></thead><tbody>'+
'<tr><td>Productor equivocado</td><td>Se paga la segunda a otra persona</td><td>Elige por código y confirma el nombre.</td></tr>'+
'<tr><td>Contar cera como segunda normal</td><td>Lote de cera sin registrar</td><td>Usa el botón Con cera.</td></tr>'+
'<tr><td>No contar terceras</td><td>El desecho del productor sale bajo</td><td>Cuenta los tres tipos.</td></tr>'+
'<tr><td>Guardar dos veces lo mismo</td><td>Conteo duplicado</td><td>Revisa Mis conteos y anula el repetido.</td></tr></tbody></table>'},
{t:'Responsabilidad y uso legal',h:legal('Conteo de segundas')}
]};

var ORD=['jefe','inocuidad','tarima','segundas'];
ORD.forEach(function(k){M[k].titulo='Manual de usuario · '+M[k].puesto;M[k].sub='Alcance, importancia y restricciones para usuarios de la App de Vera Limes / Citrilar';});
window.MANUALES={ver:VER,vig:VIG,ord:ORD,M:M};
})();
