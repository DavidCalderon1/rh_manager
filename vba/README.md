# Instalación de las macros actualizadas

Estos archivos van en `Formato Carrera de drones.xlsm`. Ninguno se puede "pegar" directo en el
archivo -- se importan desde el editor de VBA.

## 1. Reemplazar RenderRaceTables.bas y ManageRace.bas, agregar OfflineData.bas

1. Abre el Excel, `Alt+F11` para abrir el editor de VBA.
2. Busca el módulo **RenderRaceTables**. Clic derecho → **Remove RenderRaceTables...** → cuando
   pregunte si quieres exportarlo antes de quitarlo, dile que no. `File > Import File...` →
   `RenderRaceTables.bas`.
3. Repite lo mismo para **ManageRace** (se redujo bastante: ahora solo tiene `EditHeatPilotsById`,
   que abre `frmManageHeat` -- ver sección 4 más abajo, es un paso adicional. `EditHeatPilots`,
   `RemixClass`, `DeleteClass` y `DeleteAllHeats` se eliminaron, ver sección 9).
4. **Nuevo módulo `OfflineData.bas`** -- este no reemplaza nada existente, es la primera vez que se
   importa. `File > Import File...` → `OfflineData.bas`. Ver sección 15 para qué hace (modo sin
   conexión).
5. Guarda (`Ctrl+S`, mantener como `.xlsm`).

**Nota sobre caracteres:** todos los `.bas` están guardados en ASCII puro (sin tildes ni símbolos
especiales) a propósito -- esta clase de archivos causó errores de compilación confusos
("variable no definida") cuando tenían caracteres fuera de ese rango. Si en algún momento editas
estos archivos a mano, evita tildes/eñes/símbolos raros dentro del código (los `MsgBox`/textos
sí pueden llevarlos sin problema, es solo el código en sí el que conviene mantener en ASCII).

## 2. Backend: una sola clase por generación + agregar a clase existente

Sin nada que instalar de tu lado -- ya está corriendo en el contenedor. Cambios:

- **"Generar Heats" con varios grupos (G1, G2, G3...) ahora crea UNA sola clase**, con todos los
  heats de todos los grupos adentro (antes creaba una clase separada por cada grupo).
- **Nuevo: agregar heats a una clase que ya existe.** Ver sección 3 -- requiere actualizar
  `frmCreateHeats` con un checkbox nuevo.

## 3. frmCreateHeats: opción "agregar a clase existente"

Este formulario YA EXISTE en tu Excel -- no se reimporta, se actualiza a mano porque el diseño
visual (posiciones de los controles) vive en un archivo binario que no puedo generar yo. Sigue
las instrucciones completas en **`frmCreateHeats_updates.txt`** (en esta misma carpeta): agregar
un CheckBox + ComboBox nuevos, y reemplazar/agregar 4 procedimientos de código (te doy el texto
completo, solo hay que pegarlo).

Resultado: al generar heats, puedes marcar "Agregar a clase existente", elegir la clase de una
lista, y los grupos nuevos se agregan ahí (siguiendo la numeración, p.ej. si ya tiene G1-G3 el
nuevo llega como G4) en vez de crear una clase aparte.

## 4. Nuevo formulario: frmManageHeat (edición de heats más amigable)

Reemplaza la cadena de `InputBox` que tenía el botón **E** de cada heat por un formulario real
con listas. Sigue las instrucciones completas en **`frmManageHeat_instructions.txt`**: crear el
formulario con sus controles (10 minutos) y pegar el código que te doy ahí. `ManageRace.bas` (ya
reimportado en el paso 1) queda listo para abrirlo en cuanto exista.

## 5. Recordatorio de rondas anteriores -- qué hay en la hoja "Carrera"

- Columna A reservada para botones fijos: Generar Heats, Actualizar Todo, Cargar Todo RH. El grid
  arranca en la columna B. ("Editar Piloto Heat", "Eliminar Clase" y "Borrar Todos Heats" se
  quitaron -- ver sección 9: los botones por heat/grupo ya cubren lo mismo sin pedir un id.)
- Cada grupo tiene tres botones en su encabezado: **R** (refresca), **MX** (remix con el
  generador nativo "Balanced Random Fill"), **X** (elimina el grupo por completo).
- Cada heat tiene **R** (refresca), **E** (editar pilotos, ahora abre `frmManageHeat`) y **X**
  (elimina ese heat individual -- ver sección 7).
- Columna **Fr** con la frecuencia real asignada por RotorHazard.
- Columnas de ronda dinámicas según `win_condition` de la clase (tiempo / puntos / posición).
- Bloque **"Posiciones"** debajo de cada grupo con la clasificación general de la clase (columna
  del piloto y de la posición ya intercambiadas esta ronda -- antes el nombre quedaba en el
  espacio angosto de "Fr" y se veía cortado).
- Refresco quirúrgico (**R**) vía las hojas ocultas `RaceMeta`/`RaceMetaGroups`; **MX**/**X**
  redibujan porque cambian la estructura. El **R** de grupo ahora también actualiza el bloque
  "Posiciones" (ver sección 6).

## 6. Esta ronda: 9 arreglos

1. **Error "argumento o llamada a procedimiento no válida" al generar heats.** Pasaba si no
   seleccionabas ninguna frecuencia (o ningún piloto) en las listas de `frmCreateHeats`: el texto
   armado quedaba vacío y `Left(texto, Len(texto) - 1)` truena con longitud -1. Ahora valida antes
   y te avisa con un mensaje claro ("Selecciona al menos una frecuencia/piloto") en vez de tronar.
   Ya está en `frmCreateHeats_updates.txt` -- si ya habías pegado la versión anterior de
   `CommandButtonGenerate_Click`, reemplázala por la nueva.
2. **`ComboBoxExistingClass` mostraba el id en vez del nombre de la clase.** El combo tiene 2
   columnas (id oculto, nombre visible) pero faltaba decirle que ocultara la columna del id.
   `LoadExistingClasses` ahora fija `ColumnWidths = "0 pt;120 pt"` y `BoundColumn = 1`, así que se
   ve el nombre pero `.Value` sigue devolviendo el id (lo que necesita el payload). Mismo archivo,
   reemplaza `LoadExistingClasses` por la nueva versión.
3. **El botón "R" de cada grupo no actualizaba "Posiciones".** Antes solo tocaba las celdas de
   heats/pilotos; el bloque de posiciones se quedaba desactualizado hasta el próximo redibujado
   completo. `RenderRaceTables.bas` ahora guarda en la hoja oculta `RaceMetaGroups` en qué filas
   vive el bloque de posiciones de cada clase, y `RefreshGroupCells` lo reescribe con los datos
   `standings` que ya venían en la misma respuesta de la API (sin pedir nada extra al servidor).
   **Requiere reimportar `RenderRaceTables.bas`** (paso 1 de este README).
4. **"Generar Heats" borraba las tablas de las demás clases, o redibujaba toda la hoja de más.**
   Al generar (o agregar heats a una clase existente) mientras había otras clases ya dibujadas en
   "Carrera", el redibujado solo conocía la clase recién creada/actualizada -- el resto
   desaparecía. Primero se arregló trayendo de nuevo TODAS las demás clases y redibujando la hoja
   completa con todo junto (ningún dato se perdía, pero seguía siendo un redibujado completo cada
   vez que agregabas un heat a una clase existente). Ahora es quirúrgico de verdad:
   `CommandButtonGenerate_Click` llama a la nueva rutina `RenderSingleClassInPlace` (en
   `RenderRaceTables.bas`), que solo toca la franja de columnas de la clase afectada -- borra y
   vuelve a dibujar SOLO ese bloque (en el mismo lugar si ya existía, o al final si es una clase
   nueva) sin tocar ni un pixel de las demás clases. (Si marcaste "Borrar todo" sí se hace un
   redibujado completo, porque en ese caso es lo correcto: todas las clases anteriores
   efectivamente dejaron de existir en RotorHazard). **Requiere reimportar `RenderRaceTables.bas`
   Y reemplazar el bloque final de `CommandButtonGenerate_Click`** (texto actualizado en
   `frmCreateHeats_updates.txt`).
5. **Editar pilotos de un heat (o editarlo directo en RotorHazard) y refrescar no mostraba el
   cambio.** El refresco quirúrgico (`ApplyNodeResultRow`) solo reescribía la frecuencia y las
   columnas de ronda -- **nunca tocaba la celda del nombre del piloto**. Si cambiabas qué piloto
   ocupa un slot ya ocupado, el nombre viejo se quedaba pegado en pantalla aunque el resto de datos
   sí se actualizara. Ahora también reescribe esa celda. Ver también el punto 7 más abajo (cómo se
   maneja un slot que pasa de vacío a ocupado, o viceversa). **Requiere reimportar
   `RenderRaceTables.bas`.**
6. **Botones de heat/grupo (R/E/MX/X) sin explicación de qué hacen.** Ahora, al pasar el mouse
   por encima de cualquiera de esos botones (o de los botones fijos de la columna A), aparece un
   tooltip explicando su función ("Recargar este heat", "Editar los pilotos de este heat",
   "Remix: reordenar pilotos con el generador nativo...", etc.), vía `ActionSettings(...).ScreenTip`.
   No se cambiaron los símbolos/letras visibles de los botones (R/E/MX/X) para no complicar el
   renderizado en botones tan chicos -- si prefieres iconos en vez de o además del tooltip, dilo y
   los agregamos. **Requiere reimportar `RenderRaceTables.bas`.**
7. **"Render error: uso no válido de null" + slots vacíos mostrando fila "(vacio)".** Dos cosas
   relacionadas, ambas en `RenderRaceTables.bas`:
   - Causa del error: `IsNull(pilotID) Or CStr(pilotID) = ""` -- en VBA el operador `Or` NO hace
     short-circuit (evalúa los dos lados siempre), así que `CStr(pilotID)` se ejecutaba igual
     aunque `pilotID` fuera `Null`, y eso truena con "uso no válido de null". Se reemplazó por una
     función `HasPilotValue` con `If/ElseIf` real.
   - De paso, se volvió a la idea original de NO dibujar fila para los slots vacíos (nada de
     "(vacio)" en la tabla) -- un heat con 8 nodos y solo 4 pilotos ahora solo muestra 4 filas, no
     8. La excepción es cuando un slot CAMBIA de vacío a ocupado (o al revés): como eso significa
     que hay que agregar o quitar una fila -- y eso no se puede hacer con un ajuste quirúrgico de
     celdas -- `RefreshGroupCells`/`RefreshHeatCells` ahora lo detectan solas y hacen
     automáticamente un redibujado completo de esa clase, avisando con un `MsgBox`.
   **Requiere reimportar `RenderRaceTables.bas`.**
8. **Los tiempos no se veían con el formato correcto.** En `RenderHeat`, la celda de cada ronda
   recibía `.Value = "0:45.243"` ANTES de fijar `.NumberFormat = "@"` (texto). Excel reconoce ese
   texto como una hora y lo convierte a un número de serie de hora al asignarlo sobre una celda en
   formato General -- y una vez hecho eso, poner el formato de texto después ya no recupera el
   texto original, solo muestra el número crudo. Se invirtió el orden (`NumberFormat` primero,
   `Value` después), que es la forma correcta en VBA para este tipo de valores tipo-fecha/hora.
   **Requiere reimportar `RenderRaceTables.bas`.**
9. **Metadatos del bloque "Posiciones" poco legibles.** El detalle extra de cada fila (`heat`,
   `heat_rank`, etc.) se veía pegado, sin comas: `heat=Practica G3-H1  heat_rank=4`. Ahora lleva
   espacios alrededor del `=` y una coma entre pares: `heat = Practica G3-H1 , heat_rank = 4`.
   **Requiere reimportar `RenderRaceTables.bas`.**

## 7. Nueva ronda: elegir grupos a remixar + eliminar heat individual

**Backend** (sin nada que instalar de tu lado, ya está en el contenedor):

- **`MX` ahora remixa por grupo, no toda la clase.** Antes, remixar CUALQUIER heat de una clase
  borraba y regeneraba TODOS los heats de la clase completa, mezclando pilotos entre grupos y
  perdiendo la separación G1/G2/G3. Ahora `remix_raceclass_heats` remixa cada grupo por
  separado: si le pasas `group_ids`, solo esos grupos se borran y regeneran (con sus propios
  pilotos, sin mezclarse con los de otro grupo); los demás grupos de la clase quedan
  completamente intactos (mismos heats, mismos ids, mismos pilotos). Verificado contra el
  contenedor: remixar el grupo 1 de una clase con G1 y G2 dejó G2 con los mismos heat ids y
  pilotos de antes, mientras G1 recibió heats nuevos con los pilotos reordenados.
- **Nuevo endpoint** `GET /api/rhm/raceclass/<id>/groups` -- lista los grupos (G1, G2, ...) de
  una clase con sus heats, para poder elegir cuáles remixar.
- **Nuevo endpoint** `DELETE /api/rhm/heat/<id>` -- elimina un heat individual (antes solo se
  podía borrar TODOS los heats o la clase completa).

**VBA** (`RenderRaceTables.bas`, hay que reimportarlo):

- El botón **MX** de cada grupo ahora primero consulta los grupos de esa clase. Si la clase tiene
  un solo grupo, remixa directo (como antes). Si tiene varios, te muestra una lista
  (`InputBox`) con cada grupo y su cantidad de heats, y puedes escribir los números de los que
  quieres remixar separados por coma (ej: `1,3`), o dejarlo vacío para remixar todos. Después de
  remixar, solo se redibuja esa clase (quirúrgico, no toca las demás).
- Cada heat ahora tiene un tercer botón **X** (además de **R** y **E**) para eliminarlo
  individualmente. Como eliminar un heat cambia cuántas filas se ven, se redibuja esa clase
  completa (quirúrgico -- las demás clases no se tocan), no solo esas celdas.

## 8. Nueva ronda: eliminar clase/heat sin errores ni falsos "ok"

**Backend** (ya está en el contenedor):

- **Eliminar una clase generaba errores en RotorHazard.** Confirmado en logs: al borrar una
  clase, `DELETE FROM race_class` a veces tronaba con `FOREIGN KEY constraint failed` -- incluso
  en una clase de prueba recién creada, sin ninguna carrera corrida. Causa: las filas de
  `RaceClassAttribute` de esa clase no siempre se borran de la base de datos ANTES que la fila de
  `race_class` dentro del mismo commit de SQLAlchemy (no hay una relación de ORM declarada entre
  esas dos tablas que le diga a SQLAlchemy en qué orden hacerlo). `delete_class_safe` (nuevo, en
  `manager_uc.py`) borra esas filas de atributo en su propio commit ANTES de pedirle a RotorHazard
  que borre la clase, evitando el problema de raíz. Verificado en vivo: se creó una clase de
  prueba y se eliminó sin ningún error en los logs (antes sí aparecía).
- **Eliminar una clase/heat con una carrera guardada fallaba en silencio.** RotorHazard protege
  el historial de carreras: si un heat o una clase ya tiene una carrera guardada, rechaza
  eliminarlo (sin lanzar una excepción, solo devuelve "no se pudo" internamente) -- pero nuestras
  rutas de `DELETE` no revisaban ese resultado y respondían `{"status":"ok"}` igual, aunque en
  realidad no se hubiera borrado nada. Ahora las rutas de eliminar heat/clase/todos-los-heats
  revisan el resultado real y devuelven un error claro (con el nombre del heat problemático) en
  vez de un falso "ok". Lo mismo se aplicó al remix: si un heat de un grupo tiene una carrera
  guardada, ahora se detecta ANTES de borrar nada de ese grupo (para no terminar con heats
  duplicados a medio hacer).

**VBA** (`RenderRaceTables.bas`, hay que reimportarlo):

- **La alerta de eliminar un heat solo mostraba el id, no el nombre.** `DeleteHeatAndRedraw` ahora
  consulta el heat antes de preguntar, y la confirmación muestra su nombre real (ej. "TestRemix
  G1-H1") en vez de solo el número.

## 9. Nueva ronda: quitar botones de columna A que ya no hacían falta

**VBA** (`RenderRaceTables.bas` y `ManageRace.bas`, hay que reimportar ambos):

- Se quitaron los botones fijos **"Editar Piloto Heat"**, **"Eliminar Clase"** y **"Borrar Todos
  Heats"** de la columna A -- pedían el id de la clase/heat por `InputBox`, y ya existen botones
  dedicados que hacen lo mismo directamente sobre el heat/grupo que estás viendo: **E** por heat
  (editar pilotos), **X** por heat (eliminar ese heat), **X** por grupo (eliminar la clase
  completa). Solo quedan **Generar Heats**, **Actualizar Todo** y **Cargar Todo RH** en la
  columna A (los tres cubren algo que ningún botón de heat/grupo puede hacer: crear/agregar
  heats, o recargar/traer clases enteras).
- Junto con los botones, se eliminó la lógica que ya no se usaba en `ManageRace.bas`:
  `EditHeatPilots` (la cadena de `InputBox` que terminaba abriendo `frmManageHeat`), `DeleteClass`,
  `DeleteAllHeats`, y `RemixClass` (esta última ya estaba huérfana -- nunca tuvo un botón asignado
  desde que se agregó el **MX** por grupo). `ManageRace.bas` quedó reducido a solo
  `EditHeatPilotsById`, que es lo que el botón **E** de cada heat sigue usando.

## 10. Nueva ronda: formulario real para elegir que grupos remixar

El botón **MX** de cada grupo pedía los números de los grupos a remixar escribiéndolos en un
`InputBox` (separados por coma). Ahora, si la clase tiene más de un grupo, se abre un formulario
real (`frmRemixGroups`) con una lista donde eliges los grupos con un clic (sin necesitar Ctrl ni
Shift), un botón "Seleccionar todos", y "Remixar"/"Cancelar". Si la clase tiene un solo grupo,
sigue remixando directo -- no tiene sentido preguntar cuando no hay nada que elegir.

Este formulario YA NO EXISTE en tu Excel -- hay que crearlo a mano. Sigue las instrucciones
completas en **`frmRemixGroups_instructions.txt`** (en esta misma carpeta): crear el formulario
con sus controles (10 minutos) y pegar el código que te doy ahí. `RenderRaceTables.bas` (ya
reimportado en el paso 1) queda listo para abrirlo en cuanto exista -- mientras tanto (si no has
creado el formulario todavía), el botón **MX** en una clase con varios grupos va a fallar con
"no se puede encontrar el proyecto o la biblioteca" hasta que lo crees.

**Requiere reimportar `RenderRaceTables.bas` Y crear `frmRemixGroups`** (sección 10 de este README).

## 11. Nueva ronda: "No se pudo conectar con RotorHazard" al recargar tras un remix

Este mensaje engañoso salía al darle **R** a un grupo/heat después de remixar una clase cuando NO
tienes las 8 frecuencias configuradas (por ejemplo, solo 4 canales disponibles) -- RotorHazard deja
algún slot sin nodo/frecuencia asignada (`node_index = null`), pendiente de que confirmes
manualmente con "Seed Now" en su propia interfaz. Verificado en vivo contra un heat real con
exactamente ese estado.

- **Causa real:** en varios lugares de `RenderRaceTables.bas` se hacía `CLng(nd("node_index"))` o
  `CStr(nd("node_index"))` sin revisar si venía `Null` primero. En VBA, `CLng(Null)`/`CStr(Null)`
  truenan con "uso no válido de null" -- ese error, atrapado por el manejador genérico de la
  rutina, es justo lo que se mostraba como "No se pudo conectar con RotorHazard" (el mensaje no
  tiene nada que ver con la conexión real).
- **Fix VBA:** nueva función `HasNodeIndex` que revisa `Null` seguro (sin el problema de
  short-circuit de `Or` de rondas anteriores). Un piloto sembrado pero sin nodo asignado todavía
  SE MUESTRA en la tabla (no se oculta en silencio) -- la columna **Fr** muestra `?` con un
  comentario ("Sin nodo/frecuencia asignada -- resuelve 'Seed Now' en RotorHazard, luego
  'Actualizar Todo'") en vez de tronar. Esa fila no queda registrada para refresco quirúrgico
  (no hay forma estable de identificarla hasta que RotorHazard le asigne un nodo real), así que
  usa **Actualizar Todo** para verla actualizada una vez resuelvas "Seed Now" allá.
- **Fix backend** (`manager_uc.py`): de paso, se encontró que si un heat tiene MÁS DE UN slot sin
  nodo asignado, el código armaba un diccionario indexado por `node_index` -- y como `None` no es
  una clave única, se perdían todos menos el último. Verificado en vivo: un heat con 4 slots sin
  nodo (`node_index: null`) ahora los conserva los 4, antes solo hubiera mostrado 1.

**Requiere reimportar `RenderRaceTables.bas`.** El cambio de `manager_uc.py` ya está corriendo en
el contenedor (reiniciado y verificado).

## 12. Nueva ronda: el redibujado tras "Seed Now" tocaba TODA la hoja, no solo la clase

Al confirmar las bandas en el modal de "Heat Plan" de RotorHazard (resolviendo el "Seed Now" del
punto anterior) y luego recargar ese heat o grupo en Excel, salía el aviso de "se hace un
redibujado completo" -- pero ese redibujado usaba `RefreshResults`, que vuelve a traer y redibujar
**todas** las clases de la hoja, no solo la que cambió. Ahora usa `RenderSingleClassInPlace`
(la misma rutina quirúrgica de rondas anteriores) para tocar solo la clase afectada:

- `RefreshHeatCells` (botón **R** de un heat): ya sabe el `class_id` del heat (viene en la misma
  respuesta que ya trae), así que redibuja solo esa clase.
- `RefreshGroupCells` (botón **R** de un grupo): ya tenía los datos frescos de esa clase en
  memoria (los mismos que usa para las celdas y "Posiciones"), así que los reutiliza para
  redibujar en el sitio, sin pedir nada extra al servidor.

**Requiere reimportar `RenderRaceTables.bas`.**

## 13. Nueva ronda: modo sin conexión (`Config!E14`)

El Excel ahora puede manejarse **sin conexión a RotorHazard** -- `Config!E14` (leído por la función
`ThisWorkbook.WorkOffline()`, que ya existía y ya se usaba como guarda en todos lados) sigue siendo
la celda que decide el modo. La diferencia es que ahora, cuando está en `TRUE`, **todos los botones
siguen funcionando** en vez de simplemente no hacer nada: cada uno usa una lógica local en vez de
llamar a RotorHazard.

### La idea central

`JsonConverter.ParseJson` (la librería que ya se usa en todo el proyecto) devuelve los objetos JSON
como `Scripting.Dictionary`/`Collection` normales -- tipos de VBA, no algo propio de la librería. El
nuevo módulo **`OfflineData.bas`** arma esos mismos árboles de datos A MANO (sin ningún texto JSON
de por medio) y se los pasa tal cual a `RenderGroup`/`RenderSingleClassInPlace`/`RenderStandings` --
la misma canalización de dibujo de siempre. Nada se dibuja "distinto" sin conexión, solo cambia de
dónde sale el dato de entrada.

### Qué se ingresa a mano y qué se calcula solo

Las celdas R1..Rn de cada heat (tiempo, puntos o posición según el `win_condition` de la clase) ya
eran valores planos editables -- eso no cambió. Lo que sí es nuevo: al darle **R** a un heat o a un
grupo (o **Actualizar Todo**) sin conexión, en vez de tronar o no hacer nada, el Excel **lee lo que
escribiste en esas celdas y recalcula el bloque "Posiciones"** solo, con los mismos 3 métodos que ya
existían:

- `Cumulative_Points` -- suma los puntos del piloto en todas sus rondas, orden descendente.
- `Last_Heat_Position` -- toma la última posición registrada del piloto, orden ascendente.
- `Laps_Time__Best_X_Rounds` -- suma las mejores rondas del piloto (cuántas: celda nueva
  **`Config!L7`**, "Rondas a contar", default 3 si la dejas vacía), orden ascendente por tiempo.

**Importante:** esto es una aproximación fiel, no una copia exacta del motor de RotorHazard (que
tiene más configuración/casos especiales de los que una hoja llenada a mano puede expresar) --
suficiente para datos ingresados manualmente, pero no esperes que coincida al centésimo con lo que
RotorHazard mostraría si estuviera corriendo la misma carrera.

### Qué hace cada botón sin conexión

| Botón | Con conexión | Sin conexión |
|---|---|---|
| **Generar Heats** | Crea la clase/heats en RotorHazard | Mezcla los pilotos localmente (mismo algoritmo que el backend, portado a VBA) y dibuja la tabla -- sin pedir frecuencias (la columna Fr queda con etiquetas genéricas "N1", "N2"...) |
| **R** (heat/grupo) | Trae resultados frescos del servidor | Relee R1..Rn de la hoja, recalcula "Posiciones", redibuja esa clase |
| **Actualizar Todo** | Igual que R pero para todas las clases | Igual, recorriendo todas las clases ya dibujadas |
| **Cargar Todo RH** | Trae TODAS las clases de RotorHazard | No tiene equivalente -- muestra un aviso, no hace nada |
| **E** (editar heat) | Cambia el piloto vía API | Cambia el piloto escribiendo directo en la celda, recalcula y redibuja |
| **MX** (remix) | Reordena vía el generador nativo de RotorHazard | Reordena localmente (mismo algoritmo), limpia los tiempos de esos heats, redibuja |
| **X** (heat) | Borra el heat en RotorHazard | Quita ese heat de la tabla y recalcula "Posiciones" |
| **X** (grupo/clase) | Borra la clase en RotorHazard | Limpia esa franja de columnas de la hoja |

### Limitaciones que conviene conocer

- **Los pilotos y las frecuencias vienen de `Config!A:C`**, no de una API -- si un piloto no está
  ahí, no puedes elegirlo sin conexión. Las frecuencias no se piden en absoluto (no representan
  nada real sin hardware conectado).
- **IDs locales**: las clases/heats creados sin conexión reciben ids **negativos** (RotorHazard
  siempre usa positivos), así nunca chocan -- pero solo tienen sentido dentro de este Excel, no
  existen en ningún RotorHazard.
- **Un slot vacío no se puede llenar por primera vez desde `frmManageHeat` sin conexión.** Como los
  slots vacíos no se dibujan ni se registran (ver sección 7 de una ronda anterior), tampoco se
  "cosechan" de vuelta sin conexión -- solo se puede reasignar/quitar un piloto de un slot que YA
  estaba ocupado. Para llenar un slot vacío, usa **Generar Heats** o **MX** (remix), que sí arman
  la tabla desde cero.
- El bloque "Posiciones" sin conexión dice "(local)" en el nombre del método (ej. "Cumulative
  Points (local)") para dejar claro que es el cálculo local, no el de RotorHazard.

### Archivos que cambiaron

- **`OfflineData.bas`** (nuevo) -- toda la lógica local: `ShuffleHeatsLocal` (puerto de
  `generate_heats_logic`), `HarvestClassFromSheet` (relee la hoja), `ComputeStandingsLocal` (los 3
  métodos), `GenerateHeatsLocal`/`PerformRemixLocal`/`DeleteHeatLocal` (las acciones), y varios
  helpers de pilotos/clases locales.
- **`RenderRaceTables.bas`** -- cada rutina que antes hacía `If WorkOffline() Then Exit Sub` ahora
  llama a lo de arriba en su lugar.
- **`frmCreateHeats_updates.txt`** -- `CommandButtonGenerate_Click` ahora arma la clase localmente
  sin conexión; `LoadExistingClasses` lista las clases ya dibujadas en la hoja; nuevo
  `LoadPilotsOffline` (una línea nueva en tu `UserForm_Initialize`).
- **`frmManageHeat_instructions.txt`** -- `LoadHeat`/`LoadPilotChoices`/`CommandButtonSave_Click`
  ahora funcionan sin conexión.
- **`frmRemixGroups_instructions.txt`** -- `LoadClass` ahora lista los grupos sin conexión (parsea
  "G\<n\>" de los nombres de heat ya en la hoja).

## 14. Prueba end-to-end sugerida (esta ronda)

1. Abre "Generar Heats" sin seleccionar ninguna frecuencia o ningún piloto -- confirma que sale un
   `MsgBox` claro en vez del error de "argumento no válido".
2. Marca "Agregar a clase existente" -- confirma que el combo muestra los NOMBRES de las clases,
   no sus ids.
3. Corre una ronda, dale clic al botón **R** del grupo (no al de un heat individual) -- confirma
   que el bloque "Posiciones" se actualiza con los resultados nuevos sin que se vea un redibujado
   completo de la hoja.
4. Con al menos 2 clases ya dibujadas en "Carrera", genera heats de una tercera (o agrega un grupo
   a una de las existentes) -- confirma que las clases anteriores siguen ahí, no solo la nueva, y
   que la clase afectada no "salta" a otra posición ni desordena las demás.
5. Pasa el mouse (sin hacer clic) sobre los botones **R**/**E** de un heat y **R**/**MX**/**X** de
   un grupo, y sobre los botones de la columna A -- confirma que aparece un tooltip explicando qué
   hace cada uno.
6. Genera heats con más nodos que pilotos (p.ej. 8 nodos, 4 pilotos) -- confirma que NO sale ningún
   error de "uso no válido de null" y que la tabla solo muestra las filas de los pilotos asignados
   (sin filas "(vacio)").
7. Abre `frmManageHeat` (botón **E**) en un heat, cambia el piloto de un slot ya ocupado -- confirma
   que el nombre nuevo aparece al guardar (o con el botón **R** del heat/grupo).
8. Asigna un piloto a un slot que estaba vacío directamente en la interfaz de RotorHazard (no en
   Excel), luego dale **R** al heat o al grupo en Excel -- confirma que sale el aviso de
   redibujado completo y que el piloto nuevo aparece (con su propia fila nueva).
9. Corre una clase con `win_condition` de tiempo ("Laps/Time: Best X Rounds") -- confirma que la
   columna de tiempo muestra algo como "0:45.243" y NO un número raro tipo "0.031514...".
10. (Opcional) Crea un heat nuevo directo en RotorHazard dentro de una clase ya dibujada, dale **R**
    al grupo -- confirma que sale el aviso correspondiente y que el redibujado completo lo muestra.
11. Mira el detalle de una fila en "Posiciones" -- confirma que se ve como
    `heat = ... , heat_rank = ...` (con espacios y coma), no pegado.
12. En una clase con 2+ grupos (G1, G2...), dale **MX** -- confirma que se abre el formulario
    `frmRemixGroups` con la lista de grupos. Selecciona UNO solo y dale "Remixar" -- confirma que
    solo ese grupo cambió de pilotos (los heats del otro grupo siguen con los mismos pilotos e ids
    de heat), y que el formulario se cerró solo. En una clase con un solo grupo, confirma que
    **MX** remixa directo sin abrir ningún formulario.
13. Dale **X** a un heat individual -- confirma que se elimina y que la tabla de esa clase se
    redibuja sin él, sin afectar otras clases: la confirmación debe mostrar el NOMBRE del heat, no
    solo su id.
14. Crea una clase de prueba desechable y elimínala con el **X** de su grupo -- confirma que
    desaparece de "Cargar Todo RH" y que no queda huérfana (revisa los logs de RotorHazard si
    tienes acceso: ya no debe salir "FOREIGN KEY constraint failed").
15. Revisa la columna A -- confirma que solo quedan **Generar Heats**, **Actualizar Todo** y
    **Cargar Todo RH** (sin "Editar Piloto Heat", "Eliminar Clase" ni "Borrar Todos Heats").
16. Si tienes menos de 8 frecuencias configuradas: genera/remixa una clase, deja que RotorHazard
    pida "Seed Now" para algún piloto, y dale **R** al heat o al grupo en Excel -- confirma que YA
    NO sale "No se pudo conectar con RotorHazard", y que ese piloto aparece con "?" en la columna
    Fr (con un comentario si pasas el mouse encima) en vez de tronar o desaparecer.
17. Confirma las bandas en el modal de "Heat Plan" de RotorHazard (resolviendo un "Seed Now"),
    luego dale **R** a ese heat o grupo en Excel -- confirma que el aviso ahora dice "se redibuja
    esta clase" (no "redibujado completo"), y que las DEMÁS clases en la hoja no parpadean ni se
    redibujan.
18. Pon `Config!E14` en modo sin conexión. Genera una clase nueva -- confirma que arma la tabla
    localmente (pilotos mezclados, Fr con etiquetas genéricas), sin ningún error de red.
19. Sin conexión, escribe a mano tiempos/puntos/posiciones en R1..Rn de esa clase (según su
    `win_condition`) y dale **R** al grupo -- confirma que "Posiciones" se recalcula con el método
    correcto (agrega `Config!L7` si usas el método de tiempo, para probar que cuenta las X mejores
    rondas).
20. Sin conexión, prueba **E** (cambiar un piloto), **MX** (remixar), **X** de heat, y **X** de
    grupo -- confirma que los 4 funcionan sin ningún error de red y que "Posiciones" queda
    consistente después de cada uno.
21. Sin conexión, marca "Agregar a clase existente" en "Generar Heats" -- confirma que el combo
    muestra las clases ya dibujadas en la hoja, y que agregar heats a una de ellas no borra lo que
    ya tenía.
22. Vuelve `Config!E14` a modo con conexión y repite un par de acciones (generar, refrescar) --
    confirma que todo sigue funcionando igual que antes de este cambio.

## 15. Prueba end-to-end de rondas anteriores

1. Genera heats con 2+ grupos y confirma (por ejemplo con "Cargar Todo RH") que quedaron bajo
   **una sola clase**, no una por grupo.
2. Sigue `frmCreateHeats_updates.txt`, marca "Agregar a clase existente", elige esa misma clase,
   genera 1 grupo más -- confirma que llega como el siguiente número de grupo, sin crear una
   clase aparte ni pisar los grupos anteriores.
3. Sigue `frmManageHeat_instructions.txt`, prueba el botón **E** de un heat -- confirma que abre
   el formulario con los slots/pilotos reales, que guardar funciona, y que la celda en "Carrera"
   se actualiza sola sin cerrar el formulario.
4. Confirma que el bloque "Posiciones" ahora muestra el nombre del piloto en la columna ancha
   (no cortado) y la posición en la columna angosta de al lado.
