# ucList para twinBASIC (32/64 bits)

Reescritura desde cero del control **ucList** (y su clase **ClsShadow**) de ucListPlus para twinBASIC.
La interfaz pública es la del original: los mismos eventos, propiedades, métodos, enumeraciones y nombres en el PropertyBag. Solo cambian a `LongPtr` los parámetros que son handles (`hWnd`, `hDC`, `hBmp` y `oTextBox`), para que funcione en 64 bits.

## Archivos (`Sources/`)

| Archivo | Tipo | Contenido |
|---|---|---|
| `ucList.twin` + `ucList.tbcontrol` | UserControl (`Class ucList … End Class`) | El control. El `.tbcontrol` es el JSON del diseñador; el control no tiene controles hijos de diseño. |
| `ClsShadow.twin` | Clase (`Class ClsShadow … End Class`) | Sombra y borde redondeado. |
| `Forms/FrmTest.twin` + `FrmTest.tbform` | Formulario | Banco de pruebas (ver más abajo). |
| `mUcListHelpers.twin` | Módulo (`Module mUcListHelpers … End Module`) | Subclasificación (SetWindowSubclass), clase de ventana propia, GDI+, ganchos de teclado y ratón. |

Todos los archivos usan saltos de línea CRLF (`\r\n`), igual que los que genera twinBASIC.

Uso: añade los cuatro archivos al proyecto (o arrastra la carpeta `Sources`) y coloca ucList en el formulario.
Las propiedades de diseño (ListStyle, ItemHeight, colores, etc.) se configuran en la ventana de propiedades, como en VB6.

## Formulario de prueba (`Sources/Forms/FrmTest.twin` + `FrmTest.tbform`)

Banco de pruebas que ejercita **todas** las propiedades, métodos y eventos de ucList (y la clase ClsShadow a través de `.Shadow`):

| Control | Para qué |
|---|---|
| `ucList1` | ListBox con 30 elementos, iconos, grupos, separador, atajo, colores y un elemento deshabilitado. Es el control que se prueba. Clic derecho: menú. |
| `ucProps` | ListBox `MultiSelect`: cada casilla activa o desactiva una propiedad de `ucList1` o de `ucDrop`. |
| `ucActions` | ListBox con filtro (`ShowFilter`, EDIT): cada elemento llama a un método. Clic o Intro para ejecutarlo; `txtParam` sirve de parámetro. Las consultas escriben en el registro. |
| `cmdDrop` + `ucDrop` | DropDown con filtro, abierto como botón conmutador. |
| `cmdMenu` + `ucMenu`/`ucSub`/`ucSub2` | Menú con atajos, aceleradores `&`, elemento deshabilitado y dos niveles de submenú. |
| `txtAuto1`/`txtAuto2` + `ucAuto` | Autocompletar: un mismo ucList para dos TextBox. Sirve para probar el caso de Intro y luego Tab. |
| `txtLog` | Registro de todos los eventos. |

Cómo montarlo:
1. Crea un proyecto *Standard EXE* llamado **`ucListPlus`**. En el `.tbform` el control se identifica como `ucListPlus.ucList`; si tu proyecto tiene otro nombre, cambia ese prefijo con buscar y reemplazar.
2. Importa los archivos de `Sources` y de `Sources/Forms`.
3. En *Project → Settings*, elige **`FrmTest`** como objeto de inicio.

twinBASIC guarda las propiedades de los UserControl en el `.tbform` como un bloque binario que no se puede generar a mano. Por eso el `.tbform` solo guarda la posición y el tamaño de cada ucList, y `Form_Load` asigna `ListStyle` y el resto de propiedades por código.
En diseño, el ucList ya no fuerza el tamaño de 32×32, para que el formulario mantenga su geometría. Los iconos son PNG incrustados en base64, así que no hacen falta archivos externos.

## Arquitectura

* **La lista vive en una ventana propia** (clase `ucListPlusWnd`), que se pinta con doble buffer:
  * `ListStyle = StyleListBox`: ventana hija que ocupa todo el UserControl.
  * `StyleDropDown` / `StyleMenu`: ventana emergente `WS_EX_NOACTIVATE`, propiedad del formulario. En tiempo de ejecución el UserControl queda oculto, como en el original.
* **Filtro (`ShowFilter`)**: un control `EDIT` de Win32 creado por código, hijo de la ventana de la lista. Con `EN_CHANGE` se filtra y el texto encontrado se resalta con `FindTextColor`.
* **Barra de desplazamiento** interna. En `GoogleDesign` es un pulgar fino redondeado; en `MicrosoftDesign` es una barra con tema y botones. Responde a la rueda, al arrastre del pulgar, a los clics en la pista y a la repetición al mantener pulsado.
* **Sombra**: en ListBox, `ClsShadow` usa una ventana hija recortada por región. En las emergentes usa una ventana *layered* con `UpdateLayeredWindow`. La sombra sigue a su ventana cuando se mueve, se oculta o se destruye.

## Modelo de foco

* **Listas emergentes (DropDown/Menu)**: nunca toman el foco.
  * El control que lo tenía lo conserva: el botón, o el TextBox de `AutoCompleteTextBox`.
  * Mientras hay una lista abierta, un gancho de teclado *del hilo* hace lo que hacía `RegisterHotKey` en el original, sin afectar a otras aplicaciones:
    * Las teclas de navegación (↑ ↓ Inicio Fin RePág AvPág Intro Esc ← →) van a la lista.
    * Tab cierra la lista y **sigue su camino**: el foco pasa al siguiente control.
    * Con filtro, los caracteres van al EDIT, y también Supr, Ctrl+A/C/V/X.
    * Sin filtro, los caracteres sirven para buscar por la primera letra o para activar el acelerador `&` en los menús.
    * Con `AutoCompleteTextBox`, el resto de teclas siguen yendo al TextBox.
  * Un gancho de ratón detecta los clics fuera de las listas abiertas. El cierre se **publica**, así que el clic llega antes a su destino.
  * La lista también se cierra cuando:
    * el formulario se desactiva, se mueve o cambia de tamaño;
    * el TextBox de autocompletar pierde el foco.
* **ListBox**:
  * El foco del UserControl se pasa a la ventana de la lista, o al EDIT si `ShowFilter = True`.
  * Tab sale del control con normalidad.
  * Al tener el foco, la sombra y el borde toman `SelectionColor`; al perderlo vuelven a `BorderColor`.
  * Desde el filtro, ↑ ↓ RePág AvPág e Intro mueven la selección de la lista.
* Autocompletar: cuando se elige un elemento (Intro o clic), el texto se escribe en el TextBox con `WM_SETTEXT` y se lanza `PostAutoComplete`.
  * El texto asignado por código no vuelve a abrir la lista.
  * Pulsar Tab después **no** reabre la lista en el TextBox anterior.

## Eventos

`IndexChange`, `ItemClick`, `HideList`, `ScrollBottom`, `PostAutoComplete`, `DblClick`, `MouseDown/Move/Up`, `PostPaint(hdc)`, `ItemCustomDraw(...)`, `KeyPress`, `KeyUp`, `KeyDown`, `OLECompleteDrag`, `OLEDragDrop`, `OLEDragOver`, `OLEGiveFeedback`, `OLESetData` y `OLEStartDrag`.

Coordenadas de los eventos de ratón: píxeles de la ventana de la lista, que son los mismos que espera `ItemHitest`.

## Honestidad sobre la verificación

En el entorno donde se escribió este código **no hay compilador de twinBASIC ni Windows**, así que no se ha podido compilar ni ejecutar. Se comprobó de forma automática:

* el equilibrio de bloques (`If/For/Do/Select/With/Sub/Function/Property/Type/Enum`);
* la cobertura de la API pública frente a `ucList.ctl` y `ClsShadow.cls` (no falta ningún miembro);
* que las firmas coinciden con el original (solo cambian a `LongPtr` los handles);
* que no hay identificadores sin declarar;
* el envoltorio `Class/Module … End Class/Module`;
* los saltos de línea CRLF y la validez del JSON del `.tbcontrol`.

Además se repasaron a mano los escenarios de foco, teclado y ratón descritos arriba.
Si al compilar aparece algún error, indícame la línea y el mensaje.
