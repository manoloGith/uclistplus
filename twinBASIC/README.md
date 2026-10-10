# ucListPlus para twinBASIC

Conversión del control **ucList** (ucListPlus, de Leandro Ascierto) de VB6 a
[twinBASIC](https://twinbasic.com). Mantiene los mismos **eventos, propiedades y
métodos públicos** que el `ucList.ctl` original y compila tanto en **32 como en 64 bits**.

## Archivos

| Archivo | Descripción |
|---|---|
| `Sources/ucList.twin` | Código del UserControl `ucList` |
| `Sources/ucList.tbcontrol` | Diseñador del UserControl (vacío: no contiene controles hijos) |
| `Sources/ClsShadow.twin` | Clase de sombra/borde (`ucList.Shadow`), convertida a 64 bits |
| `Sources/mUcListHelpers.twin` | Módulo de subclasificación (`SetWindowSubclass`) |
| `Sources/ucSelect.twin` + `.tbcontrol` | Control ComboBox de ejemplo (usa un `ucList` interno) |
| `Sources/JsonBag.twin` | Parser JSON usado por los ejemplos `FrmComboBox` y `FrmGoogleSearch` |
| `Sources/Forms/*.tbform` + `*.twin` | Los 10 formularios de ejemplo recreados |
| `Demo/*.bmp` | Imágenes de `FrmGoogleSearch` extraídas de su `.frx` |

`ucScrollbar.ctl` y `clsSubClass.cls` **ya no son necesarios**.

## Cómo usarlo

1. Crea un proyecto nuevo en twinBASIC (por ejemplo *Standard EXE*).
2. En el *Project Explorer*, clic derecho sobre la carpeta `Sources` → **Import** →
   selecciona los 4 archivos de `twinBASIC/Sources`.
3. El control `ucList` aparece en la caja de herramientas. Úsalo igual que en VB6:

```vb
Private Sub Form_Load()
    With ucList1
        .ImageListInitialize 16, 16
        .ImageListAddImage App.Path & "\icono.png"
        .AddItem "Copiar" & vbTab & "Ctrl+C", 0
        .AddItem "-"
        .AddItem "Pegar" & vbTab & "Ctrl+V"
        .Shadow.ShadowOffsetY = 2
    End With
End Sub

Private Sub Command1_Click()
    Dim R As RECT
    GetWindowRect Command1.hWnd, R
    ucList1.ShowList R.Left, R.Bottom, , R.Top
End Sub

Private Sub ucList1_ItemClick()
    Debug.Print ucList1.Text
End Sub
```

## Proyecto de ejemplo (formularios)

Los 10 formularios de ejemplo están recreados en `Sources/Forms`. La disposición de
los controles es la del `.frm` original, convertida a píxeles.

| Formulario | Qué muestra |
|---|---|
| `FrmMain` | Menú principal: lista que abre los demás ejemplos (**objeto de inicio**) |
| `FrmAutocomplete` | Autocompletar varios `TextBox` y un `ucSelect` con lista de países |
| `FrmComboBox` | Tres `ucSelect` encadenados (provincias/departamentos/localidades) con datos de una API |
| `FrmCustomDraw` | Elementos dibujados a mano con el evento `ItemCustomDraw` |
| `FrmDragDrop` | Arrastrar y soltar elementos entre dos listas (OLE) |
| `FrmDropDownList` | Cinco listas desplegables lanzadas desde botones |
| `FrmGoogleSearch` | Búsqueda de imágenes con autocompletado en tiempo real y dictado por voz |
| `FrmListBox` | Cinco `ucList` con `ListStyle = StyleListBox` |
| `FrmMenu` | Menús y submenús con `ListStyle = StyleMenu` |
| `FrmShadow` | Uso de `ClsShadow` en un formulario sin bordes |

### Cómo montarlo
1. Crea un proyecto *Standard EXE* llamado **`ucListPlus`**. En los `.tbform`, los
   controles de usuario se identifican como `ucListPlus.ucList` y
   `ucListPlus.ucSelect`. Si tu proyecto tiene otro nombre, cambia ese prefijo en
   los `.tbform` con buscar y reemplazar.
2. Importa todos los archivos de `Sources` y de `Sources/Forms`.
3. En *Project → Settings*, elige **`FrmMain`** como objeto de inicio.
4. Copia junto al ejecutable (o en la carpeta del proyecto, que es `App.Path` al
   depurar) los archivos de datos del ejemplo original. No están en este
   repositorio: `VB6.png`, `Personal.txt`, `Nombres.txt`, `Groups.txt`,
   `Country.txt` y las carpetas `Caras\` y `Vectores PNG\`. Copia también los dos
   `.bmp` de `Demo/`.

### Propiedades de diseño de los `ucList`
twinBASIC guarda las propiedades de un UserControl en un `.tbform` como un bloque
binario (`_IPersistStreamData`) con un formato interno que no se puede generar de
forma fiable a mano. Por eso los `.tbform` solo guardan la posición, el tamaño y la
visibilidad de cada `ucList`/`ucSelect`. Las propiedades que en VB6 estaban en el `.frm`
(`ListStyle`, colores, `ListWidth`, `OLEDropMode`, …) se aplican por código en
`ApplyDesignProperties`, al principio de cada `Form_Load`.

Para que esto funcione, `ucList` tiene tres cambios:
- En ejecución, `ListStyle` se puede asignar por código. Muestra u oculta el control
  y no cambia su tamaño.
- Si se crea sin propiedades guardadas (`InitProperties` en ejecución), arranca en
  modo desplegable, oculto y sin los elementos de ejemplo.
- En el diseñador ya no se fuerza el tamaño de 32×32 en los estilos desplegable y menú,
  para que se mantenga la geometría del formulario.

Si abres un formulario en el diseñador de twinBASIC y lo guardas, el IDE añadirá el
bloque binario y a partir de ahí las propiedades se pueden editar también en la
ventana de propiedades.

### Otros cambios en los formularios
- `FrmShadow`: el `CommonDialog` (`COMDLG32.OCX`) se reemplazó por la API
  `ChooseColorW`, y las dos `Line` por `Label` de 1 píxel de alto.
- `FrmGoogleSearch`: las imágenes del `.frx` se cargan en `Form_Load` desde
  `GoogleSearch_Microphone.bmp` y `GoogleSearch_Search.bmp`. `mciSendString` y la
  codificación base64 usan las versiones Unicode (`W`) de la API.
- Todas las `Declare` de los formularios usan `PtrSafe`/`LongPtr`.
- Se corrigieron pequeños errores del original: `EOF(1)` en lugar de `EOF(FF)`,
  `.ListStyle = GoogleDesign` en lugar de `.ListDesign`, `CloseThemeData` que nunca
  se llamaba y el radio de la región al cambiar el ancho en `FrmShadow`.

## Qué cambió respecto a la versión VB6

### Cambios necesarios para twinBASIC / 64 bits
- Todas las `Declare` usan `PtrSafe` y `LongPtr` para handles y punteros.
- **Subclasificación**: `clsSubClass` usaba thunks en código máquina de 32 bits,
  que no funcionan en 64 bits. Ahora se usa `SetWindowSubclass`/`DefSubclassProc`
  de comctl32 (`mUcListHelpers.twin`), tanto en `ucList` como en `ClsShadow`.
- **Barra de desplazamiento**: `ucScrollbar` (también con thunks) se reemplazó por una
  barra interna dibujada por el propio control:
  - `ListDesign = GoogleDesign`: barra fina con *thumb* redondeado (igual que el
    antiguo `ucScrollbar1_OnPaint`).
  - `ListDesign = MicrosoftDesign`: barra con botones y tema de Windows (uxtheme), o
    clásica si los temas están desactivados.
  - Soporta arrastre del *thumb*, clic en la pista, botones con autorrepetición
    y rueda del ratón.
- **Timer1** se reemplazó por `SetTimer`/`WM_TIMER` sobre la ventana del control.
- **Text1** (el `TextBox` del filtro) se reemplazó por un control `EDIT` de Win32
  creado por código con `CreateWindowExW`:
  - `EN_CHANGE` (vía `WM_COMMAND` al padre) dispara la lógica de `Text1_Change`.
  - `WM_KEYDOWN`/`WM_CHAR` del EDIT se subclasifican y llaman a la lógica de
    `Text1_KeyDown`/`Text1_KeyPress`. Si la tecla se consume, se descarta el `WM_CHAR`
    para evitar el pitido.
  - `WM_GETDLGCODE` pide flechas, Enter y Escape (en modo desplegable) para que no
    los consuma el formulario. Tab sigue moviendo el foco.
  - Los colores se aplican con `WM_CTLCOLOREDIT` y la fuente con `WM_SETFONT`.
  - El EDIT solo existe en tiempo de ejecución. En diseño se dibuja solo el marco.
- `VarPtrArray` (msvbvm60.dll, que no existe en 64 bits) se eliminó; las imágenes
  desde arrays de bytes se cargan con `GlobalAlloc` + `CreateStreamOnHGlobal`.
- La imagen de diseño (`ucList.ctx`) se sustituyó por un icono dibujado por código
  en tiempo de diseño.

### Cambios en tipos de la interfaz pública (por 64 bits)
Estos cambios no afectan al código en 32 bits:
- `hWnd` devuelve `LongPtr`.
- `AutoCompleteTextBox(oTextBox As LongPtr, ...)`.
- Eventos `PostPaint(ByVal hdc As LongPtr)` e
  `ItemCustomDraw(Index, ByVal hdc As LongPtr, Left, Top, Width, Height)`.
- `ImageListDraw(ByVal hdc As LongPtr, ...)`.
- `ImageListAddImage` acepta también handles `LongLong` (HBITMAP/HICON en 64 bits).

### Errores del original corregidos
- `BSort SortAscendant` ordenaba en forma descendente (las dos ramas eran iguales).
- `SelectionRadius` marcaba como modificada la propiedad `ImagesRadius`.
- `AddItem` con `Index` dejaba restos del elemento desplazado en el nuevo
  (`IsGroup`, `SubMenu`, `ItemBold`).
- `RemoveItem` no actualizaba el contador de separadores y no comprobaba el rango.
- `ItemSelected` (Get) redibujaba el control innecesariamente.
- El evento `PostPaint` estaba declarado pero nunca se disparaba; ahora se dispara
  al terminar cada repintado.

## Notas
- Este código no se pudo compilar en el entorno en que se convirtió, porque
  twinBASIC solo funciona en Windows. Si al importarlo aparece algún error de
  compilación, suele bastar con ajustes menores.
- Los `.tbform` se generaron a partir de los `.frm` originales. El CLSID del
  `PictureBox` (`{33AD4ED0-6699-11CF-B70C-00AA0060D393}`) se dedujo de la secuencia de
  GUID de los controles de VB6 y no se pudo comprobar en un `.tbform` real. Si
  twinBASIC no reconoce los `PictureBox` al abrir un formulario, borra y vuelve a
  añadir esos controles en el diseñador con el mismo nombre e índice.
