# ucListPlus para twinBASIC

Conversión del control **ucList** (ucListPlus, de Leandro Ascierto) de VB6 a
[twinBASIC](https://twinbasic.com). Mantiene los mismos **eventos, propiedades y
métodos públicos** que el `ucList.ctl` original y compila tanto en **32 como en 64 bits**.

## Archivos

| Archivo | Descripción |
|---|---|
| `Sources/ucList.twin` | Código del UserControl `ucList` |
| `Sources/ucList.tbcontrol` | Diseñador del UserControl (incluye el `TextBox` `Text1` del filtro) |
| `Sources/ClsShadow.twin` | Clase de sombra/borde (`ucList.Shadow`), convertida a 64 bits |
| `Sources/mUcListHelpers.twin` | Módulo de subclasificación (`SetWindowSubclass`) |

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
- Los formularios de ejemplo de VB6 (`Frm*.frm`) se pueden importar en twinBASIC
  con su importador de proyectos VB6. Cambia `Proyecto1.ucList` por este `ucList`.
