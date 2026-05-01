# Arabic poetry formatting (Word)

VBA macros for Microsoft Word: after you type a verse with **`**`** between the two hemistichs and press **Enter**, the line becomes a borderless RTL two-column table (صدر | عجز).

## Use

1. **Alt+F11** → **File → Import File…** → `macros/FormatPoem.bas` → save **Normal.dotm** if you want it in all documents.
2. **Alt+F8** → run **`ToggleArabicPoetryTable`** (ON). Run again to turn OFF.
3. Example: `الصدر ** العجز` then **Enter**.

Macros need a macro-enabled document/template (e.g. `.docm` / `.dotm`) and allowed macro settings in Word’s Trust Center.
