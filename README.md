# Práctica 2 – Automatización de Microsoft Excel mediante COM Interop

Aplicación de consola en C# que controla Microsoft Excel a través de COM Interop (Tema 4. COM/DCOM).
Genera un reporte de ventas con fórmulas, formato y un gráfico, lo guarda como `.xlsx` y libera
todas las referencias COM para no dejar procesos `EXCEL.EXE` huérfanos.

## Requisitos

- Windows con **Microsoft Excel** instalado (probado con Office 365, Excel 16.0).
- **Visual Studio 2019** con la carga de trabajo *Desarrollo de escritorio con .NET*.
- **.NET Framework 4.8** (Developer Pack).
- Primary Interop Assembly de Excel (`Microsoft.Office.Interop.Excel` 15.0) en el GAC.
  Se instala junto con Office / Visual Studio.

## Estructura

```
ExcelComInterop.sln
ExcelComInterop/
├── ExcelComInterop.csproj
├── Program.cs               # Toda la lógica de la práctica
├── App.config
└── Properties/AssemblyInfo.cs
```

## Compilar

**Desde Visual Studio:** abrir `ExcelComInterop.sln` y usar *Compilar → Recompilar solución*
(configuración `Debug` / `Any CPU`).

**Desde la terminal:**

```powershell
& "C:\Program Files (x86)\Microsoft Visual Studio\2019\Enterprise\MSBuild\Current\Bin\MSBuild.exe" `
    ExcelComInterop.sln /restore:false /t:Rebuild /p:Configuration=Debug
```

(Ajustar `Enterprise` por `Community` o `Professional` según la edición instalada.)

El ejecutable queda en `ExcelComInterop\bin\Debug\ExcelComInterop.exe`.

## Ejecutar

Desde Visual Studio con **F5** / **Ctrl+F5**, o desde la terminal:

```powershell
.\ExcelComInterop\bin\Debug\ExcelComInterop.exe
```

Por defecto el archivo se guarda en:

```
%USERPROFILE%\Documents\Practica2_ExcelCOM\Reporte_Ventas_Trimestral.xlsx
```

Para usar otra ruta, pasarla como primer argumento:

```powershell
.\ExcelComInterop\bin\Debug\ExcelComInterop.exe "C:\Temp\reporte.xlsx"
```

Al terminar, el programa espera una tecla (se omite si la entrada está redirigida).
El código de salida es `0` si todo salió bien, `1` si hubo un error y `2` si `EXCEL.EXE`
no se cerró.

## Qué hace

| Paso | Descripción |
|------|-------------|
| a | Crea una instancia invisible de `Excel.Application` |
| b | Agrega un libro y selecciona la hoja activa (`Ventas`) |
| c | Escribe el encabezado, una tabla de 10 productos, fórmulas (`Total Ventas`, `Valor del Inventario`, sumas totales) y aplica formato (negritas, color, bordes, moneda) |
| d | Crea un gráfico de columnas enlazado a la tabla, embebido en la hoja |
| e | Guarda el libro como XLSX y lo cierra |
| f | Llama a `Quit()`, libera cada objeto COM con `Marshal.ReleaseComObject`, fuerza el GC y comprueba que el proceso `EXCEL.EXE` haya terminado |

## Notas

- Excel puede tardar **10–15 segundos** en cerrarse tras `Quit()`; el programa espera hasta 60 s
  y lo indica en pantalla.
- La referencia al PIA en el `.csproj` apunta a la ruta del GAC de este equipo
  (`C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Excel\15.0.0.0__71e9bce111e9429c\`).
  En otro equipo, quitar esa referencia y agregarla de nuevo desde
  *Agregar referencia → Ensamblados → Extensiones* (o *COM → Microsoft Excel 16.0 Object Library*).
- Si aparece `Skipped: 1` al compilar en VS, revisar en *Configuration Manager* que el proyecto
  tenga marcada la casilla *Build* para `Debug | Any CPU`.
