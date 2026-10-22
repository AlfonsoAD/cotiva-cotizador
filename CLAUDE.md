# Cotiva · cotizador Excel/VBA y sus productos — flujo con bflow

La sesión principal conduce las tareas con la skill `bflow`: pregunta al humano y lanza a los agentes `bflow-*`. No implementa; el código lo escribe `bflow-implementer`.

- Avanzar una tarea, retomar trabajo o "¿qué sigue?" → skill `bflow` (o `/bflow <ID>`). bflow lleva el estado en GitHub Issues (etiquetas `bflow:*`), las ramas, el check y el PR.
- Rama base: la define `vcs.base_branch` en el `bflow.yaml` de cada rama. bflow saca de ahí las ramas de cada tarea y hace PR y merge de vuelta a ella.
  - `dev` → `main`: el cotizador genérico.
  - `auto-repair-shop`: la plantilla genérica de Cotiva para talleres mecánicos, construida sobre el cotizador.
- Lo que entra a una rama base debe servir a cualquier cliente de ese producto (en `auto-repair-shop`, a cualquier taller): aunque una mejora nazca de un cliente, se implementa genérica. Lo propio de un taller (nombre, logos, colores, datos fiscales, precios, catálogos) es configuración (hoja "Mi empresa", `assets/`, hojas `BD_*`), nunca constantes en el código ni textos con su nombre. Si algo de verdad solo aplica a un cliente, va en su rama `cliente/<nombre>`, que sale de la rama base de su producto y nunca se mergea de vuelta.
- Escribir o revisar código VBA → skill `vba-cotizador` (arquitectura por módulo, clean code, maquetado de formularios y plantillas).
- Preguntas de exploración, dudas de arquitectura o mantenimiento de `.claude/`, `CLAUDE.md` o `bflow.yaml` → respóndelas o edítalas directo, sin flujo.
- Las fechas límite las pone el humano en el tracker; nunca las muevas.

## Mapa
| Ruta | Qué es |
|---|---|
| `bflow.yaml` | Configuración de bflow: tracker, rama base, check |
| `specs/<ID>-<slug>/spec.md` | Specs de las tareas hechas con bflow |
| `src/modules/` | Módulos `.bas`; `modData` es el único que toca las hojas `BD_*` |
| `src/classes/` | Modelo (`Quote`, `QuoteItem`), sin acceso a hojas |
| `src/forms/` | Código de formularios; la interfaz se construye por código |
| `src/build/modBuild.bas` | Arma el libro durante el build y se elimina del `.xlsm` final |
| `build/` | `build.ps1` (src → `.xlsm`) y `export.ps1` (`.xlsm` → src) |
| `assets/` | Logos opcionales para el build |
| `.claude/skills/vba-cotizador/` | Reglas de código VBA del repo |
| `README.md` | Uso, distribución, seguridad, macros de administrador y migración de datos |

## Comandos
- Build: `powershell -ExecutionPolicy Bypass -File build\build.ps1` → `dist/Cotiva-Cotizador.xlsm`. Requiere Excel de escritorio y `src/modules/modSecrets.bas` (copia de `modSecrets.example.bas`; nunca se sube). Abre Excel visible: un diálogo de VBA es un error de compilación, se corrige en `src/`.
- Export: `powershell -ExecutionPolicy Bypass -File build\export.ps1 [-Workbook ruta.xlsm]` para traer a `src/` lo editado en el editor de VBA.
- Nunca edites el `.xlsm` a mano ni lo subas al repo (`dist/` está ignorado). `.gitattributes` convierte el código a Windows-1252 con CRLF en la copia de trabajo: no cambies la codificación de los archivos de `src/`.
- No hay tests automáticos: la verificación es compilar, correr `build.ps1` y probar en Excel nuevo, guardar, cargar, PDF y datos de versiones anteriores.
