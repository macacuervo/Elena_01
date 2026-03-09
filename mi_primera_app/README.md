# Automatización PDF -> Word (versión simple)

## 1) Archivos encontrados en `referencias/`

- `pat67524_2026.pdf`
- `Modelo InvAT con macro importar Delt@ (v6 octubre 2025).docm`

## 2) Estructura del PDF (resumen para principiantes)

El PDF está dividido en páginas de formulario (parte de accidente). Internamente guarda texto en bloques comprimidos (`FlateDecode`).

En el contenido se ven pares tipo **etiqueta -> valor**, por ejemplo:

- `Nombre:` -> `PEDRO`
- `Apellido 1º:` -> `LOPEZ`
- `Fecha de accidente:` -> `09/02/2026`
- `Nombre o Razón Social:` -> `RECOBAL ALEJOS BAHILLO SL`

La app usa esas etiquetas para localizar los datos.

## 3) Estructura del Word `.docm`

El documento `.docm` (Word con macros) contiene:

- **macros VBA** (`word/vbaProject.bin`),
- **campos legacy de formulario** (`FORMTEXT`, `FORMCHECKBOX`),
- y además varios **placeholders de texto** en minúsculas dentro de `word/document.xml`, por ejemplo:
  - `trabnombre`
  - `trabapellido1`
  - `trabapellido2`
  - `fechaacci`
  - `empresa`
  - `direccion1`, `localidad1`, `provincia1`

Para mantenerlo simple y robusto sin romper macros, esta primera versión reemplaza solo esos placeholders de texto.

## 4) Solución técnica propuesta (simple y robusta)

1. Leer el PDF como binario.
2. Descomprimir los streams de texto (`zlib`).
3. Extraer cadenas de texto del contenido PDF.
4. Buscar valores por etiqueta con expresiones regulares sencillas.
5. Copiar la plantilla `.docm` a un archivo de salida.
6. Abrir el ZIP del `.docm` y reemplazar placeholders en `word/document.xml`.

Ventajas:
- no depende de librerías externas,
- conserva macros,
- fácil de depurar y ampliar.

## 5) Tabla de correspondencias (incluye ambigüedades)

| Placeholder Word | Fuente en PDF | Regla aplicada | Nota |
|---|---|---|---|
| `trabnombre` | `Nombre:` | primera línea tras etiqueta | clara |
| `trabapellido1` | `Apellido 1º:` | primera línea tras etiqueta | clara |
| `trabapellido2` | `Apellido 2º:` | primera línea tras etiqueta | clara |
| `fechaacci` | `Fecha de accidente:` | primera línea tras etiqueta | clara |
| `fechanac` | `Fecha nacimiento:` | primera línea tras etiqueta | clara |
| `empresa` | `Nombre o Razón Social:` | primera aparición | **ambigua** (puede haber varias empresas) |
| `direccion1` | `Domicilio:` | concatena 1ª y 2ª línea | **ambigua** por formato variable |
| `localidad1` | `Municipio:` | primera aparición | puede ser trabajador o centro |
| `provincia1` | `Provincia:` | primera aparición | puede ser trabajador o centro |
| `ocupacion` | `Ocupación:` | primera línea tras etiqueta | clara |
| `fechaingreso` | `Fecha de ingreso en la empresa:` | primera línea tras etiqueta | clara |
| `empresacalle` | bloque centro de trabajo | línea tras `Domicilio:` en bloque empresa | heurística |
| `empresamunicipio` | bloque centro de trabajo | línea tras `Municipio:` en bloque empresa | heurística |
| `empresaprovincia` | bloque centro de trabajo | línea tras `Provincia:` en bloque empresa | heurística |

> Recomendación: si luego quieres precisión total, el siguiente paso es mapear por sección (Trabajador, Empresa, Centro, Accidente) con parser por bloques.

## 6) Cómo ejecutar la primera versión funcional

```bash
python mi_primera_app/app.py \
  --pdf "mi_primera_app/referencias/pat67524_2026.pdf" \
  --template "mi_primera_app/referencias/Modelo InvAT con macro importar Delt@ (v6 octubre 2025).docm" \
  --output "mi_primera_app/salida.docm"
```

El script imprimirá los campos detectados y generará `mi_primera_app/salida.docm`.

## 7) Qué mejora esta versión reforzada

- Valida que existan los archivos de entrada y que sus extensiones sean correctas.
- Si hay un problema, muestra errores cortos y fáciles de entender (pensados para principiantes).
- Genera un informe claro con:
  - campos encontrados,
  - campos obligatorios faltantes,
  - campos opcionales faltantes,
  - cuántos placeholders del Word se han reemplazado.

Esto te permite detectar rápidamente si el PDF viene con un formato distinto y qué regla habría que ajustar.
