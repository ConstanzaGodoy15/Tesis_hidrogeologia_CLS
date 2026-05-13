# Sistema de Codificación Litológica (CLS)
 
Script desarrollado para la memoria de título *"Caracterización sedimentaria del Acuífero de Santiago mediante un Sistema de Codificación Litológica automatizado"*, Escuela de Geología, Facultad de Ciencias Físicas y Matemáticas, Universidad de Chile.
 
---
 
## ¿Qué hace?
 
El Sistema CLS transforma descripciones estratigráficas textuales heterogéneas —provenientes de expedientes de perforación de pozos— en un código litológico estructurado de cuatro caracteres denominado **TDSM**:
 
| Posición | Nombre | Descripción |
|---|---|---|
| T | Tipo de unidad | Sedimento (U), Roca (R), Suelo (S) o Sin información (0) |
| D | Dominante | Material principal del intervalo |
| S | Secundario | Material secundario del intervalo |
| M | Matriz/Accesorio | Tercer componente o material de relleno |
 
El script procesa cada intervalo estratigráfico, identifica los materiales presentes mediante un diccionario litológico controlado, aplica reglas de interpretación sedimentológica y genera automáticamente el código correspondiente junto con un conjunto de variables interpretativas adicionales.
 
---
 
## Archivos de entrada requeridos
 
El script requiere tres archivos Excel con la siguiente estructura:
 
| Archivo | Hojas requeridas |
|---|---|
| `Pozos_memoria_CGC.xlsx` | `collar`, `estratigrafia_bruta` |
| `Diccionario_estratigrafia.xlsx` | `LITOLOGIA`, `MODIFICADORES`, `INDICADORES`, `CLASES_TAMANO` |
| `Template_resultado_CLS.xlsx` | `collar`, `codificacion_estandar`, `desglose_codificacion`, `proporcion_sedimento`, `Diccionario` |
 
> Los archivos de entrada no se incluyen en este repositorio debido a su volumen. Para solicitarlos, ver la sección de contacto.
 
---
 
## Archivos de salida
 
El script genera un archivo Excel (`Resultados_CLS_FINAL.xlsx`) con cuatro hojas:
 
- **collar**: información general de cada pozo
- **codificacion_estandar**: código TDSM y variables interpretativas por intervalo
- **desglose_codificacion**: descomposición del código con atributos adicionales (tamaño, forma, color, estructuras, etc.)
- **proporcion_sedimento**: espesor equivalente y porcentaje por clase granulométrica por pozo
---
 
## Uso
 
1. Clona el repositorio o descarga el script `COD_CLS_FINAL.py`
2. Instala las dependencias:
```bash
pip install pandas numpy openpyxl
```
3. Edita las rutas al inicio del script:
```python
DB_PATH       = Path("ruta/a/Pozos_memoria_CGC.xlsx")
DICT_PATH     = Path("ruta/a/Diccionario_estratigrafia.xlsx")
TEMPLATE_PATH = Path("ruta/a/Template_resultado_CLS.xlsx")
OUTPUT_PATH   = Path("ruta/a/Resultados_CLS_FINAL.xlsx")
```
4. Ejecuta el script:
```bash
python COD_CLS_FINAL.py
```
 
---
 
## Requisitos
 
- Python 3.10+
- pandas
- numpy
- openpyxl
---
 
## Contacto
 
Para solicitar los archivos de entrada o ante cualquier consulta, contactarse con constanza.godoy.c@ug.uchile.cl
