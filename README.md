# Automatizador de Documentos Word

Sistema completo para automatizar la generación de documentos Word basados en plantillas con campos personalizables.

## Características

- **Interfaz gráfica intuitiva** con tkinter
- **Sistema de plantillas** con marcadores personalizables en formato `{{CAMPO}}`
- **Generación automática** de documentos Word
- **Guardar y cargar configuraciones** para reutilizar datos comunes
- **Asistente para crear plantillas** desde documentos existentes
- **Soporte para múltiples plantillas** diferentes

## Estructura del Proyecto

```
proyecto-automatizado-documentos/
│
├── main.py                   # Script principal para ejecutar la aplicación
├── requirements.txt          # Dependencias del proyecto
├── README.md                 # Este archivo
│
├── src/                      # Código fuente
│   ├── interfaz_grafica.py  # Interfaz gráfica principal
│   ├── procesador_plantillas.py  # Motor de procesamiento de plantillas
│   └── crear_plantilla.py   # Asistente para crear plantillas
│
├── templates/                # Plantillas de documentos Word
│   └── 9393653-Memoria_tecvf.docx  # Ejemplo de documento
│
├── output/                   # Documentos generados (se crea automáticamente)
│
├── config/                   # Archivos de configuración
│   └── ejemplo_memoria_tecnica.json  # Configuración de ejemplo
│
└── docs/                     # Documentación adicional

```

## Requisitos

- Python 3.7 o superior
- pip (gestor de paquetes de Python)

### Dependencias

Las dependencias se instalan automáticamente con pip:

```bash
pip install -r requirements.txt
```

**Nota para Linux:** Si tkinter no está instalado, ejecuta:

```bash
# Ubuntu/Debian
sudo apt-get install python3-tk

# Fedora
sudo dnf install python3-tkinter
```

## Instalación

1. **Clonar o descargar el repositorio**

```bash
git clone https://github.com/pipex360/proyecto-automatizado-documentos.git
cd proyecto-automatizado-documentos
```

2. **Instalar dependencias**

```bash
pip install -r requirements.txt
```

3. **Verificar instalación**

```bash
python main.py
```

## Uso

### Opción 1: Usar la Interfaz Gráfica (Recomendado)

1. **Ejecutar la aplicación**

```bash
python main.py
```

2. **Seleccionar una plantilla**
   - Click en "Examinar..."
   - Selecciona tu archivo .docx con marcadores

3. **Cargar campos**
   - Click en "Cargar Campos"
   - La aplicación detectará automáticamente todos los marcadores `{{CAMPO}}`

4. **Rellenar datos**
   - Completa los campos mostrados en el formulario

5. **Generar documento**
   - Click en "Generar Documento"
   - Elige dónde guardar el archivo
   - ¡Listo! Tu documento se generará automáticamente

### Opción 2: Crear Tu Propia Plantilla

#### Método Manual

1. Abre tu documento Word existente
2. Identifica los datos que quieres automatizar
3. Reemplázalos con marcadores en formato `{{NOMBRE_MARCADOR}}`

**Ejemplo:**

Texto original:
```
La empresa ACME S.A., RUT 12.345.678-9, ubicada en Av. Principal 123
```

Texto con marcadores:
```
La empresa {{EMPRESA}}, RUT {{RUT}}, ubicada en {{DIRECCION}}
```

4. Guarda el documento en la carpeta `templates/`

#### Método Asistido

Usa el asistente para analizar documentos existentes:

```bash
cd src
python crear_plantilla.py
```

Opciones disponibles:
- **Analizar un documento existente**: Te sugerirá qué campos automatizar
- **Crear plantilla de ejemplo**: Genera una plantilla base para empezar

### Opción 3: Uso Programático

Puedes usar el módulo directamente en tus scripts:

```python
from src.procesador_plantillas import ProcesadorPlantillas

# Crear procesador
procesador = ProcesadorPlantillas("templates/mi_plantilla.docx")

# Definir datos
datos = {
    "EMPRESA": "Mi Empresa S.A.",
    "RUT": "12.345.678-9",
    "DIRECCION": "Calle Principal 123"
}

# Generar documento
procesador.procesar(datos, "output/documento_generado.docx")
```

## Funcionalidades Avanzadas

### Guardar y Cargar Configuraciones

**Guardar configuración:**
1. Rellena los campos en la interfaz
2. Click en "Guardar Configuración"
3. Elige nombre y ubicación del archivo .json

**Cargar configuración:**
1. Click en "Cargar Configuración"
2. Selecciona el archivo .json guardado previamente
3. Los campos se rellenarán automáticamente

Esto es útil para:
- Reutilizar datos comunes de tu empresa
- Crear diferentes configuraciones para distintos clientes
- Compartir configuraciones con tu equipo

### Marcadores Especiales

Puedes usar cualquier nombre para tus marcadores, pero se recomienda:

- **Usar MAYÚSCULAS**: `{{EMPRESA}}` en vez de `{{empresa}}`
- **Usar guiones bajos** para separar palabras: `{{SUPERFICIE_TOTAL}}`
- **Nombres descriptivos**: `{{FECHA_INICIO}}` en vez de `{{F1}}`

**Marcadores comunes recomendados:**

```
{{EMPRESA}}              - Nombre de la empresa
{{RUT}}                  - RUT o identificador fiscal
{{DIRECCION}}            - Dirección principal
{{COMUNA}}               - Comuna/Ciudad
{{REGION}}               - Región/Provincia/Estado
{{FECHA}}                - Fecha del documento
{{PERSONAL}}             - Número de trabajadores
{{HORARIO}}              - Horario de trabajo
{{SUPERFICIE_TOTAL}}     - Superficie total
{{SUPERFICIE_CONSTRUIDA}} - Superficie construida
{{ACTIVIDAD}}            - Descripción de actividad
{{CONTACTO}}             - Persona de contacto
{{TELEFONO}}             - Teléfono de contacto
{{EMAIL}}                - Correo electrónico
```

## Ejemplos

### Ejemplo 1: Memoria Técnica

El proyecto incluye un ejemplo completo de memoria técnica industrial.

**Plantilla:** `templates/9393653-Memoria_tecvf.docx`
**Configuración:** `config/ejemplo_memoria_tecnica.json`

Para usarlo:
1. Convierte el documento en plantilla (reemplaza datos específicos con marcadores)
2. Carga la plantilla en la aplicación
3. Usa la configuración de ejemplo o crea la tuya

### Ejemplo 2: Crear Plantilla Simple

```python
from docx import Document

# Crear documento
doc = Document()
doc.add_heading('Carta de Presentación', 0)
doc.add_paragraph(f'Estimado/a {{{{DESTINATARIO}}}},')
doc.add_paragraph(
    f'Por medio de la presente, {{{{EMPRESA}}}} desea expresar su interés '
    f'en {{{{PROPOSITO}}}}.'
)
doc.add_paragraph('Atentamente,')
doc.add_paragraph('{{REMITENTE}}')
doc.add_paragraph('{{CARGO}}')

# Guardar
doc.save('templates/carta_presentacion.docx')
```

## Solución de Problemas

### Error: "No se encontró la plantilla"
- Verifica que la ruta del archivo sea correcta
- Asegúrate de que el archivo tenga extensión .docx

### Error: "No se encontraron marcadores"
- Verifica que los marcadores estén en formato `{{MARCADOR}}`
- Usa llaves dobles: `{{}}`, no simples `{}`
- Los marcadores deben estar en MAYÚSCULAS

### Error: "python-docx no está instalado"
```bash
pip install python-docx
```

### Error: "tkinter no está disponible" (Linux)
```bash
sudo apt-get install python3-tk  # Ubuntu/Debian
sudo dnf install python3-tkinter  # Fedora
```

### El documento generado no mantiene el formato
- Asegúrate de que los marcadores estén completos en un solo "run" de Word
- Evita aplicar formato (negrita, cursiva) a parte del marcador
- Si es necesario, copia el marcador completo y pega sin formato

## Contribuir

Las contribuciones son bienvenidas. Por favor:

1. Fork el repositorio
2. Crea una rama para tu feature (`git checkout -b feature/nueva-funcionalidad`)
3. Commit tus cambios (`git commit -am 'Agregar nueva funcionalidad'`)
4. Push a la rama (`git push origin feature/nueva-funcionalidad`)
5. Crea un Pull Request

## Licencia

Este proyecto está bajo la licencia MIT. Ver el archivo LICENSE para más detalles.

## Contacto

Para preguntas o sugerencias, por favor abre un issue en el repositorio de GitHub.

## Próximas Funcionalidades

- [ ] Soporte para imágenes dinámicas
- [ ] Integración con Excel para cargar múltiples registros
- [ ] Generación masiva de documentos
- [ ] Plantillas para diferentes tipos de documentos (cartas, informes, contratos)
- [ ] Validación de campos (formatos, campos requeridos)
- [ ] Previsualización antes de generar
- [ ] Exportación a PDF
- [ ] Historial de documentos generados

---

**Versión:** 1.0.0
**Autor:** Proyecto Automatizado Documentos
**Última actualización:** Noviembre 2025