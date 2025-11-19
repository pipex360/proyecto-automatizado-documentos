# Inicio Rápido

Guía para comenzar a usar el automatizador en 5 minutos.

## 1. Instalación (1 minuto)

```bash
# Instalar dependencias
pip install python-docx
```

## 2. Ejecutar la Aplicación (1 minuto)

```bash
# Ejecutar desde la raíz del proyecto
python main.py
```

Se abrirá la interfaz gráfica.

## 3. Primera Prueba con Plantilla de Ejemplo (3 minutos)

### Opción A: Crear Plantilla de Ejemplo

1. En una terminal, ejecuta:
```bash
cd src
python crear_plantilla.py
```

2. Selecciona opción **2** (Crear plantilla de ejemplo)

3. Se creará `templates/plantilla_ejemplo.docx`

### Opción B: Convertir Tu Documento

Si ya tienes el documento `9393653-Memoria_tecvf.docx`:

1. Abre el documento en Word
2. Busca datos específicos como:
   - "DAVID BROWN SANTASALO SOUTH AMERICA S.A"
   - "76.130.276-0"
   - "Pudahuel Poniente 1107"
   - "36 trabajadores"
   - etc.

3. Reemplaza estos datos con marcadores:
   - "DAVID BROWN..." → `{{EMPRESA}}`
   - "76.130.276-0" → `{{RUT}}`
   - "Pudahuel..." → `{{DIRECCION}}`
   - "36" → `{{PERSONAL}}`

4. Guarda como `templates/plantilla_memoria.docx`

## 4. Usar la Interfaz Gráfica

1. **Cargar plantilla:**
   - Click "Examinar..."
   - Selecciona tu plantilla
   - Click "Cargar Campos"

2. **Rellenar datos:**
   - Completa los campos que aparecen

3. **Generar documento:**
   - Click "Generar Documento"
   - Elige dónde guardar
   - ¡Listo!

## Ejemplo de Datos para Prueba

Si creaste la plantilla de ejemplo, usa estos datos:

```
EMPRESA: Mi Empresa de Prueba S.A.
RUT: 99.999.999-9
DIRECCION: Calle Falsa 123
COMUNA: Santiago
REGION: Región Metropolitana
PERSONAL: 50
HORARIO: Lunes a Viernes, 9:00 a 18:00
SUPERFICIE_TOTAL: 1000
SUPERFICIE_CONSTRUIDA: 500
DESCRIPCION_ACTIVIDAD: Actividades de prueba para demostración del sistema
```

## Solución Rápida de Problemas

**Error: "No module named 'docx'"**
```bash
pip install python-docx
```

**Error: "No se encontraron marcadores"**
- Verifica que usaste doble llave: `{{MARCADOR}}`
- Deben estar en MAYÚSCULAS

**La interfaz no se abre**
- Linux: `sudo apt-get install python3-tk`
- Verifica Python 3.7+: `python --version`

## Próximos Pasos

- Lee el [README.md](README.md) completo para funcionalidades avanzadas
- Crea tus propias plantillas personalizadas
- Guarda configuraciones para reutilizar datos comunes
- Explora el asistente de plantillas: `python src/crear_plantilla.py`

¡Disfruta automatizando tus documentos!
