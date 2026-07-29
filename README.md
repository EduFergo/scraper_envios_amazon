# Amazon Delivery Time Checker

Script en Python que consulta automáticamente el tiempo estimado de entrega de un producto de Amazon.es para varios códigos postales y guarda los resultados en un archivo Excel.

---

## Requisitos

- Python 3 instalado.
- Dependencias del proyecto.

Instalar las dependencias:

```bash
pip install playwright openpyxl
```

Instalar el navegador utilizado por Playwright:

```bash
playwright install chromium
```

---

## Estructura del proyecto

```
Proyecto/
│
├── main.py
├── ejecutar.bat
└── tiempos_envio.xlsx    # Se crea automáticamente
```

---

## Configuración

En `main.py` se encuentra el diccionario de códigos postales que se utilizarán para realizar las consultas.

Ejemplo:

```python
CODIGOS_POSTALES = {
    "A Coruña": "15001",
    "Madrid": "28001",
    "Barcelona": "08001",
}
```

Puedes añadir, modificar o eliminar códigos postales según tus necesidades.

---

## Ejecución

La forma recomendada de utilizar la aplicación es mediante el archivo:

```
ejecutar.bat
```

No es necesario abrir una consola ni ejecutar comandos manualmente.

### Pasos

1. Abrir el producto de Amazon que se desea consultar.
2. Copiar la URL del producto.
3. Hacer doble clic sobre `ejecutar.bat`.
4. Cuando aparezca el mensaje:

```
Pega la URL del producto y pulsa ENTER:
```

5. Pegar la URL (`Ctrl + V`) y pulsar **Enter**.

El programa abrirá automáticamente el navegador, consultará todos los códigos postales configurados y generará el archivo Excel con los resultados.

---

## Resultado

Al finalizar se generará (o actualizará) el archivo:

```
tiempos_envio.xlsx
```

Este archivo contiene un histórico con:

- Fecha y hora de la consulta.
- Provincia.
- Código postal.
- Estimación de entrega mostrada por Amazon.

---

## Notas

- La consulta puede tardar entre 2 y 3 minutos dependiendo del número de códigos postales configurados.
- Durante la ejecución no se recomienda interactuar con el navegador que abre el programa.
- Si Amazon modifica el diseño de su página, puede ser necesario actualizar los selectores utilizados por el script.

---

## Solución de problemas

### Python no se reconoce como comando

Comprobar que Python está instalado correctamente ejecutando:

```bash
python --version
```

### Error relacionado con Playwright

Volver a instalar Chromium:

```bash
playwright install chromium
```

### No se genera el archivo Excel

Comprobar los mensajes mostrados durante la ejecución. Si todas las consultas fallan, es posible que Amazon haya modificado la estructura de la página y sea necesario actualizar los selectores del script.
