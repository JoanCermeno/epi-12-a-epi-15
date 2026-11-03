# EPI-12 a EPI-15

Script para completar hojas vacías de un libro EPI-15 mensual a partir de los EPI-12 semanales de cada CPT.

## Requisitos

- Python 3.9 o posterior.
- `openpyxl`.

En Manjaro:

```bash
sudo pacman -S python-openpyxl
```

También se puede instalar con pip:

```bash
pip install openpyxl
```

En Windows, instala Python desde [python.org](https://www.python.org/) y activa **Add Python to PATH** durante la instalación. Después abre el símbolo del sistema y ejecuta `pip install openpyxl`.

## Uso mensual

Coloca el script y el EPI-15 del mes en una carpeta, y guarda los EPI-12 semanales en una subcarpeta:

```text
epi-12-15/
├── epi12_a_epi15.py
├── EPI-15 NUEVO SEPTIEMBRE.xlsx
└── epi-12/
    ├── EPI-12 semana 1.xlsx
    ├── EPI-12 semana 2.xlsx
    ├── EPI-12 semana 3.xlsx
    └── EPI-12 semana 4.xlsx
```

Abre una terminal en esa carpeta y ejecuta:

```bash
python epi12_a_epi15.py "EPI-15 NUEVO SEPTIEMBRE.xlsx" epi-12/ --mes SEPTIEMBRE
```

En Windows puedes usar `py` en lugar de `python`.

El script genera un archivo nuevo con el sufijo `_AUTO.xlsx` y un reporte `_reporte.txt`. No modifica los archivos originales.

## Opciones

- `--mes OCTUBRE`: escribe el mes en las hojas que llena.
- `--solo "CPTII EL LICEO,CPTII GARCITA"`: procesa solo esos CPT.
- `--otras-total`: pone en OTRAS CAUSAS el total atendido completo, sin restar los diagnósticos.
- `--salida nombre.xlsx`: elige el nombre del archivo de salida.

## Qué hace

- Lee los EPI-12 y suma los casos por CPT y grupo de edad.
- Solo llena las hojas del EPI-15 que considera vacías; las hojas con datos no se tocan.
- Completa los diagnósticos definidos en la tabla de equivalencias del script.
- Calcula OTRAS CAUSAS DE CONSULTA para que el total coincida con los pacientes atendidos.
- Escribe el mes indicado y una nota en la hoja procesada.

El reporte avisa sobre semanas faltantes o repetidas, CPT sin pacientes, casos con edad ignorada y bloques sin una hoja correspondiente.

## Limitaciones

- Los diagnósticos sin equivalencia quedan sin llenar y se contabilizan en OTRAS CAUSAS.
- No completa los datos de identificación del encabezado.
- No corrige errores de `LISTA_CPT` en los EPI-12.

> Antes de usarlo con reportes oficiales, comprueba los resultados contra copias de tus formularios. La estructura de las hojas y las columnas debe coincidir con la que espera el script.
