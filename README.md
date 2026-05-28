# Monitor de Precios de Hardware

Scraping de precios de productos propios vs. competidores y generación de dashboard en Excel con envío por correo.

## Requisitos

- Python 3.10+
- pip

## Instalación

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Configuración

Copia y completa el archivo `.env`:

```
GMAIL_SENDER=tu_correo@gmail.com
GMAIL_APP_PASSWORD=tu_contraseña_de_aplicacion
GMAIL_RECIPIENT=destinatario@ejemplo.com
```

La contraseña de aplicación se genera en: https://myaccount.google.com/apppasswords

## Uso

1. Prepara `mapeo_hardware.xlsx` con las columnas:
   `SKU_Interno`, `Nombre_Producto`, `Precio_Propio`, `Stock_Propio`,
   `URL_Competidor_A`, `URL_Competidor_B`

2. Ejecuta:

```bash
python3 hardware_price_monitor.py
```

Genera `Dashboard_Hardware_YYYYMMDD.xlsx` y envía el reporte por correo con el archivo adjunto.
