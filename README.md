# BATSEJ_OPEN_FINANCE

# Automatización de Cálculo de Comisiones para BATSEJ OPEN FINANCE

Este proyecto implementa una automatización en Python para calcular las comisiones que las empresas deben pagar por el uso de una API de verificación de cuentas bancarias. Los datos se extraen de una base de datos SQLite y los resultados se exportan a un archivo Excel para su análisis y facturación.

---

## **Funcionalidades**

- **Carga de datos:** Extrae los registros de la base de datos SQLite.
- **Limpieza de datos:** Filtra las empresas activas y las fechas dentro del rango solicitado (julio y agosto de 2024).
- **Cálculo de comisiones:** Aplica la lógica de cobro basada en el contrato de cada empresa, incluyendo descuentos y el IVA del 19%.
- **Exportación de resultados:** Guarda los resultados en un archivo Excel dentro de la carpeta `resultado/`.
- **Envío de correos:** Permite enviar el archivo Excel por correo electrónico (opcional).

---

## 📋 **Requisitos**

Para ejecutar este proyecto, necesitas tener instalado lo siguiente:

- **Python 3.7 o superior:** [Descargar Python](https://www.python.org/downloads/)
- **SQLite3:** Viene incluido con Python, no es necesario instalarlo por separado.

---

## 🛠️ **Instalación de Dependencias**

Puedes instalar las dependencias necesarias de dos maneras:

### Opción 1: Usando `setup.py`
Ejecuta el siguiente comando para instalar las dependencias automáticamente:
```bash
python setup.py
