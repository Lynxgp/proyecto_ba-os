# Registro de salidas

Aplicación web de Google Apps Script para registrar las salidas del alumnado al baño, con hoja de cálculo de Google como base de datos.
Desarrollada para el IES Campanillas (Málaga).

## Contenido
- `Code.gs`: servidor (Apps Script).
- `Index.html`: interfaz web.
- `docs/Manual_Registro_de_salidas.pdf`: manual de usuario (profesorado, admin y puesta en marcha).
- `LICENSE`: GNU GPL v3.

## Puesta en marcha
1. Crea una hoja de cálculo de Google y abre **Extensiones → Apps Script**.
2. Pega `Code.gs` y crea un archivo HTML llamado `Index` con el contenido de `Index.html`.
3. Ejecuta `preparar` una vez (crea las pestañas y el usuario `admin`; la contraseña inicial aparece una sola vez).
4. **Implementar → Nueva implementación → Aplicación web**.

Consulta el manual (Anexos A, B y C) para la administración, el inicio de curso y la importación desde Séneca.

## Licencia
Software libre bajo GNU GPL v3 o posterior. © 2026 IES Campanillas (Málaga).
