# Registro de salidas

Aplicación web de Google Apps Script para registrar las salidas del alumnado al baño, con hoja de cálculo de Google como base de datos.
Desarrollada para el IES Campanillas (Málaga).

## Contenido
- `Code.gs`: servidor (Apps Script).
- `Index.html`: interfaz web.
- `docs/Manual_Registro_de_salidas.pdf`: manual de usuario (profesorado, admin y puesta en marcha).
- `LICENSE`: GNU GPL v3.

## Funciones
- Registro de salidas al baño con cuenta atrás, máximo semanal, incidencias y observaciones.
- El alumnado se apunta solo con su código o **buscando su nombre** (se muestra el código).
- **Ficha del alumno** con resumen, observaciones, incidencias y últimas 10 semanas.
- **Contraseñas cifradas** (SHA-256 con sal) cuando las guarda el sistema; el administrador puede **restablecer la contraseña de un profesor al DNI**.
- El administrador puede **borrar el historial hasta una fecha** (vista previa y confirmación) e importar las listas de Séneca.

## Puesta en marcha
1. Crea una hoja de cálculo de Google y abre **Extensiones → Apps Script**.
2. Pega `Code.gs` y crea un archivo HTML llamado `Index` con el contenido de `Index.html`.
3. Ejecuta `preparar` una vez (crea las pestañas y el usuario `admin`; la contraseña inicial aparece una sola vez).
4. **Implementar → Nueva implementación → Aplicación web**.

Consulta el manual (Anexos A, B y C) para la administración, el inicio de curso y la importación desde Séneca.

## Licencia
Software libre bajo GNU GPL v3 o posterior. © 2026 IES Campanillas (Málaga).
