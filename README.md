# SIREB-VIPI

## Sistema de Registro y Estadísticas de Biblioteca - UNELLEZ San Carlos (VIPI)

## Descripción del Proyecto

SIREB-VIPI es una aplicación web de arquitectura Serverless desarrollada para modernizar la Biblioteca "Ramón Villegas Izquiel". Permite la gestión digital de préstamos de libros, control automático de morosidad y un portal de autogestión estudiantil para la emisión instantánea de Solvencias de Biblioteca en formato PDF.

## Equipo de Desarrollo

* Victor Bello
* Neudy Chacón
* Nelson Chacón
* Anibal Santana
* Angel Salazar

## Stack Tecnológico y Arquitectura

El proyecto está construido bajo un modelo Cliente-Servidor integrado al ecosistema de Google Workspace, lo que garantiza cero costos de hosting y alta disponibilidad:

* **Frontend (Cliente):** HTML5, CSS3, Tailwind / Bootstrap 5, JavaScript Vanilla.
* **Backend (Servidor):** Google Apps Script (JavaScript V8 Engine).
* **Base de Datos:** Google Sheets (actuando como DB Relacional).
* **APIs Externas Integradas:** Google Drive API (Generación de PDFs) y Google Docs API (Plantillas).

## Estructura del Repositorio

* **/frontend:** Contiene las interfaces gráficas (Catálogo Público, Panel Administrativo y Portal Estudiantil).
* **/backend:** Scripts .js que controlan las rutas, el CRUD de la base de datos y los Triggers de multas.
* **/docs:** Manual de usuario del sistema.

Desarrollado para la Feria de Software - UNELLEZ 2026
