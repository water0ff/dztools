@"
# Changelog

Todos los cambios notables en este proyecto serán documentados en este archivo.

El formato está basado en [Keep a Changelog](https://keepachangelog.com/es-ES/1.0.0/),
y este proyecto adhiere a [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- Estructura modular del proyecto
- Módulo GUI para creación de formularios
- Módulo Database para operaciones SQL
- GitHub Actions para CI/CD
- Pruebas con Pester

### Changed
- Reorganización del código en módulos
- Mejor manejo de dependencias

### Fixed
- Problemas de compatibilidad con PS5.0

## [0.1.0] - 2024-01-15
### Added
- Proyecto inicial con todas las funcionalidades
- Interfaz gráfica completa
- Conexión a SQL Server
- Herramientas de instalación
- Gestión de impresoras
- Sistema de backup

## Cómo actualizar el changelog

### Added
Para nuevas funcionalidades.

### Changed
Para cambios en funcionalidades existentes.

### Deprecated
Para funcionalidades que serán removidas en versiones futuras.

### Removed
Para funcionalidades removidas.

### Fixed
Para corrección de bugs.

### Security
En caso de vulnerabilidades.
"@ | Out-File -FilePath "CHANGELOG.md" -Encoding UTF8