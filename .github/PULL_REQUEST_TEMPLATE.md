## Resumen breve
- Issue/Tarea relacionada: <!-- enlaza al issue o describe el contexto -->
- Alcance principal (marca lo que aplique):
  - [ ] Interfaz GUI (Windows Forms)
  - [ ] Funciones SQL Server (consultas/backup)
  - [ ] Utilidades de sistema/instaladores
  - [ ] Configuración de CI/CD o tooling

## Descripción detallada
<!-- Explica el problema, la solución propuesta y los cambios principales. Incluye cualquier decisión técnica relevante. -->

## Tipo de cambio
- [ ] Nueva funcionalidad (sin romper compatibilidad)
- [ ] Corrección de bug
- [ ] Breaking change (cambia comportamiento existente)
- [ ] Refactorización interna
- [ ] Documentación o guías de uso

## Compatibilidad con PowerShell 5.0
- [ ] Ejecutado en Windows 10/Server 2016 con PowerShell 5.0/5.1
- [ ] Sin cmdlets ni APIs exclusivos de versiones superiores
- [ ] Validé dependencias externas (módulos, .NET) disponibles en esos entornos

## Pruebas y calidad
- [ ] `Invoke-Pester ./tests/ -Output Detailed`
- [ ] `Invoke-ScriptAnalyzer` sin violaciones críticas
- [ ] Pruebas manuales de las rutas principales impactadas
- [ ] Capturas actualizadas si hay cambios visibles en la GUI

## Documentación y versionado
- [ ] Actualicé README/ayuda en línea si el comportamiento cambia
- [ ] Ajusté `CHANGELOG.md` o notas de release si aplica
- [ ] Revisé/actualicé `src/version.json` cuando cambia la versión de la herramienta

## Checklist final
- [ ] El código sigue el estilo y convenciones del proyecto
- [ ] Sin código de depuración ni trazas temporales
- [ ] No rompe funcionalidades existentes (probado en escenarios relevantes)
- [ ] CI local o remoto revisado cuando corresponde

## Contexto adicional
<!-- Información extra, riesgos conocidos o pasos de despliegue. -->