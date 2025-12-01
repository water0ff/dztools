@"
## Descripción
<!-- Describe los cambios realizados -->

## Tipo de cambio
- [ ] Nueva funcionalidad (non-breaking change)
- [ ] Bug fix (non-breaking change)
- [ ] Breaking change (cambia funcionalidad existente)
- [ ] Refactorización
- [ ] Documentación

## Compatibilidad PS5.0
- [ ] He verificado que funciona en PowerShell 5.0
- [ ] No uso cmdlets de versiones superiores
- [ ] He probado en Windows 10/Server 2016

## Testing
- [ ] He ejecutado las pruebas existentes
- [ ] He añadido nuevas pruebas
- [ ] He probado manualmente las funcionalidades

## Checklist
- [ ] Mi código sigue el estilo del proyecto
- [ ] He comentado mi código donde sea necesario
- [ ] No he dejado código de debugging
- [ ] He actualizado la documentación
- [ ] No rompe funcionalidades existentes

## Screenshots (si aplica)
<!-- Capturas de pantalla de los cambios -->

## Contexto adicional
<!-- Información adicional sobre el PR -->
"@ | Out-File -FilePath ".github/PULL_REQUEST_TEMPLATE.md" -Encoding UTF8