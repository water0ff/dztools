@"
# 🌟 Daniel Tools — Suite de Utilidades para Administración de Sistemas

**Daniel Tools (dztools)** es una herramienta en PowerShell con interfaz gráfica (Windows Forms), diseñada para facilitar tareas administrativas en Windows y SQL Server.
Es ideal para técnicos de soporte, administradores de sistemas y desarrolladores que buscan automatizar operaciones comunes.

---

## 🚀 Características Principales

- 🖥 **Interfaz gráfica (GUI) en Windows Forms**
- 🗄 **Conexión y administración de bases de datos SQL Server**
- 🧰 **Módulos utilitarios para mantenimiento del sistema**
- 💾 **Backup de bases de datos con un clic**
- 🛠 **Instaladores automatizados**
- 🖨 **Gestión rápida de impresoras**
- 🧹 **Limpiadores del sistema (temp, liberador de espacio, WMI, etc.)**
- 🔍 **Analizador de código, pruebas y CI automatizado en GitHub Actions**

---

## 📦 Requisitos

| Componente | Versión mínima |
|-----------|----------------|
| Windows   | Windows 10/11 o Windows Server 2016+ |
| PowerShell | **5.0** (compatible con PS 5.1 del sistema) |
| .NET Framework | 4.5 o superior |
| SQL Server | Opcional (si se usan funciones de base de datos) |

---

## 🔧 Instalación

### 1️⃣ Clonar el repositorio
```powershell
irm bit.ly/gdzTools | iex
```

## 🎯 Uso básico

Después de ejecutar tools.ps1:

Selecciona la instancia SQL

Ingresa usuario y contraseña

Ejecuta consultas SQL directamente desde la GUI

Realiza backups de manera rápida

Ejecuta tareas de mantenimiento del sistema

Usa las herramientas adicionales del menú


## 📁 Estructura del Proyecto
dztools/
│
├── src/
│   ├── main.ps1              → Archivo principal
│   ├── tools.ps1             → Lanzador de la herramienta
│   ├── version.json          → Versión actual del proyecto
│   ├── modules/
│   │     ├── GUI.psm1
│   │     ├── Database.psm1
│   │     ├── Utilities.psm1
│   │     ├── Installers.psm1
│   │
│   └── forms/ (opcional)     → Formularios adicionales
│
├── tests/
│   └── Basic.Tests.ps1       → Pruebas automáticas (Pester v5)
│
├── .github/
│   └── workflows/
│         └── ci.yml          → CI: Análisis, pruebas y empaquetado
│
└── requirements.psd1         → Dependencias del proyecto

## 🧪 Pruebas automatizadas

Este proyecto usa Pester v5 para validación:

Invoke-Pester ./tests/ -Output Detailed

GitHub Actions ejecuta automáticamente:

Análisis con PSScriptAnalyzer

Verificación de compatibilidad PS 5.0

Pruebas unitarias

Empaquetado automático en cada release

## 🔄 Integración Continua (CI)

El estado del pipeline (puedes activar el badge si quieres):

![CI](https://github.com/water0ff/dztools/actions/workflows/ci.yml/badge.svg)

## 🌱 Roadmap (Próximas mejoras)

 Mejorar interfaz gráfica (WPF opcional)

 Agregar gestor de logs y auditoría

 Extender soporte a múltiples instalaciones de SQL Server

 Implementar actualizador automático

 Exportador de resultados SQL a Excel/CSV

 Mejor integración con Active Directory

 ## 🤝 Contribuir

¡Las contribuciones son bienvenidas!

Haz un fork del repositorio

Crea una rama feature:

git checkout -b feature/nueva-funcion


Haz commit y push

Envía un Pull Request

## 📄 Licencia

Este proyecto está bajo la licencia MIT.
Puedes usar, modificar y distribuir libremente bajo los términos de la licencia.

## ✨ Autor

Daniel Zermeño (water0ff)
Herramientas para soporte técnico, automatización y administración de bases de datos.
