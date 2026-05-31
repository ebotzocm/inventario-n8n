# Sistema de Gestión de Inventario Automatizado con n8n

## Descripción del Proyecto

Este proyecto consiste en el desarrollo de un sistema web de gestión de inventario automatizado utilizando:

* Python Flask
* PostgreSQL
* n8n
* Docker Compose
* HTML, CSS y JavaScript

El sistema permite administrar productos, ventas, alertas, auditoría, reportes y automatizaciones de inventario mediante workflows desarrollados en n8n.

El proyecto fue desarrollado como parte del curso de Ingeniería de Software y Automatización de Procesos.

---

# Características principales

## Gestión de productos

* Agregar productos
* Editar productos
* Eliminar productos
* Control de stock mínimo
* Búsqueda y paginación

## Gestión de ventas

* Registro de ventas
* Validación automática de stock
* Actualización automática de inventario
* Historial de ventas

## Sistema de alertas

* Alertas automáticas de bajo stock
* Prevención de alertas duplicadas
* Estados de alertas:

  * Pendiente
  * Reabastecida
  * Cerrada

## Reportes

* Reportes exportables en Excel
* Ventas diarias
* Ventas semanales
* Inventario crítico
* Alertas pendientes

## Auditoría

* Registro de eventos automáticos
* Registro de ventas
* Registro de reabastecimientos
* Registro de generación de reportes

## Dashboard

* Estadísticas generales
* Productos críticos
* Ventas por producto
* Stock actual
* Alertas pendientes

## Gestión de usuarios

Roles disponibles:

* Administrador
* Inventario
* Ventas
* Supervisor

---

# Tecnologías utilizadas

| Tecnología     | Uso            |
| -------------- | -------------- |
| Python Flask   | Aplicación web |
| PostgreSQL     | Base de datos  |
| n8n            | Automatización |
| Docker Compose | Contenedores   |
| HTML/CSS/JS    | Frontend       |
| Bootstrap      | Estilos        |
| Chart.js       | Gráficas       |

---

# Arquitectura del sistema

El sistema funciona mediante una arquitectura local basada en contenedores Docker:

Usuario → Flask → n8n → PostgreSQL

n8n se encarga de:

* Procesamiento automático
* Validaciones
* Alertas
* Reportes
* Auditoría

---

# Requisitos

## Software requerido

* Docker Desktop
* Python 3.10 o superior
* Git
* Navegador web

---

# Instalación del proyecto

## 1. Clonar repositorio

```bash
git clone https://github.com/usuario/inventario-n8n.git
cd inventario-n8n
```

---

## 2. Crear entorno virtual

```bash
python -m venv .venv
```

---

## 3. Activar entorno virtual

### Windows

```bash
.venv\Scripts\activate
```

### Linux / Mac

```bash
source .venv/bin/activate
```

---

## 4. Instalar dependencias

```bash
pip install -r requirements.txt
```

---

# Configuración de Docker

## Ejecutar contenedores

```bash
docker compose up -d
```

---

## Verificar contenedores

```bash
docker ps
```

Debe aparecer:

* PostgreSQL
* n8n

---

# Configuración de la base de datos

## Acceder a PostgreSQL

```bash
docker exec -it inv_postgres psql -U n8n -d inventario
```

---

## Ver tablas

```sql
\dt
```

---

# Ejecución del sistema web

## Iniciar Flask

```bash
python app.py
```

---

# Acceso al sistema

## Aplicación web

```text
http://127.0.0.1:5000
```

## n8n

```text
http://localhost:5678
```

---

# Usuarios del sistema

## Administrador

Usuario:

```text
admin
```

Contraseña:

```text
admin123
```

---

# Workflows implementados en n8n

## WF1 – Registro de ventas

Funciones:

* Recepción de venta
* Validación de producto
* Registro en base de datos

---

## WF2 – Procesamiento de stock y alertas

Funciones:

* Actualización automática de inventario
* Verificación de stock crítico
* Generación de alertas
* Envío de correo
* Auditoría

---

## WF3 – Reportes automáticos

Funciones:

* Generación de reportes
* Exportación Excel/CSV
* Registro de auditoría

---

# Uso del sistema

## Productos

Permite:

* Registrar productos
* Editar stock
* Definir stock mínimo

---

## Ventas

Permite:

* Registrar ventas
* Actualizar inventario automáticamente

---

## Alertas

Permite:

* Visualizar productos con bajo stock
* Resolver alertas

---

## Reportes

Permite:

* Descargar reportes Excel
* Consultar historial de reportes

---

## Auditoría

Permite:

* Consultar eventos del sistema
* Ver trazabilidad de operaciones

---

# Pruebas realizadas

## Prueba de registro de ventas

Resultado:

* Venta registrada correctamente
* Inventario actualizado

---

## Prueba de bajo stock

Resultado:

* Generación automática de alerta
* Registro en auditoría

---

## Prueba de reportes

Resultado:

* Archivo Excel generado correctamente

---

## Prueba de reabastecimiento

Resultado:

* Stock actualizado
* Alerta cerrada automáticamente

---

# Estructura del proyecto

```text
inventario_web/
│
├── app.py
├── config.py
├── requirements.txt
├── docker-compose.yml
│
├── templates/
├── static/
│
├── workflows/
│   ├── WF1_RegistroVentas.json
│   ├── WF2_StockAlertas.json
│   └── WF3_Reportes.json
│
├── docs/
│   ├── Arquitectura.png
│   ├── WorkflowPrincipal.png
│   └── DocumentacionTecnica.pdf
│
└── reportes_generados/
```

---

# Seguridad implementada

* Login con roles
* Restricción por permisos
* Validación de formularios
* Manejo de errores
* Auditoría de eventos

---

# Consideraciones importantes

* El sistema funciona completamente en entorno local.
* Docker debe estar encendido.
* n8n debe estar ejecutándose para automatizaciones.
* PostgreSQL debe estar activo para consultas.

---

# Autor

Proyecto desarrollado por:

Eddy Botzoc
Ingeniería en Sistemas
Universidad Mariano Gálvez

---

# Licencia

Proyecto académico con fines educativos.
