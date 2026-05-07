# Smart-Document-Filling-Agent 🤖📄
### De 4 horas de trabajo manual a 3 minutos con un solo clic.

**Smart-Document-Filling-Agent** es una solución de automatización local diseñada para eliminar el cuello de botella del ingreso manual de datos. Este agente transforma bases de datos (Excel/CSV) en documentos finales profesionales con precisión quirúrgica, preservando formatos originales y garantizando la privacidad total de la información.

---

## 🔴 El Problema: "La Trampa Administrativa"

Muchas empresas pierden cientos de horas anuales en tareas repetitivas:

1. **Digitación Manual:** Copiar datos de extractos bancarios o bases de datos a formatos de recibos, contratos o facturas.
2. **Error Humano:** Riesgo constante de errores en montos, fechas y numeración correlativa.
3. **Privacidad:** Muchas soluciones en la nube exponen datos financieros sensibles a terceros.

---

## 🟢 La Solución: "Automatización Pixel-Perfect"

Este agente de escritorio utiliza **Python y COM Automation** para interactuar directamente con Microsoft Excel, permitiendo:

- **Generación Masiva:** Procesar cientos de registros en segundos.
- **Fidelidad de Formato:** A diferencia de otros sistemas, este agente respeta logos, celdas combinadas y estilos de plantillas pre-existentes.
- **Ejecución Offline:** Los datos nunca salen del computador del cliente.

---

## 📊 Caso de Éxito: Optimización de Tesorería (Logística Delfín S.A.S.)

En este escenario real, el agente automatizó la gestión de caja menor:

| Métrica | Antes | Después |
|---|---|---|
| Tiempo por ciclo contable | 2 – 4 horas | ~3 minutos |
| Errores de numeración | Frecuentes | **0** (automatizado) |
| Errores de monto | Ocasionales | **0** (directo del extracto) |
| Privacidad de datos | Riesgo en la nube | **100% local** |

> **Impacto:** Reducción del **98%** en el tiempo de procesamiento y eliminación total de errores de numeración.

---

## 🛠️ Stack Tecnológico

| Componente | Tecnología |
|---|---|
| **Core** | Python 3.10+ |
| **Data Processing** | `pandas` — Limpieza y filtrado inteligente |
| **Automation** | `win32com` — Interacción nativa con MS Excel |
| **UI** | `customtkinter` — Interfaz moderna y minimalista |
| **Distribution** | PyInstaller — Empaquetado autónomo en `.exe` para Windows |

---

## ⚙️ Arquitectura

```
[Excel/CSV de entrada]  →  [Agente]  →  [Documento maestro .xlsx]
                               ↕
                    ┌──────────────────────┐
                    │  Filtrado inteligente │  ← Excluye transferencias,
                    │  Reglas de negocio    │    seg. social, cuotas
                    │  Numeración auto      │    y aplica conceptos custom
                    │  COM pixel-perfect    │
                    └──────────────────────┘
```

**Procesamiento 100% local** — Cero exposición de datos financieros a servidores externos.

---

## 🚀 Versatilidad: ¿Dónde más se puede aplicar?

- **Contabilidad:** Generación de recibos, cuentas de cobro y conciliaciones.
- **Recursos Humanos:** Creación masiva de certificados y contratos laborales.
- **Ventas:** Generación de cotizaciones y órdenes de pedido personalizadas.
- **Logística:** Reportes de entrega, manifiestos y guías de transporte.

---

## 🔒 Seguridad y Privacidad

> ❌ No requiere conexión a internet.  
> ❌ No envía datos a APIs externas.  
> ✅ Toda la información permanece en el equipo del cliente.  
> ✅ Compatible con políticas de confidencialidad corporativa (DIAN, Supersociedades).

---

## 📋 Requisitos

- Windows 10 / 11
- Microsoft Excel 2016 o superior (instalado localmente)
- Python 3.10+ *(solo si se ejecuta el script directamente)*

---

## 📈 ¿Buscas optimizar la rentabilidad de tu negocio?

Si tu equipo está atrapado en tareas repetitivas, estás perdiendo dinero.  
Ofrezco **Auditorías de Procesos gratuitas (15 min)** para identificar oportunidades de automatización.

---

**Kevin Seryeit Castañeda Aldana**  
*Ingeniero de Sistemas | Especialista en Automatización de Procesos e IA*

[![LinkedIn](https://img.shields.io/badge/LinkedIn-Contactar-blue?style=flat&logo=linkedin)](https://www.linkedin.com/in/kevinseryeit)
