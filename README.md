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
| Errores de numeración | Frecuentes | 0 |
| Errores de monto | Ocasionales | 0 |
| Formato / imágenes preservados | Manual | 100% automático |

> **Impacto:** Reducción del **98%** en el tiempo de procesamiento y eliminación total de errores de numeración.

---

## 🛠️ Stack Tecnológico

| Capa | Tecnología | Rol |
|---|---|---|
| Core | Python 3.10+ | Lógica central y portabilidad |
| Data Processing | `pandas` | Limpieza, filtrado y ordenamiento |
| Document Engine | `win32com` | Interacción nativa con MS Excel (preserva imágenes y estilos) |
| UI | `customtkinter` | Interfaz moderna y minimalista |
| Distribution | `PyInstaller` | Empaquetado autónomo en `.exe` para Windows |

---

## 🚀 Versatilidad: ¿Dónde más se puede aplicar?

- **Contabilidad:** Generación de recibos, cuentas de cobro y conciliaciones.
- **Recursos Humanos:** Creación masiva de certificados y contratos laborales.
- **Ventas:** Generación de cotizaciones y órdenes de pedido personalizadas.
- **Logística:** Listas de despacho, actas de entrega, manifiestos de carga.

---

## ⚙️ Instalación y Uso (Modo Desarrollo)

```bash
# 1. Clonar el repositorio
git clone https://github.com/Kecasta/Smart-Document-Filling-Agent.git
cd Smart-Document-Filling-Agent

# 2. Instalar dependencias
pip install -r requirements.txt

# 3. Ejecutar el agente
python caja_menor_pro.py
```

> **Requisito del sistema:** Microsoft Excel instalado (requerido por `win32com` para la generación de documentos).

### Generar el ejecutable .exe

```bash
pyinstaller --noconfirm --onefile --windowed \
  --name "Agente_Caja_Menor" --icon "recibo.ico" \
  --add-data "Formato_Caja_Menor_Logistica_Delfin.xlsx;." \
  --add-data "recibo.ico;." --add-data "recibo.png;." \
  --hidden-import "win32com" --hidden-import "win32com.client" \
  --hidden-import "pythoncom" --hidden-import "pywintypes" --hidden-import "winreg" \
  --collect-data "customtkinter" --collect-data "tkcalendar" \
  caja_menor_pro.py
```

---

## 🔒 Seguridad y Privacidad

- ❌ Sin conexión a internet requerida.
- ❌ Sin envío de datos a APIs o servidores de terceros.
- ✅ Procesamiento 100% local en el equipo del cliente.
- ✅ Compatible con políticas corporativas de confidencialidad financiera.

---

## 📈 ¿Buscas optimizar la rentabilidad de tu negocio?

Si tu equipo está atrapado en tareas repetitivas, estás perdiendo dinero. Ofrezco **Auditorías de Procesos gratuitas (15 min)** para identificar oportunidades de automatización.

**Kevin Seryeit Castañeda Aldana**  
Ingeniero de Sistemas | Especialista en Automatización de Procesos e IA

[![LinkedIn](https://img.shields.io/badge/LinkedIn-Connect-blue?style=flat&logo=linkedin)](https://linkedin.com)
[![GitHub](https://img.shields.io/badge/GitHub-Kecasta-black?style=flat&logo=github)](https://github.com/Kecasta)

---

*Desarrollado con Python 3.x | Windows 10/11 | Microsoft Excel*
