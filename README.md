# 🔐 Validador de Coherencia - Matriz LA/FT/FPADM

Sistema de validación integral para matrices de riesgo **Lavado de Activos (LA) / Financiación del Terrorismo (FT) / Financiación de Proliferación de Armas de Destrucción Masiva (FPADM)** basado en la **Guía de Identificación de Riesgos LAFT V3** de la Superintendencia Financiera de Colombia (SFC).

## 🎯 Propósito

Validar la coherencia estructural entre **Riesgos, Causas y Controles** en matrices de riesgos de conformidad, asegurando:

- ✓ Exposiciones claras a LA/FT/FPADM
- ✓ Causas con escenarios específicos de conducta delictiva
- ✓ Controles con 4 componentes: Alertamiento, Análisis, Reporte, Toma de Decisiones
- ✓ Coherencia Control ≤ Causa (especificidad)
- ✓ Cumplimiento de criterios SFC para indicador de coherencia ≥ 80%

---

## 📋 Requisitos

- **Node.js** v14+ ([descargar](https://nodejs.org/))
- **OpenAI API Key** con acceso a GPT Nano (gpt-4o-mini) ([obtener aquí](https://platform.openai.com/api-keys))
- Navegador moderno (Chrome, Edge, Firefox, Safari)

---

## ⚙️ Instalación

### 1. Clonar o descargar el proyecto

```bash
cd aris-coherencia-matriz
```

### 2. Instalar dependencias

```bash
npm install
```

### 3. Configurar API Key

Crea un archivo `.env` en la raíz del proyecto:

```bash
# Opción A: Crear desde .env.example
cp .env.example .env
```

Luego edita `.env` y reemplaza `tu_api_key_aqui` con tu API key de OpenAI:

```env
OPENAI_API_KEY=sk-proj-xxxxxxxxxxxxxxxxxxxxx
PORT=3000
```

---

## 🚀 Ejecución

### Iniciar el servidor

```bash
npm start
```

Deberías ver:

```
╔════════════════════════════════════════════════════╗
║  🔐 Validador de Coherencia - Matriz LA/FT/FPADM  ║
╚════════════════════════════════════════════════════╝

📍 Servidor corriendo en: http://localhost:3000
🔑 API Key: ✓ Configurada

Abre http://localhost:3000 en tu navegador.
```

### Acceder a la interfaz

Abre en tu navegador: **http://localhost:3000**

---

## 📝 Estructura de Entrada

### RIESGO
**Formato esperado:**
```
La posibilidad de que [ENTIDAD] sea utilizada como instrumento para [LA/FT/FPADM], 
derivado de [delitos subyacentes específicos]...
```

**Ejemplo:**
```
Posibilidad de que Seguros Bolívar sea utilizada para lavado de activos provenientes 
de narcotráfico internacional, a través de clientes vinculados a redes criminales que 
estructuran pagos de primas de seguros en efectivo.
```

### CAUSA
**Estructura obligatoria:**
```
[SUJETO] [ACCIÓN] [OBJETO DETERMINADO] [DELITO FUENTE] [TIPOLOGÍA]
```

- **SUJETO:** ¿Quién? (Clientes PNJ, intermediarios, corredores, etc.)
- **ACCIÓN:** Verbo activo (realiza, contrata, transfiere, paga, etc.)
- **OBJETO:** Activo ESPECÍFICO (primas en efectivo, cuotas, transferencias internacionales, pólizas, no genérico como "fondos")
- **DELITO FUENTE:** Acto delictivo presumido (narcotráfico, corrupción, minería ilegal, etc.)
- **TIPOLOGÍA:** Método o circunstancia (pitufeo, estructuración, empresas fachada, etc.)

**Ejemplo:**
```
Clientes personas naturales vinculadas a redes criminales realizan pagos de primas 
de seguros en efectivo mediante múltiples pólizas menores (pitufeo) para disimular 
origen ilícito de narcotráfico internacional.
```

### CONTROL
**Estructura obligatoria (4 componentes):**
```
[ALERTAMIENTO]. Mediante [ANÁLISIS]. Como resultado [REPORTE]. Para [TOMA DE DECISIONES].
```

- **ALERTAMIENTO:** Patrones/hechos que disparan la alerta (elementos de causa)
- **ANÁLISIS:** "Mediante" examen detallado para confirmar operación inusual/sospechosa
- **REPORTE:** "Como resultado" documentación interno/externo
- **TOMA DE DECISIONES:** "Para" medidas (restricción, cancelación, seguimiento, etc.)

**Ejemplo:**
```
Sistema de alertas identifica pagos de primas en efectivo > USD 5,000 de clientes 
nuevos sin historial. Mediante análisis de patrones de comportamiento e identificación 
de posible vinculación con redes criminales. Como resultado, elevación a Oficina de 
Cumplimiento para validación y generación de reporte a UIAF si corresponde. Para 
implementar restricciones de canal de pago y seguimiento reforzado de las pólizas.
```

---

## 🔍 Validación Realizada por GPT Nano

El sistema evalúa:

### RIESGO ✓
- [ ] ¿Expresa claramente exposición a LA/FT/FPADM?
- [ ] ¿Especifica delitos subyacentes o contexto?

### CAUSA ✓
- [ ] ¿Tiene sujeto activo explícito?
- [ ] ¿Usa verbo activo (no pasivo)?
- [ ] ¿El objeto es DETERMINADO o indeterminado?
- [ ] ¿Especifica delito fuente?
- [ ] ¿Describe tipología/método?
- [ ] ¿Incluye ≥3 factores de riesgo?
- [ ] ¿Modela conducta delictiva (no falla operativa)?

### CONTROL ✓
- [ ] ¿Incluye 4 componentes (alertamiento, análisis, reporte, decisión)?
- [ ] ¿El alertamiento menciona elementos de la causa?
- [ ] ¿Está diseñado para mitigar conducta delictiva?
- [ ] ¿Es coherente con la causa?

### COHERENCIA ✓
- [ ] Control ≤ Causa en especificidad
- [ ] Trío forma un flujo lógico completo

---

## 📊 Respuesta del Sistema

El servidor retorna:

```json
{
  "analysis": {
    "overall_coherence": "Alto/Medio/Bajo",
    "riesgo_assessment": {
      "valid": true/false,
      "score": 1-3,
      "findings": "..."
    },
    "causa_assessment": {
      "valid": true/false,
      "score": 1-3,
      "sujeto": "detectado/no detectado",
      "accion": "detectada/no detectada",
      "objeto": "determinado/indeterminado/no detectado",
      "delito_fuente": "especificado/no especificado",
      "tipologia": "descrita/no descrita",
      "factores_riesgo": "cantidad detectada",
      "findings": "..."
    },
    "control_assessment": {
      "valid": true/false,
      "score": 1-3,
      "componentes": ["alertamiento", "análisis", "reporte", "decisión"],
      "coherencia_con_causa": "alta/media/baja",
      "findings": "..."
    },
    "strengths": ["..."],
    "weaknesses": ["..."],
    "recommendations": ["..."]
  },
  "recommendations": "Resumen de fortalezas y mejoras"
}
```

---

## 🔧 Estructura del Proyecto

```
aris-coherencia-matriz/
├── server.js              # Backend Express + OpenAI
├── index.html             # Frontend interactivo
├── package.json           # Dependencias Node.js
├── .env                   # Variables de entorno (NO VERSIONADO)
├── .env.example           # Plantilla .env
├── README.md              # Esta guía
├── requerimiento.md       # Especificación SFC detallada
└── bd-matriz-estructura.xlsx  # Datos de 110 registros analizados
```

---

## 🛠️ Solución de Problemas

### Error: "API key no configurada"
- Verifica que exista el archivo `.env` en la raíz
- Confirma que `OPENAI_API_KEY=sk-proj-...` esté presente

### Error: "API key inválida"
- Verifica que la clave sea correcta en platform.openai.com
- Comprueba que tenga créditos disponibles

### Error: "No se pudo conectar"
- Verifica que el servidor esté corriendo (`npm start`)
- Comprueba que http://localhost:3000 sea accesible

### Respuesta vacía o error 500
- Revisa la consola del servidor (`npm start`) para detalles
- Verifica conexión a Internet (para llamada a OpenAI)
- Comprueba logs de error

---

## 📚 Referencias

- **Guía de Identificación de Riesgos LAFT V3** (SFC, Junio 2026)
- **Estatuto Orgánico de la SFC** (Decreto 333/2024)
- **Metodología de Evaluación de Riesgos LAFT** (SFC)
- **Escala de Calificación 1-3 (Débil/Mejora/Adecuado)**

---

## 📞 Contacto & Soporte

- **Usuario:** Javier Ubaque (javier.ubaque@segurosbolivar.com)
- **Empresa:** Seguros Bolívar S.A.
- **Marco Legal:** Superintendencia Financiera de Colombia (SFC)

---

## 📄 Licencia

ISC - Uso interno Seguros Bolívar

---

**Última actualización:** Junio 2026  
**Versión:** 1.0.0
