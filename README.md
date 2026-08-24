
# 📑 Auto-corrección de Prácticos – Metodología de la Investigación

Aplicación en **Streamlit** para que los alumnos suban sus prácticos, se corrijan automáticamente según rúbrica básica y reciban sus notas y comentarios por correo electrónico. **La devolución se envía al correo de quien entrega y una copia a la cátedra.**

---

## 🚀 Uso
1. Clonar este repositorio.
2. Instalar dependencias:

```bash
pip install -r requirements.txt
```

En Streamlit Cloud, configurar secretos:

- `SENDGRID_API_KEY`, `SENDER_EMAIL`, `SENDER_NAME`, `REPLY_TO`
- o fallback Gmail: `EMAIL_USER`, `EMAIL_PASS`
- copia a la cátedra: `TEACHER_EMAIL` (recomendado) y/o `TEACHER_BCC`

Si no hay `TEACHER_EMAIL` ni `TEACHER_BCC`, la copia va a `investigacion@uccuyo.edu.ar`.

## 🧪 Prácticos y rúbricas evaluadas

- **Práctico Nº 1 — IA en la escritura del proyecto**  
- **Práctico Nº 2 — Establecimiento de Métodos de Recolección de Datos y Tipos de Muestreos. Tamaño de la Muestra**  
  - Evalúa: tipo de muestreo y justificación, instrumentos y adecuación, validez/fiabilidad, tamaño muestral.
- **Práctico Nº 3 — Operacionalización de Variables y Determinación de Métodos de Análisis de Datos**  
  - Evalúa: cuadro de operacionalización, métodos de análisis, validación, ética.
- **Práctico Nº 4 — Introducción + Marco teórico + Búsqueda (≈500 palabras en total)**
- **Módulo 5 — Mendeley: citas en Word y bibliografía**
- **Módulo 6 — Estilos de Word e índice automático**
- **Práctico Nº 7 — Análisis cuantitativo**
- **Práctico Nº 8 — Análisis cualitativo**
