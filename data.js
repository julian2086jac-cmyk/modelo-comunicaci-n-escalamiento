/**
 * data.js — Fuente maestra del formulario
 * Modelo Corporativo de Comunicación y Escalamiento
 *
 * Generado desde: Plantilla_Forms_Modelo_Comunicacion_Escalamiento.xlsx
 *
 * Para actualizar: edita únicamente este archivo.
 * El formulario se construye dinámicamente desde aquí.
 */

const PROCESOS = [
  "Calidad Control",
  "Calidad Aseguramiento",
  "Técnico Regulatorio",
  "Ambiental",
  "Excelencia Corporativa",
  "Sistemas de Gestión"
];

const TRANSVERSAL_KEY = "Transversal";

const SECCIONES = {
  "Calidad Control": {
    titulo: "Calidad Control",
    icono: "🔬",
    situaciones: [
      "Resultados analíticos fuera de especificación",
      "Retrasos en la liberación de materiales o productos que impactan la programación",
      "Equipos críticos del laboratorio fuera de servicio o con desviaciones que comprometan la confiabilidad de los resultados o la continuidad de la operación",
      "Tendencias analíticas desfavorables o resultados atípicos que requieran investigación",
      "Desviaciones analíticas que comprometan la confiabilidad de los resultados"
    ],
    preguntasAbiertas: [
      "¿Qué otra situación importante genera comunicación, coordinación o escalamiento en este proceso?",
      "¿Cuál considera que hoy es la principal dificultad de comunicación o escalamiento de este proceso?"
    ]
  },

  "Calidad Aseguramiento": {
    titulo: "Calidad Aseguramiento",
    icono: "✅",
    situaciones: [
      "Verificación de disponibilidad de documentación vigente antes del inicio de producción",
      "Verificación del cumplimiento de prerrequisitos para el inicio de producción",
      "Disponibilidad de estándares y condiciones requeridas para nuevas producciones",
      "Acompañamiento técnico durante primeras producciones",
      "Desviaciones identificadas durante primeras producciones",
      "Ajustes de proceso requeridos durante nuevos proyectos",
      "Evaluación de riesgos para la calidad del producto",
      "Cambios que impactan el estado validado de procesos, equipos o sistemas",
      "Gestión de Cambios (Change Control)",
      "Evaluación del impacto de cambios en materias primas, procesos o equipos",
      "Decisiones sobre disposición de producto derivadas de primeras producciones (liberar, retener o rechazar)",
      "Identificación y comunicación de riesgos durante nuevos proyectos o escalamientos industriales",
      "Escalamiento de riesgos de calidad al negocio"
    ],
    preguntasAbiertas: [
      "¿Qué otra situación importante genera comunicación, coordinación o escalamiento en este proceso?",
      "¿Cuál considera que hoy es la principal dificultad de comunicación o escalamiento de este proceso?"
    ]
  },

  "Técnico Regulatorio": {
    titulo: "Técnico Regulatorio",
    icono: "📋",
    situaciones: [
      "Cambios regulatorios con impacto en productos, instalaciones, registros o certificaciones existentes",
      "Observaciones o requerimientos de autoridades sanitarias",
      "Riesgos de incumplimiento regulatorio identificados durante revisiones técnicas",
      "Retrasos en la aprobación de artes que impacten lanzamientos o producción",
      "Retrasos en trámites regulatorios que impacten el lanzamiento o la continuidad de la producción",
      "Consultas técnicas de clientes que requieran análisis transversal"
    ],
    preguntasAbiertas: [
      "¿Qué otra situación importante genera comunicación, coordinación o escalamiento en este proceso?",
      "¿Cuál considera que hoy es la principal dificultad de comunicación o escalamiento de este proceso?"
    ]
  },

  "Ambiental": {
    titulo: "Ambiental",
    icono: "🌿",
    situaciones: [
      "Incidentes o emergencias ambientales",
      "Incumplimientos de requisitos legales ambientales",
      "Quejas ambientales de comunidades, clientes o autoridades",
      "Hallazgos o requerimientos de autoridades ambientales",
      "Tendencias desfavorables en la gestión de residuos, vertimientos o emisiones",
      "Riesgos ambientales que puedan afectar la continuidad de la operación o el cumplimiento legal"
    ],
    preguntasAbiertas: [
      "¿Qué otra situación importante genera comunicación, coordinación o escalamiento en este proceso?",
      "¿Cuál considera que hoy es la principal dificultad de comunicación o escalamiento de este proceso?"
    ]
  },

  "Excelencia Corporativa": {
    titulo: "Excelencia Corporativa",
    icono: "⭐",
    situaciones: [
      "Indicadores con desviaciones significativas frente a la meta",
      "Retrasos en proyectos transversales que afecten otros procesos",
      "Necesidades de articulación entre procesos",
      "Seguimiento a compromisos estratégicos con desviaciones relevantes",
      "Falta de apropiación o cumplimiento de metodologías, modelos o estándares corporativos",
      "Falta de evolución o madurez de los modelos de gestión implementados",
      "Identificación de oportunidades de mejora que requieran intervención transversal",
      "Necesidad de definir o redefinir lineamientos corporativos para asegurar criterios homogéneos entre procesos"
    ],
    preguntasAbiertas: [
      "¿Qué otra situación importante genera comunicación, coordinación o escalamiento en este proceso?",
      "¿Cuál considera que hoy es la principal dificultad de comunicación o escalamiento de este proceso?"
    ]
  },

  "Sistemas de Gestión": {
    titulo: "Sistemas de Gestión",
    icono: "⚙️",
    situaciones: [
      "Hallazgos críticos o mayores identificados durante auditorías internas o externas",
      "Acciones correctivas próximas a vencer o vencidas sin avance",
      "Tendencias desfavorables en indicadores del Sistema de Gestión",
      "Incremento significativo de desviaciones, no conformidades o reclamaciones",
      "Riesgos y oportunidades del sistema que requieran intervención de otros procesos o de la Dirección",
      "Riesgos estratégicos identificados que requieran gestión transversal o escalamiento",
      "Cambios en el Sistema Integrado de Gestión que impacten uno o varios procesos",
      "Incumplimientos del Sistema Integrado de Gestión que puedan afectar certificaciones o auditorías",
      "Desviaciones recurrentes que evidencien fallas sistémicas y requieran acciones transversales"
    ],
    preguntasAbiertas: [
      "¿Qué otra situación importante genera comunicación, coordinación o escalamiento en este proceso?",
      "¿Cuál considera que hoy es la principal dificultad de comunicación o escalamiento de este proceso?"
    ]
  },

  "Transversal": {
    titulo: "Situaciones Transversales",
    icono: "🔄",
    situaciones: [
      "Escalamiento de decisiones que superan el nivel de autonomía del proceso",
      "Definición o ajuste de prioridades entre actividades, proyectos o recursos",
      "Conflictos de prioridad entre dos o más procesos",
      "Solicitudes que requieren la participación coordinada de varios procesos",
      "Situaciones que requieren comunicación inmediata con la Vicepresidencia",
      "Desviaciones significativas entre el presupuesto aprobado y el gasto ejecutado que requieran revisión o toma de decisiones",
      "Necesidad de ajuste, redistribución u optimización de recursos (humanos, técnicos, financieros o de infraestructura)",
      "Desviaciones entre la capacidad disponible y la carga de trabajo que requieran redistribución, incremento o reducción de recursos, o ajuste de prioridades para optimizar la operación"
    ],
    preguntasAbiertas: [
      "¿Qué otra situación importante genera comunicación, coordinación o escalamiento en este proceso?",
      "¿Cuál considera que hoy es la principal dificultad de comunicación o escalamiento de este proceso?"
    ]
  }
};
