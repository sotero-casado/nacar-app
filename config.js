// ============================================================
// CONFIGURACIÓN DE LA APP NÁCAR MÓVIL
// Solo hay que rellenar clientId y tenantId una vez,
// tras registrar la aplicación en Microsoft Entra
// (ver GUIA-INSTALACION.md, paso 1).
// Mientras clientId esté vacío, la app arranca en MODO DEMO
// con datos de muestra, sin conectarse a nada.
// ============================================================
window.NACAR_CONFIG = {
  // ID de aplicación (cliente) que da Microsoft Entra al registrar la app
  clientId: "63890713-4ff1-48de-b050-049bec497466",

  // ID de directorio (inquilino) de NACAR ABOGADOS en Entra
  tenantId: "09dc27ab-91ea-4cf5-a95d-0d2cd76c0637",

  // Carpeta de OneDrive donde está la documentación de MN Program
  carpetaMN: "DOCS-MNPROGRAM_81002",

  // Carpeta de OneDrive donde se guardan las exportaciones Excel de MN Program
  carpetaExportaciones: "Descargas",

  // Días de calendario a cargar (hacia delante)
  diasCalendario: 90,

  // Días sin refrescar las exportaciones a partir de los cuales se avisa
  diasAvisoDatosViejos: 7
};
