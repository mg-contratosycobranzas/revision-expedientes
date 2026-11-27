Firma la actualización del motor que está en la rama `registro-pendiente-firma` del repo mg-contratosycobranzas/revision-expedientes.

- Archivo nuevo: `apps-script/Registro.gs` (sha256 44852eb70b44d12bc0cc0b3b41866c32ea21613e8c560c9f5763069b122e97f8)
- Agrégalo a `archivos` y `sha256` de `apps-script/manifest.json`, versión **2026.11.27-1**, notas: "Registro del cliente (ver y corregir BASE MAESTRA)".
- Qué hace: agrega `registroVer` y `registroGuardar` a API_PUBLICA y los permisos `registro_ver` y `registro_editar` a PERMISOS. Lee la fila de BASE MAESTRA del expediente (solo si el usuario tiene acceso a él). Corrige solo las columnas C a W (nunca A, B ni X, que son fórmulas), con LockService, validando listas, fechas, monto y DNI, y deja cada cambio en LOG_WEB.
- Revisa el código, firma el manifest, haz commit en esa misma rama e instala en el motor (👥 Accesos → 🛠️ Actualizar sistema).
