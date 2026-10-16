# Verificación de expedientes en ControlDoc

Revisa cada radicado en ControlDoc y lo compara con la base `BASE DILIGENCIAS 2026.xlsx`
(pestañas EXPEDIENTES 25 y 26): número de expediente, serie y subserie.
**Fase 1: solo lectura**, no cambia nada en ControlDoc.

## Instalación (una sola vez, en Windows)
1. Instalar Python 3 desde https://www.python.org (marcar "Add python to PATH").
2. En esta carpeta, abrir una terminal y ejecutar:
   ```
   pip install -r requirements.txt
   ```
   Usa el Google Chrome que ya tiene instalado.

## Uso
1. Descargar la base desde SharePoint (Archivo > Descargar copia) y el Excel exportado de ControlDoc
   con los radicados a revisar (o un `.txt` con un radicado por línea). Ponerlos en esta carpeta.
2. Prueba con 3 radicados:
   ```
   python verificar_expedientes.py --base "BASE DILIGENCIAS 2026.xlsx" --radicados DocumentosExpediente.xlsx --limite 3
   ```
3. Se abre Chrome: inicie sesión en ControlDoc y presione ENTER en la terminal.
4. Al final queda un archivo `resultado_AAAAMMDD_HHMM.xlsx` con la columna **ACCION** para cada radicado:
   - `OK`
   - `ACTUALIZAR TIPOLOGIA (expediente vacío y serie no corresponde)`
   - `ASOCIAR EXPEDIENTE (serie correcta, expediente vacío)`
   - `REVISAR: ...` / `NO ESTA EN LA BASE - revisar manualmente`

Con `--pausar`, en cada radicado que requiere actualizar tipología el script deja abierto
Documentos > Actualizar Tipología y espera a que usted haga el cambio y presione ENTER.

## Cómo busca en la base
Primero busca el radicado tal cual en la columna RADICADO. Si no está (por ejemplo un E- que responde
a un R-), busca los radicados citados en el ASUNTO ("RESPUESTA DEL RADICADO CONTROLDOC NO. R-...").

Las equivalencias de serie (p. ej. `PROCESOS (INTERVENCION)` en la base = `PROCESOS DE INTERVENCION` en
ControlDoc) están en `EQUIVALENCIAS_SERIE` dentro del script, y los selectores de la página en `SEL`.
