# Verificación de expedientes en ControlDoc

Revisa cada radicado en ControlDoc y compara el número de expediente, la serie y la subserie
con lo que debería tener. El expediente correcto se obtiene así:

`DocumentosExpediente.xlsx` (radicado) → `Segui_Perso.xlsx` (caso al que pertenece y sus demás
radicados, cédula y número de proceso) → `BASE DILIGENCIAS 2026.xlsx`, pestañas EXPEDIENTES 25 y 26
(No. EXPEDIENTE y SERIE).
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
   python verificar_expedientes.py --base "BASE DILIGENCIAS 2026.xlsx" --seguimiento Segui_Perso.xlsx --radicados DocumentosExpediente.xlsx --limite 3
   ```
3. Se abre Chrome: inicie sesión en ControlDoc y presione ENTER en la terminal.
4. Al final queda un archivo `resultado_AAAAMMDD_HHMM.xlsx` con la columna **ACCION** para cada radicado:
   - `OK`
   - `ACTUALIZAR TIPOLOGIA (expediente vacío y serie no corresponde)`
   - `ASOCIAR EXPEDIENTE (serie correcta, expediente vacío)`
   - `REVISAR: ...` / `NO ESTA EN LA BASE - revisar manualmente`

Con `--pausar`, en cada radicado que requiere actualizar tipología el script deja abierto
Documentos > Actualizar Tipología y espera a que usted haga el cambio y presione ENTER.

Para ver solo el cruce de los Excel, sin abrir ControlDoc, agregue `--solo-cruce`.

## Cómo busca el expediente
1. Busca el radicado en todas las pestañas de Segui_Perso y toma la(s) fila(s) del caso.
2. En EXPEDIENTES 25/26 busca, en este orden: el propio radicado, los radicados citados en el asunto,
   los demás radicados de la fila de Segui_Perso, los números de proceso (23 dígitos) y la cédula.
3. La columna VINCULO CON BASE del resultado dice cuál de esos datos coincidió.
Use siempre una copia reciente de la base: los casos que aún no tienen expediente creado salen como
`SIN EXPEDIENTE EN LA BASE`.

Las equivalencias de serie (p. ej. `PROCESOS (INTERVENCION)` en la base = `PROCESOS DE INTERVENCION` en
ControlDoc) están en `EQUIVALENCIAS_SERIE` dentro del script, y los selectores de la página en `SEL`.
