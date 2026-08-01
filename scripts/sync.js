// scripts/sync.js
// Lee tareas directamente desde Microsoft Planner vía Graph API
// y las guarda en tareas.json con el mismo esquema que usaba el Excel.
//
// Requiere permiso: Tasks.Read.All (Application) en la app de Azure.
// Plan: fiwX3ttKokqag77b7Mw2W2UAGDg7

const fetch = require('node-fetch');
const fs    = require('fs');

const TENANT        = process.env.AZURE_TENANT_ID;
const CLIENT_ID     = process.env.AZURE_CLIENT_ID;
const CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;

const PLAN_ID = 'fiwX3ttKokqag77b7Mw2W2UAGDg7';
const GRAPH   = 'https://graph.microsoft.com/v1.0';

if (!TENANT || !CLIENT_ID || !CLIENT_SECRET) {
  console.error('ERROR: Faltan secrets de Azure.');
  process.exit(1);
}

// ── Autenticación ─────────────────────────────────────────────────────────────
async function getToken() {
  const r = await fetch(
    `https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: CLIENT_ID, client_secret: CLIENT_SECRET,
        scope: 'https://graph.microsoft.com/.default',
        grant_type: 'client_credentials'
      })
    }
  );
  if (!r.ok) throw new Error(`Token failed: ${r.status} — ${await r.text()}`);
  return (await r.json()).access_token;
}

// ── Helpers de Graph ──────────────────────────────────────────────────────────
async function graphGet(token, url) {
  const r = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!r.ok) throw new Error(`Graph GET ${url} → ${r.status}: ${await r.text()}`);
  return r.json();
}

// Paginación automática: sigue @odata.nextLink hasta agotar resultados
async function graphGetAll(token, url) {
  const items = [];
  let next = url;
  while (next) {
    const json = await graphGet(token, next);
    items.push(...(json.value || []));
    next = json['@odata.nextLink'] || null;
  }
  return items;
}

// ── Conversión de fecha Excel serial → ISO (para compatibilidad con el JSON viejo) ──
function serialToISO(v) {
  if (!v) return '';
  if (typeof v === 'string') return v; // ya es string/ISO
  // número serial de Excel
  const d = new Date(Math.round((v - 25569) * 86400 * 1000));
  return d.toISOString().slice(0, 10);
}

// ── Conversión de prioridad numérica de Planner → texto ──────────────────────
function parsePrioridad(p) {
  // Planner: 0=Urgente, 1-3=Importante, 5-7=Media, 8-9=Baja (null/undefined=Sin prioridad)
  if (p === null || p === undefined) return 'Sin prioridad';
  if (p === 0)          return 'Urgente';
  if (p >= 1 && p <= 3) return 'Importante';
  if (p >= 5 && p <= 7) return 'Media';
  return 'Baja';
}

// ── Conversión de % completado → etiqueta de estado ──────────────────────────
function parseEstado(pct) {
  if (pct === 100) return 'Completed';
  if (pct > 0)     return 'In progress';
  return 'Not started';
}

// ── Días restantes hasta vencimiento ─────────────────────────────────────────
function diasRestantes(dueDateStr) {
  if (!dueDateStr) return '';
  const hoy  = new Date(); hoy.setHours(0, 0, 0, 0);
  const venc = new Date(dueDateStr); venc.setHours(0, 0, 0, 0);
  return Math.ceil((venc - hoy) / (1000 * 60 * 60 * 24));
}

// ── Semáforo ──────────────────────────────────────────────────────────────────
function semaforo(dias, pct) {
  if (pct === 100) return 'Completada';
  if (!dias && dias !== 0) return 'Sin fecha';
  if (dias < 0)  return 'Vencida';
  if (dias <= 3) return 'Próxima a vencer';
  return 'En tiempo';
}

// ── Main ──────────────────────────────────────────────────────────────────────
(async () => {
  try {
    console.log('[1/5] Obteniendo token de Microsoft...');
    const token = await getToken();
    console.log('      OK');

    // ── Buckets ──────────────────────────────────────────────────────────────
    console.log('[2/5] Leyendo buckets del plan...');
    const buckets = await graphGetAll(token, `${GRAPH}/planner/plans/${PLAN_ID}/buckets`);
    const bucketMap = {};
    buckets.forEach(b => { bucketMap[b.id] = b.name; });
    console.log(`      ${buckets.length} buckets: ${buckets.map(b => b.name).join(', ')}`);

    // ── Tareas ────────────────────────────────────────────────────────────────
    console.log('[3/5] Leyendo tareas del plan...');
    const tareas = await graphGetAll(token, `${GRAPH}/planner/plans/${PLAN_ID}/tasks`);
    console.log(`      ${tareas.length} tareas encontradas`);

    // ── Detalles (descripción) — en lotes de 10 para no saturar ──────────────
    console.log('[4/5] Obteniendo detalles de tareas...');
    const detailMap = {};
    const LOTE = 10;
    for (let i = 0; i < tareas.length; i += LOTE) {
      const lote = tareas.slice(i, i + LOTE);
      await Promise.all(lote.map(async t => {
        try {
          const d = await graphGet(token, `${GRAPH}/planner/tasks/${t.id}/details`);
          detailMap[t.id] = d.description || '';
        } catch { detailMap[t.id] = ''; }
      }));
      if (i + LOTE < tareas.length) await new Promise(r => setTimeout(r, 300));
    }
    console.log(`      Detalles obtenidos`);

    // ── Construir filas (una por asignado, igual que el Excel anterior) ───────
    console.log('[5/5] Construyendo tareas.json...');
    const filas = [];
    let idRegistro = 1;

    for (const t of tareas) {
      const asignados = Object.keys(t.assignments || {});
      const bucket    = bucketMap[t.bucketId] || '';
      const pct       = t.percentComplete || 0;
      const estado    = parseEstado(pct);
      const prioridad = parsePrioridad(t.priority);
      const fechaCreacion  = t.createdDateTime ? t.createdDateTime.slice(0, 10) : '';
      const fechaVence     = t.dueDateTime     ? t.dueDateTime.slice(0, 10)     : '';
      const fechaCompletada= t.completedDateTime ? t.completedDateTime.slice(0, 10) : '';
      const dias  = diasRestantes(fechaVence);
      const sem   = semaforo(diasRestantes(fechaVence) !== '' ? diasRestantes(fechaVence) : null, pct);
      const desc  = detailMap[t.id] || '';

      if (asignados.length === 0) {
        // tarea sin asignar: una fila con asignado vacío
        filas.push({
          ID_Registro:      idRegistro++,
          ID_Tarea:         t.id,
          Titulo_Tarea:     t.title || '',
          Descripcion:      desc,
          Asignado_Nombre:  '',
          Asignado_email:   '',
          Fecha_Creacion:   fechaCreacion,
          Fecha_vencimiento: fechaVence,
          Prioridad:        prioridad,
          Bucket:           bucket,
          Estado:           estado,
          Completada:       fechaCompletada,
          'Dias para vencer': typeof dias === 'number' ? dias : '',
          Semaforo:         sem,
          Canal_teams:      ''
        });
      } else {
        // una fila por asignado (mismo esquema que el Excel)
        for (const userId of asignados) {
          const assignment = t.assignments[userId];
          filas.push({
            ID_Registro:      idRegistro++,
            ID_Tarea:         t.id,
            Titulo_Tarea:     t.title || '',
            Descripcion:      desc,
            Asignado_Nombre:  userId,   // se resuelve a nombre en el paso de usuarios
            Asignado_email:   '',
            Fecha_Creacion:   fechaCreacion,
            Fecha_vencimiento: fechaVence,
            Prioridad:        prioridad,
            Bucket:           bucket,
            Estado:           estado,
            Completada:       fechaCompletada,
            'Dias para vencer': typeof dias === 'number' ? dias : '',
            Semaforo:         sem,
            Canal_teams:      ''
          });
        }
      }
    }

    // ── Resolver nombres de usuario ───────────────────────────────────────────
    // Recolectar IDs únicos
    const userIds = [...new Set(filas.map(f => f.Asignado_Nombre).filter(Boolean))];
    const userMap = {};
    console.log(`      Resolviendo ${userIds.length} usuarios...`);
    await Promise.all(userIds.map(async uid => {
      try {
        const u = await graphGet(token, `${GRAPH}/users/${uid}`);
        userMap[uid] = { nombre: u.displayName || uid, email: u.mail || u.userPrincipalName || '' };
      } catch {
        userMap[uid] = { nombre: uid, email: '' };
      }
    }));

    // Sustituir IDs por nombres
    filas.forEach(f => {
      if (f.Asignado_Nombre && userMap[f.Asignado_Nombre]) {
        f.Asignado_email  = userMap[f.Asignado_Nombre].email;
        f.Asignado_Nombre = userMap[f.Asignado_Nombre].nombre;
      }
    });

    // ── Escribir JSON ─────────────────────────────────────────────────────────
    const out = {
      actualizado: new Date().toISOString(),
      total:       filas.length,
      datos:       filas
    };
    fs.writeFileSync('tareas.json', JSON.stringify(out, null, 2));
    console.log(`      ${filas.length} filas escritas (${tareas.length} tareas, ${userIds.length} usuarios)`);
    console.log('LISTO.');

  } catch (e) {
    console.error('FALLÓ:', e.message);
    process.exit(1);
  }
})();
