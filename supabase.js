// Módulo de persistencia a Supabase vía Management API (SQL).
// Escribe el desglose completo de cada validación en public.matriz_riesgo.
const https = require('https');

const TOKEN = process.env.SUPABASE_PAT;
const PROJECT = process.env.SUPABASE_PROJECT_REF || 'lgycqjytcwqjmlwystse';

function runSQL(query) {
  return new Promise((resolve, reject) => {
    if (!TOKEN) return reject(new Error('SUPABASE_PAT no configurado'));
    const payload = JSON.stringify({ query });
    const req = https.request({
      host: 'api.supabase.com',
      path: `/v1/projects/${PROJECT}/database/query`,
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${TOKEN}`,
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(payload)
      }
    }, res => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        if (res.statusCode >= 200 && res.statusCode < 300) resolve(data);
        else reject(new Error(`HTTP ${res.statusCode}: ${data.substring(0, 300)}`));
      });
    });
    req.on('error', reject);
    req.write(payload);
    req.end();
  });
}

function v(x) {
  if (x === null || x === undefined || x === '') return 'null';
  if (typeof x === 'number') return String(x);
  if (typeof x === 'boolean') return x ? 'true' : 'false';
  return `'${String(x).replace(/'/g, "''")}'`;
}
function jsonb(x) {
  if (x === null || x === undefined) return 'null';
  return `'${JSON.stringify(x).replace(/'/g, "''")}'::jsonb`;
}

/**
 * Upsert de un registro con su desglose de validación.
 * @param {object} reg - registro (id, riesgo, causa, control, calif_*, completado)
 * @param {object} a   - objeto analysis de la IA (opcional; si viene, guarda el desglose)
 */
async function upsertRegistro(reg, a) {
  const total = (Number(reg.calif_riesgo) || 0) + (Number(reg.calif_causa) || 0) + (Number(reg.calif_control) || 0);
  const coh = Math.round(total / 9 * 100);

  const ra = a?.riesgo_assessment || {};
  const ca = a?.causa_assessment || {};
  const co = a?.control_assessment || {};
  const comp = co.componentes || {};
  const ae = co.alertamiento_estructura || {};
  const rs = a?.redaccion_sugerida || {};
  const conteoSe = parseInt(String(co.conteo_se).match(/\d+/)?.[0] ?? '', 10);

  // Campos de análisis solo si viene 'a'
  const analisisCols = a ? `,
    coherencia_general = ${v(a.overall_coherence)},
    riesgo_score = ${v(Number(ra.score) || null)}, riesgo_findings = ${v(ra.findings)},
    causa_score = ${v(Number(ca.score) || null)}, causa_nota = ${v(ca.nota_aplicada)},
    causa_sujeto = ${v(ca.sujeto)}, causa_accion = ${v(ca.accion)}, causa_objeto = ${v(ca.objeto)},
    causa_delito_fuente = ${v(ca.delito_fuente)}, causa_tipologia = ${v(ca.tipologia)},
    causa_factores_riesgo = ${v(ca.factores_riesgo)}, causa_findings = ${v(ca.findings)},
    control_score = ${v(Number(co.score) || null)}, control_nota = ${v(co.nota_aplicada)},
    control_alertamiento = ${v(comp.alertamiento === true)}, control_analisis = ${v(comp.analisis === true)},
    control_reporte = ${v(comp.reporte === true)}, control_decision = ${v(comp.decision === true)},
    control_conteo_se = ${v(isNaN(conteoSe) ? null : conteoSe)},
    control_coherencia_causa = ${v(co.coherencia_con_causa)},
    control_error_categorico = ${v(co.error_categorico === true)}, control_findings = ${v(co.findings)},
    alert_nucleo_verbal = ${v(ae.nucleo_verbal)}, alert_complemento = ${v(ae.complemento)},
    alert_senal_alerta = ${v(ae.senal_alerta)},
    sugerido_riesgo = ${v(rs.riesgo)}, sugerido_causa = ${v(rs.causa)}, sugerido_control = ${v(rs.control)},
    fortalezas = ${jsonb(a.strengths)}, debilidades = ${jsonb(a.weaknesses)},
    recomendaciones = ${jsonb(a.recommendations)}, analysis_raw = ${jsonb(a)},
    validado_at = now()` : '';

  const sql = `insert into public.matriz_riesgo
    (id, riesgo, causa, control, calif_riesgo, calif_causa, calif_control, coherencia_pct, estado, updated_at)
    values (${v(reg.id)}, ${v(reg.riesgo)}, ${v(reg.causa)}, ${v(reg.control)},
      ${v(reg.calif_riesgo)}, ${v(reg.calif_causa)}, ${v(reg.calif_control)},
      ${v(coh)}, ${v(reg.completado ? 'Revisado' : 'Pendiente')}, now())
    on conflict (id) do update set
      riesgo = excluded.riesgo, causa = excluded.causa, control = excluded.control,
      calif_riesgo = excluded.calif_riesgo, calif_causa = excluded.calif_causa, calif_control = excluded.calif_control,
      coherencia_pct = excluded.coherencia_pct, estado = excluded.estado, updated_at = now()${analisisCols};`;

  return runSQL(sql);
}

const enabled = !!TOKEN;

module.exports = { upsertRegistro, runSQL, enabled };
