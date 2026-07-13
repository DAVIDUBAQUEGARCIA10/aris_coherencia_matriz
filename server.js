const express = require('express');
const axios = require('axios');
require('dotenv').config();
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const supabase = require('./supabase');

const DATA_PATH = path.join(__dirname, 'data.json');

// Persiste un registro (y su desglose de análisis) en Supabase, sin bloquear la respuesta.
function syncSupabase(reg, analysis) {
    if (!supabase.enabled) return;
    supabase.upsertRegistro(reg, analysis)
        .then(() => console.log(`☁️  Supabase: registro ${reg.id} sincronizado`))
        .catch(e => console.warn(`⚠️  Supabase (registro ${reg.id}):`, e.message));
}

/**
 * Salvaguarda conservadora: elimina el pronombre relativo "que" más común
 * en la redacción de causas (Anexo 1 V2 Jun 2026), sin alterar locuciones
 * válidas como "para que" o "de modo que".
 */
function limpiarQue(texto) {
    if (!texto) return texto;
    let t = texto;
    // "Clientes que realizan" -> "Clientes realizan"  (sujeto + que + verbo)
    t = t.replace(/\b([A-Za-zÁÉÍÓÚáéíóúÑñ]+(?:es|os|as|s)?)\s+que\s+(realiza|realizan|efectúa|efectúan|ejecuta|ejecutan|ingresa|ingresan|transfiere|transfieren|deposita|depositan|recibe|reciben|paga|pagan|adquiere|adquieren|canaliza|canalizan|dispersa|dispersan|ordena|ordenan|utiliza|utilizan)\b/gi,
        '$1 $2');
    // "montos que superan" -> "montos superiores a" (caso frecuente)
    t = t.replace(/\bque\s+supera(n)?\b/gi, 'superiores a');
    t = t.replace(/\bque\s+excede(n)?\b/gi, 'superiores a');
    // "sustantivo que se realizan/hacen" -> "sustantivo realizados/hechos" (que + se + verbo -> participio o quita "que se")
    t = t.replace(/\bque\s+se\s+(realiza|realizan)\b/gi, 'realizadas');
    t = t.replace(/\bque\s+se\s+(efectúa|efectúan)\b/gi, 'efectuadas');
    t = t.replace(/\bque\s+se\s+(ejecuta|ejecutan)\b/gi, 'ejecutadas');
    // "que indican/reflejan/evidencian/muestran X" -> "indicando/reflejando..." (verbo declarativo -> gerundio)
    // "movimientos que simulan" -> "movimientos simulando" (sustantivo + que + verbo conjugado -> gerundio)
    const gerundios = {
        'simulan': 'simulando', 'simula': 'simulando',
        'ocultan': 'ocultando', 'oculta': 'ocultando',
        'dificultan': 'dificultando', 'dificulta': 'dificultando',
        'permiten': 'permitiendo', 'permite': 'permitiendo',
        'facilitan': 'facilitando', 'facilita': 'facilitando',
        'evitan': 'evitando', 'evita': 'evitando',
        'integran': 'integrando', 'integra': 'integrando',
        'encubren': 'encubriendo', 'encubre': 'encubriendo',
        'disimulan': 'disimulando', 'disimula': 'disimulando',
        'buscan': 'buscando', 'busca': 'buscando',
        'indican': 'indicando', 'indica': 'indicando',
        'reflejan': 'reflejando', 'refleja': 'reflejando',
        'evidencian': 'evidenciando', 'evidencia': 'evidenciando',
        'muestran': 'mostrando', 'muestra': 'mostrando',
        'sugieren': 'sugiriendo', 'sugiere': 'sugiriendo',
        'representan': 'representando', 'representa': 'representando',
        'generan': 'generando', 'genera': 'generando',
        'involucran': 'involucrando', 'involucra': 'involucrando',
        'presentan': 'presentando', 'presenta': 'presentando'
    };
    t = t.replace(/\bque\s+([a-záéíóúñ]+)\b/gi, (m, verbo) => {
        const g = gerundios[verbo.toLowerCase()];
        return g ? g : m;
    });
    // Limpiar dobles espacios resultantes
    t = t.replace(/\s{2,}/g, ' ').trim();
    return t;
}

/**
 * Detecta si un texto AÚN contiene "que" como pronombre relativo problemático.
 * Excluye locuciones válidas: "para que", "con el fin de que", "de modo que", etc.
 */
function quedanQueRelativos(texto) {
    if (!texto) return false;
    // Quitar locuciones donde "que" es válido
    let t = texto
        .replace(/\bpara\s+que\b/gi, '')
        .replace(/\bcon\s+el\s+fin\s+de\s+que\b/gi, '')
        .replace(/\bde\s+(?:modo|manera|forma)\s+que\b/gi, '')
        .replace(/\ba\s+fin\s+de\s+que\b/gi, '')
        .replace(/\btales?\s+que\b/gi, '')
        .replace(/\bpuesto\s+que\b/gi, '')
        .replace(/\bdado\s+que\b/gi, '')
        .replace(/\bya\s+que\b/gi, '')
        .replace(/\bde\s+tal\s+(?:modo|forma|manera)\s+que\b/gi, '');
    return /\bque\b/i.test(t);
}

/**
 * Corrige combinaciones verbo-objeto incoherentes que puede dejar la reescritura.
 * P.ej. "realizan pólizas" -> "suscriben pólizas". Solo casos claros y seguros.
 */
function corregirVerboObjeto(texto) {
    if (!texto) return texto;
    let t = texto;
    // realizar/hacer + póliza(s) -> suscribir + póliza(s)
    t = t.replace(/\b(realiza|realizan|hace|hacen|efectúa|efectúan|ejecuta|ejecutan)\s+((?:una?\s+|las?\s+|los?\s+|múltiples\s+|varias?\s+)?p[óo]liza)/gi,
        (m, verbo, obj) => (/n$/i.test(verbo) ? 'suscriben ' : 'suscribe ') + obj);
    // realizar/hacer + título(s) de capitalización -> suscribir
    t = t.replace(/\b(realiza|realizan|hace|hacen)\s+((?:una?\s+|el\s+|los?\s+|múltiples\s+)?t[íi]tulos?\s+de\s+capitalizaci[óo]n)/gi,
        (m, verbo, obj) => (/n$/i.test(verbo) ? 'suscriben ' : 'suscribe ') + obj);
    // realizar/hacer + seguro(s) -> adquirir
    t = t.replace(/\b(realiza|realizan|hace|hacen)\s+((?:una?\s+|el\s+|los?\s+)?seguros?)/gi,
        (m, verbo, obj) => (/n$/i.test(verbo) ? 'adquieren ' : 'adquiere ') + obj);
    return t.replace(/\s{2,}/g, ' ').trim();
}

/**
 * Elimina redundancias léxicas comunes en la redacción del control
 * (p.ej. "Alertamiento de alertas") y expande abreviaturas coloquiales.
 */
function limpiarRedundancias(texto) {
    if (!texto) return texto;
    let t = texto;

    // ── Normalización de FLUIDEZ del inicio del alertamiento ──
    // Casos rotos como "Alertamiento se generan automáticas ...", "Alertamiento genera alertas
    // automáticas ...", "Alertamiento identifica automáticas ..." -> "Alertamiento generado automáticamente sobre ..."
    // 1) Normalizar el inicio roto: "Alertamiento [se] genera(n)/genera/identifica/... [alertas] automátic(a/as/o/os/amente) [parametrizadas]"
    //    Se captura el CONECTOR o el conteni­do que sigue para no partir palabras.
    t = t.replace(
        /^\s*Alertamiento\s+(?:se\s+)?(?:generan?|genera|generado|identifica|verifica|detecta|monitorea|revisa|activan?|emiten?)\s+(?:las\s+|los\s+)?(?:alert(?:a|as)\s+)?autom[áa]tic(?:[oa]s?|amente)\b(?:\s+parametrizad[oa]s?)?/i,
        'Alertamiento generado automáticamente'
    );
    // 2) Asegurar un conector válido tras "automáticamente" (sobre/en/de/cuando/por/para/respecto).
    //    Si la palabra siguiente NO es un conector, insertar "sobre".
    t = t.replace(/^(Alertamiento generado automáticamente)\s+(?!(?:sobre|en|de|del|por|para|cuando|respecto|acerca|los|las|el|la)\b)/i,
        '$1 sobre ');
    // 3) Verbo suelto restante tras "Alertamiento" (sin "automático"): "Alertamiento identifica X" -> "Alertamiento sobre X"
    t = t.replace(/^\s*Alertamiento\s+(?:se\s+)?(?:identifica|verifica|genera|detecta|monitorea|revisa)\s+(?!autom|alert)/i, 'Alertamiento sobre ');
    // 4) Limpieza de residuos que puedan haber quedado de reemplazos previos
    t = t.replace(/automáticamente\s+sobre\s+mente\w*/gi, 'automáticamente');
    t = t.replace(/\bsobre\s+mente\w*/gi, 'sobre');
    t = t.replace(/\bmente\w*\s+(en|sobre|de|del|cuando)\b/gi, '$1');
    t = t.replace(/\bsobre\s+sobre\b/gi, 'sobre');

    // ── Redundancia "Alertamiento de alertas [adjetivos] <conector> ..." ──
    // Toma todo desde "Alertamiento de [art] alert-X" + cualquier adjetivo hasta el
    // conector que introduce el contenido real, y colapsa a "Alertamiento <conector> ...".
    // Adjetivos permitidos entre "alertas" y el conector (letras/espacios, máx pocas palabras).
    t = t.replace(
        /\bAlertamiento\s+de\s+(?:la|las|el|los)\s+alert(?:a|as|amiento|amientos)((?:\s+[a-záéíóúñ]+){0,3}?)\s+(cuando|al|si|de|sobre|por|en|para|mediante)\b/gi,
        (m, medio, con) => {
            const c = con.toLowerCase();
            // "generadas/emitidas cuando" -> "Alertamiento generado cuando"
            if (/\b(generad|emitid|detectad)/i.test(medio) && (c === 'cuando' || c === 'al' || c === 'si')) {
                return `Alertamiento generado ${c}`;
            }
            return c === 'de' ? 'Alertamiento de' : `Alertamiento ${c}`;
        }
    );
    // Sin artículo: "Alertamiento de alertas <conector>"
    t = t.replace(
        /\bAlertamiento\s+de\s+alert(?:a|as|amiento|amientos)((?:\s+[a-záéíóúñ]+){0,3}?)\s+(cuando|al|si|de|sobre|por|en|para|mediante)\b/gi,
        (m, medio, con) => {
            const c = con.toLowerCase();
            if (/\b(generad|emitid|detectad)/i.test(medio) && (c === 'cuando' || c === 'al' || c === 'si')) {
                return `Alertamiento generado ${c}`;
            }
            return c === 'de' ? 'Alertamiento de' : `Alertamiento ${c}`;
        }
    );
    // "Alertamiento de verificar/verificación/generar de alertas ..." -> "Alertamiento de "
    // (quita el verbo/infinitivo de monitoreo redundante justo tras "Alertamiento de")
    t = t.replace(/\bAlertamiento\s+de\s+(?:verificar|verificaci[óo]n\s+de|generar|generaci[óo]n\s+de|detectar|detecci[óo]n\s+de|identificar|monitorear|revisar)\s+(?=alert|las\s+alert)/gi, 'Alertamiento de ');
    // "Alertamiento de verificar <sustantivo>" (infinitivo suelto) -> "Alertamiento de <sustantivo>"
    t = t.replace(/\bAlertamiento\s+de\s+(?:verificar|generar|detectar|identificar|monitorear|revisar)\s+/gi, 'Alertamiento de ');
    // "Alertamiento de [las] alertas <participio> o <resto>" -> "Alertamiento de <resto>"
    // (ej: "de alertas detectadas o anomalías detectadas" -> "de anomalías detectadas")
    t = t.replace(/\bAlertamiento\s+de\s+(?:la|las|el|los)?\s*alert(?:a|as)\s+(?:detectad|generad|identificad|emitid|activad|parametrizad)\w*\s+(?:o|y)\s+/gi, 'Alertamiento de ');
    // "Alertamiento de [las] alerta(s) <participio>" -> "Alertamiento de <participio-cosa>" (deja el hecho)
    t = t.replace(/\bAlertamiento\s+de\s+(?:la|las|el|los)?\s*alert(?:a|as)\s+(detectad|generad|identificad|emitid)\w*\s+(?=(?:de|en|sobre|por|cuando|del)\b)/gi, 'Alertamiento de ');
    // "Alertamiento de [las] alerta(s)" simple sin más -> "Alertamiento de "
    t = t.replace(/\bAlertamiento\s+de\s+(?:la|las|el|los)?\s*alert(?:a|as|amiento|amientos)\b\s*/gi, 'Alertamiento de ');
    // "Alertamiento de verificación/generación/detección DE ALERTAS ..." -> quitar "de [las] alertas"
    t = t.replace(/\b(Alertamiento\s+de\s+[a-záéíóúñ]+(?:ci[óo]n|ado|miento|eo))\s+de\s+(?:la|las|los|el)?\s*alert(?:a|as|amiento|amientos)\b/gi, '$1');
    // "Alertamiento ... de alertas cuando/de/sobre X" en medio de la oración principal
    t = t.replace(/\bde\s+(?:la|las|los|el)?\s*alert(?:a|as|amiento|amientos)\s+(cuando|de|sobre|por|en|para|generad|emitid)/gi, '$1');
    // Residuo huérfano: "Alertamiento de detectadas/generadas o X" -> "Alertamiento de X"
    t = t.replace(/\bAlertamiento\s+de\s+(?:detectad|generad|identificad|emitid|activad)\w*\s+(?:o|y)\s+/gi, 'Alertamiento de ');
    // Residuo: "Alertamiento de detectadas en/de/sobre X" -> "Alertamiento de X"
    t = t.replace(/\bAlertamiento\s+de\s+(?:detectad|generad|identificad|emitid|activad)\w*\s+(?:en|de|sobre|por)\s+/gi, 'Alertamiento de ');
    // Redundancias análisis/reporte
    t = t.replace(/\banálisis\s+del?\s+análisis\b/gi, 'análisis');
    t = t.replace(/\breporte\s+del?\s+reporte\b/gi, 'reporte');
    // Expandir abreviaturas coloquiales frecuentes (bordes de palabra)
    const abrevs = [
        [/\btx\b/gi, 'transacción'], [/\bops?\b/gi, 'operación'],
        [/\bdocs?\b/gi, 'documento'], [/\binfo\b/gi, 'información'],
        [/\baprox\.?\b/gi, 'aproximadamente'], [/\bvlr\b/gi, 'valor'],
        [/\bc\/u\b/gi, 'cada uno'], [/\bxq\b/gi, 'porque'], [/\bpq\b/gi, 'porque']
    ];
    abrevs.forEach(([re, rep]) => { t = t.replace(re, rep); });
    // Limpiar dobles espacios
    t = t.replace(/\s{2,}/g, ' ').trim();
    return t;
}

/**
 * Expande las siglas oficiales a su forma completa (solo para redacción sugerida).
 * Maneja variantes con la sigla entre paréntesis para no duplicar.
 */
function expandirSiglas(texto) {
    if (!texto) return texto;
    let t = texto;
    const siglas = [
        [/\bROS\b/g, 'reporte de operación sospechosa'],
        [/\bUIAF\b/g, 'Unidad de Información y Análisis Financiero'],
        [/\bPEP\b/g, 'persona expuesta políticamente'],
        [/\bOFAC\b/g, 'Oficina de Control de Activos Extranjeros del Departamento del Tesoro de los Estados Unidos'],
        [/\bDIAN\b/g, 'Dirección de Impuestos y Aduanas Nacionales'],
        [/\bSFC\b/g, 'Superintendencia Financiera de Colombia'],
        [/\bGAFI\b/g, 'Grupo de Acción Financiera Internacional'],
        [/\bLA\/FT\/FPADM\b/g, 'lavado de activos, financiación del terrorismo y financiación de la proliferación de armas de destrucción masiva'],
        [/\bLAFT\b/g, 'lavado de activos y financiación del terrorismo']
    ];
    // Primero, quitar sigla entre paréntesis tras el término completo: "... financiero (UIAF)" -> "... financiero"
    t = t.replace(/\s*\((?:ROS|UIAF|PEP|OFAC|DIAN|SFC|GAFI|LAFT|LA\/FT\/FPADM)\)/g, '');
    // ONU: solo cuando va sola, no dentro de "ONU s" ni nombres propios ya expandidos
    t = t.replace(/\bONU\b(?!\s+de|\w)/g, 'Organización de las Naciones Unidas');
    siglas.forEach(([re, rep]) => { t = t.replace(re, rep); });
    t = t.replace(/\s{2,}/g, ' ').replace(/\s+([.,;])/g, '$1').trim();
    return t;
}

/**
 * Corrige el uso incorrecto de "lavado de activos DE <delito>" y prefiere la
 * expresión genérica "delitos fuente de lavado de activos" sobre "corrupción",
 * salvo que el texto ORIGINAL ya mencionara ese delito específico.
 * @param {string} textoOriginal - texto fuente para no borrar delitos que sí estaban.
 */
function corregirDelitos(texto, textoOriginal) {
    if (!texto) return texto;
    let t = texto;
    const delitos = 'corrupci[óo]n|narcotr[áa]fico|miner[íi]a\\s+ilegal|extorsi[óo]n|contrabando|cohecho|soborno|trata\\s+de\\s+personas|tr[áa]fico\\s+de\\s+\\w+';
    // "lavado de activos de <delito>" -> "lavado de activos y <delito>"
    t = t.replace(new RegExp(`(lavado\\s+de\\s+activos)\\s+de\\s+(${delitos})`, 'gi'), '$1 y $2');

    // Preferir "delitos fuente de lavado de activos" sobre "corrupción",
    // pero SOLO si el original NO mencionaba corrupción (para no borrar dato real).
    const origMencionaCorrupcion = textoOriginal && /corrupci[óo]n|cohecho|soborno|dcap|contratos?\s+p[úu]blic/i.test(textoOriginal);
    if (!origMencionaCorrupcion) {
        // "provenientes de corrupción" / "origen en corrupción" -> "de delitos fuente de lavado de activos"
        t = t.replace(/\b(provenientes?|proveniente|producto|origen|derivad[oa]s?)\s+de\s+(?:la\s+)?corrupci[óo]n\b/gi,
            '$1 de delitos fuente de lavado de activos');
        t = t.replace(/\bde\s+(?:la\s+)?corrupci[óo]n\s+(administrativa|p[úu]blica)\b/gi, 'de delitos fuente de lavado de activos');
        // "recursos de corrupción" / "capitales de corrupción"
        t = t.replace(/\b(recursos?|capitales?|activos?|fondos?|dineros?)\s+(?:de|por)\s+(?:la\s+)?corrupci[óo]n\b/gi,
            '$1 provenientes de delitos fuente de lavado de activos');
        // "lavado de activos y corrupción" (si corrupción no estaba) -> dejar solo genérico
        t = t.replace(/\blavado\s+de\s+activos\s+y\s+corrupci[óo]n\b/gi, 'delitos fuente de lavado de activos');
    }
    return t.replace(/\s{2,}/g, ' ').trim();
}

/**
 * Corrige el término de "sistema" según la naturaleza de la entidad.
 * Capitalizadora NO es aseguradora: se le quita cualquier "asegurador/financiero".
 */
function corregirSistema(texto, esCapitalizadora) {
    if (!texto || !esCapitalizadora) return texto;
    let t = texto;
    // Variantes genéricas o incorrectas -> término correcto de capitalización
    t = t.replace(/sistema\s+financiero\s+(?:o|y|u)\s+asegurador/gi, 'sistema de capitalización');
    t = t.replace(/sistema\s+asegurador\s+(?:o|y|u)\s+financiero/gi, 'sistema de capitalización');
    t = t.replace(/sistema\s+asegurador/gi, 'sistema de capitalización');
    t = t.replace(/sistema\s+financiero/gi, 'sistema de capitalización');
    t = t.replace(/sector\s+asegurador/gi, 'sector de capitalización');
    t = t.replace(/uso\s+del\s+seguro\b/gi, 'uso del producto de capitalización');
    return t;
}

// Fórmula genérica de señal de alerta (evita inventar % o umbrales numéricos).
const SENAL_GENERICA = 'que presentan cumplimiento en las señales de alerta parametrizadas con las lógicas y umbrales definidos en su parametrización';

/**
 * Inserta una señal de alerta GENÉRICA en el alertamiento cuando este carece de ella.
 * NO usa porcentajes ni umbrales inventados: remite a la parametrización definida.
 * Trabaja solo sobre la primera oración (el alertamiento), sin tocar análisis/reporte/decisión.
 */
function completarSenalAlerta(controlSugerido) {
    if (!controlSugerido) return controlSugerido;
    // Separar la primera oración (alertamiento) del resto (Mediante/Como/Para)
    const m = controlSugerido.match(/^([\s\S]*?)(\.\s+(?:Mediante|Como|Para)[\s\S]*)$/i);
    let alerta = m ? m[1].trim() : controlSugerido.replace(/\.\s*$/, '').trim();
    const resto = m ? m[2] : '.';

    // Quitar punto final del alertamiento si lo tiene
    alerta = alerta.replace(/\.\s*$/, '').trim();
    // Insertar la señal genérica. Si termina con complemento, se enlaza de forma natural.
    // Evitar duplicar un "que" previo.
    const conector = /,\s*$/.test(alerta) ? ' ' : ', ';
    alerta = alerta.replace(/[,;\s]+$/, '');
    const nuevaAlerta = `${alerta}${conector}${SENAL_GENERICA}`;
    return `${nuevaAlerta}${resto.startsWith('.') ? resto : '. ' + resto}`;
}

/**
 * Reescribe el control sugerido REINCORPORANDO el contexto/detalles del original
 * que se hubieran perdido, manteniendo la estructura (Alertamiento + Mediante/se + Como/se + Para/se).
 */
async function recuperarContexto(controlSugerido, controlOriginal) {
    const apiKey = process.env['API-KEY'];
    if (!apiKey) return controlSugerido;
    const model = process.env.OPENAI_MODEL || 'gpt-4o-mini';
    const prompt = `El CONTROL SUGERIDO perdió detalles importantes que SÍ estaban en el CONTROL ORIGINAL (por ejemplo: listas de ramos/productos específicos, jurisdicciones o territorios de alto riesgo, características del sujeto, montos, plazos). Reescribe el control sugerido REINCORPORANDO TODOS esos detalles del original, SIN perder ninguno. Reglas obligatorias:
- Mantén la estructura: inicia con "Alertamiento", luego "Mediante el cual se ..." (análisis), "Como resultado, se ..." (reporte), "Para lo cual se ..." (decisión). Los tres "se".
- Conserva TODAS las listas de ramos/productos, jurisdicciones, territorios, montos y plazos que aparezcan en el original.
- Conserva OBLIGATORIAMENTE los complementos de perfil del sujeto/beneficiario: persona expuesta políticamente (PEP), manejo de recurso público, contratista del Estado, consorcio/unión temporal, tipo de cliente (natural/jurídica), beneficiario final. Estos son factores de riesgo LAFT valiosos: no los elimines, enriquécelos si es posible.
- No inventes datos nuevos; solo recupera los del original y organízalos bien.
- El resultado debe ser tan detallado o MÁS que el original, nunca más corto.
- Devuelve SOLO el control reescrito, sin comillas ni explicaciones.

CONTROL ORIGINAL (fuente de los detalles): ${controlOriginal}

CONTROL SUGERIDO (a enriquecer): ${controlSugerido}`;
    try {
        const resp = await axios.post('https://api.openai.com/v1/chat/completions', {
            model,
            messages: [
                { role: 'system', content: 'Reescribes controles LAFT en español conservando todo el detalle, sin markdown ni comillas.' },
                { role: 'user', content: prompt }
            ],
            temperature: 0.2,
            max_tokens: 900
        }, { headers: { 'Authorization': `Bearer ${apiKey}`, 'Content-Type': 'application/json' } });
        let out = (resp.data.choices[0].message.content || '').trim().replace(/^["'`]+|["'`]+$/g, '');
        return out && /^alertamiento/i.test(out) && out.length >= controlSugerido.length ? out : controlSugerido;
    } catch (e) {
        console.warn('⚠️ No se pudo recuperar contexto:', e.message);
        return controlSugerido;
    }
}

/**
 * Verifica las reglas de una CAUSA sugerida y devuelve lista de violaciones.
 */
function verificarCausa(causa) {
    const v = [];
    if (quedanQueRelativos(causa)) v.push('contiene el pronombre relativo "que" (debe eliminarse)');
    if (/lavado\s+de\s+activos\s+de\s+(corrupci|narcotr|miner|extorsi|contrabando|cohecho|soborno)/i.test(causa))
        v.push('usa "lavado de activos DE <delito>" (debe ser "y")');
    return v;
}

/**
 * Verifica las reglas de un CONTROL sugerido y devuelve lista de violaciones.
 */
function verificarControl(control) {
    const v = [];
    const c = control || '';
    if (!/^alertamiento\b/i.test(c.trim())) v.push('no inicia con "Alertamiento"');
    if (!/mediante/i.test(c)) v.push('no usa el conector "Mediante"');
    const ses = (c.match(/\bse\s+[a-záéíóúñ]+/gi) || []).length;
    if (ses < 3) v.push(`solo tiene ${ses} construcciones "se+verbo" (se esperan 3)`);
    if (!/como\s+resultado/i.test(c)) v.push('falta el componente de reporte ("Como resultado, se ...")');
    // Toma de decisiones: debe existir "Para ... se <verbo de resolución>" (no solo "para prevenir/evitar")
    const tieneDecisionSe = /\bpara\s+(?:lo\s+cual\s+|el\s+cual\s+)?se\s+(restring|cancel|suspend|bloque|declin|aplica|ejecuta|adopta|decid|marca|activa|reduce|inactiva|no\s+renov)/i.test(c);
    if (!tieneDecisionSe) v.push('falta el componente de toma de decisiones ("Para lo cual se restringe/cancela/aplica ...")');
    if (quedanQueRelativos(c)) v.push('contiene el pronombre relativo "que"');
    if (/alertamiento\s+de\s+(?:la|las|el|los)?\s*alert/i.test(c)) v.push('redundancia "Alertamiento de alertas"');
    return v;
}

/**
 * DOBLE VALIDACIÓN: verifica las reglas de causa y control; si hay violaciones,
 * pide a la IA una pasada de corrección enfocada SOLO en arreglarlas.
 * Devuelve { causa, control } corregidos.
 */
async function validarYCorregir(causa, control) {
    const apiKey = process.env['API-KEY'];
    let cOut = causa, ctrlOut = control;

    const viCausa = verificarCausa(causa);
    const viControl = verificarControl(control);
    if ((!viCausa.length && !viControl.length) || !apiKey) return { causa: cOut, control: ctrlOut };

    const model = process.env.OPENAI_MODEL || 'gpt-4o-mini';
    const prompt = `Corrige el siguiente texto de una matriz LAFT para que cumpla EXACTAMENTE estas reglas, SIN perder información ni contexto (misma o mayor completitud):

REGLAS DE LA CAUSA:
- PROHIBIDO el pronombre relativo "que". Reescribe con gerundios, participios u oraciones directas. (Permitido solo en locuciones "para que", "con el fin de que").
- "lavado de activos de <delito>" es INCORRECTO: usa "provenientes de <delito>" o "lavado de activos y <delito>".
- Conserva sujeto, acción, objeto determinado, delito fuente, tipología y factores de riesgo.

REGLAS DEL CONTROL:
- Inicia OBLIGATORIAMENTE con "Alertamiento" + núcleo verbal + complemento + señal de alerta.
- Estructura de 4 componentes con los TRES "se": "Alertamiento ... [señal]. Mediante el cual se [analiza/verifica] ... Como resultado, se [reporta/documenta] ... Para lo cual se [restringe/cancela/suspende/aplica/ejecuta/adopta] ...".
- El componente de TOMA DE DECISIONES es OBLIGATORIO y debe redactarse "Para lo cual se <verbo de resolución>" (restringe, cancela, suspende, bloquea, aplica, ejecuta, adopta, decide). NO uses solo "para prevenir/evitar" sin el "se + verbo de resolución".
- Ninguna parte del control debe usar el pronombre relativo "que".
- No repitas "alerta" tras "Alertamiento" (nada de "Alertamiento de alertas").
- Conserva TODO el detalle (ramos, jurisdicciones, montos, plazos).

VIOLACIONES DETECTADAS A CORREGIR:
${viCausa.length ? 'CAUSA: ' + viCausa.join('; ') : 'CAUSA: (ok)'}
${viControl.length ? 'CONTROL: ' + viControl.join('; ') : 'CONTROL: (ok)'}

Devuelve EXCLUSIVAMENTE un JSON: {"causa": "...", "control": "..."} con los textos corregidos (sin markdown).

CAUSA ACTUAL: ${causa}
CONTROL ACTUAL: ${control}`;

    try {
        const resp = await axios.post('https://api.openai.com/v1/chat/completions', {
            model,
            messages: [
                { role: 'system', content: 'Corriges textos LAFT en español cumpliendo reglas estrictas. Respondes SOLO JSON válido.' },
                { role: 'user', content: prompt }
            ],
            temperature: 0.1,
            max_tokens: 1200,
            response_format: { type: 'json_object' }
        }, { headers: { 'Authorization': `Bearer ${apiKey}`, 'Content-Type': 'application/json' } });
        let txt = (resp.data.choices[0].message.content || '').replace(/^\s*```(?:json)?\s*/i, '').replace(/\s*```\s*$/i, '').trim();
        const obj = JSON.parse(txt);
        if (obj.causa && obj.causa.length > 20) cOut = obj.causa;
        if (obj.control && obj.control.length > 20) ctrlOut = obj.control;
    } catch (e) {
        console.warn('⚠️ validarYCorregir falló:', e.message);
    }
    return { causa: cOut, control: ctrlOut };
}

/**
 * Completa el DELITO FUENTE (y concreta el objeto si es indeterminado) en la causa,
 * de forma coherente con el resto de su redacción. Devuelve la causa reescrita.
 */
async function completarDelitoFuente(causaSugerida) {
    const apiKey = process.env['API-KEY'];
    if (!apiKey) return causaSugerida;
    const model = process.env.OPENAI_MODEL || 'gpt-4o-mini';
    const prompt = `La siguiente CAUSA de una matriz de riesgo LAFT NO especifica claramente el DELITO FUENTE (el delito subyacente que origina los recursos ilícitos). Reescribe la causa AGREGANDO un delito fuente COHERENTE con el resto de su redacción (sujeto, objeto, tipología y escenario descrito). Reglas:
- USA PREFERENTEMENTE la expresión genérica "delitos fuente de lavado de activos". Ejemplos: "recursos provenientes de delitos fuente de lavado de activos", "activos ilícitos asociados a delitos fuente de lavado de activos". NO nombres "corrupción" ni otro delito específico salvo que el texto original ya lo mencione explícitamente.
- IMPORTANTE: el "lavado de activos" (el proceso de legitimar recursos) es distinto del "delito fuente" (el que origina los recursos). La expresión correcta es "delitos fuente de lavado de activos". NUNCA escribas "provenientes de lavado de activos de corrupción" (incorrecto).
- Si el objeto es indeterminado ("recursos", "comportamiento transaccional", "información financiera"), concrétalo con el producto real (título de capitalización, póliza, cuenta, prima).
- NO uses el pronombre relativo "que".
- Conserva todos los hechos existentes; solo completa lo que falta.
- Devuelve SOLO la causa reescrita, sin comillas ni explicaciones.

CAUSA A COMPLETAR: ${causaSugerida}`;
    try {
        const resp = await axios.post('https://api.openai.com/v1/chat/completions', {
            model,
            messages: [
                { role: 'system', content: 'Reescribes causas LAFT en español, sin markdown, sin comillas y sin el pronombre relativo "que".' },
                { role: 'user', content: prompt }
            ],
            temperature: 0.2,
            max_tokens: 600
        }, { headers: { 'Authorization': `Bearer ${apiKey}`, 'Content-Type': 'application/json' } });
        let out = (resp.data.choices[0].message.content || '').trim().replace(/^["'`]+|["'`]+$/g, '');
        return out && out.length > 20 ? out : causaSugerida;
    } catch (e) {
        console.warn('⚠️ No se pudo completar el delito fuente:', e.message);
        return causaSugerida;
    }
}

/**
 * Llama a OpenAI con el prompt de validación y devuelve el objeto analysis parseado.
 */
async function analizarTrio(riesgo, causa, control) {
    const apiKey = process.env['API-KEY'];
    if (!apiKey) throw Object.assign(new Error('API-KEY no configurada. Verifica el archivo .env'), { statusCode: 500 });

    const prompt = VALIDATION_PROMPT
        .replace('{riesgo}', riesgo)
        .replace('{causa}', causa)
        .replace('{control}', control);

    const model = process.env.OPENAI_MODEL || 'gpt-4o-mini';
    const response = await axios.post('https://api.openai.com/v1/chat/completions', {
        model: model,
        messages: [
            { role: 'system', content: 'Eres un evaluador SFC que aplica los flujogramas LAFT. Respondes SIEMPRE en JSON válido, sin markdown.' },
            { role: 'user', content: prompt }
        ],
        temperature: 0.2,
        max_tokens: 2500,
        response_format: { type: 'json_object' }
    }, {
        headers: { 'Authorization': `Bearer ${apiKey}`, 'Content-Type': 'application/json' }
    });

    let txt = response.data.choices[0].message.content
        .replace(/^\s*```(?:json)?\s*/i, '')
        .replace(/\s*```\s*$/i, '')
        .trim();

    let analysis;
    try { analysis = JSON.parse(txt); }
    catch (e) {
        const m = txt.match(/\{[\s\S]*\}/);
        if (m) { try { analysis = JSON.parse(m[0]); } catch (e2) {} }
    }
    if (!analysis) throw Object.assign(new Error('GPT no retornó JSON válido'), { statusCode: 502 });

    // Aplicar regla de tope: control nunca supera a causa
    const cs = Number(analysis.causa_assessment?.score);
    const cos = Number(analysis.control_assessment?.score);
    if (analysis.control_assessment && cs && cos && cos > cs) {
        analysis.control_assessment.score = cs;
        analysis.control_assessment.findings =
            (analysis.control_assessment.findings || '') + ' [Ajuste regla de tope: control ≤ causa]';
    }

    // Salvaguarda: completar el conteo de "se" y el desglose de componentes del control
    // evaluando el texto ORIGINAL, por si la IA no los devolvió consistentes.
    if (analysis.control_assessment) {
        const ca = analysis.control_assessment;
        const seCount = (control.match(/\bse\s+[a-záéíóúñ]+/gi) || []).length;
        if (ca.conteo_se === undefined || ca.conteo_se === null || String(ca.conteo_se).match(/\d+/)?.[0] === undefined) {
            ca.conteo_se = `${seCount}`;
        }
        // Detección léxica de componentes como respaldo (no sobrescribe si la IA ya marcó false intencional)
        ca.componentes = ca.componentes || {};
        const det = {
            alertamiento: /alert/i.test(control),
            analisis: /analiz|evalú|evalua|verific|examin|contrast/i.test(control),
            reporte: /report|document|registr|ROS|UIAF|eleva/i.test(control),
            decision: /restring|cancel|suspend|bloque|declin|seguimiento|decid|no renovaci/i.test(control)
        };
        ['alertamiento','analisis','reporte','decision'].forEach(k => {
            if (ca.componentes[k] === undefined) ca.componentes[k] = det[k];
        });
    }

    // Detectar naturaleza: ¿el registro es de Capitalizadora? (no aseguradora)
    const textoOriginal = `${riesgo} ${causa} ${control}`;
    const esCapitalizadora = /capitalizadora/i.test(textoOriginal);

    // Salvaguarda de redacción: quitar "que" residual + corregir verbo-objeto en la causa sugerida
    if (analysis.redaccion_sugerida?.causa) {
        let cs = limpiarQue(analysis.redaccion_sugerida.causa);
        cs = corregirVerboObjeto(cs); // "realizan pólizas" -> "suscriben pólizas"
        analysis.redaccion_sugerida.causa = cs;
    }

    // Salvaguarda de naturaleza: corregir "sistema asegurador/financiero" en registros de Capitalizadora
    if (esCapitalizadora && analysis.redaccion_sugerida) {
        ['riesgo', 'causa', 'control'].forEach(k => {
            if (analysis.redaccion_sugerida[k]) {
                analysis.redaccion_sugerida[k] = corregirSistema(analysis.redaccion_sugerida[k], true);
            }
        });
    }

    // Salvaguarda: el control sugerido DEBE iniciar con "Alertamiento" (sin redundancias)
    if (analysis.redaccion_sugerida?.control) {
        let c = analysis.redaccion_sugerida.control.trim();
        if (!/^alertamiento\b/i.test(c)) {
            c = c.charAt(0).toLowerCase() + c.slice(1);
            // Evitar "Alertamiento asociado a alertas..." -> si ya empieza por "alerta", no duplicar
            c = /^alert/i.test(c) ? `Alertamiento de ${c.replace(/^alert(?:a|as|amiento|amientos)\s+(?:de\s+|sobre\s+)?/i, '')}` : `Alertamiento asociado a ${c}`;
        } else {
            c = c.charAt(0).toUpperCase() + c.slice(1);
        }
        analysis.redaccion_sugerida.control = c;
    }

    // Salvaguarda: eliminar redundancias/abreviaturas y EXPANDIR SIGLAS en la redacción sugerida
    if (analysis.redaccion_sugerida) {
        const origMap = { riesgo, causa, control };
        ['riesgo', 'causa', 'control'].forEach(k => {
            if (analysis.redaccion_sugerida[k]) {
                let v = limpiarRedundancias(analysis.redaccion_sugerida[k]);
                v = expandirSiglas(v); // siglas completas: ROS, UIAF, PEP, etc.
                v = corregirDelitos(v, origMap[k]); // corrige "lavado de activos DE" y "corrupción" genérico
                analysis.redaccion_sugerida[k] = v;
            }
        });
    }

    // Salvaguarda: el control sugerido debe contener los tres "se" (análisis, reporte, decisión)
    if (analysis.redaccion_sugerida?.control) {
        const ctrl = analysis.redaccion_sugerida.control;
        const ses = (ctrl.match(/\bse\s+[a-záéíóúñ]+/gi) || []).length;
        const mediante = /mediante/i.test(ctrl);
        // Detección de los 4 componentes por señales léxicas
        const tieneAlerta = /alert/i.test(ctrl);
        const tieneAnalisis = /analiz|evalú|evalua|verific|examin|contrast/i.test(ctrl);
        const tieneReporte = /report|document|registr|ROS|UIAF|eleva/i.test(ctrl);
        const tieneDecision = /restring|cancel|suspend|bloque|declin|aplica seguimiento|decid|no renovaci/i.test(ctrl);
        const faltantes = [];
        if (!tieneAlerta) faltantes.push('alertamiento');
        if (!tieneAnalisis) faltantes.push('análisis');
        if (!tieneReporte) faltantes.push('reporte');
        if (!tieneDecision) faltantes.push('toma de decisiones');
        if (ses < 3 || !mediante || faltantes.length) {
            analysis.redaccion_sugerida._advertencia_control =
                `Control: ${ses} "se+verbo" (se esperan 3), ${mediante ? 'usa' : 'NO usa'} "Mediante"` +
                (faltantes.length ? `, faltan componentes: ${faltantes.join(', ')}` : ', 4 componentes presentes') + '.';
        }

        // Verificar estructura del ALERTAMIENTO (oración principal): núcleo verbal + complemento + señal de alerta
        // Incluye: umbrales ($, %), patrones (coincidencias, listas, anomalías), condiciones (cuando, si, por)
        const señalRegex = /\b(superior|inferior|mayor|menor|excede|super[ae]|umbral|tope|límite|limite|patr[óo]n|variaci[óo]n|inusual|at[íi]pic|nuev[ao]|sin\s+justificaci|alto\s+riesgo|fracciona|pitufeo|frecuen|parametrizad|cumplimiento\s+en\s+las\s+señal|coincidencias|listas\s+(?:restrictivas|sospechosas|negras|sancionadas)|anomal|inconsisten|ausencia\s+de|incongruencia|cuando\s+(?:el|la|los|las)|\bsi\s+(?:el|la|los|las)|red flag|flag|alert|>\s*\d|\d+\s*%)/i;
        let alerta = (ctrl.split(/\.\s+(?=Mediante|Como|Para)/i)[0] || ctrl).trim();

        // ¿El control ORIGINAL tenía señales CONCRETAS (montos, plazos, %, umbrales específicos)?
        const señalConcretaRegex = /\$\s*[\d.,]+|[\d.,]+\s*(millones|mil|smmlv|salarios|d[óo]lares|usd|pesos|cop)|\d+\s*%|(?:menor|mayor|superior|inferior)\s+a\s+[\w$.,]+|periodo\s+(?:menor|inferior|de)\s+\w+|(?:tres|dos|un|seis|doce|\d+)\s+(?:d[íi]as|meses|años)|desde\s+la\s+emisi[óo]n|fraccionamiento|pitufeo|jurisdicci[óo]n(?:es)?\s+de\s+(?:alto\s+riesgo|riesgo)/i;
        const originalTieneSenalConcreta = señalConcretaRegex.test(control || '');
        const sugeridaTieneSenalConcreta = señalConcretaRegex.test(ctrl);

        // Si NO hay señal en la sugerencia, insertar fórmula genérica.
        // PERO si el original tenía señales concretas y la sugerencia las perdió, advertir (no empobrecer).
        if (!señalRegex.test(alerta) && !originalTieneSenalConcreta) {
            const nuevoControl = completarSenalAlerta(ctrl);
            if (nuevoControl && nuevoControl !== ctrl) {
                let cc = limpiarRedundancias(nuevoControl);
                cc = expandirSiglas(cc);
                analysis.redaccion_sugerida.control = cc;
            }
        } else if (originalTieneSenalConcreta && !sugeridaTieneSenalConcreta) {
            // El original tenía umbral/monto concreto y la sugerencia lo perdió -> advertir
            analysis.redaccion_sugerida._advertencia_senal =
                'El control original tenía señales de alerta concretas (montos/plazos/umbrales) que la sugerencia simplificó. Revisa que no se hayan perdido: ' +
                (control.match(señalConcretaRegex)?.[0] ? '"' + control.match(señalConcretaRegex)[0] + '..."' : '');
        }

        // Detección de PÉRDIDA DE CONTEXTO: si el original tiene detalles específicos que
        // la sugerencia perdió (listas entre paréntesis, ramos, jurisdicciones), advertir.
        const sugActual = analysis.redaccion_sugerida.control || '';
        const detallesPerdidos = [];
        // Listas entre paréntesis del original (ej: "(ramos cumplimiento, transporte, ...)")
        const parentOrig = (control || '').match(/\(([^)]{15,})\)/g) || [];
        parentOrig.forEach(p => {
            const contenido = p.replace(/[()]/g, '').trim();
            // ¿alguna palabra clave de esa lista sobrevive en la sugerencia?
            const palabras = contenido.split(/[,\s]+/).filter(w => w.length > 4);
            const sobrevive = palabras.some(w => new RegExp(w.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'i').test(sugActual));
            if (!sobrevive && palabras.length) detallesPerdidos.push(p);
        });
        // Menciones de jurisdicción/territorio de alto riesgo
        const jurisOrig = /jurisdicci[óo]n(?:es)?\s+(?:cataloga\w*\s+)?(?:como\s+)?de\s+(?:alto\s+)?riesgo|territorios?\s+con\s+exposici[óo]n|nacionalidad\s+o\s+residencia/i;
        if (jurisOrig.test(control || '') && !jurisOrig.test(sugActual)) {
            detallesPerdidos.push('mención de jurisdicción/territorio de alto riesgo');
        }
        // Complementos clave frecuentes que NO deben perderse (PEP, tipo de beneficiario, recurso público)
        const complementosClave = [
            [/persona\s+expuesta\s+pol[íi]ticamente|\bPEP\b|expuesto\s+pol[íi]tic/i, 'persona expuesta políticamente (PEP)'],
            [/recurso\s+p[úu]blico|manejo\s+de\s+recurso/i, 'manejo de recurso público'],
            [/beneficiario\s+(?:final|real)|beneficiario\s+est[áa]/i, 'característica del beneficiario'],
            [/consorcio|uni[óo]n\s+temporal|UT\b/i, 'figura de consorcio/unión temporal'],
            [/contratista\s+del\s+estado|contratos?\s+p[úu]blic/i, 'contratista del Estado / contratos públicos'],
            [/persona\s+natural(?:es)?\s+y\s+jur[íi]dica|personas?\s+jur[íi]dica/i, 'tipo de cliente (natural/jurídica)']
        ];
        complementosClave.forEach(([re, etiqueta]) => {
            if (re.test(control || '') && !re.test(sugActual)) detallesPerdidos.push(etiqueta);
        });
        // Heurística de longitud: sugerencia mucho más corta que el original (perdió contenido)
        const lenOrig = (control || '').length;
        const lenSug = sugActual.length;
        const muchoMasCorta = lenOrig > 200 && lenSug < lenOrig * 0.6;

        if (detallesPerdidos.length || muchoMasCorta) {
            // Intentar recuperar el contexto perdido con una segunda pasada
            const recuperado = await recuperarContexto(analysis.redaccion_sugerida.control, control);
            if (recuperado && recuperado !== analysis.redaccion_sugerida.control) {
                let cc = limpiarRedundancias(recuperado);
                cc = expandirSiglas(cc);
                cc = corregirDelitos(cc, control);
                analysis.redaccion_sugerida.control = cc;
            }
            // Re-evaluar si aún falta algo tras la recuperación
            const sugFinal = analysis.redaccion_sugerida.control || '';
            const aunFalta = detallesPerdidos.filter(p => {
                const palabras = p.replace(/[()]/g, '').split(/[,\s]+/).filter(w => w.length > 4);
                return palabras.length && !palabras.some(w => new RegExp(w.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'i').test(sugFinal));
            });
            if (aunFalta.length) {
                analysis.redaccion_sugerida._advertencia_contexto =
                    'Revisa: puede faltar contexto del control original (' + aunFalta.slice(0, 2).join('; ') + '). Complétalo antes de guardar.';
            }
        }

        // Re-verificar estructura del alertamiento sobre el control final (el sugerido)
        const ctrlFinal = analysis.redaccion_sugerida.control;
        alerta = (ctrlFinal.split(/\.\s+(?=Mediante|Como|Para)/i)[0] || ctrlFinal).trim();
        const iniciaAlert = /^alertamiento\b/i.test(alerta);
        const nucleo = /\b(generad|generaci|identific|detect|deteccion|marcaci|activaci|monitore|verificaci|validaci|revisi|control|seguimiento|alertamiento\s+de)/i.test(alerta);
        const senal = señalRegex.test(alerta);
        const faltAlert = [];
        if (!iniciaAlert) faltAlert.push('no inicia con "Alertamiento"');
        if (!nucleo) faltAlert.push('sin núcleo verbal claro');
        if (!senal) faltAlert.push('sin señal de alerta (umbral/patrón)');
        if (faltAlert.length) {
            analysis.redaccion_sugerida._advertencia_alertamiento =
                `Alertamiento: ${faltAlert.join('; ')}. Estructura esperada: "Alertamiento" + núcleo verbal + complemento + señal de alerta.`;
        }

        // IMPORTANTE: sincronizar el desglose mostrado con la REDACCIÓN SUGERIDA final,
        // porque es lo que el usuario ve/guarda (evita que muestre datos del control original).
        if (analysis.control_assessment) {
            const partes = ctrlFinal.split(/\.\s+(?=Mediante|Como|Para)/i);
            const co2 = analysis.control_assessment;
            co2.desglose = co2.desglose || {};
            if (partes[0]) co2.desglose.alertamiento = partes[0].trim();
            partes.forEach(p => {
                if (/^mediante/i.test(p.trim())) co2.desglose.analisis = p.trim();
                else if (/^como resultado/i.test(p.trim())) co2.desglose.reporte = p.trim();
                else if (/^para\b/i.test(p.trim())) co2.desglose.decision = p.trim();
            });
            // Núcleo verbal: normalizar al infinitivo/sustantivo base (evita "detectadas")
            const nucMatch = alerta.match(/\b(generaci[óo]n|generad\w*|identificaci[óo]n|identific\w*|detecci[óo]n|detect\w*|marcaci[óo]n|activaci[óo]n|monitoreo|verificaci[óo]n|validaci[óo]n|revisi[óo]n)/i);
            const nucBase = {
                'generadas':'generación','generado':'generación','generados':'generación','generada':'generación',
                'identifica':'identificación','detectadas':'detección','detectado':'detección','detecta':'detección',
                'valida':'validación','verifica':'verificación','revisa':'revisión','activa':'activación','marca':'marcación'
            };
            let nuc = nucMatch ? nucMatch[0] : '';
            nuc = nucBase[nuc.toLowerCase()] || nuc;
            // Señal de alerta: extraer el fragmento del umbral/patrón
            const senMatch = alerta.match(/(cuando[^.]*|superiores?\s+a[^.,]*|mayor(?:es)?[^.,]*|con\s+variaci[óo]n[^.,]*|patr[óo]n[^.,]*|sin\s+justificaci[óo]n[^.,]*|>\s*\d[^.,]*|\d+\s*%[^.,]*|inusual[^.,]*|at[íi]pic[^.,]*)/i);
            co2.alertamiento_estructura = {
                inicia_con_alertamiento: iniciaAlert,
                nucleo_verbal: nuc || (nucleo ? 'presente' : 'no detectado'),
                complemento: nucleo ? 'presente' : 'no detectado',
                senal_alerta: senal ? (senMatch ? senMatch[0].trim() : 'presente') : 'no detectada'
            };
        }
    }

    // Salvaguarda: si la causa sugerida NO tiene delito fuente, completarlo coherente
    if (analysis.redaccion_sugerida?.causa) {
        const cLow = analysis.redaccion_sugerida.causa.toLowerCase();
        const delitoRegex = /narcotr[áa]fico|corrupci[óo]n|dcap|miner[íi]a\s+ilegal|extorsi[óo]n|contrabando|trata\s+de\s+personas|cohecho|soborno|financiaci[óo]n\s+del\s+terrorismo|tr[áa]fico\s+de|delito\s+fuente|actividades?\s+il[íi]citas?\s+de/i;
        const iaDiceNoDelito = /no\s+especificad/i.test(String(analysis.causa_assessment?.delito_fuente || ''));
        if (!delitoRegex.test(cLow) || iaDiceNoDelito) {
            const nuevaCausa = await completarDelitoFuente(analysis.redaccion_sugerida.causa);
            if (nuevaCausa && nuevaCausa !== analysis.redaccion_sugerida.causa) {
                let cc = limpiarQue(nuevaCausa);
                cc = corregirVerboObjeto(cc);
                cc = expandirSiglas(cc);
                cc = corregirDelitos(cc, causa); // corrige "lavado de activos DE" y "corrupción" genérico
                analysis.redaccion_sugerida.causa = cc;
            }
        }
    }

    // Salvaguarda: la causa sugerida con score 3 debe tener ≥3 factores de riesgo
    if (analysis.redaccion_sugerida?.causa) {
        const c = analysis.redaccion_sugerida.causa.toLowerCase();
        const factores = {
            cliente: /client|usuari|persona|contratant|asegurad|tomador|benefici|fideicomit/i.test(c),
            producto: /título|titulo|capitaliz|póliza|poliza|cuenta|prima|cuota|seguro|producto|rescate/i.test(c),
            canal: /canal|digital|presencial|corresponsal|ventanilla|cajero|sucursal|línea|linea|app|web/i.test(c),
            jurisdiccion: /jurisdicci|país|pais|municipio|zona|frontera|región|region|exterior|internacional|extranjer/i.test(c)
        };
        const n = Object.values(factores).filter(Boolean).length;
        const causaScore = Number(analysis.causa_assessment?.score);
        if (causaScore === 3 && n < 3) {
            analysis.redaccion_sugerida._advertencia_causa =
                `La causa sugerida se calificó 3 pero solo se detectan ${n} de 4 factores de riesgo (${Object.entries(factores).filter(([,v])=>v).map(([k])=>k).join(', ') || 'ninguno'}). Debería incluir al menos 3.`;
        }
    }

    // Salvaguarda de COHERENCIA INTERNA: filtrar debilidades que contradicen los scores.
    // Si la causa tiene score 3, no tiene sentido listar "la causa no especifica objeto/contexto".
    if (Array.isArray(analysis.weaknesses)) {
        const causaScore = Number(analysis.causa_assessment?.score);
        const controlScore = Number(analysis.control_assessment?.score);
        analysis.weaknesses = analysis.weaknesses.filter(w => {
            const t = String(w).toLowerCase();
            const esSobreCausa = /\bcausa\b/.test(t);
            const esSobreControl = /\bcontrol\b/.test(t);
            const dicefalta = /no\s+(especifica|detalla|incluye|determina|describe|expresa|contempla)|carece|falta|ausencia|no\s+es\s+claro/.test(t);
            // Descartar debilidad sobre la causa si la causa YA es 3
            if (esSobreCausa && !esSobreControl && dicefalta && causaScore === 3) return false;
            // Descartar debilidad sobre el control si el control YA es 3
            if (esSobreControl && !esSobreCausa && dicefalta && controlScore === 3) return false;
            return true;
        });
    }

    // ===== DOBLE VALIDACIÓN FINAL =====
    // Verifica todas las reglas sobre la redacción sugerida ya procesada.
    // Si aún hay violaciones (ej: "que" residual, falta componente), pide corrección a la IA
    // y aplica de nuevo las limpiezas locales. Máximo 2 intentos.
    if (analysis.redaccion_sugerida) {
        for (let intento = 0; intento < 2; intento++) {
            const causaAct = analysis.redaccion_sugerida.causa || '';
            const controlAct = analysis.redaccion_sugerida.control || '';
            const viC = verificarCausa(causaAct);
            const viCtrl = verificarControl(controlAct);
            if (!viC.length && !viCtrl.length) break; // todo cumple

            const corr = await validarYCorregir(causaAct, controlAct);
            // Re-aplicar limpiezas locales al resultado corregido
            let nc = limpiarQue(corr.causa); nc = corregirVerboObjeto(nc); nc = expandirSiglas(nc); nc = corregirDelitos(nc, causa);
            let nk = limpiarRedundancias(corr.control); nk = expandirSiglas(nk); nk = corregirDelitos(nk, control);
            analysis.redaccion_sugerida.causa = nc;
            analysis.redaccion_sugerida.control = nk;
        }
        // Advertencia final: recalcular desde cero. Si ya cumple, LIMPIAR advertencias viejas.
        const restC = verificarCausa(analysis.redaccion_sugerida.causa || '');
        const restK = verificarControl(analysis.redaccion_sugerida.control || '');
        if (restC.length) analysis.redaccion_sugerida._advertencia_causa = 'Revisar causa: ' + restC.join('; ') + '.';
        else delete analysis.redaccion_sugerida._advertencia_causa;
        if (restK.length) analysis.redaccion_sugerida._advertencia_control = 'Revisar control: ' + restK.join('; ') + '.';
        else delete analysis.redaccion_sugerida._advertencia_control;
    }

    return analysis;
}

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());
app.use(express.static(path.join(__dirname)));

// Servir data.json explícitamente
app.get('/data.json', (req, res) => {
    res.sendFile(path.join(__dirname, 'data.json'));
});

/**
 * PROMPT PARA VALIDACIÓN CON GPT NANO
 * Basado en: Guía de Identificación de Riesgos LAFT V3 (Jun 2026)
 */
const VALIDATION_PROMPT = `Eres un evaluador experto de la Superintendencia Financiera de Colombia (SFC) que aplica EXACTAMENTE los flujogramas oficiales de evaluación de causas y controles LA/FT/FPADM (Anexos 1 y 2, V2 Junio 2026). Evalúas matrices de riesgo del grupo Bolívar (que incluye distintas compañías: Capitalizadora Bolívar, Seguros Bolívar, Comerciales Bolívar, etc.).

═══════════════════════════════════════════
REGLA CRÍTICA DE FIDELIDAD — NO INVENTES NI SUSTITUYAS DATOS
═══════════════════════════════════════════
Debes evaluar y reescribir SOLO con base en el texto del registro que se te entrega. PROHIBIDO:
- Cambiar el NOMBRE DE LA ENTIDAD/COMPAÑÍA. Usa EXACTAMENTE el que aparezca en el texto original (si dice "Comerciales Bolívar", NO lo cambies a "Capitalizadora Bolívar"; si dice "Seguros Bolívar", consérvalo). Si el texto no nombra entidad, usa un término genérico neutro como "la Compañía" o "la entidad".
- Inventar productos, montos, jurisdicciones, delitos o tipologías que NO estén en el texto original.
- Introducir "títulos de capitalización" u otro producto si el registro NO lo menciona.
La redacción sugerida debe COMPLETAR la estructura (sujeto/acción/objeto/contexto y los componentes del control) usando ÚNICAMENTE los hechos ya presentes en el registro, reescritos con mejor sintaxis. Si falta un elemento, descríbelo de forma genérica sin inventar especificidad que no exista.

═══════════════════════════════════════════
NATURALEZA DE CADA COMPAÑÍA — USA EL TÉRMINO CORRECTO (NO GENERALICES)
═══════════════════════════════════════════
Según la entidad que aparezca en el registro, el "sistema" al que pertenece es distinto. NUNCA escribas "sistema financiero o asegurador" de forma genérica ni mezcles ambos:
- "CAPITALIZADORA BOLÍVAR" → es una SOCIEDAD DE CAPITALIZACIÓN. NO es aseguradora. PROHIBIDO decir "sistema asegurador" o "sistema financiero". Usa "sector de capitalización", "el sistema de capitalización" o simplemente "la Compañía"/"la entidad".
- "SEGUROS BOLÍVAR" → pertenece al SISTEMA ASEGURADOR. Usa "sistema asegurador".
- "COMERCIALES BOLÍVAR" → pertenece al SISTEMA ASEGURADOR. Usa "sistema asegurador".
- Si el registro no nombra entidad, usa un término neutro ("la Compañía", "la entidad") sin atribuir naturaleza aseguradora ni financiera.

═══════════════════════════════════════════
CONTEXTO DE NEGOCIO — "TÍTULOS" (aplica SOLO si el texto menciona títulos)
═══════════════════════════════════════════
SI (y solo si) el registro menciona "títulos", interprétalos como TÍTULOS DE CAPITALIZACIÓN (el producto de la Capitalizadora), NO como títulos valores/acciones/bonos.
- En ese caso, el título de capitalización cuenta para el factor de riesgo "producto/servicio".
- Nómbralo correctamente: "pagos de cuotas del título de capitalización", "rescate/cancelación anticipada del título", "suscripción de títulos de capitalización", etc.
- Si el registro NO habla de títulos, NO los menciones.

═══════════════════════════════════════════
FLUJOGRAMA DE CAUSAS (calificación 1, 2 o 3)
═══════════════════════════════════════════
Responde en orden:

P1. ¿La causa propone CÓMO la entidad puede ser utilizada para LA/FT/FPADM e incluye los elementos esperados (SUJETO, ACCIÓN, OBJETO, CONTEXTO delictivo)?
  → NO (Nota 1, PUNTAJE 1) si: modela fallas operativas de controles; NO propone un sujeto; está descrita como un riesgo; o le falta uno o más elementos esenciales (sujeto, acción, objeto o contexto).

P2 (solo si P1=Sí). ¿La causa determina el OBJETO y detalla el CONTEXTO delictivo (circunstancias específicas, métodos o herramientas)?
  → Si NO determina objeto Y/O no detalla contexto → preguntar: ¿La causa es GENÉRICA?
     • GENÉRICA (Nota 2, PUNTAJE 1) = sujeto + acción sobre OBJETO INDETERMINADO (recursos, recursos ilícitos, transacciones, operaciones, bienes, activos, fondos) Y contexto delictivo que NO describe circunstancias específicas.
     • NO genérica pero incompleta (Nota 3, PUNTAJE 2): tiene UNA de las dos: (a) objeto indeterminado pero contexto SÍ específico (pitufeo, estructuración, etc.), o (b) objeto determinado pero contexto NO específico.

P3 (solo si P2=Sí, objeto determinado Y contexto específico). ¿La causa identifica al menos 3 de los 4 factores de riesgo (cliente, producto/servicio, canal, jurisdicción), incluyendo SIEMPRE el cliente?
  → SÍ (Nota 4, PUNTAJE 3): sujeto + acción + objeto determinado + contexto específico + ≥3 factores de riesgo.
  → NO (PUNTAJE 2): tiene objeto determinado y contexto específico pero menos de 3 factores de riesgo.

═══════════════════════════════════════════
FLUJOGRAMA DE CONTROLES (calificación 1, 2 o 3)
═══════════════════════════════════════════
Responde en orden:

C1. ¿El control está DISEÑADO para mitigar una conducta delictiva asociada al riesgo de LA/FT (NO un evento operativo)?
  → NO (Nota 1, PUNTAJE 1) si: incurre en ERROR CATEGÓRICO (control operativo: control de acceso, doble control de inventario, recolectar documentación, capacitaciones, validar antecedentes, control del control); o NO describe explícitamente cómo mitiga un escenario de conducta delictiva/tipología.
  → REGLA OBLIGATORIA: si la CAUSA fue calificada en 1, el control NO puede superar 1 (también es 1).

C2 (solo si C1=Sí). ¿El control incluye los CUATRO componentes mínimos: ALERTAMIENTO + ANÁLISIS + REPORTE + TOMA DE DECISIONES?
  → NO (Nota 2, PUNTAJE 2) si falta uno o varios de los cuatro componentes.

C3 (solo si C2=Sí). ¿El control incluye los ELEMENTOS CLAVE de la causa (mismo sujeto, objeto, tipología) y evidencia cómo reduce probabilidad/impacto?
  → NO (Nota 3, PUNTAJE 2) si no guarda coherencia con la causa definida.
  → SÍ (Nota 4, PUNTAJE 3): diseñado para mitigar conducta delictiva + 4 componentes + coherente con la causa.

REGLA DE TOPE: el puntaje del control NUNCA puede superar el de la causa (control_score ≤ causa_score).

═══════════════════════════════════════════
RIESGO
═══════════════════════════════════════════
Debe expresar "Posibilidad de que [ENTIDAD] sea utilizada para LA/FT/FPADM". Score 3 si es claro y menciona delitos subyacentes; 2 si es claro pero genérico; 1 si no expresa exposición a LAFT.

═══════════════════════════════════════════
REDACCIÓN SUGERIDA — DEBE ALCANZAR PUNTAJE 3 DE VERDAD (COMPLETAR, NUNCA REDUCIR)
═══════════════════════════════════════════
La redacción sugerida NO es un resumen: es la versión ÓPTIMA que efectivamente merece puntaje 3. Regla de oro: COMPLETA los elementos faltantes, jamás elimines los que ya existen. La versión sugerida debe ser igual o MÁS COMPLETA que la original, nunca más corta en contenido.

⚠️ REGLA DE CONSERVACIÓN DE CONTEXTO (CRÍTICA, aplica a CAUSA y CONTROL): DEBES PRESERVAR TODOS los detalles específicos del texto original. Está PROHIBIDO resumir, generalizar o eliminar información concreta como:
  - Listas de ramos/productos específicos (ej: "ramos cumplimiento, transporte, maquinaria y equipo, navegación").
  - Jurisdicciones, territorios, países o zonas mencionadas (ej: "jurisdicción de alto riesgo", "territorios con exposición elevada").
  - Montos, plazos, porcentajes, umbrales concretos.
  - Características del sujeto/beneficiario (ej: "tomador con nacionalidad o residencia en jurisdicción de alto riesgo", "beneficiario catalogado como persona expuesta políticamente (PEP) por manejo de recurso público", "contratista del Estado bajo figura de consorcio/unión temporal"). Estos perfiles son factores de riesgo LAFT: NUNCA los elimines.
  - Cualquier condición, patrón o circunstancia detallada.
Si el control original tiene MUCHO contexto, la sugerencia debe MANTENER TODO ese contexto y solo reorganizarlo en la estructura sintáctica correcta (Alertamiento + Mediante/se + Como resultado/se + Para/se). Reescribir NO significa acortar: significa reordenar conservando el 100% de la información. Un buen resultado es tan detallado o más que el original.

Para la CAUSA sugerida, DEBE contener y ser verificable que incluye TODOS estos elementos (checklist obligatorio para puntaje 3):
  1) SUJETO explícito (quién).
  2) ACCIÓN (verbo activo).
  3) OBJETO DETERMINADO (producto/activo específico, no genérico). Si el objeto original es indeterminado ("recursos", "comportamiento transaccional", "información financiera", "operaciones"), DEBES concretarlo con el producto real de la entidad (título de capitalización, póliza, cuenta, prima, cuota) coherente con el resto de la causa.
  4) DELITO FUENTE (OBLIGATORIO, NUNCA OMITIR): el delito subyacente que origina los recursos. PREFERENCIA DE REDACCIÓN: usa la expresión genérica "delitos fuente de lavado de activos" (o "recursos provenientes de delitos fuente de lavado de activos") en lugar de nombrar un único delito específico como "corrupción". Solo nombra un delito puntual (narcotráfico, minería ilegal, extorsión, etc.) cuando el propio texto original YA lo mencione explícitamente; en ese caso puedes conservarlo. Si la causa original NO especifica delito fuente, agrega "delitos fuente de lavado de activos" de forma coherente con el escenario. Ejemplos: "recursos provenientes de delitos fuente de lavado de activos", "activos ilícitos asociados a delitos fuente de lavado de activos". NUNCA dejes la causa sin delito fuente explícito, pero NO fuerces "corrupción" si no estaba en el original.
  5) TIPOLOGÍA/circunstancias específicas (método: pitufeo, estructuración, empresas fachada, etc.).
  6) AL MENOS 3 DE LOS 4 FACTORES DE RIESGO, incluyendo SIEMPRE el cliente. Los 4 factores son:
     • CLIENTE (tipo de cliente/sujeto) — obligatorio.
     • PRODUCTO/SERVICIO (título de capitalización, póliza, cuenta, etc.).
     • CANAL (digital, presencial, corresponsal, ventanilla, etc.).
     • JURISDICCIÓN (zona/país/municipio de riesgo).
  → Si la causa original solo tenía 2 factores, tu redacción sugerida DEBE agregar explícitamente al menos un tercer factor (producto, canal o jurisdicción) de forma coherente con los hechos, para que realmente cumpla puntaje 3. Nombra los factores de manera que se identifiquen claramente en el texto.
  → En "factores_riesgo" reporta cuántos y cuáles quedan en la REDACCIÓN SUGERIDA (deben ser ≥3 si el score es 3).

Para el CONTROL sugerido, DEBE contener explícitamente los CUATRO componentes (alertamiento + análisis + reporte + toma de decisiones), con la estructura sintáctica de "Mediante" + los tres "se". Si al original le faltaba un componente, tu redacción sugerida DEBE agregarlo. No propongas un control de puntaje 3 al que le falte alguno de los cuatro.

COHERENCIA OBLIGATORIA entre score y redacción sugerida:
- Si el score sugerido implícito es 3, la redacción sugerida DEBE cumplir el checklist completo de arriba. Si no puedes completarlo con los hechos disponibles sin inventar, entonces el elemento NO merece 3: baja el score y explícalo en findings, pero no entregues una redacción "de 3" incompleta.
- Antes de responder, RE-VERIFICA que la causa sugerida tenga ≥3 factores y el control sugerido los 4 componentes. Si falta algo, complétalo.

DEVUELVE EXCLUSIVAMENTE este JSON (sin texto adicional, sin markdown):
{
  "overall_coherence": "Alto|Medio|Bajo",
  "riesgo_assessment": { "score": 1, "findings": "..." },
  "causa_assessment": {
    "score": 1,
    "nota_aplicada": "Nota 1|2|3|4",
    "sujeto": "texto del sujeto detectado o 'no detectado'",
    "accion": "verbo detectado o 'no detectado'",
    "objeto": "determinado: <cual> | indeterminado | no detectado",
    "delito_fuente": "especificado: <cual> | no especificado",
    "tipologia": "descrita: <cual> | no descrita",
    "factores_riesgo": "N de 4 (cliente/producto/canal/jurisdiccion presentes)",
    "findings": "razón de la calificación según el flujograma"
  },
  "control_assessment": {
    "score": 1,
    "nota_aplicada": "Nota 1|2|3|4",
    "componentes": { "alertamiento": true, "analisis": true, "reporte": true, "decision": true },
    "desglose": {
      "alertamiento": "fragmento del control que corresponde al alertamiento, o 'no detectado'",
      "analisis": "fragmento del control que corresponde al análisis (debe llevar 'se'+verbo), o 'no detectado'",
      "reporte": "fragmento del control que corresponde al reporte (debe llevar 'se'+verbo), o 'no detectado'",
      "decision": "fragmento del control que corresponde a la toma de decisiones (debe llevar 'se'+verbo), o 'no detectado'"
    },
    "alertamiento_estructura": {
      "inicia_con_alertamiento": true,
      "nucleo_verbal": "verbo/sustantivo verbal detectado (p.ej. 'generación', 'identificación', 'detección'), o 'no detectado'",
      "complemento": "complemento directo o nominal detectado (qué se alerta), o 'no detectado'",
      "senal_alerta": "señal/umbral/patrón que dispara la alerta (p.ej. 'superiores al tope', 'variación > 50%'), o 'no detectada'"
    },
    "conteo_se": "N (cantidad de construcciones 'se'+verbo detectadas; se esperan 3)",
    "coherencia_con_causa": "alta|media|baja",
    "error_categorico": true,
    "findings": "razón de la calificación según el flujograma"
  },
  "redaccion_sugerida": {
    "riesgo": "redacción mejorada del riesgo",
    "causa": "redacción mejorada de la causa (puntaje 3), SIN el pronombre relativo 'que'",
    "control": "redacción mejorada del control (puntaje 3, 4 componentes), usando el conector 'Mediante' para introducir el análisis"
  },
  "strengths": ["..."],
  "weaknesses": ["..."],
  "recommendations": ["..."]
}

═══════════════════════════════════════════
REGLAS DE REDACCIÓN OBLIGATORIAS (aplican a redaccion_sugerida Y a los textos citados en findings/recommendations)
═══════════════════════════════════════════
1. CAUSA: NO uses el pronombre relativo "que". Reescribe con construcciones directas (gerundios, participios, "con el fin de", "para", oraciones independientes). Alineado con el Anexo 1 V2 (Jun 2026) que eliminó el "que" de los ejemplos.
   • Mal: "Clientes que realizan pagos que superan el tope..."
   • Bien: "Clientes realizan pagos superiores al tope..."
   • COHERENCIA VERBO-OBJETO (CRÍTICO): el verbo DEBE concordar semánticamente con el objeto. Al quitar el "que" NO dejes verbos genéricos incoherentes. Ejemplos:
     - Una PÓLIZA se SUSCRIBE / ADQUIERE / CONTRATA / TOMA (NO "se realiza una póliza").
     - Un TÍTULO DE CAPITALIZACIÓN se SUSCRIBE / ADQUIERE / CONSTITUYE.
     - Una CUOTA / PAGO se PAGA / REALIZA / ABONA / CANCELA.
     - Una TRANSFERENCIA / OPERACIÓN se REALIZA / EFECTÚA / ORDENA.
     - Un RESCATE / RECLAMACIÓN se SOLICITA / TRAMITA / EFECTÚA.
     - Mal: "Clientes realizan pólizas...". Bien: "Clientes suscriben pólizas..." o "Clientes adquieren pólizas...".
   • Verifica que la oración sea gramaticalmente natural en español antes de entregarla.
2. CONTROL: respeta la estructura sintáctica oficial (PPT Evaluación Controles LAFT). Una ORACIÓN PRINCIPAL (alertamiento) + TRES ORACIONES SUBORDINADAS (análisis, reporte, toma de decisiones). Cada subordinada DEBE construirse como: CONECTOR + "se" + VERBO CONJUGADO. Es decir, el control SIEMPRE debe contener los TRES "se" (uno por subordinada).
   ⚠️ FLUIDEZ GRAMATICAL (CRÍTICO): el alertamiento debe ser una frase natural y bien formada en español. PROHIBIDO construcciones rotas como "Alertamiento se generan automáticas", "Alertamiento genera alertas automáticas", "Alertamiento identifica automáticas". Redacta el inicio de forma fluida, por ejemplo: "Alertamiento generado automáticamente sobre [hecho]...", "Alertamiento sobre [hecho] detectado en el aplicativo de monitoreo...", "Alertamiento generado cuando [condición]...". Lee la frase completa y asegúrate de que suene natural antes de entregarla.
   • ALERTAMIENTO (oración principal): el control DEBE INICIAR OBLIGATORIAMENTE con la palabra "Alertamiento" y contener, en este orden, TRES sub-elementos (estructura de la PPT):
       1) NÚCLEO VERBAL: sustantivo/verbo de acción de monitoreo (generación, identificación, detección, marcación, activación de una alerta). Ej: "Alertamiento generado por...", "Alertamiento que identifica...", "Alertamiento de detección de...".
       2) COMPLEMENTO (CD o CN): QUÉ se está alertando, tomado de la causa (el sujeto/objeto/operación). Ej: "...pagos anticipados de cuotas del título de capitalización...".
       3) SEÑAL DE ALERTA (OBLIGATORIA, NUNCA OMITIR): la condición que dispara la alerta.
          ⚠️ REGLA DE PRESERVACIÓN (MUY IMPORTANTE): si el CONTROL ORIGINAL YA menciona umbrales, montos, plazos, porcentajes o patrones concretos (ej: "valor superior a $50.000.000", "periodo menor a tres meses", "desde la emisión de la póliza", "pagos de siniestros", jurisdicciones, fraccionamiento), DEBES CONSERVARLOS TAL CUAL en la sugerencia. NO los reemplaces por la fórmula genérica ni los resumas ni los elimines. La sugerencia debe mantener TODO el detalle del original y solo mejorar la estructura sintáctica. NO empobrezcas el control.
          SOLO si el control original NO trae ninguna señal concreta, usa la fórmula genérica EXACTA "que presentan cumplimiento en las señales de alerta parametrizadas con las lógicas y umbrales definidos en su parametrización". PROHIBIDO inventar valores numéricos ("50%", "30 días", "USD 5.000") que no estuvieran en el original.
     Ejemplo con señal genérica: "Alertamiento generado por anomalías detectadas en el comportamiento transaccional de clientes que presentan cumplimiento en las señales de alerta parametrizadas con las lógicas y umbrales definidos en su parametrización."
     Los TRES sub-elementos deben estar presentes SIEMPRE. NUNCA entregues un alertamiento sin señal de alerta. EVITA CUALQUIER REDUNDANCIA con la raíz "alerta": está PROHIBIDO "Alertamiento de alertas", "Alertamiento de verificación de alertas", "Alertamiento de generación de alertas", "Alertamiento de detección de alertas". Después de "Alertamiento" NO vuelvas a mencionar "alerta(s)"; usa directamente el núcleo verbal + el hecho monitoreado (ej: "Alertamiento de verificación de la identidad...", "Alertamiento generado por anomalías en...").
   • ANÁLISIS (subordinada): conector "Mediante" + "se" + verbo de evaluación conjugado (se analiza, se evalúa, se verifica, se examina, se contrasta). Ej: "Mediante el cual se analiza el patrón de pago y se verifica el origen de los recursos."
   • REPORTE (subordinada): conector "como resultado" + "se" + verbo de registro conjugado (se reporta, se documenta, se registra, se genera, se eleva). Ej: "Como resultado, se reporta internamente al oficial de cumplimiento y, de corresponder, se documenta el ROS ante la UIAF."
   • TOMA DE DECISIONES (subordinada): conector de finalidad "para" + "se" + verbo de resolución conjugado (se restringe, se cancela, se suspende, se bloquea, se aplica, se decide). Ej: "Para lo cual se restringe el canal de pago y se aplica seguimiento reforzado al título."
   • EJEMPLO COMPLETO (con los 3 "se"): "Alertamiento de pagos anticipados de cuotas en efectivo superiores al tope. Mediante el cual se analiza el patrón de pago y se verifica el origen de los recursos. Como resultado, se reporta al oficial de cumplimiento y se documenta el ROS ante la UIAF. Para lo cual se restringe el canal de pago y se aplica seguimiento reforzado al título."
   • REGLA DURA: si el control sugerido NO contiene los tres "se" (análisis + reporte + decisión), reescríbelo hasta cumplirlo.
   • SIN REDUNDANCIAS (CRÍTICO): jamás repitas la raíz "alert" tras "Alertamiento". PROHIBIDO absolutamente: "Alertamiento de alertas", "Alertamiento de alertas automáticas", "Alertamiento de alertas parametrizadas", "Alertamiento de las alertas generadas", "Alertamiento de la alerta", "Alertamiento de alertamiento". Después de "Alertamiento" va DIRECTAMENTE el hecho/patrón o un conector: "Alertamiento de pagos anticipados...", "Alertamiento de transferencias internacionales superiores a...", "Alertamiento generado cuando se detecta...", "Alertamiento en el aplicativo de monitoreo cuando...". NUNCA menciones "alertas" como objeto del alertamiento. Tampoco repitas "análisis del análisis", "reporte del reporte".
3. SIN ABREVIATURAS NI SIGLAS EN LA REDACCIÓN SUGERIDA: en los campos de redaccion_sugerida (riesgo, causa, control) escribe TODO en palabras completas y desarrolladas. NO uses siglas: desarróllalas siempre. En particular:
   • "ROS" → escribe "reporte de operación sospechosa".
   • "UIAF" → escribe "Unidad de Información y Análisis Financiero".
   • "PEP" → escribe "persona expuesta políticamente".
   • "LA/FT/FPADM" → escribe "lavado de activos, financiación del terrorismo y financiación de la proliferación de armas de destrucción masiva" (o el subconjunto que aplique, p.ej. "lavado de activos").
   • "OFAC" → "Oficina de Control de Activos Extranjeros del Departamento del Tesoro de los Estados Unidos".
   • "ONU" → "Organización de las Naciones Unidas".
   • "DIAN" → "Dirección de Impuestos y Aduanas Nacionales".
   • "SFC" → "Superintendencia Financiera de Colombia".
   También evita abreviaturas coloquiales: "tx" (transacción), "op" (operación), "doc" (documento), "info" (información), "aprox" (aproximadamente), "vlr" (valor), "c/u", "xq"/"pq" (porque). Ninguna palabra recortada. Puedes, si lo deseas, poner la sigla entre paréntesis DESPUÉS del término completo, pero nunca la sigla sola.

ANALIZA Y CALIFICA EL SIGUIENTE TRÍO:

RIESGO:
{riesgo}

CAUSA:
{causa}

CONTROL:
{control}`;

/**
 * Endpoint de validación
 */
app.post('/api/validate', async (req, res) => {
    const { riesgo, causa, control } = req.body;

    // Validar que todos los campos estén presentes
    if (!riesgo || !causa || !control) {
        return res.status(400).json({
            error: 'Riesgo, Causa y Control son requeridos'
        });
    }

    try {
        console.log('📤 Analizando trío con IA (vía analizarTrio)...');
        // Usa la misma función que recalcular/guardar -> aplica TODAS las salvaguardas
        // (verbo-objeto, siglas, redundancias, "que", naturaleza, alertamiento, etc.)
        const analysis = await analizarTrio(riesgo, causa, control);

        const recommendations = [
            analysis.strengths?.map(s => `✓ ${s}`).join('\n'),
            analysis.weaknesses?.map(w => `✗ ${w}`).join('\n'),
            analysis.recommendations?.map(r => `→ ${r}`).join('\n')
        ].filter(Boolean).join('\n\n');

        console.log('✓ Validación completada exitosamente');
        res.json({ analysis, recommendations, message: 'Validación completada' });

    } catch (error) {
        const status = error.statusCode || error.response?.status || 500;
        console.error('❌ Error:', error.response?.data || error.message);
        if (status === 401) return res.status(401).json({ error: 'API key de OpenAI inválida o expirada' });
        if (status === 429) return res.status(429).json({ error: 'Límite de rate limit alcanzado. Intenta de nuevo en unos momentos.' });
        res.status(status).json({ error: `Error en validación: ${error.message}` });
    }
});

/**
 * Endpoint: Recalcular calificación con IA y GUARDAR en data.json.
 * Body: { id, aplicarRedaccion?: bool }
 * Si aplicarRedaccion=true, reemplaza el texto del registro por la redacción sugerida.
 */
app.post('/api/recalcular', async (req, res) => {
    const { id, aplicarRedaccion } = req.body;
    if (id === undefined || id === null) {
        return res.status(400).json({ error: 'Falta el id del registro' });
    }

    try {
        const data = JSON.parse(fs.readFileSync(DATA_PATH, 'utf8'));
        const reg = data.find(r => String(r.id) === String(id));
        if (!reg) return res.status(404).json({ error: `Registro ${id} no encontrado` });

        const analysis = await analizarTrio(reg.riesgo, reg.causa, reg.control);

        // Guardar calificaciones recalculadas
        reg.calif_riesgo = Number(analysis.riesgo_assessment?.score) || reg.calif_riesgo;
        reg.calif_causa = Number(analysis.causa_assessment?.score) || reg.calif_causa;
        reg.calif_control = Number(analysis.control_assessment?.score) || reg.calif_control;

        // Opcional: aplicar redacción mejorada
        if (aplicarRedaccion && analysis.redaccion_sugerida) {
            if (analysis.redaccion_sugerida.riesgo) reg.riesgo = analysis.redaccion_sugerida.riesgo;
            if (analysis.redaccion_sugerida.causa) reg.causa = analysis.redaccion_sugerida.causa;
            if (analysis.redaccion_sugerida.control) reg.control = analysis.redaccion_sugerida.control;
        }

        fs.writeFileSync(DATA_PATH, JSON.stringify(data, null, 2), 'utf8');
        console.log(`✓ Registro ${id} recalculado: R=${reg.calif_riesgo} C=${reg.calif_causa} Ctrl=${reg.calif_control}${aplicarRedaccion ? ' (redacción aplicada)' : ''}`);
        syncSupabase(reg, analysis);

        res.json({ ok: true, registro: reg, analysis });
    } catch (error) {
        const status = error.statusCode || error.response?.status || 500;
        console.error('❌ Error recalcular:', error.response?.data || error.message);
        res.status(status).json({ error: `Error al recalcular: ${error.message}` });
    }
});

/**
 * Endpoint: Guardar la redacción editada por el usuario y RECALIFICAR ese texto.
 * Guarda el texto que envía el cliente y la calificación que corresponde a ESE texto
 * (recalculada con IA), para que nota y redacción sean siempre coherentes.
 * Body: { id, riesgo, causa, control, recalcular?: bool (default true),
 *         calif_riesgo?, calif_causa?, calif_control? (usados si recalcular=false) }
 */
app.post('/api/guardar', async (req, res) => {
    const { id, calif_riesgo, calif_causa, calif_control, riesgo, causa, control, recalcular = true } = req.body;
    if (id === undefined || id === null) {
        return res.status(400).json({ error: 'Falta el id del registro' });
    }
    try {
        const data = JSON.parse(fs.readFileSync(DATA_PATH, 'utf8'));
        const reg = data.find(r => String(r.id) === String(id));
        if (!reg) return res.status(404).json({ error: `Registro ${id} no encontrado` });

        // Texto final a guardar (editado por el usuario, o el actual si no viene)
        const finalRiesgo = (typeof riesgo === 'string' && riesgo.trim()) ? riesgo.trim() : reg.riesgo;
        const finalCausa = (typeof causa === 'string' && causa.trim()) ? causa.trim() : reg.causa;
        const finalControl = (typeof control === 'string' && control.trim()) ? control.trim() : reg.control;

        reg.riesgo = finalRiesgo;
        reg.causa = finalCausa;
        reg.control = finalControl;

        const toScore = v => {
            const n = Number(v);
            return (!isNaN(n) && n >= 1 && n <= 3) ? n : null;
        };

        let analysis = null;
        if (recalcular) {
            // Recalificar el TEXTO FINAL para que la nota corresponda a lo que se guarda
            analysis = await analizarTrio(finalRiesgo, finalCausa, finalControl);
            reg.calif_riesgo = Number(analysis.riesgo_assessment?.score) || reg.calif_riesgo;
            reg.calif_causa = Number(analysis.causa_assessment?.score) || reg.calif_causa;
            reg.calif_control = Number(analysis.control_assessment?.score) || reg.calif_control;
        } else {
            // Usar las calificaciones que manda el cliente (sin llamar IA)
            if (toScore(calif_riesgo) !== null) reg.calif_riesgo = toScore(calif_riesgo);
            if (toScore(calif_causa) !== null) reg.calif_causa = toScore(calif_causa);
            if (toScore(calif_control) !== null) reg.calif_control = toScore(calif_control);
        }

        fs.writeFileSync(DATA_PATH, JSON.stringify(data, null, 2), 'utf8');
        console.log(`💾 Registro ${id} guardado${recalcular ? ' (recalificado)' : ''}: R=${reg.calif_riesgo} C=${reg.calif_causa} Ctrl=${reg.calif_control}`);
        syncSupabase(reg, analysis);

        res.json({ ok: true, registro: reg, analysis });
    } catch (error) {
        const status = error.statusCode || error.response?.status || 500;
        console.error('❌ Error guardar:', error.response?.data || error.message);
        res.status(status).json({ error: `Error al guardar: ${error.message}` });
    }
});

/**
 * Endpoint: Marcar un registro como completado (revisado).
 * Body: { id, completado?: bool }  (completado por defecto true)
 */
app.post('/api/completar', (req, res) => {
    const { id, completado = true } = req.body;
    if (id === undefined || id === null) {
        return res.status(400).json({ error: 'Falta el id del registro' });
    }
    try {
        const data = JSON.parse(fs.readFileSync(DATA_PATH, 'utf8'));
        const reg = data.find(r => String(r.id) === String(id));
        if (!reg) return res.status(404).json({ error: `Registro ${id} no encontrado` });

        reg.completado = !!completado;
        fs.writeFileSync(DATA_PATH, JSON.stringify(data, null, 2), 'utf8');
        console.log(`✓ Registro ${id} marcado como ${reg.completado ? 'COMPLETADO' : 'pendiente'}`);
        syncSupabase(reg, null); // solo actualiza estado/calificación

        const pendientes = data.filter(r => !r.completado).length;
        res.json({ ok: true, registro: reg, pendientes, total: data.length });
    } catch (error) {
        console.error('❌ Error completar:', error.message);
        res.status(500).json({ error: `Error al completar: ${error.message}` });
    }
});

/**
 * Endpoint: Exportar data.json a un Excel con las 3 columnas enriquecidas.
 * Descarga directa de bd-matriz-enriquecida.xlsx
 */
app.get('/api/exportar', (req, res) => {
    try {
        const data = JSON.parse(fs.readFileSync(DATA_PATH, 'utf8'));

        const filas = data.map(r => ({
            'ID': r.id,
            'RIESGO_DESCRIPCION': r.riesgo || '',
            'CAUSA_DESCRIPCION': r.causa || '',
            'CONTROL_DESCRIPCION': r.control || '',
            'CALIF_RIESGO': r.calif_riesgo ?? '',
            'CALIF_CAUSA': r.calif_causa ?? '',
            'CALIF_CONTROL': r.calif_control ?? '',
            'CALIF_TOTAL': (Number(r.calif_riesgo) || 0) + (Number(r.calif_causa) || 0) + (Number(r.calif_control) || 0),
            'COHERENCIA_PCT': Math.round(((Number(r.calif_riesgo) || 0) + (Number(r.calif_causa) || 0) + (Number(r.calif_control) || 0)) / 9 * 100) + '%',
            'ESTADO': r.completado ? 'Revisado' : 'Pendiente'
        }));

        const ws = XLSX.utils.json_to_sheet(filas);
        // Anchos de columna razonables
        ws['!cols'] = [
            { wch: 6 }, { wch: 60 }, { wch: 60 }, { wch: 60 },
            { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 14 }, { wch: 12 }
        ];
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Matriz Enriquecida');

        const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
        const fecha = new Date().toISOString().slice(0, 10);
        res.setHeader('Content-Disposition', `attachment; filename="bd-matriz-enriquecida-${fecha}.xlsx"`);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(buffer);
        console.log(`📥 Excel exportado: ${filas.length} registros`);
    } catch (error) {
        console.error('❌ Error exportar:', error.message);
        res.status(500).json({ error: `Error al exportar: ${error.message}` });
    }
});

/**
 * Ruta raíz - Servir index.html
 */
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

/**
 * Iniciar servidor
 */
app.listen(PORT, () => {
    const model = process.env.OPENAI_MODEL || 'gpt-4o-mini';
    console.log(`
╔════════════════════════════════════════════════════╗
║  🔐 Validador de Coherencia - Matriz LA/FT/FPADM  ║
╚════════════════════════════════════════════════════╝

📍 Servidor corriendo en: http://localhost:${PORT}
🔑 API Key: ${process.env['API-KEY'] ? '✓ Configurada' : '✗ NO CONFIGURADA'}
🤖 Modelo: ${model}

Abre http://localhost:${PORT} en tu navegador.
    `);
});

/**
 * Manejo de errores global
 */
process.on('unhandledRejection', (reason, promise) => {
    console.error('❌ Promise rechazada sin manejo:', reason);
});
