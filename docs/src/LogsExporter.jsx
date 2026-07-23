// frontend/src/LogsExporter.jsx
//
// Componente hermano de tu Events Exporter, pero para la API de Logs de
// Instana (POST /api/logging/logs/getLogs/v1), que tiene un contrato
// distinto: paginación real (offset/retrievalSize) y SIN agregación en
// el servidor (confirmado: "API query currently does not support adding
// groups") — por eso agrupamos aquí en el cliente, igual que ya haces
// con la clasificación de entidades en el exporter de eventos.
//
// IMPORTANTE: la forma exacta en que Instana devuelve "log.custom" en la
// respuesta (¿objeto plano con todos los custom tags del log, o hay que
// pedir cada tag por separado?) no se pudo confirmar sin una llamada real.
// El parser de abajo intenta varias formas razonables (ver parseCustomTags)
// y deja un console.warn si no reconoce el formato — revisa la consola en
// tu primera prueba real y ajusta esa función si hace falta.

import { useState, useCallback } from 'react';
import {
  BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, CartesianGrid, Legend,
} from 'recharts';
import ExcelJS from 'exceljs';

const BACKEND_URL = import.meta.env.VITE_LOGS_BACKEND_URL || 'http://localhost:3002';
const MAX_PAGES_SAFETY = 50; // 50 x retrievalSize = techo duro anti-loop-infinito

const OPERATORS = ['EQUALS', 'NOT_EQUAL', 'CONTAINS', 'NOT_CONTAIN'];

function toEpochMs(dateStr) {
  return new Date(dateStr).getTime();
}

// Intenta extraer los custom tags de un item de log de varias formas
// posibles, ya que no se pudo confirmar el shape exacto sin acceso real
// a la API. Ajustar aquí si la consola muestra un formato distinto.
function parseCustomTags(item) {
  const raw = item?.tags?.['log.custom'] ?? item?.['log.custom'] ?? item?.custom;
  if (!raw) return {};
  if (typeof raw === 'object' && !Array.isArray(raw)) return raw;
  if (Array.isArray(raw)) {
    // formato tipo [{key, value}, ...]
    return Object.fromEntries(raw.map((kv) => [kv.key ?? kv.name, kv.value]));
  }
  console.warn('[LogsExporter] Formato de log.custom no reconocido, revisar item crudo:', item);
  return {};
}

export default function LogsExporter() {
  const [tenantUrl, setTenantUrl] = useState('');
  const [apiToken, setApiToken] = useState('');
  const [from, setFrom] = useState('');
  const [to, setTo] = useState('');
  const [filterKey, setFilterKey] = useState('event_id.id');
  const [filterOperator, setFilterOperator] = useState('EQUALS');
  const [filterValue, setFilterValue] = useState('');
  const [groupByTag, setGroupByTag] = useState('app.name');

  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [progress, setProgress] = useState('');
  const [rows, setRows] = useState([]);       // registros crudos normalizados
  const [grouped, setGrouped] = useState([]); // [{ key, count }] para el chart

  const runQuery = useCallback(async () => {
    setError('');
    setRows([]);
    setGrouped([]);

    if (!tenantUrl || !apiToken || !from || !to) {
      setError('Completa Tenant URL, Token y el rango de fechas.');
      return;
    }

    const toMs = toEpochMs(to);
    const fromMs = toEpochMs(from);
    if (!(toMs > fromMs)) {
      setError('El rango de fechas es inválido (to debe ser posterior a from).');
      return;
    }
    const windowSize = toMs - fromMs;

    const tagFilterExpression = filterValue
      ? {
          type: 'TAG_FILTER',
          name: 'log.custom',
          operator: filterOperator,
          entity: 'NOT_APPLICABLE',
          key: filterKey,
          value: filterValue,
        }
      : undefined;

    setLoading(true);
    const retrievalSize = 200;
    let offset = 0;
    let allItems = [];

    try {
      for (let page = 0; page < MAX_PAGES_SAFETY; page++) {
        setProgress(`Consultando página ${page + 1} (offset ${offset})...`);

        const body = {
          tenantUrl,
          apiToken,
          tagFilterExpression,
          timeConfig: {
            to: toMs,
            windowSize,
            focusedMoment: toMs,
            autoRefresh: false,
          },
          retrievalSize,
          offset,
          orderDirection: 'DESC',
          requestedTags: ['log.level', 'log.message', 'log.custom', 'log.timestamp'],
        };

        const resp = await fetch(`${BACKEND_URL}/api/logs-query`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(body),
        });

        if (!resp.ok) {
          const errBody = await resp.json().catch(() => ({}));
          throw new Error(errBody.error || `Error HTTP ${resp.status}`);
        }

        const data = await resp.json();
        // Ajustar esta línea según el shape real de la respuesta (p.ej.
        // podría ser data.items, data.logs, o directamente un array).
        const items = data.items ?? data.logs ?? (Array.isArray(data) ? data : []);

        if (!items.length) break;

        allItems = allItems.concat(items);
        offset += retrievalSize;

        if (items.length < retrievalSize) break; // última página
      }

      const normalized = allItems.map((item) => {
        const custom = parseCustomTags(item);
        return {
          timestamp: item.timestamp ?? item['log.timestamp'],
          level: item.level ?? item['log.level'],
          message: item.message ?? item['log.message'],
          ...custom,
        };
      });

      setRows(normalized);

      // Agrupamiento en cliente (la API no lo soporta)
      const counts = {};
      normalized.forEach((r) => {
        const key = r[groupByTag] ?? '(sin valor)';
        counts[key] = (counts[key] || 0) + 1;
      });
      setGrouped(Object.entries(counts).map(([key, count]) => ({ key, count })));

      setProgress(`Listo: ${normalized.length} registros obtenidos.`);
    } catch (err) {
      setError(String(err.message || err));
      setProgress('');
    } finally {
      setLoading(false);
    }
  }, [tenantUrl, apiToken, from, to, filterKey, filterOperator, filterValue, groupByTag]);

  const exportExcel = useCallback(async () => {
    if (!rows.length) return;

    const wb = new ExcelJS.Workbook();

    // Hoja 1: Resumen (agrupado)
    const summary = wb.addWorksheet('Resumen');
    summary.columns = [
      { header: groupByTag, key: 'key', width: 30 },
      { header: 'Cantidad', key: 'count', width: 15 },
    ];
    grouped.forEach((g) => summary.addRow(g));
    summary.getRow(1).font = { bold: true };
    summary.autoFilter = { from: 'A1', to: 'B1' };
    summary.views = [{ state: 'frozen', ySplit: 1 }];

    // Hoja 2: Detalle completo
    const allKeys = new Set(['timestamp', 'level', 'message']);
    rows.forEach((r) => Object.keys(r).forEach((k) => allKeys.add(k)));
    const cols = Array.from(allKeys);

    const detail = wb.addWorksheet('Detalle');
    detail.columns = cols.map((c) => ({ header: c, key: c, width: 22 }));
    rows.forEach((r) => detail.addRow(r));
    detail.getRow(1).font = { bold: true };
    detail.autoFilter = { from: 'A1', to: `${String.fromCharCode(64 + cols.length)}1` };
    detail.views = [{ state: 'frozen', ySplit: 1 }];

    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: 'application/octet-stream' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `instana_logs_export_${Date.now()}.xlsx`;
    a.click();
    URL.revokeObjectURL(url);
  }, [rows, grouped, groupByTag]);

  return (
    <div style={{ maxWidth: 960, margin: '0 auto', padding: 24, fontFamily: 'sans-serif' }}>
      <h2>Instana Logs Exporter</h2>

      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
        <label>
          Tenant URL
          <input value={tenantUrl} onChange={(e) => setTenantUrl(e.target.value)}
                 placeholder="https://empresa.instana.io" style={{ width: '100%' }} />
        </label>
        <label>
          API Token
          <input type="password" value={apiToken} onChange={(e) => setApiToken(e.target.value)}
                 style={{ width: '100%' }} />
        </label>
        <label>
          Desde
          <input type="datetime-local" value={from} onChange={(e) => setFrom(e.target.value)}
                 style={{ width: '100%' }} />
        </label>
        <label>
          Hasta
          <input type="datetime-local" value={to} onChange={(e) => setTo(e.target.value)}
                 style={{ width: '100%' }} />
        </label>
        <label>
          Filtro - clave (custom tag)
          <input value={filterKey} onChange={(e) => setFilterKey(e.target.value)}
                 placeholder="event_id.id / app.name / channel" style={{ width: '100%' }} />
        </label>
        <label>
          Operador
          <select value={filterOperator} onChange={(e) => setFilterOperator(e.target.value)} style={{ width: '100%' }}>
            {OPERATORS.map((op) => <option key={op} value={op}>{op}</option>)}
          </select>
        </label>
        <label>
          Valor del filtro (vacío = sin filtro)
          <input value={filterValue} onChange={(e) => setFilterValue(e.target.value)} style={{ width: '100%' }} />
        </label>
        <label>
          Agrupar por (para el gráfico/resumen)
          <input value={groupByTag} onChange={(e) => setGroupByTag(e.target.value)} style={{ width: '100%' }} />
        </label>
      </div>

      <div style={{ marginTop: 16 }}>
        <button onClick={runQuery} disabled={loading}>
          {loading ? 'Consultando...' : 'Consultar logs'}
        </button>
        <button onClick={exportExcel} disabled={!rows.length} style={{ marginLeft: 8 }}>
          Exportar Excel
        </button>
      </div>

      {progress && <p>{progress}</p>}
      {error && <p style={{ color: 'crimson' }}>{error}</p>}

      {grouped.length > 0 && (
        <div style={{ height: 300, marginTop: 24 }}>
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={grouped}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="key" />
              <YAxis allowDecimals={false} />
              <Tooltip />
              <Legend />
              <Bar dataKey="count" fill="#185FA5" />
            </BarChart>
          </ResponsiveContainer>
        </div>
      )}

      {rows.length > 0 && (
        <p style={{ marginTop: 12, color: '#555' }}>
          {rows.length} registros obtenidos (máximo {MAX_PAGES_SAFETY * 200} por consulta - ajustar MAX_PAGES_SAFETY si necesitas más).
        </p>
      )}
    </div>
  );
}
