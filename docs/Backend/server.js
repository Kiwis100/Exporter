// backend/server.js
// Proxy seguro para la API de Logs de Instana.
// Mismo patrón de seguridad que instana-events-exporter: el apiToken viaja
// del navegador al backend por HTTPS en el body de la petición, se usa
// una sola vez para llamar a Instana, y nunca se persiste ni se loguea.

const express = require('express');
const cors = require('cors');

const app = express();
app.use(cors());
app.use(express.json({ limit: '2mb' }));

const PORT = process.env.PORT || 3002;

// POST /api/logs-query
// body esperado: { tenantUrl, apiToken, tagFilterExpression, timeConfig,
//                   retrievalSize, offset, orderDirection, requestedTags }
// Reenvía tal cual a POST {tenantUrl}/api/logging/logs/getLogs/v1
app.post('/api/logs-query', async (req, res) => {
  const { tenantUrl, apiToken, ...queryBody } = req.body || {};

  if (!tenantUrl || !apiToken) {
    return res.status(400).json({ error: 'tenantUrl y apiToken son requeridos' });
  }
  if (!/^https:\/\//i.test(tenantUrl)) {
    return res.status(400).json({ error: 'tenantUrl debe usar https://' });
  }

  const url = `${tenantUrl.replace(/\/$/, '')}/api/logging/logs/getLogs/v1`;

  try {
    const upstream = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        // Formato de auth confirmado por el usuario: "apiToken <token>"
        authorization: `apiToken ${apiToken}`,
      },
      body: JSON.stringify(queryBody),
    });

    const text = await upstream.text();

    if (!upstream.ok) {
      // No reenviamos el body crudo de error a ciegas por si contuviera
      // algo sensible; devolvemos el status y un mensaje genérico + texto.
      return res.status(upstream.status).json({
        error: `Instana respondió ${upstream.status}`,
        detail: text.slice(0, 500),
      });
    }

    let data;
    try {
      data = JSON.parse(text);
    } catch {
      // Respuesta no-JSON inesperada; la devolvemos cruda para poder
      // inspeccionar el formato real la primera vez que se use.
      return res.status(502).json({ error: 'Respuesta no-JSON de Instana', raw: text.slice(0, 1000) });
    }

    return res.json(data);
  } catch (err) {
    return res.status(502).json({ error: 'No se pudo contactar a Instana', detail: String(err) });
  }
});

app.listen(PORT, () => {
  console.log(`Instana Logs Exporter backend escuchando en http://localhost:${PORT}`);
});
