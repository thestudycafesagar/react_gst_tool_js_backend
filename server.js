/**
 * Local-only Bank Statement / GSTR-2B / Tally Entry API.
 * 1:1 port of Backend_Tally/main.py's FastAPI app -> Express.
 * Runs fully offline except for the Tally HTTP bridge (localhost:9000)
 * and the GST portal scraper. Listens on the same port (8000) and the
 * same single-origin CORS policy as the Python service, so the existing
 * React frontend needs zero changes.
 */
const express = require('express');
const cors = require('cors');
const corsConfig = require('./config/cors');
const routes = require('./routes');

const app = express();

app.use(cors(corsConfig));
app.use(express.json({ limit: '50mb' }));

app.use('/', routes);

const PORT = process.env.PORT || 8000;
app.listen(PORT, () => {
  console.log(`Tally backend (Node) listening on port ${PORT}`);
});

module.exports = app;
