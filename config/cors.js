/** Mirrors Backend_Tally/main.py's CORSMiddleware config exactly:
 * a single allowed origin, allow_methods=["*"] (FastAPI expands this to the
 * full standard method list), allow_headers=["*"] (FastAPI reflects back
 * whatever the request's Access-Control-Request-Headers asked for, rather
 * than sending a literal "*" — the `cors` package's default behavior when
 * allowedHeaders is omitted already does the same reflection). */
module.exports = {
  origin: 'http://localhost:5173',
  methods: ['GET', 'HEAD', 'PUT', 'PATCH', 'POST', 'DELETE', 'OPTIONS'],
};
