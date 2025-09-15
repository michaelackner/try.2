// Lightweight tunnel runner using localtunnel to expose localhost:8000
// Writes the assigned public URL into tunnel_url.txt

const fs = require('fs');
const path = require('path');
const localtunnel = require('localtunnel');

(async () => {
  try {
    const port = parseInt(process.env.PORT || '8000', 10);
    // Try a readable subdomain, fallback to random if taken
    const base = 'varo-rebilling';
    let tunnel;
    let attempt = 0;
    while (!tunnel) {
      const sub = attempt === 0 ? base : `${base}-${Math.random().toString(36).slice(2, 7)}`;
      try {
        tunnel = await localtunnel({ port, subdomain: sub });
      } catch (err) {
        attempt++;
        if (attempt > 5) throw err;
      }
    }

    const url = tunnel.url;
    const file = path.join(process.cwd(), 'tunnel_url.txt');
    fs.writeFileSync(file, url + '\n', 'utf8');
    console.log(`[tunnel] Public URL: ${url}`);

    // Keep the process alive until terminated
    tunnel.on('close', () => {
      console.log('[tunnel] Closed');
      process.exit(0);
    });

    const shutdown = () => {
      console.log('\n[tunnel] Shutting down...');
      try { tunnel.close(); } catch (_) {}
      process.exit(0);
    };
    process.on('SIGINT', shutdown);
    process.on('SIGTERM', shutdown);
  } catch (err) {
    console.error('[tunnel] Failed to start:', err && err.message ? err.message : err);
    process.exit(1);
  }
})();

