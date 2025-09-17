const { spawn } = require('child_process');
const assert = require('assert');

function startServer() {
  return new Promise((resolve, reject) => {
    const server = spawn('python3', [
      '-m',
      'uvicorn',
      'server:app',
      '--host',
      '127.0.0.1',
      '--port',
      '8001'
    ]);

    const onData = (data) => {
      const text = data.toString();
      if (text.includes('Application startup complete')) {
        cleanup();
        resolve(server);
      }
    };

    const onError = (err) => {
      cleanup();
      reject(err);
    };

    const onExit = (code) => {
      cleanup();
      reject(new Error(`Server exited with code ${code}`));
    };

    const cleanup = () => {
      clearTimeout(timeout);
      server.stdout.off('data', onData);
      server.stderr.off('data', onData);
      server.off('error', onError);
      server.off('exit', onExit);
    };

    const timeout = setTimeout(() => {
      cleanup();
      server.kill('SIGTERM');
      reject(new Error('Timed out waiting for server startup'));
    }, 10000);

    server.stdout.on('data', onData);
    server.stderr.on('data', onData);
    server.on('error', onError);
    server.on('exit', onExit);
  });
}

async function stopServer(server) {
  if (!server) return;
  return new Promise((resolve) => {
    server.once('close', resolve);
    server.kill('SIGTERM');
  });
}

async function runTest() {
  const server = await startServer();

  try {
    const formData = new FormData();
    const blob = new Blob(['not-an-excel-file'], { type: 'application/octet-stream' });
    formData.append('file', blob, 'invalid.xlsx');

    const response = await fetch('http://127.0.0.1:8001/process', {
      method: 'POST',
      body: formData
    });

    assert.strictEqual(response.ok, false, 'Expected non-OK response');
    assert.strictEqual(response.status, 400, 'Expected HTTP 400 status');

    const errorBody = await response.json();
    assert.ok(errorBody.error, 'Expected error message in response body');

    console.log('Bad upload fetch test passed.');
  } finally {
    await stopServer(server);
  }
}

runTest().catch((error) => {
  console.error('Bad upload fetch test failed:', error);
  process.exitCode = 1;
});
