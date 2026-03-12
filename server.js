const http = require("http");
const handler = require("./api/index.js");

const PORT = process.env.PORT || 3000;

const server = http.createServer((req, res) => {
  handler(req, res);
});

server.listen(PORT, () => {
  console.log(`Excel API running at http://localhost:${PORT}`);
  console.log(`Health check: GET http://localhost:${PORT}/`);
  console.log(`Process file: POST http://localhost:${PORT}/`);
});