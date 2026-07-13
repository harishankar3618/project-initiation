const path = require('path');
const os = require('os');
const express = require('express');
const cors = require('cors');
const helmet = require('helmet');
const compression = require('compression');
const morgan = require('morgan');
const rateLimit = require('express-rate-limit');
require('dotenv').config();

const { buildBootstrap, initiateProject } = require('./src/sharepointService');
const { graphSearchUsers } = require('./src/graphClient');
const metadata = require('./src/metadata');

const app = express();
const port = Number(process.env.PORT || 45821);

function getLocalIPv4() {
  const interfaces = os.networkInterfaces();
  for (const name of Object.keys(interfaces)) {
    for (const iface of interfaces[name]) {
      if (iface.family === 'IPv4' && !iface.internal) {
        return iface.address;
      }
    }
  }
  return '127.0.0.1';
}

function validateEnv() {
  const required = ['TENANT_ID', 'CLIENT_ID', 'CLIENT_SECRET'];
  const missing = required.filter(function (key) { return !process.env[key]; });
  if (missing.length) {
    console.error('FATAL: Missing required environment variables: ' + missing.join(', '));
    console.error('Please check your .env file.');
    process.exit(1);
  }
}

validateEnv();

// Trust proxy (required when behind Nginx/load balancer)
app.set('trust proxy', 1);

// Request ID correlation
app.use(function (req, res, next) {
  req.id = req.headers['x-request-id'] || require('crypto').randomUUID();
  res.setHeader('X-Request-Id', req.id);
  next();
});

// Security headers
app.use(helmet({
  contentSecurityPolicy: {
    directives: {
      defaultSrc: ["'self'"],
      scriptSrc: ["'self'", "'unsafe-inline'"],
      scriptSrcAttr: ["'unsafe-inline'"],
      styleSrc: ["'self'", "'unsafe-inline'"],
      imgSrc: ["'self'", "data:", "https:"],
      connectSrc: ["'self'"],
      fontSrc: ["'self'", "data:"],
      objectSrc: ["'none'"],
      frameAncestors: ["'self'"]
    }
  },
  crossOriginEmbedderPolicy: false
}));

// Compression
app.use(compression({ threshold: 1024 }));

// CORS — restrict to allowed origins
const allowedOrigins = (process.env.ALLOWED_ORIGINS || '')
  .split(',')
  .map(function (o) { return o.trim(); })
  .filter(Boolean);

app.use(cors({
  origin: function (origin, callback) {
    if (!origin) return callback(null, true);
    if (allowedOrigins.indexOf(origin) !== -1) return callback(null, true);
    return callback(new Error('Not allowed by CORS: ' + origin), false);
  },
  credentials: true,
  maxAge: 86400
}));

// Body parsing
app.use(express.json({ limit: '1mb' }));

// Request logging
if (process.env.NODE_ENV !== 'test') {
  app.use(morgan('combined', {
    skip: function (req) { return req.url === '/health'; }
  }));
}

// Rate limiting
const generalLimiter = rateLimit({
  windowMs: Number(process.env.RATE_LIMIT_WINDOW_MS) || 60000,
  max: Number(process.env.RATE_LIMIT_MAX) || 100,
  standardHeaders: true,
  legacyHeaders: false,
  message: { error: 'Too many requests. Please try again later.' }
});

const initiateLimiter = rateLimit({
  windowMs: Number(process.env.RATE_LIMIT_INITIATE_WINDOW_MS) || 60000,
  max: Number(process.env.RATE_LIMIT_INITIATE_MAX) || 10,
  standardHeaders: true,
  legacyHeaders: false,
  message: { error: 'Too many initiation requests. Please try again later.' }
});

app.use('/api/', generalLimiter);
app.use('/api/initiate', initiateLimiter);

// Health endpoint
app.get('/health', function (_req, res) {
  res.json({
    status: 'OK',
    uptime: process.uptime(),
    version: '1.0.0'
  });
});

// Metadata-driven configuration
app.get('/api/metadata', function (_req, res) {
  res.setHeader('Cache-Control', 'no-store');
  res.json(metadata);
});

// Graph-backed people search
app.get('/api/users', async function (req, res) {
  try {
    const q = typeof req.query.q === 'string' ? req.query.q : '';
    const department = typeof req.query.department === 'string' ? req.query.department : '';
    const users = await graphSearchUsers(q, department);
    res.setHeader('Cache-Control', 'no-store');
    res.json(users);
  } catch (error) {
    res.status(500).json({
      error: 'Failed to search users',
      message: 'An unexpected error occurred.'
    });
  }
});

app.get('/api/bootstrap', async function (_req, res) {
  try {
    const data = await buildBootstrap();
    res.setHeader('Cache-Control', 'no-store');
    res.json(data);
  } catch (error) {
    res.status(500).json({
      error: 'Failed to load SharePoint bootstrap data',
      message: 'An unexpected error occurred.'
    });
  }
});

// Create Main Tracker items per department
app.post('/api/initiate', async function (req, res) {
  const payload = req.body && req.body.payload ? req.body.payload : req.body;
  try {
    if (!payload || !Array.isArray(payload.departments) || !payload.departments.length) {
      return res.status(400).json({ error: 'Invalid payload: departments array required.' });
    }
    const result = await initiateProject(payload);
    res.setHeader('Cache-Control', 'no-store');
    res.json({ ok: result.errors.length === 0, created: result.created, errors: result.errors, warnings: result.warnings, sent: result.sent, payload: payload });
  } catch (error) {
    res.status(500).json({
      error: 'Failed to save to Main Tracker',
      message: 'An unexpected error occurred.'
    });
  }
});

// Static files
app.use(express.static(__dirname));

app.get('/', function (_req, res) {
  res.sendFile(path.join(__dirname, 'ProjectInitiationForm.html'));
});

// Global error handler
app.use(function (err, req, res, _next) {
  const status = err.status || 500;
  res.status(status).json({
    error: status === 500 ? 'Internal server error' : err.message,
    requestId: req.id
  });
});

// 404 handler
app.use(function (req, res) {
  res.status(404).json({ error: 'Not found' });
});

// Start server
const localIPv4 = getLocalIPv4();
const server = app.listen(port, '0.0.0.0', function () {
  console.log('\n========================================');
  console.log('  PMO Project Initiation Portal');
  console.log('========================================');
  console.log('  Server Started Successfully\n');
  console.log('  Local:');
  console.log('  http://localhost:' + port);
  console.log('');
  console.log('  Network:');
  console.log('  http://' + localIPv4 + ':' + port);
  console.log('\n========================================\n');
});

// Graceful shutdown
function shutdown(signal) {
  console.log('Received ' + signal + '. Shutting down gracefully...');
  server.close(function () {
    console.log('HTTP server closed.');
    process.exit(0);
  });
  setTimeout(function () {
    console.error('Forced shutdown due to timeout.');
    process.exit(1);
  }, 10000);
}

process.on('SIGTERM', function () { shutdown('SIGTERM'); });
process.on('SIGINT', function () { shutdown('SIGINT'); });

process.on('uncaughtException', function (err) {
  console.error('Uncaught exception:', err);
  process.exit(1);
});

process.on('unhandledRejection', function (reason) {
  console.error('Unhandled promise rejection:', reason);
  process.exit(1);
});
