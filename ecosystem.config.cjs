const fs = require('fs');
const path = require('path');

/**
 * Safely parse a .env file with zero external dependencies.
 * Also falls back to `dotenv` if available.
 */
function loadEnv(filePath) {
  const envMap = {};

  // Try using dotenv first if available
  try {
    const dotenv = require('dotenv');
    const result = dotenv.config({ path: filePath });
    if (result.parsed) {
      return result.parsed;
    }
  } catch {
    // dotenv not installed, continue with fallback parser
  }

  // Fallback: parse .env manually
  if (fs.existsSync(filePath)) {
    const content = fs.readFileSync(filePath, 'utf-8');
    const lines = content.split(/\r?\n/);

    for (const rawLine of lines) {
      const line = rawLine.trim();
      if (!line || line.startsWith('#')) continue;

      const match = line.match(/^\s*(?:export\s+)?([\w.-]+)(?:\s*=\s*(.*))?$/);
      if (!match) continue;

      const key = match[1];
      let val = match[2] !== undefined ? match[2].trim() : '';

      // Remove single or double wrapping quotes
      if (
        (val.startsWith('"') && val.endsWith('"')) ||
        (val.startsWith("'") && val.endsWith("'"))
      ) {
        val = val.slice(1, -1);
      }

      // Unescape newlines
      val = val.replace(/\\n/g, '\n').replace(/\\r/g, '\r');
      envMap[key] = val;
    }
  }

  return envMap;
}

const envFilePath = path.resolve(__dirname, '.env');
const parsedEnv = loadEnv(envFilePath);

module.exports = {
  apps: [
    {
      name: 'smart-clm',
      script: 'dist/server.cjs',
      cwd: __dirname,

      // Single instance (fork) is recommended because in-app schedulers
      // (reminders, backups) run via internal background timers (setInterval)
      instances: 1,
      exec_mode: 'fork',
      autorestart: true,
      watch: false,
      max_memory_restart: '1G',

      // Native PM2 env file support
      env_file: '.env',

      // Logging configuration
      log_date_format: 'YYYY-MM-DD HH:mm:ss Z',
      error_file: './logs/pm2-error.log',
      out_file: './logs/pm2-out.log',
      merge_logs: true,

      // Default environment: loaded .env with production defaults
      env: {
        NODE_ENV: 'production',
        PORT: 3000,
        ...parsedEnv,
      },

      // Overrides when running: pm2 start ecosystem.config.cjs --env production
      env_production: {
        NODE_ENV: 'production',
        PORT: 3000,
        ...parsedEnv,
      },

      // Overrides when running: pm2 start ecosystem.config.cjs --env development
      env_development: {
        NODE_ENV: 'development',
        PORT: 3000,
        ...parsedEnv,
      },
    },
  ],
};
