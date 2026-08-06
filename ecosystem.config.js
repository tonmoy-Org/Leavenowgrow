module.exports = {
  apps: [
    {
      name: "leavenowgrow-admin",
      script: "./server.js",
      exec_mode: "cluster",
      instances: 2,
      autorestart: true,
      watch: false,
      max_memory_restart: "1G",
      min_uptime: "10s",
      max_restarts: 50,
      restart_delay: 2000,
      listen_timeout: 8000,
      kill_timeout: 5000,
      env: {
        NODE_ENV: "production"
      }
    }
  ]
};
