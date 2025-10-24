# Gunicorn configuration file
import multiprocessing
import os

# Server socket
bind = f"0.0.0.0:{os.environ.get('PORT', 5000)}"
backlog = 2048

# Worker processes
workers = 1  # Use only 1 worker to conserve memory on free tier
worker_class = 'sync'
worker_connections = 1000
timeout = 120  # Increase timeout for large file processing
keepalive = 30

# Memory management
max_requests = 100  # Restart worker after 100 requests to prevent memory leaks
max_requests_jitter = 10
preload_app = True

# Logging
loglevel = 'info'
accesslog = '-'
errorlog = '-'

# Process naming
proc_name = 'excel-formatter'

# Worker timeout
worker_tmp_dir = '/dev/shm'  # Use shared memory for better performance