import subprocess
import logging

def run_powershell(cmd):
    logger = logging.getLogger(__name__)
    try:
        result = subprocess.run(["powershell", "-Command", cmd],
        capture_output=True, text=True, timeout=10)
        return result.returncode, result.stdout.strip(), result.stderr.strip()
    except Exception as e:
        return 1, "", str(e)