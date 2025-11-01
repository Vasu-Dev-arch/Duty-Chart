import logging
import os

def setup_logging():
    log_dir = os.path.join(os.getenv("APPDATA"), "DutyChart")
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, "duty_chart_app.log")
    
    logging.basicConfig(
        filename=log_file,
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s"
    )
    logging.info("DutyChart app started.")