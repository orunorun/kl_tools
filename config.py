import logging
import configparser

# Configuration management setup
config = configparser.ConfigParser()
config.read('config.ini')

# Logging setup
logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)