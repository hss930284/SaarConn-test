import logging

# Configure logging
logging.basicConfig(
    filename='SAARCONN.log',
    level=logging.DEBUG,  # Changed to INFO to capture informational messages
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Example log message when main.py runs successfully
logging.info('main.py has been run successfully at this time and it successfully created arxml.')