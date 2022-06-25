import logging

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
logger.setLevel(100)

# create file handler which logs even debug messages
fh = logging.FileHandler('DAV_Quartal_2_ical.log')
fh.setLevel(logging.DEBUG)

# create console handler with a higher log level
ch = logging.StreamHandler()
ch.setLevel(logging.DEBUG)

# create formatter and add it to the handlers
formatter = logging.Formatter('%(asctime)s - %(module)s - %(threadName)s -  %(levelname)s - %(message)s')
fh.setFormatter(formatter)
ch.setFormatter(formatter)

# add the handlers to the logger
logger.addHandler(fh)
logger.addHandler(ch)
