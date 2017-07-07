import logging

#create logger
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

#log File path
f = 'log/logfile.txt'

# create file handler which logs even debug messages
fh = logging.FileHandler(f)
fh.setLevel(logging.DEBUG)

#create console handler with a higher loglevel
ch = logging.StreamHandler()
ch.setLevel(logging.ERROR)

#create formatter and add it to the handlers
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
ch.setFormatter(formatter)
fh.setFormatter(formatter)

# add the handlers to logger
logger.addHandler(ch)
logger.addHandler(fh)

#'application' code
logger.debug('debug message')
logger.info('info message')
logger.warn('warn message')
logger.error('error message')
logger.critical('critical message')
