import logging
import logger as lr

def main():
    #logger.error('test test test ERROR')
    lr.logger.debug('debug message')
    lr.logger.info('info message')
    lr.logger.warn('warn message')
    lr.logger.error('error message')
    lr.logger.critical('critical message')

if __name__ == '__main__':
    main()
