
# _logger.py

import logging

import _database as _db

__version__ = '0.0.1'

class SQLAlchemyHandler(logging.Handler):
    # A very basic logger that commits a LogRecord to the SQL Db
    def emit(self, record):
        session = _db.ConnectToDatabase()
        log = _db.Log(
            logger=record.__dict__['filename'],
            level=record.__dict__['levelname'],
            msg=record.__dict__['msg'],)
        session.add(log)
        session.commit()
        session.close()


import logging
log = logging.getLogger('root')

log.debug('Logging Module Initialized.')
