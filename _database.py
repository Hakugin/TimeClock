
# _database.py

from sqlalchemy import Column
import sys
import os

from sqlalchemy import Table, Column
from sqlalchemy import create_engine, exists
from sqlalchemy.types import DateTime, Integer
from sqlalchemy.types import String, Time
from sqlalchemy.sql import func
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker

frozen = getattr(sys, 'frozen', '')
if not frozen:
    app_root = os.path.dirname(__file__)
elif frozen in ('dll', 'console_exe', 'windows_exe'):
    app_root = os.path.dirname(sys.executable)
elif frozen in ('macosx_app',):
    app_root = os.environ['RESOURCEPATH']
db_loc = os.path.join(app_root, 'ClockPuncher3.db')

engine = create_engine("sqlite:///%s" % db_loc, echo=False)
base = declarative_base(engine)
metadata = base.metadata

import logging
log = logging.getLogger('root')

class Log(base):
    __tablename__ = 'logs'
    id = Column(Integer, primary_key=True) # auto incrementing
    logger = Column(String) # the name of the logger. (e.g. myapp.views)
    level = Column(String) # info, debug, or error?
    msg = Column(String) # any custom log you may have included
    created_at = Column(DateTime, default=func.now()) # the current timestamp

    def __init__(self, logger=None, level=None,
                    msg=None, created_at=None):
        self.logger = logger
        self.level = level
        self.msg = msg
        self.created = created_at

    def __unicode__(self):
        return self.__repr__()

    def __repr__(self):
        return "%s\t%s\t%s\t%s UTC" % (
            self.logger, self.level, self.msg, self.created_at.strftime('%m/%d/%Y-%H:%M:%S'))


class Settings(base):
    __tablename__ = 'settings'
    name = Column(String, primary_key=True)
    value = Column(String)

    def __repr__(self):
        return '<%s: %s>' % (self.name, self.value)


class PunchHistory(base):
    __tablename__ = 'timestamp_history'
    the_date = Column(String(8), primary_key=True)
    day_start = Column(String)
    lunch_start = Column(String)
    lunch_end = Column(String)
    day_end = Column(String)
    
    def __repr__(self):
        return "Timestamp History: \n\tDate:\tDay Start:\tLunch Start:\tLunch End:\tDay End:\n%s\ts%s\t%s\t%s\n" % (
                   self.the_date, self.day_start, self.lunch_start, self.lunch_end, self.day_end)


def OnUpdateTimeStamp(theDate, dayS=None,
                lunchS=None, lunchE=None, dayE=None):
    ''''''
    log.debug('Opening session with database.')
    session = ConnectToDatabase()
    log.debug('Checking database for entries matching todays date.')
    ret = session.query(exists().where(
        PunchHistory.the_date==theDate)).scalar()
    if not ret:
        log.debug('No entries found, creating new entry.')
        the_data = PunchHistory(
                the_date    = theDate,
                day_start   = dayS,
                lunch_start = lunchS,
                lunch_end   = lunchE,
                day_end     = dayE
            )
    else:
        log.debug('Existing entry found, updating existing entry.')
        the_data = session.query(PunchHistory).filter(
            PunchHistory.the_date==theDate).first()
        if dayS != None:
            log.debug('Updating \'Start of Day\' entry with: %s' % dayS)
            the_data.day_start   = dayS
        if lunchS != None:
            log.debug('Updating \'Start of Lunch\' entry with: %s' % lunchS)
            the_data.lunch_start = lunchS
        if lunchE != None:
            log.debug('Updating \'End of Lunch\' entry with: %s' % lunchE)
            the_data.lunch_end   = lunchE
        if dayE != None:
            log.debug('Updating \'End of Day\' entry with: %s' % dayE)
            the_data.day_end     = dayE
    log.debug('Adding data to session object.')
    session.add(the_data)
    log.debug('Commiting data to database.')
    session.commit()
    log.debug('Closing session.')
    session.close()


def OnLoadTimeStamps(theDate):
    ''''''
    timeStamps = {}
    session = ConnectToDatabase()
    result = session.query(PunchHistory).filter(
        PunchHistory.the_date==theDate).first()
    if result:
        if result.day_start:
            timeStamps['dayS']   = result.day_start.encode('utf8')
        else:
            timeStamps['dayS']   = ''

        if result.lunch_start:
            timeStamps['lunchS'] = result.lunch_start.encode('utf8')
        else:
            timeStamps['lunchS']   = ''

        if result.lunch_end:
            timeStamps['lunchE'] = result.lunch_end.encode('utf8')
        else:
            timeStamps['lunchE']   = ''

        if result.day_end:
            timeStamps['dayE']   = result.day_end.encode('utf8')
        else:
            timeStamps['dayE']   = ''

    else:
        timeStamps['dayS']   = ''
        timeStamps['lunchS'] = ''
        timeStamps['lunchE'] = ''
        timeStamps['dayE']   = ''
    session.close()
    return timeStamps


def ConnectToDatabase(object=None):
    '''Connect to the sqlite db and return the session object.'''
    Session = sessionmaker(bind=engine)
    session = Session()
    return session


def SaveSettings(settings):
    ''''''
    session = ConnectToDatabase()
    for key, val in settings.items():
        ret = session.query(exists().where(
            Settings.name == key)).scalar()
        if not ret:
            the_data = Settings(
                name=key, value=val)
        else:
            the_data = session.query(Settings).filter(
                Settings.name==key).first()
            the_data.value=val
        session.add(the_data)
    session.commit()
    session.close()


def LoadSettings():
    ''''''
    session = ConnectToDatabase()
    query = session.query(Settings)
    results = query.all()
    prgSettings = {}
    for option in results:
        prgSettings[option.name.encode('utf8')] = option.value.encode('utf8')
    session.close()
    return prgSettings


def ReadLogData():
    ''''''
    session = ConnectToDatabase()
    query = session.query(Log)
    results = query.all()
    session.close()
    return results


metadata.create_all()
