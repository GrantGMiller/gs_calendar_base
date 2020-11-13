import datetime
import json
import time
from extronlib.system import File, Wait, ProgramLog

offsetSeconds = time.timezone if (time.localtime().tm_isdst == 0) else time.altzone
offsetHours = offsetSeconds / 60 / 60 * -1
MY_TIME_ZONE = offsetHours


class _CalendarItem:
    '''
    An object to represent an event on the calendar
    '''

    def __init__(self, startDT, endDT, data, parentCalendar):
        '''
        :param startDT:
        :param endDT:
        :param data: dict MUST contain keys 'ItemId', 'Subject', see __str__ method
        :param parentCalendar:
        '''
        if data is None:
            data = {}
        # print('_CalendarItem data=', data)
        self._data = data.copy()  # dict like {'ItemId': 'jasfsd', 'Subject': 'SuperMeeting', ...}

        self._startDT = startDT
        self._endDT = endDT

        if 'Duration' not in self._data:
            self._CalculateDuration()

        self._hasAttachments = data.get('HasAttachments', False)
        self._parentExchange = parentCalendar

    def AddData(self, key, value):
        self._data[key] = value

    def _CalculateDuration(self):
        # Returns float in seconds
        delta = self.Get('End') - self.Get('Start')
        duration = delta.total_seconds()
        self.AddData('Duration', duration)

    def Get(self, key):
        if key == 'Start':
            return self._startDT
        elif key == 'End':
            return self._endDT
        elif key == 'Duration':
            self._CalculateDuration()
            return self._data.get(key, None)
        else:
            return self._data.get(key, None)

    def get(self, key):
        return self.Get(key)

    def __contains__(self, dt):
        '''
        allows you to compare _CalendarItem object like you would compare datetime objects

        Example:
        if datetime.datetime.now() in calItem:
            print('the datetime is within the CalendarItem start/end')

        :param dt:
        :return:
        '''
        # Note: isinstance(datetime.datetime.now(), datetime.date.today()) == True
        # Because the point in time exist in that date
        # print('CalendarItem.__contains__(', dt)
        # print('self=', self)
        # print('self._startDT=', self._startDT)
        # print('self._endDT=', self._endDT)

        if isinstance(dt, datetime.datetime):
            # if dt.tzinfo is None and self._startDT.tzinfo is not None:
            #     # dt is naive, assume its in local system timezone
            #     # dt = dt.replace(tzinfo=datetime.timezone.utc)
            #     dt = dt.astimezone()
            #     print('dt converted to local tz', dt)

            # print('self.startDT <= dt is', self._startDT <= dt)
            # print('dt <= self._endDT is', dt <= self._endDT)

            if self._startDT <= dt <= self._endDT:
                # print('87return True')
                return True
            else:
                # print('90return False')
                return False

        elif isinstance(dt, datetime.date):
            if self._startDT.year == dt.year and \
                    self._startDT.month == dt.month and \
                    self._startDT.day == dt.day:
                return True

            elif self._endDT.year == dt.year and \
                    self._endDT.month == dt.month and \
                    self._endDT.day == dt.day:
                return True

            else:
                return False

    @property
    def Attachments(self):
        return self._parentExchange.GetAttachments(self)

    def HasAttachments(self):
        return self._hasAttachments

    @property
    def Data(self):
        return self._data.copy()

    def __iter__(self):
        for k, v in self._data.items():
            yield k, v

        for key in ['Start', 'End', 'Duration']:
            yield key, self.Get(key)

    def dict(self):
        # a json safe dict()
        ret = {}
        for k, v in self._data.items():
            ret[k] = v

        for key in ['Duration']:
            ret[key] = self.Get(key)

        for key in ['Start', 'End']:
            ret[key] = self.Get(key).timestamp()

        return ret

    def __str__(self):
        return '<CalendarItem: Start={}, End={}, Duration={}, Subject={}, HasAttachements={}, OrganizerName={}, ItemId[:10]={}, RoomName={}, LocationId={}>'.format(
            self.Get('Start'),
            self.Get('End'),
            self.Get('Duration'),
            self.Get('Subject'),
            self.HasAttachments(),
            self.Get('OrganizerName'),
            self.Get('ItemId')[:10] + '...',
            self.Get('RoomName'),
            self.Get('LocationId'),
        )

    def __repr__(self):
        return str(self)

    def __eq__(self, other):
        # print('188 __eq__ \nself.Data =', self.Data, ',\nother.Data=', other.Data)
        return self.Get('ItemId') == other.Get('ItemId') and \
               self.Data == other.Data

    def __lt__(self, other):
        # print('214 __gt__', self, other)
        #
        # print('192 __lt__', self, other)
        if isinstance(other, datetime.datetime):
            #     if other.tzinfo is None and self._startDT.tzinfo is not None:
            #         # other is naive, assume its in local system timezone
            #         other = other.astimezone()

            return self._startDT < other

        elif isinstance(other, _CalendarItem):
            return self._startDT < other._startDT

        else:
            raise TypeError('unorderable types: {} < {}'.format(self, other))

    def __le__(self, other):
        # print('214 __gt__', self, other)
        #
        # print('203 __le__', self, other)
        if isinstance(other, datetime.datetime):
            # if other.tzinfo is None and self._startDT.tzinfo is not None:
            #     # other is naive, assume its in local system timezone
            #     other = other.astimezone()

            return self._startDT <= other

        elif isinstance(other, _CalendarItem):
            return self._startDT <= other._startDT

        else:
            raise TypeError('unorderable types: {} < {}'.format(self, other))

    def __gt__(self, other):
        # print('214 __gt__', self, other)

        if isinstance(other, datetime.datetime):
            # if other.tzinfo is None and self._endDT.tzinfo is not None:
            # other is naive, assume its in local system timezone
            # other = other.astimezone()

            return self._endDT > other
        elif isinstance(other, _CalendarItem):
            return self._endDT > other._endDT

        else:
            raise TypeError('unorderable types: {} < {}'.format(self, other))

    def __ge__(self, other):
        # print('223 __ge__', self, other)

        if isinstance(other, datetime.datetime):
            # if other.tzinfo is None and self._endDT.tzinfo is not None:
            #     # other is naive, assume its in local system timezone
            #     other = other.astimezone()

            return self._endDT >= other
        elif isinstance(other, _CalendarItem):
            return self._endDT >= other._endDT

        else:
            raise TypeError('unorderable types: {} < {}'.format(self, other))


class _BaseCalendar:
    '''
    The Base for all calendar types ( Exchange, AdAstra )
    Dont use this class directly, instead subclass it.
    '''

    def __init__(self, *a, **k):
        self._connectionStatus = None
        self._Connected = None
        self._Disconnected = None

        self._CalendarItemDeleted = None  # callback for when an item is deleted
        self._CalendarItemChanged = None  # callback for when an item is changed
        self._NewCalendarItem = None  # callback for when an item is created

        self._calendarItems = []  # list of _CalendarItem object

        self._persitantStorage = k.get('persistentStorage', None)  # filepath or None
        # self._debug = k.get('debug', False)

        self._waitSaveToFile = Wait(1, self.SaveCalendarItemsToFile)
        self._waitSaveToFile.Cancel()

        self.LoadCalendarItemsFromFile()

    def print(self, *a, **k):
        if self._debug:
            print(*a, **k)

    @property
    def NewCalendarItem(self):
        return self._NewCalendarItem

    @NewCalendarItem.setter
    def NewCalendarItem(self, func):
        self._NewCalendarItem = func

    ##############
    @property
    def CalendarItemChanged(self):
        return self._CalendarItemChanged

    @CalendarItemChanged.setter
    def CalendarItemChanged(self, func):
        self._CalendarItemChanged = func

    ############
    @property
    def CalendarItemDeleted(self):
        return self._CalendarItemDeleted

    @CalendarItemDeleted.setter
    def CalendarItemDeleted(self, func):
        self._CalendarItemDeleted = func

    ############
    @property
    def Connected(self):
        return self._Connected

    @Connected.setter
    def Connected(self, func):
        self._Connected = func

    #############
    @property
    def Disconnected(self):
        return self._Disconnected

    @Disconnected.setter
    def Disconnected(self, func):
        self._Disconnected = func

    def _NewConnectionStatus(self, state):
        # self.print('378 _NewConnectionStatus(', state, ', self._connectionStatus=', self._connectionStatus)
        if state != self._connectionStatus:
            # the connection status has changed
            self._connectionStatus = state
            if state == 'Connected':
                if callable(self._Connected):
                    self._Connected(self, state)
            elif state == 'Disconnected':
                if callable(self._Disconnected):
                    self._Disconnected(self, state)

    def UpdateCalendar(self, calendar=None, startDT=None, endDT=None):
        '''
        Subclasses should override this

        :param calendar: a particular calendar ( None means use the default calendar)
        :param startDT: only search for events after this date
        :param endDT: only search for events before this date
        :return:
        '''
        raise NotImplementedError

    def CreateCalendarEvent(self, subject, body, startDT, endDT):
        '''
        Subclasses should override this

        Create a new calendar item with the above info

        :param subject:
        :param body:
        :param startDT:
        :param endDT:
        :return:
        '''
        raise NotImplementedError

    def ChangeEventTime(self, calItem, newStartDT, newEndDT):
        '''
        Subclasses should override this

        Changes the time of a current event

        :param calItem:
        :param newStartDT:
        :param newEndDT:
        :return:
        '''
        raise NotImplementedError

    def DeleteEvent(self, calItem):
        '''
        Subclasses should override this

        Deletes an event from the server

        :param calItem:
        :return:
        '''
        raise NotImplementedError

    # Dont override these below (unless you dare) #########################

    def GetCalendarItemsBySubject(self, exactMatch=None, partialMatch=None):
        ret = []
        for calItem in self._calendarItems:
            # self.print('426 searching for exactMatch={}, partialMatch={}'.format(exactMatch, partialMatch))
            if calItem.Get('Subject') == exactMatch:
                calItem = self._UpdateItemFromServer(calItem)
                ret.append(calItem)

            elif partialMatch and partialMatch in calItem.Get('Subject'):
                calItem = self._UpdateItemFromServer(calItem)
                ret.append(calItem)

        return ret

    def GetCalendarItemByID(self, itemId):
        for calItem in self._calendarItems:
            # self.print('424 searching for itemId={}, thisItemId={}'.format(itemId, calItem.Get('ItemId')))
            if calItem.Get('ItemId') == itemId:
                return calItem

    def GetAllEvents(self):
        return self._calendarItems.copy()

    def GetEventAtTime(self, dt=None):
        # dt = datetime.date or datetime.datetime
        # return a list of events that occur on datetime.date or at datetime.datetime

        if dt is None:
            dt = datetime.datetime.now()

        events = []

        for calItem in self._calendarItems.copy():
            if dt in calItem:
                events.append(calItem)

        return events

    def GetEventsInRange(self, startDT, endDT):
        self.UpdateCalendar(
            startDT=startDT,
            endDT=endDT,
        )
        ret = []
        for item in self._calendarItems:
            if startDT <= item <= endDT:
                ret.append(item)

        return ret

    def GetNowCalItems(self):
        # returns list of calendar nowItems happening now

        returnCalItems = []

        nowDT = datetime.datetime.now()

        for calItem in self._calendarItems.copy():
            if nowDT in calItem:
                returnCalItems.append(calItem)

        return returnCalItems

    def GetNextCalItems(self):
        # return a list CalendarItems
        # will not return events happening now. only the nearest future event(s)
        # if multiple events start at the same time, all CalendarItems will be returned

        nowDT = datetime.datetime.now()

        nextStartDT = None
        for calItem in self._calendarItems.copy():
            startDT = calItem.Get('Start')
            if startDT > nowDT:  # its in the future
                if nextStartDT is None or startDT < nextStartDT:  # its sooner than the previous soonest one. (Wha!?)
                    nextStartDT = startDT

        if nextStartDT is None:
            return []  # no events in the future
        else:
            returnCalItems = []
            for calItem in self._calendarItems.copy():
                if nextStartDT == calItem.Get('Start'):
                    returnCalItems.append(calItem)
            return returnCalItems

    def GetPreviousCalItems(self):
        # return a list CalendarItems
        # will not return events happening now. only the nearest future event(s)
        # if multiple events start at the same time, all CalendarItems will be returned

        nowDT = datetime.datetime.now()

        previousEndDT = None
        for calItem in self._calendarItems.copy():
            thisEndDT = calItem.Get('End')
            if thisEndDT > nowDT:  # its in the past
                if previousEndDT is None or thisEndDT > previousEndDT:
                    previousEndDT = thisEndDT

        if previousEndDT is None:
            return []  # no events in the future
        else:
            returnCalItems = []
            for calItem in self._calendarItems.copy():
                if previousEndDT == calItem.Get('End'):
                    returnCalItems.append(calItem)
            return returnCalItems

    def RegisterCalendarItems(self, calItems, startDT, endDT, doCallbacks=True):
        '''

        calItems should contain ALL the items between startDT and endDT

        :param calItems:
        :param startDT:
        :param endDT:
        :return:
        '''
        # Check for new and changed items
        for thisItem in calItems:
            itemInMemory = self.GetCalendarItemByID(thisItem.get('ItemId'))

            if itemInMemory is None:
                # this is a new item
                self._calendarItems.append(thisItem)
                if callable(self._NewCalendarItem) and doCallbacks:
                    self._NewCalendarItem(self, thisItem)

            elif itemInMemory != thisItem:
                # print('465')
                # print('itemInMemory=', itemInMemory)
                # print('thisItem    =', thisItem)
                # this item exist in memory but has somehow changed
                self._calendarItems.remove(itemInMemory)
                self._calendarItems.append(thisItem)
                if callable(self._CalendarItemChanged) and doCallbacks:
                    self._CalendarItemChanged(self, thisItem)

        # check for deleted items
        for itemInMemory in self._calendarItems.copy():
            if startDT <= itemInMemory <= endDT:
                if itemInMemory not in calItems:
                    # a event was deleted from the exchange server
                    self._calendarItems.remove(itemInMemory)
                    if callable(self._CalendarItemDeleted) and doCallbacks:
                        self._CalendarItemDeleted(self, itemInMemory)

        self._waitSaveToFile.Restart()

    def SaveCalendarItemsToFile(self):
        if self._persitantStorage:
            items = []
            for item in self._calendarItems.copy():
                items.append(item.dict())
            with File(self._persitantStorage, mode='wt') as file:
                file.write(json.dumps(items, indent=2, sort_keys=True))

    def LoadCalendarItemsFromFile(self):
        if self._persitantStorage and File.Exists(self._persitantStorage):
            try:
                calItems = []

                startDT = None
                endDT = None

                with File(self._persitantStorage, mode='rt') as file:
                    items = json.loads(file.read())

                    for item in items:
                        data = {}
                        for k, v, in item.items():
                            if k not in ['Start', 'End']:
                                data[k] = v

                        thisStartDT = datetime.datetime.fromtimestamp(item['Start'])
                        thisEndDT = datetime.datetime.fromtimestamp(item['End'])

                        if startDT is None or thisStartDT < startDT:
                            startDT = thisStartDT

                        if endDT is None or thisEndDT > endDT:
                            endDT = thisEndDT

                        calItem = _CalendarItem(
                            startDT=thisStartDT,
                            endDT=thisEndDT,
                            data=data,
                            parentCalendar=self,
                        )
                        calItems.append(calItem)

                self.RegisterCalendarItems(
                    calItems,
                    startDT=startDT or datetime.datetime.now(),
                    endDT=endDT or datetime.datetime.now(),
                    doCallbacks=False,
                )
            except Exception as e:
                ProgramLog(
                    'Error loading calendar items from disk: ' + str(e),
                    'error'
                )


def ConvertDatetimeToTimeString(dt):
    # converts to UTC time string
    dt = AdjustDatetimeForTimezone(dt, fromZone='Mine')
    return dt.strftime('%Y-%m-%dT%H:%M:%SZ')


def ConvertTimeStringToDatetime(string):
    '''
    Global Scripter blocks the user of 'strptime' for some reason.
    :param string:
    :return:
    '''
    # converts from UTC time string to datetime with my timezone

    # dt = datetime.datetime.strptime(string, '%Y-%m-%dT%H:%M:%SZ')
    year, month, etc = string.split('-')
    day, etc = etc.split('T')
    hour, minute, etc = etc.split(':')
    second = etc[:-1]
    dt = datetime.datetime(
        year=int(year),
        month=int(month),
        day=int(day),
        hour=int(hour),
        minute=int(minute),
        second=int(second),
    )
    dt = AdjustDatetimeForTimezone(dt, fromZone='Exchange')
    return dt


def AdjustDatetimeForTimezone(dt, fromZone):
    delta = datetime.timedelta(hours=abs(MY_TIME_ZONE))

    ts = time.mktime(dt.timetuple())
    lt = time.localtime(ts)
    dtIsDST = lt.tm_isdst > 0

    nowIsDST = time.localtime().tm_isdst > 0

    # print('nowIsDST=', nowIsDST)
    # print('dtIsDST=', dtIsDST)

    if fromZone == 'Mine':
        dt = dt + delta
        if dtIsDST and not nowIsDST:
            dt -= datetime.timedelta(hours=1)
        elif nowIsDST and not dtIsDST:
            dt += datetime.timedelta(hours=1)

    elif fromZone == 'Exchange':
        dt = dt - delta
        if dtIsDST and not nowIsDST:
            dt += datetime.timedelta(hours=1)
        elif nowIsDST and not dtIsDST:
            dt -= datetime.timedelta(hours=1)

    return dt
