# -*- coding: utf-8 -*-
from __future__ import unicode_literals
from logs import _EventLogs


class Windows2003ServerR2Evt(_EventLogs):
    def __init__(self, params):
        super(Windows2003ServerR2Evt, self).__init__(params)

    def csv_event_logs(self):
        super(Windows2003ServerR2Evt, self)._csv_event_logs(True)
