# -*- coding: utf-8 -*-
from __future__ import unicode_literals

import os

from fileCatcher import _FileCatcher
from utils.vss import _VSS


class Windows2008ServerFiles(_FileCatcher):
    def __init__(self, params):
        super(Windows2008ServerFiles, self).__init__(params)
        drive, p = os.path.splitdrive(self.systemroot)
        self.vss = _VSS._get_instance(params, drive)

    def _changeroot(self, dir):
        drive, p = os.path.splitdrive(dir)
        path_return = self.vss._return_root() + p
        return path_return


    def csv_print_infos_files(self):
        super(Windows2008ServerFiles, self)._csv_infos_fs(self._list_files())
