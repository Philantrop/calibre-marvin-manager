#!/usr/bin/env python
# coding: utf-8

from __future__ import (unicode_literals, division, absolute_import,
                        print_function)

__license__ = 'GPL v3'
__copyright__ = '2010, Gregory Riker'
__docformat__ = 'restructuredtext en'

import re, sys

from calibre.devices.usbms.driver import debug_print
from calibre.gui2 import warning_dialog

from calibre_plugins.marvin_manager.book_status import dialog_resources_path
from calibre_plugins.marvin_manager.common_utils import Logger

try:
    from PyQt5.Qt import (QDialog, QDialogButtonBox, QIcon, QPixmap,
                          QSize)
except ImportError:
    from PyQt4.Qt import (QDialog, QDialogButtonBox, QIcon, QPixmap,
                          QSize)

# Import Ui_Form from form generated dynamically during initialization
if True:
    sys.path.insert(0, dialog_resources_path)
    from deep_view_items_ui import Ui_Dialog
    sys.path.remove(dialog_resources_path)


class DeepViewItems(QDialog, Ui_Dialog, Logger):
    '''
    Present user with a list of DV items
    items keys(): ['ID', 'Name', 'Cnt', 'Loc', 'Flag', 'Note', 'Confidence']
    '''

    def __init__(self, parent, title, items, verbose=True):
        QDialog.__init__(self, parent)
        self.items = items
        self.parent = parent
        self.result = None

        self.setupUi(self)
        self.verbose = verbose
        self._log_location()

        # Set title
        self.dvi_gb.setTitle(title)

        # Populate the combobox
        for item in items:
            self.dvi_cb.addItem("%s (%d)" % (item[b'Name'], item[b'Cnt']),
                                item[b'ID'])

        # Retrieve saved hit count
        hits = self.parent.prefs.get('deep_view_hits', 5)
        self.dvi_sb.setValue(hits)

        # Hook the button events
        self.bb.clicked.connect(self.dispatch_button_click)

    def accept(self):
        self._log_location()
        index = self.dvi_cb.currentIndex()
        item = re.sub("\s{1}\(\d+\)", '', str(self.dvi_cb.currentText()))
        self.result = {'ID':  str(self.dvi_cb.itemData(index).toString()),
                       'hits': str(self.dvi_sb.value()),
                       'item': item}

        # Save the hit count
        self.parent.prefs.set('deep_view_hits', int(self.dvi_sb.value()))

        super(DeepViewItems, self).accept()

    def close(self):
        self._log_location()
        self.result = None

        # Save the hit count
        self.parent.prefs.set('deep_view_hits', int(self.dvi_sb.value()))

        super(DeepViewItems, self).close()

    def dispatch_button_click(self, button):
        '''
        BUTTON_ROLES = ['AcceptRole', 'RejectRole', 'DestructiveRole', 'ActionRole',
                        'HelpRole', 'YesRole', 'NoRole', 'ApplyRole', 'ResetRole']
        '''
        self._log_location()
        if self.bb.buttonRole(button) == QDialogButtonBox.AcceptRole:
            # Save content
            self.accept()

        elif self.bb.buttonRole(button) == QDialogButtonBox.RejectRole:
            # Cancelled
            self.close()

    def esc(self, *args):
        self.cancel()
