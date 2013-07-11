#!/usr/bin/env python
# coding: utf-8

from __future__ import (unicode_literals, division, absolute_import,
                        print_function)

__license__ = 'GPL v3'
__copyright__ = '2010, Gregory Riker'
__docformat__ = 'restructuredtext en'

import os, sqlite3, sys
from functools import partial

from calibre import strftime
from calibre.devices.usbms.driver import debug_print
from calibre.utils.magick.draw import add_borders_to_image, thumbnail

from calibre_plugins.marvin_manager.book_status import dialog_resources_path
from calibre_plugins.marvin_manager.common_utils import SizePersistedDialog

from PyQt4.Qt import (Qt, QColor, QDialog, QDialogButtonBox, QIcon, QPalette, QPixmap,
                      QSize, QSizePolicy,
                      pyqtSignal)

# Import Ui_Form from form generated dynamically during initialization
if True:
    sys.path.insert(0, dialog_resources_path)
    from html_viewer_ui import Ui_Dialog
    sys.path.remove(dialog_resources_path)

class HTMLViewerDialog(SizePersistedDialog, Ui_Dialog):
    LOCATION_TEMPLATE = "{cls}:{func}({arg1}) {arg2}"

    marvin_device_status_changed = pyqtSignal(str)

    def accept(self):
        self._log_location()
        super(HTMLViewerDialog, self).accept()

    def close(self):
        self._log_location()
        super(HTMLViewerDialog, self).close()

    def dispatch_button_click(self, button):
        '''
        BUTTON_ROLES = ['AcceptRole', 'RejectRole', 'DestructiveRole', 'ActionRole',
                        'HelpRole', 'YesRole', 'NoRole', 'ApplyRole', 'ResetRole']
        '''
        self._log_location()
        if self.bb.buttonRole(button) == QDialogButtonBox.AcceptRole:
            self._log("AcceptRole")

            # User wants to save content


            self.accept()

        elif self.bb.buttonRole(button) == QDialogButtonBox.RejectRole:
            # User cancelled
            self._log("RejectRole")

            self.close()

    def esc(self, *args):
        self.close()

    def initialize(self, parent, titles, content, book_id, installed_book, marvin_db_path):
        '''
        __init__ is called on SizePersistedDialog()
        '''
        self.setupUi(self)
        self.book_id = book_id
        self.connected_device = parent.opts.gui.device_manager.device
        self.installed_book = installed_book
        self.marvin_db_path = marvin_db_path
        self.opts = parent.opts
        self.parent = parent
        self.stored_command = None
        self.verbose = parent.verbose
        self._log_location(installed_book.title)

        # Subscribe to Marvin driver change events
        self.connected_device.marvin_device_signals.reader_app_status_changed.connect(
            self.marvin_status_changed)

        # Set the titles
        self.setWindowTitle(titles['title'])
        self.html_gb.setTitle(titles['group_box_title'])

        # Initialize the contents of the TextBrowser
        self.html_tb.setText(content)

        # Set the bg color of the content to the dialog bg color
        bgcolor = self.palette().color(QPalette.Background)
        palette = QPalette()
        palette.setColor(QPalette.Base, bgcolor)
        self.html_tb.setPalette(palette)

        # Hook the button events
        self.bb.clicked.connect(self.dispatch_button_click)

        # Restore position
        self.resize_dialog()

    def marvin_status_changed(self, command):
        '''

        '''
        self.marvin_device_status_changed.emit(command)

        self._log_location(command)

        if command in ['disconnected', 'yanked']:
            self._log("closing dialog: %s" % command)
            self.close()

    def store_command(self, command):
        '''
        '''
        self._log_location(command)
        self.stored_command = command
        self.close()

    def _log(self, msg=None):
        '''
        Print msg to console
        '''
        if not self.verbose:
            return

        if msg:
            debug_print(" %s" % msg)
        else:
            debug_print()

    def _log_location(self, *args):
        '''
        Print location, args to console
        '''
        if not self.verbose:
            return

        arg1 = arg2 = ''

        if len(args) > 0:
            arg1 = args[0]
        if len(args) > 1:
            arg2 = args[1]

        debug_print(self.LOCATION_TEMPLATE.format(
            cls=self.__class__.__name__,
            func=sys._getframe(1).f_code.co_name,
            arg1=arg1, arg2=arg2))

