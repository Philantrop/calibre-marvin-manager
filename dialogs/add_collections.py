#!/usr/bin/env python
# coding: utf-8

from __future__ import (unicode_literals, division, absolute_import,
                        print_function)

__license__ = 'GPL v3'
__copyright__ = '2013, Gregory Riker'
__docformat__ = 'restructuredtext en'

import sys

from calibre.devices.usbms.driver import debug_print
from calibre.gui2 import warning_dialog

from calibre_plugins.marvin_manager.book_status import dialog_resources_path

from PyQt4.Qt import (QDialog, QDialogButtonBox, QIcon, QPixmap, QSize,
                      pyqtSignal)

# Import Ui_Form from form generated dynamically during initialization
if True:
    sys.path.insert(0, dialog_resources_path)
    from add_collections_ui import Ui_Dialog
    sys.path.remove(dialog_resources_path)

class AddCollectionsDialog(QDialog, Ui_Dialog):

    LOCATION_TEMPLATE = "{cls}:{func}({arg1}) {arg2}"

    marvin_device_status_changed = pyqtSignal(str)

    def __init__(self, parent, connected_device):
        QDialog.__init__(self, parent.opts.gui)

        self.parent = parent
        self.verbose = parent.verbose

        # Subscribe to Marvin driver change events
        connected_device.marvin_device_signals.reader_app_status_changed.connect(
            self.marvin_status_changed)

    def accept(self):
        self._log_location()
        super(AddCollectionsDialog, self).accept()

    def close(self):
        self._log_location()
        super(AddCollectionsDialog, self).close()

    def dispatch_button_click(self, button):
        '''
        BUTTON_ROLES = ['AcceptRole', 'RejectRole', 'DestructiveRole', 'ActionRole',
                        'HelpRole', 'YesRole', 'NoRole', 'ApplyRole', 'ResetRole']
        '''
        self._log_location()
        if self.bb.buttonRole(button) == QDialogButtonBox.AcceptRole:
            self.accept()

        elif self.bb.buttonRole(button) == QDialogButtonBox.RejectRole:
            self.close()

    def esc(self, *args):
        self.close()

    def initialize(self):
        '''
        __init__ is called on SizePersistedDialog()
        '''
        self.setupUi(self)

        self._log_location()
        self.setWindowTitle("Add collections")

        # Populate the icon
        self.icon.setText('')
        self.icon.setMaximumSize(QSize(40, 40))
        self.icon.setScaledContents(True)
        self.icon.setPixmap(QPixmap(I('plus.png')))

        # Add the Accept button
        self.accept_button = self.bb.addButton('Add', QDialogButtonBox.AcceptRole)
        self.accept_button.setDefault(True)
        self.accept_button.setEnabled(False)

        # Hook the QLineEdit box
        self.new_collection_le.textChanged.connect(self.validate_destination)

        # Hook the button events
        self.bb.clicked.connect(self.dispatch_button_click)

        # Set focus away from edit control so we can see default text
        self.bb.setFocus()

    def marvin_status_changed(self, command):
        '''

        '''
        self.marvin_device_status_changed.emit(command)

        self._log_location(command)

        if command in ['disconnected', 'yanked']:
            self._log("closing dialog: %s" % command)
            self.close()

    def validate_destination(self, new_collection):
        '''
        Confirm length of collection assignment > 0
        '''
        enabled = len(str(new_collection))
        self.accept_button.setEnabled(enabled)

    # ~~~~~~ Helpers ~~~~~~
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
