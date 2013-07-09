#!/usr/bin/env python
# coding: utf-8

from __future__ import (unicode_literals, division, absolute_import,
                        print_function)

__license__ = 'GPL v3'
__copyright__ = '2010, Gregory Riker'
__docformat__ = 'restructuredtext en'

import os, sys

from calibre.devices.usbms.driver import debug_print

from calibre_plugins.marvin_manager.book_status import dialog_resources_path
from calibre_plugins.marvin_manager.common_utils import SizePersistedDialog

from PyQt4.Qt import (QDialog, QDialogButtonBox, QIcon)

# Import Ui_Form from form generated dynamically during initialization
if True:
    sys.path.insert(0, dialog_resources_path)
    from metadata_ui import Ui_Dialog
    sys.path.remove(dialog_resources_path)

class MetadataComparisonDialog(SizePersistedDialog, Ui_Dialog):
    LOCATION_TEMPLATE = "{cls}:{func}({arg1}) {arg2}"

#     def __init__(self, parent):
#         QDialog.__init__(self, parent=None)
#         self.setupUi(self)
#         self.parent = parent
#         self.verbose = parent.verbose
#         self.initialize()

    def accept(self):
        self._log_location()
        super(MetadataComparisonDialog, self).accept()

    def close(self):
        self._log_location()
        super(MetadataComparisonDialog, self).close()

    def dispatch_button_click(self, button):
        '''
        BUTTON_ROLES = ['AcceptRole', 'RejectRole', 'DestructiveRole', 'ActionRole',
                        'HelpRole', 'YesRole', 'NoRole', 'ApplyRole', 'ResetRole']
        '''
        self._log_location()
        if self.bb.buttonRole(button) == QDialogButtonBox.AcceptRole:
            self._log("AcceptRole")
            self.accept()

        elif self.bb.buttonRole(button) == QDialogButtonBox.ActionRole:
            if button.objectName() == 'export_to_marvin_button':
                self.export_to_marvin()
            elif button.objectName() == 'import_from_marvin_button':
                self.import_from_marvin()

        elif self.bb.buttonRole(button) == QDialogButtonBox.RejectRole:
            self.close()

    def export_to_marvin(self):
        self._log_location()

    def import_from_marvin(self):
        self._log_location()

    def initialize(self, parent, selected_book):
        '''
        __init__ is called on SizePersistecDialog()
        '''
        self.setupUi(self)
        self.parent = parent
        self.verbose = parent.verbose

        self._log_location(selected_book)

        self.setWindowTitle(u'Metadata Comparison')

        self.calibre_title.setText("All about calibre")
        self.calibre_author.setText("Kovid Goyal")

        self.marvin_title.setText(selected_book['title'])
        self.marvin_author.setText(selected_book['author'])


        # ~~~~~~~~ Export to Marvin button ~~~~~~~~
        self.export_to_marvin_button = self.bb.addButton('Export to Marvin', QDialogButtonBox.ActionRole)
        self.export_to_marvin_button.setObjectName('export_to_marvin_button')
        self.export_to_marvin_button.setIcon(QIcon(os.path.join(self.parent.opts.resources_path,
                                                   'icons',
                                                   'from_calibre.png')))

        # ~~~~~~~~ Import from Marvin button ~~~~~~~~
        self.import_from_marvin_button = self.bb.addButton('Import from Marvin', QDialogButtonBox.ActionRole)
        self.import_from_marvin_button.setObjectName('import_from_marvin_button')
        self.import_from_marvin_button.setIcon(QIcon(os.path.join(self.parent.opts.resources_path,
                                                   'icons',
                                                   'from_marvin.png')))

        self.bb.clicked.connect(self.dispatch_button_click)

        # Restore position
        self.resize_dialog()

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

