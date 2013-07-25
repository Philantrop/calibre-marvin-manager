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
from calibre.gui2 import warning_dialog
from calibre.library.custom_columns import CustomColumns

from calibre_plugins.marvin_manager.book_status import dialog_resources_path

from PyQt4.Qt import (Qt, QColor, QDialog, QDialogButtonBox, QIcon, QPalette, QPixmap,
                      QSize, QSizePolicy,
                      pyqtSignal)

# Import Ui_Form from form generated dynamically during initialization
if True:
    sys.path.insert(0, dialog_resources_path)
    from cc_wizard_ui import Ui_Dialog
    sys.path.remove(dialog_resources_path)

class CustomColumnWizard(QDialog, Ui_Dialog):
    FIELDS = {
              'Highlights': {'label': 'mm_highlights',
                             'datatype': 'comments',
                             'display': {}},
              'Last read':  {'label': 'mm_last_read',
                             'datatype': 'datetime',
                             'display': {}},
              'Progress':   {'label': 'mm_progress',
                             'datatype': 'float',
                             'display': {u'number_format': u'{0:.0f}%'}}
             }

    LOCATION_TEMPLATE = "{cls}:{func}({arg1}) {arg2}"

    STEP_ONE = "1. Select metadata to be added to a custom column"
    STEP_TWO = "2. Specify a name for the custom column"
    STEP_THREE_APPLY = "3. Click Apply to create the custom column"
    STEP_THREE_RENAME = "3. Click Rename to rename the custom column"

    YELLOW_BG = '<font style="background:#FDFF99">{0}</font>'

    def __init__(self, parent, verbose=True):
        QDialog.__init__(self, parent.gui)
        self.db = parent.gui.current_db
        self.gui = parent.gui
        self.modified_columns = []

        self.setupUi(self)
        self.verbose = verbose
        self._log_location()

        # Populate the icon
        self.icon.setText('')
        self.icon.setMaximumSize(QSize(56, 56))
        self.icon.setScaledContents(True)
        self.icon.setPixmap(QPixmap(I('wizard.png')))

        # Add the Action button
        self.action_button = self.bb.addButton('Action button', QDialogButtonBox.ActionRole)
        self.action_button.setDefault(True)

        # Populate marvin_source_comboBox
        self.marvin_source_comboBox.addItems([''])
        self.marvin_source_comboBox.addItems(sorted(self.FIELDS.keys()))
        self.marvin_source_comboBox.currentIndexChanged.connect(self.source_changed)

        self.highlight_step(1)
        self.reset_action_button()

        # Hook the QLineEdit box
        self.calibre_destination_le.textChanged.connect(self.validate_destination)

        # Hook the button events
        self.bb.clicked.connect(self.dispatch_button_click)

    def accept(self):
        self._log_location()
        super(CustomColumnWizard, self).accept()

    def close(self):
        self._log_location()
        super(CustomColumnWizard, self).close()

    def custom_column_add(self, requested_name, profile):
        '''
        Add the requested custom column with profile
        '''
        self._log_location(requested_name)
        self._log(profile)
        self.db.create_custom_column(profile['label'],
                                     requested_name,
                                     profile['datatype'],
                                     False,
                                     display=profile['display'])
        self.modified_columns.append({'source': profile['source'], 'destination': requested_name})

    def custom_column_rename(self, requested_name, profile):
        '''
        The name already exists for label, update it
        '''
        self._log_location(requested_name)
        self._log(profile)

        # Find the existing
        for cf in self.db.custom_field_keys():
            #self._log(self.db.metadata_for_field(cf))
            mi = self.db.metadata_for_field(cf)
            if mi['label'] == profile['label']:
                self.db.set_custom_column_metadata(mi['colnum'],
                                                   name=requested_name,
                                                   label=mi['label'],
                                                   display=mi['display'])
                self.modified_columns.append({'source': profile['source'], 'destination': requested_name})
                break

    def dispatch_button_click(self, button):
        '''
        BUTTON_ROLES = ['AcceptRole', 'RejectRole', 'DestructiveRole', 'ActionRole',
                        'HelpRole', 'YesRole', 'NoRole', 'ApplyRole', 'ResetRole']
        '''
        self._log_location()
        if self.bb.buttonRole(button) == QDialogButtonBox.AcceptRole:
            self._log("AcceptRole")

            self.accept()

        if self.bb.buttonRole(button) == QDialogButtonBox.ActionRole:
            self._log("ActionRole")
            requested_name = str(self.calibre_destination_le.text())

            if requested_name in self.get_custom_column_names():
                self._log("'%s' already in use" % requested_name)
                warning_dialog(self.gui,
                               "Already in use",
                               "<p>'%s' is an existing custom column.</p><p>Pick a different name.</p>" % requested_name,
                               show=True, show_copy_button=False)
            else:
                source = str(self.marvin_source_comboBox.currentText())
                profile = self.FIELDS[source]
                profile['source'] = source
                if button.objectName() == 'add_button':
                    self.custom_column_add(requested_name, profile)
                elif button.objectName() == 'rename_button':
                    self.custom_column_rename(requested_name, profile)
                else:
                    self._log("ERROR: unrecognized button name")

        elif self.bb.buttonRole(button) == QDialogButtonBox.RejectRole:
            self._log("RejectRole")

            self.close()

    def esc(self, *args):
        self.close()

    def highlight_step(self, step):
        '''
        '''
        self._log_location(step)
        if step == 1:
            self.step_1.setText(self.YELLOW_BG.format(self.STEP_ONE))
            self.step_2.setText(self.STEP_TWO)
            self.step_3.setText(self.STEP_THREE_APPLY)

        elif step == 2:
            self.step_1.setText(self.STEP_ONE)
            self.step_2.setText(self.YELLOW_BG.format(self.STEP_TWO))
            self.step_3.setText(self.STEP_THREE_APPLY)

    def get_custom_column_names(self):
        '''
        '''
        self._log_location()
        existing_custom_names = []
        for cf in self.db.custom_field_keys():
            #self._log(self.db.metadata_for_field(cf))
            existing_custom_names.append(self.db.metadata_for_field(cf)['name'])
        return existing_custom_names

    def reset_action_button(self, action="add_button", enabled=False):
        '''
        '''
        self.action_button.setObjectName(action)
        if action == "add_button":
            self.action_button.setText('Add custom column')
            self.action_button.setIcon(QIcon(I('plus.png')))
        elif action == "rename_button":
            self.action_button.setText("Rename custom column")
            self.action_button.setIcon(QIcon(I('edit_input.png')))
        self.action_button.setEnabled(enabled)

    def source_changed(self, index):
        '''
        '''
        self._log_location(index)
        if index == 0:
            self.highlight_step(1)

            self.reset_action_button()

        elif index > 0:
            self.highlight_step(2)

            selected = str(self.marvin_source_comboBox.currentText())
            existing = None
            label = self.FIELDS[selected]['label']
            self._log("label: %s" % label)
            for cf in self.db.custom_field_keys():
                #self._log(self.db.metadata_for_field(cf))
                cfd = self.db.metadata_for_field(cf)
                if cfd['label'] == label:
                    existing = cfd['name']
                    break

            # Does lookup already exist?
            if existing:
                self.calibre_destination_le.setText(existing)
                self.reset_action_button(action="rename_button", enabled=True)
            else:
                # Populate the edit box with the default Column name
                self.calibre_destination_le.setText(selected)
                self.reset_action_button(action="add_button", enabled=True)

            # Select the text
            self.calibre_destination_le.selectAll()
            self.calibre_destination_le.setFocus()

    def validate_destination(self, destination):
        '''
        Confirm length of column name > 0
        '''
        enabled = len(str(destination))
        self.action_button.setEnabled(enabled)

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

