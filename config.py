#!/usr/bin/env python
# coding: utf-8

from __future__ import (unicode_literals, division, absolute_import,
                        print_function)

import cStringIO, importlib, re, os, sys

from calibre.devices.usbms.driver import debug_print
from calibre.gui2 import show_restart_warning
from calibre.gui2.ui import get_gui
from calibre.utils.config import config_dir, JSONConfig

from calibre_plugins.marvin_manager.book_status import dialog_resources_path

from PyQt4.Qt import (Qt, QCheckBox, QComboBox, QFrame, QGridLayout, QGroupBox, QIcon,
                      QLabel, QPushButton, QSizePolicy, QSpacerItem, QVBoxLayout, QWidget)

plugin_prefs = JSONConfig('plugins/Marvin Mangler')

class ConfigWidget(QWidget):
    '''
    Config dialog for iOS Reader Apps
    '''
    # Location reporting template
    LOCATION_TEMPLATE = "{cls}:{func}({arg1}) {arg2}"

    def __init__(self, plugin_action):
        self.gui = get_gui()
        self.icon = plugin_action.icon
        self.parent = plugin_action
        self.prefs = plugin_prefs
        self.resources_path = plugin_action.resources_path
        self.verbose = plugin_action.verbose
        self._log_location()

        QWidget.__init__(self)
        self.l = QVBoxLayout()
        self.setLayout(self.l)

        # ~~~~~~~~ Create the Custom field options group box ~~~~~~~~
        self.cfg_custom_fields_gb = QGroupBox(self)
        self.cfg_custom_fields_gb.setTitle('Custom columns')
        self.l.addWidget(self.cfg_custom_fields_gb)

        self.cfg_custom_fields_qgl = QGridLayout(self.cfg_custom_fields_gb)
        current_row = 0

        # Custom column Wizard
        self.cfg_cc_wizard_button = QPushButton(QIcon(I('wizard.png')), 'Custom column wizard')
        self.cfg_cc_wizard_button.clicked.connect(self.launch_cc_wizard)
        # row, column, rowSpan, columnSpan
        self.cfg_custom_fields_qgl.addWidget(self.cfg_cc_wizard_button, current_row, 0, 1, 2)
        current_row += 1

        # Horizontal line
        self.cfg_hl_1 = QFrame(self.cfg_custom_fields_gb)
        self.cfg_hl_1.setFrameShape(QFrame.HLine)
        self.cfg_hl_1.setFrameShadow(QFrame.Sunken)
        self.cfg_hl_1.setObjectName("cfg_hl_1")
        self.cfg_custom_fields_qgl.addWidget(self.cfg_hl_1, current_row, 0, 1, 2)
        current_row += 1

        # Collections
        self.cfg_collections_label = QLabel('Collections')
        self.cfg_collections_label.setAlignment(Qt.AlignLeft)
        self.cfg_custom_fields_qgl.addWidget(self.cfg_collections_label, current_row, 0)

        self.collection_field_comboBox = QComboBox(self.cfg_custom_fields_gb)
        self.collection_field_comboBox.setObjectName('collection_field_comboBox')
        self.collection_field_comboBox.setToolTip('Custom field for Marvin collections')
        self.cfg_custom_fields_qgl.addWidget(self.collection_field_comboBox, current_row, 1)
        current_row += 1

        # Date read
        self.cfg_date_read_label = QLabel("Date read")
        self.cfg_date_read_label.setAlignment(Qt.AlignLeft)
        self.cfg_custom_fields_qgl.addWidget(self.cfg_date_read_label, current_row, 0)

        self.date_read_field_comboBox = QComboBox(self.cfg_custom_fields_gb)
        self.date_read_field_comboBox.setObjectName('date_read_field_comboBox')
        self.date_read_field_comboBox.setToolTip('Custom field for Date read')
        self.cfg_custom_fields_qgl.addWidget(self.date_read_field_comboBox, current_row, 1)
        current_row += 1

        # Highlights
        self.cfg_annotations_label = QLabel('Highlights')
        self.cfg_annotations_label.setAlignment(Qt.AlignLeft)
        self.cfg_custom_fields_qgl.addWidget(self.cfg_annotations_label, current_row, 0)

        self.annotations_field_comboBox = QComboBox(self.cfg_custom_fields_gb)
        self.annotations_field_comboBox.setObjectName('annotations_field_comboBox')
        self.annotations_field_comboBox.setToolTip('Custom field for Marvin annotations and highlights')
        self.cfg_custom_fields_qgl.addWidget(self.annotations_field_comboBox, current_row, 1)
        current_row += 1

        # Progress
        self.cfg_progress_label = QLabel('Progress')
        self.cfg_progress_label.setAlignment(Qt.AlignLeft)
        self.cfg_custom_fields_qgl.addWidget(self.cfg_progress_label, current_row, 0)

        self.progress_field_comboBox = QComboBox(self.cfg_custom_fields_gb)
        self.progress_field_comboBox.setObjectName('progress_field_comboBox')
        self.progress_field_comboBox.setToolTip('Custom field for Marvin reading progress')
        self.cfg_custom_fields_qgl.addWidget(self.progress_field_comboBox, current_row, 1)
        current_row += 1

        spacerItem1 = QSpacerItem(20, 60, QSizePolicy.Minimum, QSizePolicy.Expanding)
        self.cfg_custom_fields_qgl.addItem(spacerItem1)

        # ~~~~~~~~ Create the General options group box ~~~~~~~~
        self.cfg_runtime_options_gb = QGroupBox(self)
        self.cfg_runtime_options_gb.setTitle('General options')
        self.l.addWidget(self.cfg_runtime_options_gb)
        self.cfg_runtime_options_qvl = QVBoxLayout(self.cfg_runtime_options_gb)

        # ~~~~~~~~ Progress as percentage checkbox ~~~~~~~~
        self.reading_progress_checkbox = QCheckBox('Show reading progress as percentage')
        self.reading_progress_checkbox.setObjectName('show_progress_as_percentage')
        self.reading_progress_checkbox.setToolTip('Display percentage in Progress column')
        self.cfg_runtime_options_qvl.addWidget(self.reading_progress_checkbox)

        # ~~~~~~~~ Debug logging checkbox ~~~~~~~~
        self.debug_plugin_checkbox = QCheckBox('Enable debug logging for plugin')
        self.debug_plugin_checkbox.setObjectName('debug_plugin_checkbox')
        self.debug_plugin_checkbox.setToolTip('Print plugin diagnostic messages to console')
        self.cfg_runtime_options_qvl.addWidget(self.debug_plugin_checkbox)

        self.debug_libimobiledevice_checkbox = QCheckBox('Enable debug logging for libiMobileDevice')
        self.debug_libimobiledevice_checkbox.setObjectName('debug_libimobiledevice_checkbox')
        self.debug_libimobiledevice_checkbox.setToolTip('Print libiMobileDevice diagnostic messages to console')
        self.cfg_runtime_options_qvl.addWidget(self.debug_libimobiledevice_checkbox)

        spacerItem2 = QSpacerItem(20, 60, QSizePolicy.Minimum, QSizePolicy.Expanding)
        self.cfg_runtime_options_qvl.addItem(spacerItem2)

        spacerItem3 = QSpacerItem(20, 60, QSizePolicy.Minimum, QSizePolicy.Expanding)
        self.l.addItem(spacerItem3)

        # ~~~~~~~~ End of construction zone ~~~~~~~~
        self.resize(self.sizeHint())

        # Populate/restore the Annotations comboBox
        self.eligible_annotations_fields = self.get_eligible_custom_fields(eligible_types=['comments'])
        self.annotations_field_comboBox.addItems([''])
        ecf = sorted(self.eligible_annotations_fields.keys(), key=lambda s: s.lower())
        self.annotations_field_comboBox.addItems(ecf)
        cf = self.prefs.get('annotations_field_comboBox', '')
        idx = self.annotations_field_comboBox.findText(cf)
        if idx > -1:
            self.annotations_field_comboBox.setCurrentIndex(idx)

        # Populate/restore the Collections comboBox
        self.eligible_collection_fields = self.get_eligible_custom_fields(['enumeration', 'text'])
        self.collection_field_comboBox.addItems([''])
        ecf = sorted(self.eligible_collection_fields.keys(), key=lambda s: s.lower())
        self.collection_field_comboBox.addItems(ecf)
        cf = self.prefs.get('collection_field_comboBox', '')
        idx = self.collection_field_comboBox.findText(cf)
        if idx > -1:
            self.collection_field_comboBox.setCurrentIndex(idx)

        # Populate/restore the Date read comboBox
        self.eligible_date_read_fields = self.get_eligible_custom_fields(['datetime'])
        self.date_read_field_comboBox.addItems([''])
        ecf = sorted(self.eligible_date_read_fields.keys(), key=lambda s: s.lower())
        self.date_read_field_comboBox.addItems(ecf)
        cf = self.prefs.get('date_read_field_comboBox', '')
        idx = self.date_read_field_comboBox.findText(cf)
        if idx > -1:
            self.date_read_field_comboBox.setCurrentIndex(idx)

        # Populate/restore the Progress comboBox
        self.eligible_progress_fields = self.get_eligible_custom_fields(['float'])
        self.progress_field_comboBox.addItems([''])
        ecf = sorted(self.eligible_progress_fields.keys(), key=lambda s: s.lower())
        self.progress_field_comboBox.addItems(ecf)
        cf = self.prefs.get('progress_field_comboBox', '')
        idx = self.progress_field_comboBox.findText(cf)
        if idx > -1:
            self.progress_field_comboBox.setCurrentIndex(idx)

        # Restore general settings
        self.reading_progress_checkbox.setChecked(self.prefs.get('show_progress_as_percentage', False))
        self.debug_plugin_checkbox.setChecked(self.prefs.get('debug_plugin', False))
        self.debug_libimobiledevice_checkbox.setChecked(self.prefs.get('debug_libimobiledevice', False))

    def get_eligible_custom_fields(self, eligible_types=[]):
        '''
        Discover qualifying custom fields for reading progress
        '''
        #self._log_location(eligible_types)

        eligible_custom_fields = {}
        for cf in self.gui.current_db.custom_field_keys():
            cft = self.gui.current_db.metadata_for_field(cf)['datatype']
            cfn = self.gui.current_db.metadata_for_field(cf)['name']
            #self._log("cft: %s  cfn: %s" % (cft, cfn))
            if cft in eligible_types:
                eligible_custom_fields[cfn] = cf
        return eligible_custom_fields

    def launch_cc_wizard(self):
        '''
        '''
        self._log_location()

        klass = os.path.join(dialog_resources_path, 'cc_wizard.py')
        if os.path.exists(klass):
            #self._log("importing CC Wizard dialog from '%s'" % klass)
            sys.path.insert(0, dialog_resources_path)
            this_dc = importlib.import_module('cc_wizard')
            sys.path.remove(dialog_resources_path)
            dlg = this_dc.CustomColumnWizard(self, verbose=True)
            dlg.exec_()

            self._log("modified_columns: %s" % dlg.modified_columns)

            if dlg.modified_columns:
                do_restart = show_restart_warning('Restart calibre for the changes to be applied.',
                                                   parent=self.gui)
                if do_restart:
                    self.gui.quit(restart=True)
        else:
            self._log("ERROR: Can't import from '%s'" % klass)

    def save_settings(self):
        self._log_location()

        # Save annotations field
        cf = str(self.annotations_field_comboBox.currentText())
        self.prefs.set('annotations_field_comboBox', cf)
        if cf:
            self.prefs.set('annotations_field_lookup', self.eligible_annotations_fields[cf])
        else:
            self.prefs.set('annotations_field_lookup', '')

        # Save collection field
        cf = str(self.collection_field_comboBox.currentText())
        self.prefs.set('collection_field_comboBox', cf)
        if cf:
            self.prefs.set('collection_field_lookup', self.eligible_collection_fields[cf])
        else:
            self.prefs.set('collection_field_lookup', '')

        # Save Date read field
        cf = str(self.date_read_field_comboBox.currentText())
        self.prefs.set('date_read_field_comboBox', cf)
        if cf:
            self.prefs.set('date_read_field_lookup', self.eligible_date_read_fields[cf])
        else:
            self.prefs.set('date_read_field_lookup', '')

        # Save Progress field
        cf = str(self.progress_field_comboBox.currentText())
        self.prefs.set('progress_field_comboBox', cf)
        if cf:
            self.prefs.set('progress_field_lookup', self.eligible_progress_fields[cf])
        else:
            self.prefs.set('progress_field_lookup', '')

        # Save general settings
        self.prefs.set('show_progress_as_percentage', self.reading_progress_checkbox.isChecked())
        self.prefs.set('debug_plugin', self.debug_plugin_checkbox.isChecked())
        self.prefs.set('debug_libimobiledevice', self.debug_libimobiledevice_checkbox.isChecked())

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

        debug_print(self.LOCATION_TEMPLATE.format(cls=self.__class__.__name__,
            func=sys._getframe(1).f_code.co_name,
            arg1=arg1, arg2=arg2))


# For testing ConfigWidget, run from command line:
# cd ~/Documents/calibredev/Marvin_Manager
# calibre-debug config.py
# Search 'Marvin'
if __name__ == '__main__':
    from PyQt4.Qt import QApplication
    from calibre.gui2.preferences import test_widget
    app = QApplication([])
    test_widget('Advanced', 'Plugins')

