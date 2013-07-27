#!/usr/bin/env python
# coding: utf-8

from __future__ import (unicode_literals, division, absolute_import,
                        print_function)

import importlib, os, sys
from functools import partial

from calibre.devices.usbms.driver import debug_print
from calibre.gui2 import show_restart_warning
from calibre.gui2.ui import get_gui
from calibre.utils.config import JSONConfig

from calibre_plugins.marvin_manager.book_status import dialog_resources_path

from PyQt4.Qt import (Qt, QCheckBox, QComboBox, QGridLayout, QGroupBox, QIcon,
                      QLabel, QSizePolicy, QSpacerItem, QToolButton, QVBoxLayout, QWidget)

plugin_prefs = JSONConfig('plugins/Marvin Mangler')


class ConfigWidget(QWidget):
    '''
    Config dialog for Marvin Manager
    '''
    # Location reporting template
    LOCATION_TEMPLATE = "{cls}:{func}({arg1}) {arg2}"

    WIZARD_PROFILES = {
        'Collections': {
            'label': 'mm_collections',
            'datatype': 'text',
            'display': {u'is_names': False},
            'is_multiple': True
        },
        'Highlights': {
            'label': 'mm_highlights',
            'datatype': 'comments',
            'display': {},
            'is_multiple': False
        },
        'Last read': {
            'label': 'mm_date_read',
            'datatype': 'datetime',
            'display': {},
            'is_multiple': False
        },
        'Progress': {
            'label': 'mm_progress',
            'datatype': 'float',
            'display': {u'number_format': u'{0:.0f}%'},
            'is_multiple': False
        }
    }


    def __init__(self, plugin_action):
        self.gui = get_gui()
        self.icon = plugin_action.icon
        self.parent = plugin_action
        self.prefs = plugin_prefs
        self.resources_path = plugin_action.resources_path
        self.restart_required = False
        self.verbose = plugin_action.verbose

        self._log_location()

        QWidget.__init__(self)
        self.l = QVBoxLayout()
        self.setLayout(self.l)

        # ~~~~~~~~ Create the Custom fields options group box ~~~~~~~~
        self.cfg_custom_fields_gb = QGroupBox(self)
        self.cfg_custom_fields_gb.setTitle('Custom columns')
        self.l.addWidget(self.cfg_custom_fields_gb)

        self.cfg_custom_fields_qgl = QGridLayout(self.cfg_custom_fields_gb)
        current_row = 0

        # Collections
        self.cfg_collections_label = QLabel('Collections')
        self.cfg_collections_label.setAlignment(Qt.AlignLeft)
        self.cfg_custom_fields_qgl.addWidget(self.cfg_collections_label, current_row, 0)

        self.collection_field_comboBox = QComboBox(self.cfg_custom_fields_gb)
        self.collection_field_comboBox.setObjectName('collection_field_comboBox')
        self.collection_field_comboBox.setToolTip('Custom column for Marvin collections')
        self.cfg_custom_fields_qgl.addWidget(self.collection_field_comboBox, current_row, 1)

        self.cfg_collections_wizard = QToolButton()
        self.cfg_collections_wizard.setIcon(QIcon(I('wizard.png')))
        self.cfg_collections_wizard.setToolTip("Create a custom column for Collections")
        self.cfg_collections_wizard.clicked.connect(partial(self.launch_cc_wizard, 'Collections'))
        self.cfg_custom_fields_qgl.addWidget(self.cfg_collections_wizard, current_row, 2)
        current_row += 1

        # Date read
        self.cfg_date_read_label = QLabel("Last read")
        self.cfg_date_read_label.setAlignment(Qt.AlignLeft)
        self.cfg_custom_fields_qgl.addWidget(self.cfg_date_read_label, current_row, 0)

        self.date_read_field_comboBox = QComboBox(self.cfg_custom_fields_gb)
        self.date_read_field_comboBox.setObjectName('date_read_field_comboBox')
        self.date_read_field_comboBox.setToolTip('Custom column for Last read date')
        self.cfg_custom_fields_qgl.addWidget(self.date_read_field_comboBox, current_row, 1)

        self.cfg_collections_wizard = QToolButton()
        self.cfg_collections_wizard.setIcon(QIcon(I('wizard.png')))
        self.cfg_collections_wizard.setToolTip("Create a custom column for Last read date")
        self.cfg_collections_wizard.clicked.connect(partial(self.launch_cc_wizard, 'Last read'))
        self.cfg_custom_fields_qgl.addWidget(self.cfg_collections_wizard, current_row, 2)
        current_row += 1

        # Highlights
        self.cfg_annotations_label = QLabel('Highlights')
        self.cfg_annotations_label.setAlignment(Qt.AlignLeft)
        self.cfg_custom_fields_qgl.addWidget(self.cfg_annotations_label, current_row, 0)

        self.annotations_field_comboBox = QComboBox(self.cfg_custom_fields_gb)
        self.annotations_field_comboBox.setObjectName('annotations_field_comboBox')
        self.annotations_field_comboBox.setToolTip('Custom column for Marvin annotations and highlights')
        self.cfg_custom_fields_qgl.addWidget(self.annotations_field_comboBox, current_row, 1)

        self.cfg_collections_wizard = QToolButton()
        self.cfg_collections_wizard.setIcon(QIcon(I('wizard.png')))
        self.cfg_collections_wizard.setToolTip("Create a custom column for Highlights")
        self.cfg_collections_wizard.clicked.connect(partial(self.launch_cc_wizard, 'Highlights'))
        self.cfg_custom_fields_qgl.addWidget(self.cfg_collections_wizard, current_row, 2)
        current_row += 1

        # Progress
        self.cfg_progress_label = QLabel('Progress')
        self.cfg_progress_label.setAlignment(Qt.AlignLeft)
        self.cfg_custom_fields_qgl.addWidget(self.cfg_progress_label, current_row, 0)

        self.progress_field_comboBox = QComboBox(self.cfg_custom_fields_gb)
        self.progress_field_comboBox.setObjectName('progress_field_comboBox')
        self.progress_field_comboBox.setToolTip('Custom column for Marvin reading progress')
        self.cfg_custom_fields_qgl.addWidget(self.progress_field_comboBox, current_row, 1)

        self.cfg_collections_wizard = QToolButton()
        self.cfg_collections_wizard.setIcon(QIcon(I('wizard.png')))
        self.cfg_collections_wizard.setToolTip("Create a custom column for Progress")
        self.cfg_collections_wizard.clicked.connect(partial(self.launch_cc_wizard, 'Progress'))
        self.cfg_custom_fields_qgl.addWidget(self.cfg_collections_wizard, current_row, 2)
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

        # Populate/restore the Collections comboBox
        self.populate_collections()
        cf = self.prefs.get('collection_field_comboBox', '')
        idx = self.collection_field_comboBox.findText(cf)
        if idx > -1:
            self.collection_field_comboBox.setCurrentIndex(idx)

        # Populate/restore the Highlights comboBox
        self.populate_annotations()
        cf = self.prefs.get('annotations_field_comboBox', '')
        idx = self.annotations_field_comboBox.findText(cf)
        if idx > -1:
            self.annotations_field_comboBox.setCurrentIndex(idx)

        # Populate/restore the Last read comboBox
        self.populate_date_read()
        cf = self.prefs.get('date_read_field_comboBox', '')
        idx = self.date_read_field_comboBox.findText(cf)
        if idx > -1:
            self.date_read_field_comboBox.setCurrentIndex(idx)

        # Populate/restore the Progress comboBox
        self.populate_progress()
        cf = self.prefs.get('progress_field_comboBox', '')
        idx = self.progress_field_comboBox.findText(cf)
        if idx > -1:
            self.progress_field_comboBox.setCurrentIndex(idx)

        # Restore general settings
        self.reading_progress_checkbox.setChecked(self.prefs.get('show_progress_as_percentage', False))
        self.debug_plugin_checkbox.setChecked(self.prefs.get('debug_plugin', False))
        self.debug_libimobiledevice_checkbox.setChecked(self.prefs.get('debug_libimobiledevice', False))

    def get_eligible_custom_fields(self, eligible_types=[], is_multiple=None):
        '''
        Discover qualifying custom fields for reading progress
        '''
        #self._log_location(eligible_types)

        eligible_custom_fields = {}
        for cf in self.gui.current_db.custom_field_keys():
            cft = self.gui.current_db.metadata_for_field(cf)['datatype']
            cfn = self.gui.current_db.metadata_for_field(cf)['name']
            cfim = self.gui.current_db.metadata_for_field(cf)['is_multiple']
            #self._log("cf: %s  cft: %s  cfn: %s cfim: %s" % (cf, cft, cfn, cfim))
            if cft in eligible_types:
                if is_multiple is not None:
                    if bool(cfim) == is_multiple:
                        eligible_custom_fields[cfn] = cf
                else:
                    eligible_custom_fields[cfn] = cf
        return eligible_custom_fields

    def launch_cc_wizard(self, column_type):
        '''
        '''
        def _update_combo_box(comboBox, destination, previous):
            '''
            '''
            cb = getattr(self, comboBox)

            all_items = [str(cb.itemText(i))
                         for i in range(cb.count())]
            if previous and previous in all_items:
                all_items.remove(previous)
            all_items.append(destination)

            cb.clear()
            cb.addItems(sorted(all_items, key=lambda s: s.lower()))
            idx = cb.findText(destination)
            if idx > -1:
                cb.setCurrentIndex(idx)

        klass = os.path.join(dialog_resources_path, 'cc_wizard.py')
        if os.path.exists(klass):
            #self._log("importing CC Wizard dialog from '%s'" % klass)
            sys.path.insert(0, dialog_resources_path)
            this_dc = importlib.import_module('cc_wizard')
            sys.path.remove(dialog_resources_path)
            dlg = this_dc.CustomColumnWizard(self,
                                             column_type,
                                             self.WIZARD_PROFILES[column_type],
                                             verbose=True)
            dlg.exec_()

            if dlg.modified_column:
                self._log("modified_column: %s" % dlg.modified_column)

                self.restart_required = True

                destination = dlg.modified_column['destination']
                label = dlg.modified_column['label']
                previous = dlg.modified_column['previous']
                source = dlg.modified_column['source']

                if source == 'Collections':
                    _update_combo_box("collection_field_comboBox", destination, previous)

                    # Add/update the new destination so save_settings() can find it
                    self.eligible_collection_fields[destination] = label

                    # Save Date read field manually in case user cancels
                    self.prefs.set('collection_field_comboBox', destination)
                    self.prefs.set('collection_field_lookup', label)

                elif source == 'Last read':
                    _update_combo_box("date_read_field_comboBox", destination, previous)

                    # Add/update the new destination so save_settings() can find it
                    self.eligible_date_read_fields[destination] = label

                    # Save Date read field manually in case user cancels
                    self.prefs.set('date_read_field_comboBox', destination)
                    self.prefs.set('date_read_field_lookup', label)

                elif source == "Highlights":
                    _update_combo_box("annotations_field_comboBox", destination, previous)

                    # Add/update the new destination so save_settings() can find it
                    self.eligible_annotations_fields[destination] = label

                    # Save Date read field manually in case user cancels
                    self.prefs.set('annotations_field_comboBox', destination)
                    self.prefs.set('annotations_field_lookup', label)

                elif source == "Progress":
                    _update_combo_box("aprogress_field_comboBox", destination, previous)

                    # Add/update the new destination so save_settings() can find it
                    self.eligible_progress_fields[destination] = label

                    # Save Date read field manually in case user cancels
                    self.prefs.set('progress_field_comboBox', destination)
                    self.prefs.set('progress_field_lookup', label)

        else:
            self._log("ERROR: Can't import from '%s'" % klass)

    def populate_annotations(self):
        self.eligible_annotations_fields = self.get_eligible_custom_fields(eligible_types=['comments'])
        self.annotations_field_comboBox.addItems([''])
        ecf = sorted(self.eligible_annotations_fields.keys(), key=lambda s: s.lower())
        self.annotations_field_comboBox.addItems(ecf)

    def populate_collections(self):
        self.eligible_collection_fields = self.get_eligible_custom_fields(['enumeration', 'text'],
                                                                          is_multiple=True)
        self.collection_field_comboBox.addItems([''])
        ecf = sorted(self.eligible_collection_fields.keys(), key=lambda s: s.lower())
        self.collection_field_comboBox.addItems(ecf)

    def populate_date_read(self):
        self.eligible_date_read_fields = self.get_eligible_custom_fields(['datetime'])
        self.date_read_field_comboBox.addItems([''])
        ecf = sorted(self.eligible_date_read_fields.keys(), key=lambda s: s.lower())
        self.date_read_field_comboBox.addItems(ecf)

    def populate_progress(self):
        self.eligible_progress_fields = self.get_eligible_custom_fields(['float'])
        self.progress_field_comboBox.addItems([''])
        ecf = sorted(self.eligible_progress_fields.keys(), key=lambda s: s.lower())
        self.progress_field_comboBox.addItems(ecf)

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

        # If restart needed, inform user
        if self.restart_required:
            do_restart = show_restart_warning('Restart calibre for the changes to be applied.',
                                              parent=self.gui)
            if do_restart:
                self.gui.quit(restart=True)

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
