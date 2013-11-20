#!/usr/bin/env python
# coding: utf-8

from __future__ import (unicode_literals, division, absolute_import,
                        print_function)

import hashlib, importlib, os, sys
from datetime import datetime
from functools import partial
from time import mktime

from calibre.constants import islinux, isosx, iswindows
from calibre.devices.usbms.driver import debug_print
from calibre.ebooks.BeautifulSoup import BeautifulSoup
from calibre.gui2.dialogs.message_box import MessageBox
from calibre.gui2 import show_restart_warning
from calibre.gui2.ui import get_gui
from calibre.utils.config import JSONConfig

from calibre_plugins.marvin_manager.appearance import (AnnotationsAppearance,
    default_elements, default_timestamp)

from calibre_plugins.marvin_manager.common_utils import (existing_annotations,
    get_icon, move_annotations)

from PyQt4.Qt import (Qt, QCheckBox, QComboBox, QFont, QFontMetrics, QFrame,
                      QGridLayout, QGroupBox, QIcon,
                      QLabel, QPlainTextEdit, QPushButton,
                      QSizePolicy, QSpacerItem, QThread, QTimer, QToolButton,
                      QVBoxLayout, QWidget,
                      SIGNAL)

plugin_prefs = JSONConfig('plugins/Marvin XD')


class ConfigWidget(QWidget):
    '''
    Config dialog for Marvin Manager
    '''

    # Location reporting template
    LOCATION_TEMPLATE = "{cls}:{func}({arg1}) {arg2}"

    WIZARD_PROFILES = {
        'Annotations': {
            'label': 'mm_annotations',
            'datatype': 'comments',
            'display': {},
            'is_multiple': False
        },
        'Collections': {
            'label': 'mm_collections',
            'datatype': 'text',
            'display': {u'is_names': False},
            'is_multiple': True
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
        },
        'Read': {
            'label': 'mm_read',
            'datatype': 'bool',
            'display': {},
            'is_multiple': False
        },
        'Reading list': {
            'label': 'mm_reading_list',
            'datatype': 'bool',
            'display': {},
            'is_multiple': False
        },
        'Word count': {
            'label': 'mm_word_count',
            'datatype': 'int',
            'display': {u'number_format': u'{0:n}'},
            'is_multiple': False
        }
    }

    def __init__(self, plugin_action):
        QWidget.__init__(self)
        self.parent = plugin_action

        self.gui = get_gui()
        self.icon = plugin_action.icon
        self.opts = plugin_action.opts
        self.prefs = plugin_prefs
        self.resources_path = plugin_action.resources_path
        self.verbose = plugin_action.verbose

        self.restart_required = False

        self._log_location()

        self.l = QVBoxLayout()
        self.setLayout(self.l)

        # ~~~~~~~~ Create the Custom fields options group box ~~~~~~~~
        self.cfg_custom_fields_gb = QGroupBox(self)
        self.cfg_custom_fields_gb.setTitle('Custom column assignments')
        self.l.addWidget(self.cfg_custom_fields_gb)

        self.cfg_custom_fields_qgl = QGridLayout(self.cfg_custom_fields_gb)
        current_row = 0

        # ++++++++ Labels + HLine ++++++++
        self.marvin_source_label = QLabel("Marvin source")
        self.cfg_custom_fields_qgl.addWidget(self.marvin_source_label, current_row, 0)
        self.calibre_destination_label = QLabel("calibre destination")
        self.cfg_custom_fields_qgl.addWidget(self.calibre_destination_label, current_row, 1)
        current_row += 1
        self.sd_hl = QFrame(self.cfg_custom_fields_gb)
        self.sd_hl.setFrameShape(QFrame.HLine)
        self.sd_hl.setFrameShadow(QFrame.Raised)
        self.cfg_custom_fields_qgl.addWidget(self.sd_hl, current_row, 0, 1, 3)
        current_row += 1

        # ++++++++ Annotations ++++++++
        self.cfg_annotations_label = QLabel('Annotations')
        self.cfg_annotations_label.setAlignment(Qt.AlignLeft)
        self.cfg_custom_fields_qgl.addWidget(self.cfg_annotations_label, current_row, 0)

        self.annotations_field_comboBox = QComboBox(self.cfg_custom_fields_gb)
        self.annotations_field_comboBox.setObjectName('annotations_field_comboBox')
        self.annotations_field_comboBox.setToolTip('Select a custom column to store Marvin annotations')
        self.cfg_custom_fields_qgl.addWidget(self.annotations_field_comboBox, current_row, 1)

        self.cfg_highlights_wizard = QToolButton()
        self.cfg_highlights_wizard.setIcon(QIcon(I('wizard.png')))
        self.cfg_highlights_wizard.setToolTip("Create a custom column to store Marvin annotations")
        self.cfg_highlights_wizard.clicked.connect(partial(self.launch_cc_wizard, 'Annotations'))
        self.cfg_custom_fields_qgl.addWidget(self.cfg_highlights_wizard, current_row, 2)
        current_row += 1

        # ++++++++ Collections ++++++++
        self.cfg_collections_label = QLabel('Collections')
        self.cfg_collections_label.setAlignment(Qt.AlignLeft)
        self.cfg_custom_fields_qgl.addWidget(self.cfg_collections_label, current_row, 0)

        self.collection_field_comboBox = QComboBox(self.cfg_custom_fields_gb)
        self.collection_field_comboBox.setObjectName('collection_field_comboBox')
        self.collection_field_comboBox.setToolTip('Select a custom column to store Marvin collection assignments')
        self.cfg_custom_fields_qgl.addWidget(self.collection_field_comboBox, current_row, 1)

        self.cfg_collections_wizard = QToolButton()
        self.cfg_collections_wizard.setIcon(QIcon(I('wizard.png')))
        self.cfg_collections_wizard.setToolTip("Create a custom column for Marvin collection assignments")
        self.cfg_collections_wizard.clicked.connect(partial(self.launch_cc_wizard, 'Collections'))
        self.cfg_custom_fields_qgl.addWidget(self.cfg_collections_wizard, current_row, 2)
        current_row += 1

        # ++++++++ Last read ++++++++
        self.cfg_date_read_label = QLabel("Last read")
        self.cfg_date_read_label.setAlignment(Qt.AlignLeft)
        self.cfg_custom_fields_qgl.addWidget(self.cfg_date_read_label, current_row, 0)

        self.date_read_field_comboBox = QComboBox(self.cfg_custom_fields_gb)
        self.date_read_field_comboBox.setObjectName('date_read_field_comboBox')
        self.date_read_field_comboBox.setToolTip('Select a custom column to store Last read date')
        self.cfg_custom_fields_qgl.addWidget(self.date_read_field_comboBox, current_row, 1)

        self.cfg_collections_wizard = QToolButton()
        self.cfg_collections_wizard.setIcon(QIcon(I('wizard.png')))
        self.cfg_collections_wizard.setToolTip("Create a custom column to store Last read date")
        self.cfg_collections_wizard.clicked.connect(partial(self.launch_cc_wizard, 'Last read'))
        self.cfg_custom_fields_qgl.addWidget(self.cfg_collections_wizard, current_row, 2)
        current_row += 1

        # ++++++++ Progress ++++++++
        self.cfg_progress_label = QLabel('Progress')
        self.cfg_progress_label.setAlignment(Qt.AlignLeft)
        self.cfg_custom_fields_qgl.addWidget(self.cfg_progress_label, current_row, 0)

        self.progress_field_comboBox = QComboBox(self.cfg_custom_fields_gb)
        self.progress_field_comboBox.setObjectName('progress_field_comboBox')
        self.progress_field_comboBox.setToolTip('Select a custom column to store Marvin reading progress')
        self.cfg_custom_fields_qgl.addWidget(self.progress_field_comboBox, current_row, 1)

        self.cfg_progress_wizard = QToolButton()
        self.cfg_progress_wizard.setIcon(QIcon(I('wizard.png')))
        self.cfg_progress_wizard.setToolTip("Create a custom column to store Marvin reading progress")
        self.cfg_progress_wizard.clicked.connect(partial(self.launch_cc_wizard, 'Progress'))
        self.cfg_custom_fields_qgl.addWidget(self.cfg_progress_wizard, current_row, 2)
        current_row += 1

        # ++++++++ Read flag ++++++++
        self.cfg_read_label = QLabel('Read')
        self.cfg_read_label.setAlignment(Qt.AlignLeft)
        self.cfg_custom_fields_qgl.addWidget(self.cfg_read_label, current_row, 0)

        self.read_field_comboBox = QComboBox(self.cfg_custom_fields_gb)
        self.read_field_comboBox.setObjectName('read_field_comboBox')
        self.read_field_comboBox.setToolTip('Select a custom column to store Marvin Read status')
        self.cfg_custom_fields_qgl.addWidget(self.read_field_comboBox, current_row, 1)

        self.cfg_read_wizard = QToolButton()
        self.cfg_read_wizard.setIcon(QIcon(I('wizard.png')))
        self.cfg_read_wizard.setToolTip("Create a custom column to store Marvin Read status")
        self.cfg_read_wizard.clicked.connect(partial(self.launch_cc_wizard, 'Read'))
        self.cfg_custom_fields_qgl.addWidget(self.cfg_read_wizard, current_row, 2)
        current_row += 1

        # ++++++++ Reading list flag ++++++++
        self.cfg_reading_list_label = QLabel('Reading list')
        self.cfg_reading_list_label.setAlignment(Qt.AlignLeft)
        self.cfg_custom_fields_qgl.addWidget(self.cfg_reading_list_label, current_row, 0)

        self.reading_list_field_comboBox = QComboBox(self.cfg_custom_fields_gb)
        self.reading_list_field_comboBox.setObjectName('reading_list_field_comboBox')
        self.reading_list_field_comboBox.setToolTip('Select a custom column to store Marvin Reading list status')
        self.cfg_custom_fields_qgl.addWidget(self.reading_list_field_comboBox, current_row, 1)

        self.cfg_reading_list_wizard = QToolButton()
        self.cfg_reading_list_wizard.setIcon(QIcon(I('wizard.png')))
        self.cfg_reading_list_wizard.setToolTip("Create a custom column to store Marvin Reading list status")
        self.cfg_reading_list_wizard.clicked.connect(partial(self.launch_cc_wizard, 'Reading list'))
        self.cfg_custom_fields_qgl.addWidget(self.cfg_reading_list_wizard, current_row, 2)
        current_row += 1

        # ++++++++ Word count ++++++++
        self.cfg_word_count_label = QLabel('Word count')
        self.cfg_word_count_label.setAlignment(Qt.AlignLeft)
        self.cfg_custom_fields_qgl.addWidget(self.cfg_word_count_label, current_row, 0)

        self.word_count_field_comboBox = QComboBox(self.cfg_custom_fields_gb)
        self.word_count_field_comboBox.setObjectName('word_count_field_comboBox')
        self.word_count_field_comboBox.setToolTip('Select a custom column to store Marvin word counts')
        self.cfg_custom_fields_qgl.addWidget(self.word_count_field_comboBox, current_row, 1)

        self.cfg_word_count_wizard = QToolButton()
        self.cfg_word_count_wizard.setIcon(QIcon(I('wizard.png')))
        self.cfg_word_count_wizard.setToolTip("Create a custom column to store Marvin word counts")
        self.cfg_word_count_wizard.clicked.connect(partial(self.launch_cc_wizard, 'Word count'))
        self.cfg_custom_fields_qgl.addWidget(self.cfg_word_count_wizard, current_row, 2)
        current_row += 1

        spacerItem1 = QSpacerItem(20, 20, QSizePolicy.Minimum, QSizePolicy.Expanding)
        self.cfg_custom_fields_qgl.addItem(spacerItem1)

        # ~~~~~~~~ Create the CSS group box ~~~~~~~~
        self.cfg_css_options_gb = QGroupBox(self)
        self.cfg_css_options_gb.setTitle('CSS')
        self.l.addWidget(self.cfg_css_options_gb)
        self.cfg_css_options_qgl = QGridLayout(self.cfg_css_options_gb)

        current_row = 0

        # ++++++++ Annotations appearance ++++++++
        self.annotations_icon = QIcon(os.path.join(self.resources_path, 'icons', 'annotations_hiliter.png'))
        self.cfg_annotations_appearance_toolbutton = QToolButton()
        self.cfg_annotations_appearance_toolbutton.setIcon(self.annotations_icon)
        self.cfg_annotations_appearance_toolbutton.clicked.connect(self.configure_appearance)
        self.cfg_css_options_qgl.addWidget(self.cfg_annotations_appearance_toolbutton, current_row, 0)
        self.cfg_annotations_label = ClickableQLabel("Annotations")
        self.connect(self.cfg_annotations_label, SIGNAL('clicked()'), self.configure_appearance)
        self.cfg_css_options_qgl.addWidget(self.cfg_annotations_label, current_row, 1)
        current_row += 1

        # ++++++++ Injected CSS ++++++++
        self.css_editor_icon = QIcon(I('format-text-heading.png'))
        self.cfg_css_editor_toolbutton = QToolButton()
        self.cfg_css_editor_toolbutton.setIcon(self.css_editor_icon)
        self.cfg_css_editor_toolbutton.clicked.connect(self.edit_css)
        self.cfg_css_options_qgl.addWidget(self.cfg_css_editor_toolbutton, current_row, 0)
        self.cfg_css_editor_label = ClickableQLabel("Articles, Vocabulary")
        self.connect(self.cfg_css_editor_label, SIGNAL('clicked()'), self.edit_css)
        self.cfg_css_options_qgl.addWidget(self.cfg_css_editor_label, current_row, 1)

        # ~~~~~~~~ Create the General options group box ~~~~~~~~
        self.cfg_runtime_options_gb = QGroupBox(self)
        self.cfg_runtime_options_gb.setTitle('General options')
        self.l.addWidget(self.cfg_runtime_options_gb)
        self.cfg_runtime_options_qvl = QVBoxLayout(self.cfg_runtime_options_gb)

        # ++++++++ Auto refresh checkbox ++++++++
        self.auto_refresh_checkbox = QCheckBox('Automatically refresh custom column content')
        self.auto_refresh_checkbox.setObjectName('auto_refresh_at_startup')
        self.auto_refresh_checkbox.setToolTip('Update calibre custom column when Marvin XD is opened')
        self.cfg_runtime_options_qvl.addWidget(self.auto_refresh_checkbox)

        # ++++++++ Progress as percentage checkbox ++++++++
        self.reading_progress_checkbox = QCheckBox('Show reading progress as percentage')
        self.reading_progress_checkbox.setObjectName('show_progress_as_percentage')
        self.reading_progress_checkbox.setToolTip('Display percentage in Progress column')
        self.cfg_runtime_options_qvl.addWidget(self.reading_progress_checkbox)

        # ~~~~~~~~ Create the Debug options group box ~~~~~~~~
        self.cfg_debug_options_gb = QGroupBox(self)
        self.cfg_debug_options_gb.setTitle('Debug options')
        self.l.addWidget(self.cfg_debug_options_gb)
        self.cfg_debug_options_qvl = QVBoxLayout(self.cfg_debug_options_gb)

        # ++++++++ Debug logging checkboxes ++++++++
        self.debug_plugin_checkbox = QCheckBox('Enable debug logging for Marvin XD')
        self.debug_plugin_checkbox.setObjectName('debug_plugin_checkbox')
        self.debug_plugin_checkbox.setToolTip('Print plugin diagnostic messages to console')
        self.cfg_debug_options_qvl.addWidget(self.debug_plugin_checkbox)

        self.debug_libimobiledevice_checkbox = QCheckBox('Enable debug logging for libiMobileDevice')
        self.debug_libimobiledevice_checkbox.setObjectName('debug_libimobiledevice_checkbox')
        self.debug_libimobiledevice_checkbox.setToolTip('Print libiMobileDevice diagnostic messages to console')
        self.cfg_debug_options_qvl.addWidget(self.debug_libimobiledevice_checkbox)

        # Group box spacer
        spacerItem2 = QSpacerItem(20, 20, QSizePolicy.Minimum, QSizePolicy.Expanding)
        self.cfg_runtime_options_qvl.addItem(spacerItem2)

        # Widget spacer
        spacerItem3 = QSpacerItem(20, 20, QSizePolicy.Minimum, QSizePolicy.Expanding)
        self.l.addItem(spacerItem3)

        # ~~~~~~~~ End of construction zone ~~~~~~~~
        self.resize(self.sizeHint())

        # Populate/restore the Collections comboBox
        self.populate_collections()
        cf = self.prefs.get('collection_field_comboBox', '')
        idx = self.collection_field_comboBox.findText(cf)
        if idx > -1:
            self.collection_field_comboBox.setCurrentIndex(idx)

        # Populate/restore the Annotations comboBox
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

        # Populate/restore the Read comboBox
        self.populate_read()
        cf = self.prefs.get('read_field_comboBox', '')
        idx = self.read_field_comboBox.findText(cf)
        if idx > -1:
            self.read_field_comboBox.setCurrentIndex(idx)

        # Populate/restore the Reading list comboBox
        self.populate_reading_list()
        cf = self.prefs.get('reading_list_field_comboBox', '')
        idx = self.reading_list_field_comboBox.findText(cf)
        if idx > -1:
            self.reading_list_field_comboBox.setCurrentIndex(idx)

        # Populate/restore the Word count comboBox
        self.populate_word_count()
        cf = self.prefs.get('word_count_field_comboBox', '')
        idx = self.word_count_field_comboBox.findText(cf)
        if idx > -1:
            self.word_count_field_comboBox.setCurrentIndex(idx)

        # Restore general settings
        self.auto_refresh_checkbox.setChecked(self.prefs.get('auto_refresh_at_startup', False))
        self.reading_progress_checkbox.setChecked(self.prefs.get('show_progress_as_percentage', False))
        self.debug_plugin_checkbox.setChecked(self.prefs.get('debug_plugin', False))
        self.debug_libimobiledevice_checkbox.setChecked(self.prefs.get('debug_libimobiledevice', False))

        # Hook changes to diagnostic checkboxes
        self.debug_plugin_checkbox.stateChanged.connect(self.set_restart_required)
        self.debug_libimobiledevice_checkbox.stateChanged.connect(self.set_restart_required)

        # Hook changes to Annotations comboBox
        self.annotations_field_comboBox.currentIndexChanged.connect(
            partial(self.save_combobox_setting, 'annotations_field_comboBox'))

        # Launch the annotated_books_scanner
        field = plugin_prefs.get('annotations_field_lookup', None)
        self.annotated_books_scanner = InventoryAnnotatedBooks(self.gui, field)
        self.connect(self.annotated_books_scanner, self.annotated_books_scanner.signal,
            self.inventory_complete)
        QTimer.singleShot(1, self.start_inventory)

    def configure_appearance(self):
        '''
        '''
        self._log_location()
        appearance_settings = {
                                'appearance_css': default_elements,
                                'appearance_hr_checkbox': False,
                                'appearance_timestamp_format': default_timestamp
                              }

        # Save, hash the original settings
        original_settings = {}
        osh = hashlib.md5()
        for setting in appearance_settings:
            original_settings[setting] = plugin_prefs.get(setting, appearance_settings[setting])
            osh.update(repr(plugin_prefs.get(setting, appearance_settings[setting])))

        # Display the Annotations appearance dialog
        aa = AnnotationsAppearance(self, self.annotations_icon, plugin_prefs)
        cancelled = False
        if aa.exec_():
            # appearance_hr_checkbox and appearance_timestamp_format changed live to prefs during previews
            plugin_prefs.set('appearance_css', aa.elements_table.get_data())
            # Generate a new hash
            nsh = hashlib.md5()
            for setting in appearance_settings:
                nsh.update(repr(plugin_prefs.get(setting, appearance_settings[setting])))
        else:
            for setting in appearance_settings:
                plugin_prefs.set(setting, original_settings[setting])
            nsh = osh

        # If there were changes, and there are existing annotations,
        # and there is an active Annotations field, offer to re-render
        field = plugin_prefs.get("annotations_field_lookup", None)
        if osh.digest() != nsh.digest() and existing_annotations(self.parent, field):
            title = 'Update annotations?'
            msg = '<p>Update existing annotations to new appearance settings?</p>'
            d = MessageBox(MessageBox.QUESTION,
                           title, msg,
                           show_copy_button=False)
            self._log_location("QUESTION: %s" % msg)
            if d.exec_():
                self._log_location("Updating existing annotations to modified appearance")

                # Wait for indexing to complete
                while not self.annotated_books_scanner.isFinished():
                    Application.processEvents()

                move_annotations(self, self.annotated_books_scanner.annotation_map,
                    field, field, window_title="Updating appearance")

    def edit_css(self):
        '''
        '''
        self._log_location()
        from calibre_plugins.marvin_manager.book_status import dialog_resources_path
        klass = os.path.join(dialog_resources_path, 'css_editor.py')
        if os.path.exists(klass):
            sys.path.insert(0, dialog_resources_path)
            this_dc = importlib.import_module('css_editor')
            sys.path.remove(dialog_resources_path)
            dlg = this_dc.CSSEditorDialog(self, 'css_editor')
            dlg.initialize(self)
            dlg.exec_()

    def get_eligible_custom_fields(self, eligible_types=[], is_multiple=None):
        '''
        Discover qualifying custom fields for eligible_types[]
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

    def inventory_complete(self, msg):
        self._log_location(msg)

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

        from calibre_plugins.marvin_manager.book_status import dialog_resources_path

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

                if source == "Annotations":
                    _update_combo_box("annotations_field_comboBox", destination, previous)

                    # Add/update the new destination so save_settings() can find it
                    self.eligible_annotations_fields[destination] = label

                    # Save Date read field manually in case user cancels
                    self.prefs.set('annotations_field_comboBox', destination)
                    self.prefs.set('annotations_field_lookup', label)

                elif source == 'Collections':
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

                elif source == "Progress":
                    _update_combo_box("progress_field_comboBox", destination, previous)

                    # Add/update the new destination so save_settings() can find it
                    self.eligible_progress_fields[destination] = label

                    # Save Progress field manually in case user cancels
                    self.prefs.set('progress_field_comboBox', destination)
                    self.prefs.set('progress_field_lookup', label)

                elif source == "Read":
                    _update_combo_box("read_field_comboBox", destination, previous)

                    # Add/update the new destination so save_settings() can find it
                    self.eligible_read_fields[destination] = label

                    # Save Read field manually in case user cancels
                    self.prefs.set('read_field_comboBox', destination)
                    self.prefs.set('read_field_lookup', label)

                elif source == "Reading list":
                    _update_combo_box("reading_list_field_comboBox", destination, previous)

                    # Add/update the new destination so save_settings() can find it
                    self.eligible_reading_list_fields[destination] = label

                    # Save Reading list field manually in case user cancels
                    self.prefs.set('reading_list_field_comboBox', destination)
                    self.prefs.set('reading_list_field_lookup', label)

                elif source == "Word count":
                    _update_combo_box("word_count_field_comboBox", destination, previous)

                    # Add/update the new destination so save_settings() can find it
                    self.eligible_word_count_fields[destination] = label

                    # Save Word count field manually in case user cancels
                    self.prefs.set('word_count_field_comboBox', destination)
                    self.prefs.set('word_count_field_lookup', label)
        else:
            self._log("ERROR: Can't import from '%s'" % klass)

    def populate_annotations(self):
        #self.eligible_annotations_fields = self.get_eligible_custom_fields(eligible_types=['comments'])
        datatype = self.WIZARD_PROFILES['Annotations']['datatype']
        self.eligible_annotations_fields = self.get_eligible_custom_fields([datatype])
        self.annotations_field_comboBox.addItems([''])
        ecf = sorted(self.eligible_annotations_fields.keys(), key=lambda s: s.lower())
        self.annotations_field_comboBox.addItems(ecf)

    def populate_collections(self):
        #self.eligible_collection_fields = self.get_eligible_custom_fields(['enumeration', 'text'],
        #                                                                          is_multiple=True)
        datatype = self.WIZARD_PROFILES['Collections']['datatype']
        self.eligible_collection_fields = self.get_eligible_custom_fields([datatype],
                                                                          is_multiple=True)
        self.collection_field_comboBox.addItems([''])
        ecf = sorted(self.eligible_collection_fields.keys(), key=lambda s: s.lower())
        self.collection_field_comboBox.addItems(ecf)

    def populate_date_read(self):
        #self.eligible_date_read_fields = self.get_eligible_custom_fields(['datetime'])
        datatype = self.WIZARD_PROFILES['Last read']['datatype']
        self.eligible_date_read_fields = self.get_eligible_custom_fields([datatype])
        self.date_read_field_comboBox.addItems([''])
        ecf = sorted(self.eligible_date_read_fields.keys(), key=lambda s: s.lower())
        self.date_read_field_comboBox.addItems(ecf)

    def populate_progress(self):
        #self.eligible_progress_fields = self.get_eligible_custom_fields(['float'])
        datatype = self.WIZARD_PROFILES['Progress']['datatype']
        self.eligible_progress_fields = self.get_eligible_custom_fields([datatype])
        self.progress_field_comboBox.addItems([''])
        ecf = sorted(self.eligible_progress_fields.keys(), key=lambda s: s.lower())
        self.progress_field_comboBox.addItems(ecf)

    def populate_read(self):
        datatype = self.WIZARD_PROFILES['Read']['datatype']
        self.eligible_read_fields = self.get_eligible_custom_fields([datatype])
        self.read_field_comboBox.addItems([''])
        ecf = sorted(self.eligible_read_fields.keys(), key=lambda s: s.lower())
        self.read_field_comboBox.addItems(ecf)

    def populate_reading_list(self):
        datatype = self.WIZARD_PROFILES['Reading list']['datatype']
        self.eligible_reading_list_fields = self.get_eligible_custom_fields([datatype])
        self.reading_list_field_comboBox.addItems([''])
        ecf = sorted(self.eligible_reading_list_fields.keys(), key=lambda s: s.lower())
        self.reading_list_field_comboBox.addItems(ecf)

    def populate_word_count(self):
        #self.eligible_word_count_fields = self.get_eligible_custom_fields(['int'])
        datatype = self.WIZARD_PROFILES['Word count']['datatype']
        self.eligible_word_count_fields = self.get_eligible_custom_fields([datatype])
        self.word_count_field_comboBox.addItems([''])
        ecf = sorted(self.eligible_word_count_fields.keys(), key=lambda s: s.lower())
        self.word_count_field_comboBox.addItems(ecf)

    def set_restart_required(self, state):
        '''
        Set restart_required flag to show show dialog when closing dialog
        '''
        self.restart_required = True

    def save_combobox_setting(self, cb, index):
        '''
        Apply changes immediately
        '''
        cf = str(getattr(self, cb).currentText())
        self._log_location("%s => %s" % (cb, repr(cf)))

        if cb == 'annotations_field_comboBox':
            self.prefs.set(cb, cf)
            if cf:
                self.prefs.set('annotations_field_lookup', self.eligible_annotations_fields[cf])
            else:
                self.prefs.set('annotations_field_lookup', '')
            self.prefs.commit()

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

        # Save Read field
        cf = str(self.read_field_comboBox.currentText())
        self.prefs.set('read_field_comboBox', cf)
        if cf:
            self.prefs.set('read_field_lookup', self.eligible_read_fields[cf])
        else:
            self.prefs.set('read_field_lookup', '')

        # Save Reading list field
        cf = str(self.reading_list_field_comboBox.currentText())
        self.prefs.set('reading_list_field_comboBox', cf)
        if cf:
            self.prefs.set('reading_list_field_lookup', self.eligible_reading_list_fields[cf])
        else:
            self.prefs.set('reading_list_field_lookup', '')

        # Save Word count field
        cf = str(self.word_count_field_comboBox.currentText())
        self.prefs.set('word_count_field_comboBox', cf)
        if cf:
            self.prefs.set('word_count_field_lookup', self.eligible_word_count_fields[cf])
        else:
            self.prefs.set('word_count_field_lookup', '')

        # Save general settings
        self.prefs.set('auto_refresh_at_startup', self.auto_refresh_checkbox.isChecked())
        self.prefs.set('show_progress_as_percentage', self.reading_progress_checkbox.isChecked())

        # Save debug settings
        self.prefs.set('debug_plugin', self.debug_plugin_checkbox.isChecked())
        self.prefs.set('debug_libimobiledevice', self.debug_libimobiledevice_checkbox.isChecked())

        # If restart needed, inform user
        if self.restart_required:
            do_restart = show_restart_warning('Restart calibre for the changes to be applied.',
                                              parent=self.gui)
            if do_restart:
                self.gui.quit(restart=True)

    def start_inventory(self):
        self._log_location()
        self.annotated_books_scanner.start()

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


class ClickableQLabel(QLabel):

    def __init__(self, parent):
        QLabel.__init__(self, parent)

    def mouseReleaseEvent(self, event):
        self.emit(SIGNAL('clicked()'))


class InventoryAnnotatedBooks(QThread):

    def __init__(self, gui, field, get_date_range=False):
        QThread.__init__(self, gui)
        self.annotation_map = []
        self.cdb = gui.current_db
        self.get_date_range = get_date_range
        self.newest_annotation = 0
        self.oldest_annotation = mktime(datetime.today().timetuple())
        self.field = field
        self.signal = SIGNAL("inventory_complete")

    def run(self):
        self.find_all_annotated_books()
        if self.get_date_range:
            self.get_annotations_date_range()
        self.emit(self.signal, "%d annotated books" % len(self.annotation_map))

    def find_all_annotated_books(self):
        '''
        Find all annotated books in library
        '''
        cids = self.cdb.search_getting_ids('formats:EPUB', '')
        for cid in cids:
            mi = self.cdb.get_metadata(cid, index_is_id=True)
            if self.field == 'Comments':
                if mi.comments:
                    soup = BeautifulSoup(mi.comments)
                else:
                    continue
            else:
                raw = mi.get_user_metadata(self.field, False)
                if hasattr(raw, '#value#') and raw['#value#']:
                    soup = BeautifulSoup(raw['#value#'])
                    if soup.find('div', 'user_annotations') is not None:
                        self.annotation_map.append(mi.id)

    def get_annotations_date_range(self):
        '''
        Find oldest, newest annotation in annotated books
        initial values of self.oldest, self.newest are reversed to allow update comparisons
        if no annotations, restore to correct values
        '''
        annotations_found = False

        for cid in self.annotation_map:
            mi = self.cdb.get_metadata(cid, index_is_id=True)
            if self.field == 'Comments':
                soup = BeautifulSoup(mi.comments)
            else:
                soup = BeautifulSoup(mi.get_user_metadata(self.field, False)['#value#'])

            uas = soup.findAll('div', 'annotation')
            for ua in uas:
                annotations_found = True
                timestamp = float(ua.find('td', 'timestamp')['uts'])
                if timestamp < self.oldest_annotation:
                    self.oldest_annotation = timestamp
                if timestamp > self.newest_annotation:
                    self.newest_annotation = timestamp

        if not annotations_found:
            temp = self.newest_annotation
            self.newest_annotation = self.oldest_annotation
            self.oldest_annotation = temp

# For testing ConfigWidget, run from command line:
# cd ~/Documents/calibredev/Marvin_Manager
# calibre-debug config.py
# Search 'Marvin'
if __name__ == '__main__':
    from PyQt4.Qt import QApplication
    from calibre.gui2.preferences import test_widget
    app = QApplication([])
    test_widget('Advanced', 'Plugins')
