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

from calibre_plugins.marvin_manager.common_utils import (Logger,
    existing_annotations, get_cc_mapping, get_icon, move_annotations, set_cc_mapping)

try:
    from PyQt5.Qt import (Qt, QCheckBox, QComboBox, QFont, QFontMetrics, QFrame,
                          QGridLayout, QGroupBox, QFileDialog, QIcon,
                          QLabel, QLineEdit, QPlainTextEdit, QPushButton,
                          QSizePolicy, QSpacerItem, QThread, QTimer, QToolButton,
                          QVBoxLayout, QWidget,
                          pyqtSignal)
except ImportError:
    from PyQt4.Qt import (Qt, QCheckBox, QComboBox, QFont, QFontMetrics, QFrame,
                          QGridLayout, QGroupBox, QFileDialog, QIcon,
                          QLabel, QLineEdit, QPlainTextEdit, QPushButton,
                          QSizePolicy, QSpacerItem, QThread, QTimer, QToolButton,
                          QVBoxLayout, QWidget,
                          pyqtSignal)

plugin_prefs = JSONConfig('plugins/Marvin XD')


class ConfigWidget(QWidget, Logger):
    '''
    Config dialog for Marvin Manager
    '''

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
        'Locked': {
            'label': 'mm_locked',
            'datatype': 'bool',
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

        self.l = QGridLayout()
        self.setLayout(self.l)
        self.column1_layout = QVBoxLayout()
        self.l.addLayout(self.column1_layout, 0, 0)
        self.column2_layout = QVBoxLayout()
        self.l.addLayout(self.column2_layout, 0, 1)

        # ----------------------------- Column 1 -----------------------------
        # ~~~~~~~~ Create the Custom fields options group box ~~~~~~~~
        self.cfg_custom_fields_gb = QGroupBox(self)
        self.cfg_custom_fields_gb.setTitle('Custom column assignments')
        self.column1_layout.addWidget(self.cfg_custom_fields_gb)

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

        # ++++++++ Locked ++++++++
        self.cfg_locked_label = QLabel("Locked")
        self.cfg_locked_label.setAlignment(Qt.AlignLeft)
        self.cfg_custom_fields_qgl.addWidget(self.cfg_locked_label, current_row, 0)

        self.locked_field_comboBox = QComboBox(self.cfg_custom_fields_gb)
        self.locked_field_comboBox.setObjectName('locked_field_comboBox')
        self.locked_field_comboBox.setToolTip('Select a custom column to store Locked status')
        self.cfg_custom_fields_qgl.addWidget(self.locked_field_comboBox, current_row, 1)

        self.cfg_locked_wizard = QToolButton()
        self.cfg_locked_wizard.setIcon(QIcon(I('wizard.png')))
        self.cfg_locked_wizard.setToolTip("Create a custom column to store Locked status")
        self.cfg_locked_wizard.clicked.connect(partial(self.launch_cc_wizard, 'Locked'))
        self.cfg_custom_fields_qgl.addWidget(self.cfg_locked_wizard, current_row, 2)
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

        self.spacerItem1 = QSpacerItem(20, 20, QSizePolicy.Minimum, QSizePolicy.Expanding)
        self.column1_layout.addItem(self.spacerItem1)

        # ----------------------------- Column 2 -----------------------------
        # ~~~~~~~~ Create the CSS group box ~~~~~~~~
        self.cfg_css_options_gb = QGroupBox(self)
        self.cfg_css_options_gb.setTitle('CSS')
        self.column2_layout.addWidget(self.cfg_css_options_gb)
        self.cfg_css_options_qgl = QGridLayout(self.cfg_css_options_gb)

        current_row = 0

        # ++++++++ Annotations appearance ++++++++
        self.annotations_icon = QIcon(os.path.join(self.resources_path, 'icons', 'annotations_hiliter.png'))
        self.cfg_annotations_appearance_toolbutton = QToolButton()
        self.cfg_annotations_appearance_toolbutton.setIcon(self.annotations_icon)
        self.cfg_annotations_appearance_toolbutton.clicked.connect(self.configure_appearance)
        self.cfg_css_options_qgl.addWidget(self.cfg_annotations_appearance_toolbutton, current_row, 0)
        self.cfg_annotations_label = ClickableQLabel("Book notes, Bookmark notes and Annotations")
        self.cfg_annotations_label.clicked.connect(self.configure_appearance)
        self.cfg_css_options_qgl.addWidget(self.cfg_annotations_label, current_row, 1)
        current_row += 1

        # ++++++++ Injected CSS ++++++++
        self.css_editor_icon = QIcon(I('format-text-heading.png'))
        self.cfg_css_editor_toolbutton = QToolButton()
        self.cfg_css_editor_toolbutton.setIcon(self.css_editor_icon)
        self.cfg_css_editor_toolbutton.clicked.connect(self.edit_css)
        self.cfg_css_options_qgl.addWidget(self.cfg_css_editor_toolbutton, current_row, 0)
        self.cfg_css_editor_label = ClickableQLabel("Articles, Vocabulary")
        self.cfg_css_editor_label.clicked.connect(self.edit_css)
        self.cfg_css_options_qgl.addWidget(self.cfg_css_editor_label, current_row, 1)


        """
        # ~~~~~~~~ Create the Dropbox syncing group box ~~~~~~~~
        self.cfg_dropbox_syncing_gb = QGroupBox(self)
        self.cfg_dropbox_syncing_gb.setTitle('Dropbox')
        self.column2_layout.addWidget(self.cfg_dropbox_syncing_gb)
        self.cfg_dropbox_syncing_qgl = QGridLayout(self.cfg_dropbox_syncing_gb)
        current_row = 0

        # ++++++++ Syncing enabled checkbox ++++++++
        self.dropbox_syncing_checkbox = QCheckBox('Enable Dropbox updates')
        self.dropbox_syncing_checkbox.setObjectName('dropbox_syncing')
        self.dropbox_syncing_checkbox.setToolTip('Refresh custom column content from Marvin metadata')
        self.cfg_dropbox_syncing_qgl.addWidget(self.dropbox_syncing_checkbox,
            current_row, 0, 1, 3)
        current_row += 1

        # ++++++++ Dropbox folder picker ++++++++
        self.dropbox_folder_icon = QIcon(os.path.join(self.resources_path, 'icons', 'dropbox.png'))
        self.cfg_dropbox_folder_toolbutton = QToolButton()
        self.cfg_dropbox_folder_toolbutton.setIcon(self.dropbox_folder_icon)
        self.cfg_dropbox_folder_toolbutton.setToolTip("Specify Dropbox folder location on your computer")
        self.cfg_dropbox_folder_toolbutton.clicked.connect(self.select_dropbox_folder)
        self.cfg_dropbox_syncing_qgl.addWidget(self.cfg_dropbox_folder_toolbutton,
            current_row, 1)

        # ++++++++ Dropbox location lineedit ++++++++
        self.dropbox_location_lineedit = QLineEdit()
        self.dropbox_location_lineedit.setPlaceholderText("Dropbox folder location")
        self.cfg_dropbox_syncing_qgl.addWidget(self.dropbox_location_lineedit,
            current_row, 2)
        """

        # ~~~~~~~~ Create the General options group box ~~~~~~~~
        self.cfg_runtime_options_gb = QGroupBox(self)
        self.cfg_runtime_options_gb.setTitle('General options')
        self.column2_layout.addWidget(self.cfg_runtime_options_gb)
        self.cfg_runtime_options_qvl = QVBoxLayout(self.cfg_runtime_options_gb)

        # ++++++++ Temporary markers: Duplicates ++++++++
        self.duplicate_markers_checkbox = QCheckBox('Apply temporary markers to duplicate books')
        self.duplicate_markers_checkbox.setObjectName('apply_markers_to_duplicates')
        self.duplicate_markers_checkbox.setToolTip('Books with identical content will be flagged in the Library window')
        self.cfg_runtime_options_qvl.addWidget(self.duplicate_markers_checkbox)

        # ++++++++ Temporary markers: Updated ++++++++
        self.updated_markers_checkbox = QCheckBox('Apply temporary markers to books with updated content')
        self.updated_markers_checkbox.setObjectName('apply_markers_to_updated')
        self.updated_markers_checkbox.setToolTip('Books with updated content will be flagged in the Library window')
        self.cfg_runtime_options_qvl.addWidget(self.updated_markers_checkbox)

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
        self.column2_layout.addWidget(self.cfg_debug_options_gb)
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

        self.spacerItem2 = QSpacerItem(20, 20, QSizePolicy.Minimum, QSizePolicy.Expanding)
        self.column2_layout.addItem(self.spacerItem2)

        # ~~~~~~~~ End of construction zone ~~~~~~~~
        self.resize(self.sizeHint())

        # ~~~~~~~~ Populate/restore config options ~~~~~~~~
        #  Annotations comboBox
        self.populate_annotations()
        self.populate_collections()
        self.populate_date_read()
        self.populate_locked()
        self.populate_progress()
        self.populate_read()
        self.populate_reading_list()
        self.populate_word_count()

        """
        # Restore Dropbox settings, hook changes
        dropbox_syncing = self.prefs.get('dropbox_syncing', False)
        self.dropbox_syncing_checkbox.setChecked(dropbox_syncing)
        self.set_dropbox_syncing(dropbox_syncing)
        self.dropbox_syncing_checkbox.clicked.connect(partial(self.set_dropbox_syncing))
        self.dropbox_location_lineedit.setText(self.prefs.get('dropbox_folder', ''))
        """

        # Restore general settings
        self.duplicate_markers_checkbox.setChecked(self.prefs.get('apply_markers_to_duplicates', True))
        self.updated_markers_checkbox.setChecked(self.prefs.get('apply_markers_to_updated', True))
        self.auto_refresh_checkbox.setChecked(self.prefs.get('auto_refresh_at_startup', False))
        self.reading_progress_checkbox.setChecked(self.prefs.get('show_progress_as_percentage', False))

        # Restore debug settings, hook changes
        self.debug_plugin_checkbox.setChecked(self.prefs.get('debug_plugin', False))
        self.debug_plugin_checkbox.stateChanged.connect(self.set_restart_required)
        self.debug_libimobiledevice_checkbox.setChecked(self.prefs.get('debug_libimobiledevice', False))
        self.debug_libimobiledevice_checkbox.stateChanged.connect(self.set_restart_required)

        # Hook changes to Annotations comboBox
#         self.annotations_field_comboBox.currentIndexChanged.connect(
#             partial(self.save_combobox_setting, 'annotations_field_comboBox'))
#        self.connect(self.annotations_field_comboBox,
#                     SIGNAL('currentIndexChanged(const QString &)'),
#                     self.annotations_destination_changed)
        self.annotations_field_comboBox.currentIndexChanged.connect(self.annotations_destination_changed)
        # Launch the annotated_books_scanner
        field = get_cc_mapping('annotations', 'field', None)
        self.annotated_books_scanner = InventoryAnnotatedBooks(self.gui, field)
        self.annotated_books_scanner.signal.connect(self.inventory_complete)
        QTimer.singleShot(1, self.start_inventory)

    def annotations_destination_changed(self, qs_new_destination_name):
        '''
        If the destination field changes, move all existing annotations from old to new
        '''
        self._log_location(str(qs_new_destination_name))
        self._log("self.eligible_annotations_fields: %s" % self.eligible_annotations_fields)

        old_destination_field = get_cc_mapping('annotations', 'field', None)
        old_destination_name = get_cc_mapping('annotations', 'combobox', None)

        self._log("old_destination_field: %s" % old_destination_field)
        self._log("old_destination_name: %s" % old_destination_name)

#        new_destination_name = unicode(qs_new_destination_name)
        # Signnls available have changed. Now receivin an indec rather than a name. Can get the name
        # from the combobox
        new_destination_name = unicode(self.annotations_field_comboBox.currentText())
        self._log("new_destination_name: %s" % repr(new_destination_name))

        if old_destination_name == new_destination_name:
            self._log_location("old_destination_name = new_destination_name, no changes")
            return

        if new_destination_name == '':
            self._log_location("annotations storage disabled")
            set_cc_mapping('annotations', field=None, combobox=new_destination_name)
            return

        new_destination_field = self.eligible_annotations_fields[new_destination_name]

        if existing_annotations(self.parent, old_destination_field):
            command = self.launch_new_destination_dialog(old_destination_name, new_destination_name)

            if command == 'move':
                set_cc_mapping('annotations', field=new_destination_field, combobox=new_destination_name)

                if self.annotated_books_scanner.isRunning():
                    self.annotated_books_scanner.wait()
                move_annotations(self, self.annotated_books_scanner.annotation_map,
                    old_destination_field, new_destination_field)

            elif command == 'change':
                # Keep the updated destination field, but don't move annotations
                pass

            elif command == 'cancel':
                # Restore previous destination
                self.annotations_field_comboBox.blockSignals(True)
                old_index = self.annotations_field_comboBox.findText(old_destination_name)
                self.annotations_field_comboBox.setCurrentIndex(old_index)
                self.annotations_field_comboBox.blockSignals(False)

        else:
            # No existing annotations, just update prefs
            self._log("no existing annotations, updating destination to '{0}'".format(new_destination_name))
            set_cc_mapping('annotations', field=new_destination_field, combobox=new_destination_name)

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
        field = get_cc_mapping('annotations', 'field', None)
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
            cb.blockSignals(True)
            all_items = [str(cb.itemText(i))
                         for i in range(cb.count())]
            if previous and previous in all_items:
                all_items.remove(previous)
            all_items.append(destination)

            cb.clear()
            cb.addItems(sorted(all_items, key=lambda s: s.lower()))

            # Select the new destination in the comboBox
            idx = cb.findText(destination)
            if idx > -1:
                cb.setCurrentIndex(idx)

            cb.blockSignals(False)

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

                    # Save manually in case user cancels
                    set_cc_mapping('annotations', combobox=destination, field=label)

                elif source == 'Collections':
                    _update_combo_box("collection_field_comboBox", destination, previous)

                    # Add/update the new destination so save_settings() can find it
                    self.eligible_collection_fields[destination] = label

                    # Save manually in case user cancels
                    set_cc_mapping('collections', combobox=destination, field=label)

                elif source == 'Last read':
                    _update_combo_box("date_read_field_comboBox", destination, previous)

                    # Add/update the new destination so save_settings() can find it
                    self.eligible_date_read_fields[destination] = label

                    # Save manually in case user cancels
                    set_cc_mapping('date_read', combobox=destination, field=label)

                elif source == 'Locked':
                    _update_combo_box("locked_field_comboBox", destination, previous)

                    # Add/update the new destination so save_settings() can find it
                    self.eligible_locked_fields[destination] = label

                    # Save manually in case user cancels
                    set_cc_mapping('locked', combobox=destination, field=label)

                elif source == "Progress":
                    _update_combo_box("progress_field_comboBox", destination, previous)

                    # Add/update the new destination so save_settings() can find it
                    self.eligible_progress_fields[destination] = label

                    # Save manually in case user cancels
                    set_cc_mapping('progress', combobox=destination, field=label)

                elif source == "Read":
                    _update_combo_box("read_field_comboBox", destination, previous)

                    # Add/update the new destination so save_settings() can find it
                    self.eligible_read_fields[destination] = label

                    # Save manually in case user cancels
                    set_cc_mapping('read', combobox=destination, field=label)

                elif source == "Reading list":
                    _update_combo_box("reading_list_field_comboBox", destination, previous)

                    # Add/update the new destination so save_settings() can find it
                    self.eligible_reading_list_fields[destination] = label

                    # Save manually in case user cancels
                    set_cc_mapping('reading_list', combobox=destination, field=label)

                elif source == "Word count":
                    _update_combo_box("word_count_field_comboBox", destination, previous)

                    # Add/update the new destination so save_settings() can find it
                    self.eligible_word_count_fields[destination] = label

                    # Save manually in case user cancels
                    set_cc_mapping('word_count', combobox=destination, field=label)
        else:
            self._log("ERROR: Can't import from '%s'" % klass)

    def launch_new_destination_dialog(self, old, new):
        '''
        Return 'move', 'change' or 'cancel'
        '''
        from calibre_plugins.marvin_manager.book_status import dialog_resources_path
        self._log_location()
        klass = os.path.join(dialog_resources_path, 'new_destination.py')
        if os.path.exists(klass):
            self._log("importing new destination dialog from '%s'" % klass)
            sys.path.insert(0, dialog_resources_path)
            this_dc = importlib.import_module('new_destination')
            sys.path.remove(dialog_resources_path)
            dlg = this_dc.NewDestinationDialog(self, old, new)
            dlg.exec_()
            return dlg.command

    def populate_annotations(self):
        datatype = self.WIZARD_PROFILES['Annotations']['datatype']
        self.eligible_annotations_fields = self.get_eligible_custom_fields([datatype])
        self.annotations_field_comboBox.addItems([''])
        ecf = sorted(self.eligible_annotations_fields.keys(), key=lambda s: s.lower())
        self.annotations_field_comboBox.addItems(ecf)

        # Retrieve stored value
        existing = get_cc_mapping('annotations', 'combobox')
        if existing:
            ci = self.annotations_field_comboBox.findText(existing)
            self.annotations_field_comboBox.setCurrentIndex(ci)

    def populate_collections(self):
        datatype = self.WIZARD_PROFILES['Collections']['datatype']
        self.eligible_collection_fields = self.get_eligible_custom_fields([datatype],
                                                                          is_multiple=True)
        self.collection_field_comboBox.addItems([''])
        ecf = sorted(self.eligible_collection_fields.keys(), key=lambda s: s.lower())
        self.collection_field_comboBox.addItems(ecf)

        # Retrieve stored value
        existing = get_cc_mapping('collections', 'combobox')
        if existing:
            ci = self.collection_field_comboBox.findText(existing)
            self.collection_field_comboBox.setCurrentIndex(ci)

    def populate_date_read(self):
        #self.eligible_date_read_fields = self.get_eligible_custom_fields(['datetime'])
        datatype = self.WIZARD_PROFILES['Last read']['datatype']
        self.eligible_date_read_fields = self.get_eligible_custom_fields([datatype])
        self.date_read_field_comboBox.addItems([''])
        ecf = sorted(self.eligible_date_read_fields.keys(), key=lambda s: s.lower())
        self.date_read_field_comboBox.addItems(ecf)

        # Retrieve stored value
        existing = get_cc_mapping('date_read', 'combobox')
        if existing:
            ci = self.date_read_field_comboBox.findText(existing)
            self.date_read_field_comboBox.setCurrentIndex(ci)

    def populate_locked(self):
        datatype = self.WIZARD_PROFILES['Locked']['datatype']
        self.eligible_locked_fields = self.get_eligible_custom_fields([datatype])
        self.locked_field_comboBox.addItems([''])
        ecf = sorted(self.eligible_locked_fields.keys(), key=lambda s: s.lower())
        self.locked_field_comboBox.addItems(ecf)

        # Retrieve stored value
        existing = get_cc_mapping('locked', 'combobox')
        if existing:
            ci = self.locked_field_comboBox.findText(existing)
            self.locked_field_comboBox.setCurrentIndex(ci)

    def populate_progress(self):
        #self.eligible_progress_fields = self.get_eligible_custom_fields(['float'])
        datatype = self.WIZARD_PROFILES['Progress']['datatype']
        self.eligible_progress_fields = self.get_eligible_custom_fields([datatype])
        self.progress_field_comboBox.addItems([''])
        ecf = sorted(self.eligible_progress_fields.keys(), key=lambda s: s.lower())
        self.progress_field_comboBox.addItems(ecf)

        # Retrieve stored value
        existing = get_cc_mapping('progress', 'combobox')
        if existing:
            ci = self.progress_field_comboBox.findText(existing)
            self.progress_field_comboBox.setCurrentIndex(ci)

    def populate_read(self):
        datatype = self.WIZARD_PROFILES['Read']['datatype']
        self.eligible_read_fields = self.get_eligible_custom_fields([datatype])
        self.read_field_comboBox.addItems([''])
        ecf = sorted(self.eligible_read_fields.keys(), key=lambda s: s.lower())
        self.read_field_comboBox.addItems(ecf)

        # Retrieve stored value
        existing = get_cc_mapping('read', 'combobox')
        if existing:
            ci = self.read_field_comboBox.findText(existing)
            self.read_field_comboBox.setCurrentIndex(ci)

    def populate_reading_list(self):
        datatype = self.WIZARD_PROFILES['Reading list']['datatype']
        self.eligible_reading_list_fields = self.get_eligible_custom_fields([datatype])
        self.reading_list_field_comboBox.addItems([''])
        ecf = sorted(self.eligible_reading_list_fields.keys(), key=lambda s: s.lower())
        self.reading_list_field_comboBox.addItems(ecf)

        # Retrieve stored value
        existing = get_cc_mapping('reading_list', 'combobox')
        if existing:
            ci = self.reading_list_field_comboBox.findText(existing)
            self.reading_list_field_comboBox.setCurrentIndex(ci)

    def populate_word_count(self):
        #self.eligible_word_count_fields = self.get_eligible_custom_fields(['int'])
        datatype = self.WIZARD_PROFILES['Word count']['datatype']
        self.eligible_word_count_fields = self.get_eligible_custom_fields([datatype])
        self.word_count_field_comboBox.addItems([''])
        ecf = sorted(self.eligible_word_count_fields.keys(), key=lambda s: s.lower())
        self.word_count_field_comboBox.addItems(ecf)

        # Retrieve stored value
        existing = get_cc_mapping('word_count', 'combobox')
        if existing:
            ci = self.word_count_field_comboBox.findText(existing)
            self.word_count_field_comboBox.setCurrentIndex(ci)

    """
    def select_dropbox_folder(self):
        '''
        '''
        self._log_location()
        dropbox_location = QFileDialog.getExistingDirectory(
            self,
            "Dropbox folder",
            os.path.expanduser("~"),
            QFileDialog.ShowDirsOnly | QFileDialog.DontResolveSymlinks)
        self.dropbox_location_lineedit.setText(unicode(dropbox_location))

    def set_dropbox_syncing(self, state):
        '''
        Called when checkbox changes state, or when restoring state
        Set enabled state of Dropbox folder picker to match
        '''
        self.cfg_dropbox_folder_toolbutton.setEnabled(state)
        self.dropbox_location_lineedit.setEnabled(state)
    """

    def set_restart_required(self, state):
        '''
        Set restart_required flag to show show dialog when closing dialog
        '''
        self.restart_required = True

    """
    def save_combobox_setting(self, cb, index):
        '''
        Apply changes immediately
        '''
        cf = str(getattr(self, cb).currentText())
        self._log_location("%s => %s" % (cb, repr(cf)))

        if cb == 'annotations_field_comboBox':
            field = None
            if cf:
                field = self.eligible_annotations_fields[cf]
            set_cc_mapping('annotations', combobox=cf, field=field)
    """

    def save_settings(self):
        self._log_location()

        # Annotations field
        cf = unicode(self.annotations_field_comboBox.currentText())
        field = None
        if cf:
            field = self.eligible_annotations_fields[cf]
        set_cc_mapping('annotations', combobox=cf, field=field)

        # Collections field
        cf = unicode(self.collection_field_comboBox.currentText())
        field = None
        if cf:
            field = self.eligible_collection_fields[cf]
        set_cc_mapping('collections', combobox=cf, field=field)

        # Date read field
        cf = unicode(self.date_read_field_comboBox.currentText())
        field = None
        if cf:
            field = self.eligible_date_read_fields[cf]
        set_cc_mapping('date_read', combobox=cf, field=field)

        # Locked field
        cf = unicode(self.locked_field_comboBox.currentText())
        field = None
        if cf:
            field = self.eligible_locked_fields[cf]
        set_cc_mapping('locked', combobox=cf, field=field)

        # Progress field
        cf = unicode(self.progress_field_comboBox.currentText())
        field = None
        if cf:
            field = self.eligible_progress_fields[cf]
        set_cc_mapping('progress', combobox=cf, field=field)

        # Read field
        cf = unicode(self.read_field_comboBox.currentText())
        field = None
        if cf:
            field = self.eligible_read_fields[cf]
        set_cc_mapping('read', combobox=cf, field=field)

        # Reading list field
        cf = unicode(self.reading_list_field_comboBox.currentText())
        field = None
        if cf:
            field = self.eligible_reading_list_fields[cf]
        set_cc_mapping('reading_list', combobox=cf, field=field)

        # Word count field
        cf = unicode(self.word_count_field_comboBox.currentText())
        field = None
        if cf:
            field = self.eligible_word_count_fields[cf]
        set_cc_mapping('word_count', combobox=cf, field=field)

        '''
        # Save Dropbox settings
        self.prefs.set('dropbox_syncing', self.dropbox_syncing_checkbox.isChecked())
        self.prefs.set('dropbox_folder', unicode(self.dropbox_location_lineedit.text()))
        '''

        # Save general settings
        self.prefs.set('apply_markers_to_duplicates', self.duplicate_markers_checkbox.isChecked())
        self.prefs.set('apply_markers_to_updated', self.updated_markers_checkbox.isChecked())
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


class ClickableQLabel(QLabel):

    clicked = pyqtSignal(object)
    def __init__(self, parent):
        QLabel.__init__(self, parent)

    def mouseReleaseEvent(self, event):
        self.clicked.emit('clicked')


class InventoryAnnotatedBooks(QThread, Logger):

    signal = pyqtSignal(object, object)#SIGNAL()

    def __init__(self, gui, field, get_date_range=False):
        QThread.__init__(self, gui)
        self.annotation_map = []
        self.cdb = gui.current_db
        self.get_date_range = get_date_range
        self.newest_annotation = 0
        self.oldest_annotation = mktime(datetime.today().timetuple())
        self.field = field

    def run(self):
        if self.field is not None:
            self.find_all_annotated_books()
        else:
            self._log_location("No annotations field specified")
        if self.annotation_map and self.get_date_range:
            self.get_annotations_date_range()
        self.signal.emit("inventory_complete", "{0} annotated books".format(len(self.annotation_map)))

    def find_all_annotated_books(self):
        '''
        Find all annotated books in library
        '''
        self._log_location("field: {0}".format(self.field))
        cids = self.cdb.search_getting_ids('formats:EPUB', '')
        for cid in cids:
            mi = self.cdb.get_metadata(cid, index_is_id=True)
            raw = mi.get_user_metadata(self.field, False)
            if raw['#value#'] is not None:
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
    try:
        from PyQt5.Qt import QApplication
    except ImportError:
        from PyQt4.Qt import QApplication

    from calibre.gui2.preferences import test_widget
    app = QApplication([])
    test_widget('Advanced', 'Plugins')
