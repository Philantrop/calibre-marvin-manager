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
from calibre.utils.icu import sort_key
from calibre.utils.magick.draw import add_borders_to_image, thumbnail

from calibre_plugins.marvin_manager.book_status import dialog_resources_path
from calibre_plugins.marvin_manager.common_utils import SizePersistedDialog

from PyQt4.Qt import (Qt, QAbstractItemModel, QAbstractListModel, QColor,
                      QDialog, QDialogButtonBox, QIcon,
                      QModelIndex, QPalette, QPixmap, QSize, QSizePolicy, QVariant,
                      pyqtSignal)

# Import Ui_Form from form generated dynamically during initialization
if True:
    sys.path.insert(0, dialog_resources_path)
    from view_collections_ui import Ui_Dialog
    sys.path.remove(dialog_resources_path)

class MyListModel(QAbstractListModel):
    # http://www.saltycrane.com/blog/2008/01/pyqt-43-simple-qabstractlistmodel/

    def __init__(self, datain, parent=None, *args):
        '''
        datain: a list where each item is a row
        '''
        QAbstractItemModel.__init__(self, parent, *args)
        self.listdata = datain

    def rowCount(self, parent=QModelIndex()):
        return len(self.listdata)

    def data(self, index, role):
        if index.isValid() and role == Qt.DisplayRole:
            return QVariant(self.listdata[index.row()])
        else:
            return QVariant()

class CollectionsViewerDialog(SizePersistedDialog, Ui_Dialog):
    LOCATION_TEMPLATE = "{cls}:{func}({arg1}) {arg2}"

    marvin_device_status_changed = pyqtSignal(str)

    def accept(self):
        self._log_location()
        self.updated_calibre_collections = self._get_calibre_collections()
        self.updated_marvin_collections = self._get_calibre_collections()
        self.results = {
                        'command': self.stored_command,
                        'initial_calibre_collections': self.initial_calibre_collections,
                        'initial_marvin_collections': self.initial_marvin_collections,
                        'updated_calibre_collections': self.updated_calibre_collections,
                        'updated_marvin_collections': self.updated_marvin_collections
                       }

        super(CollectionsViewerDialog, self).accept()

    def close(self):
        self._log_location()
        super(CollectionsViewerDialog, self).close()

    def dispatch_button_click(self, button):
        '''
        BUTTON_ROLES = ['AcceptRole', 'RejectRole', 'DestructiveRole', 'ActionRole',
                        'HelpRole', 'YesRole', 'NoRole', 'ApplyRole', 'ResetRole']
        '''
        self._log_location()
        if self.bb.buttonRole(button) == QDialogButtonBox.AcceptRole:
            self._log("AcceptRole")
            self._log("initial calibre collections: %s" % self.initial_calibre_collections)
            self._log("initial marvin collections: %s" % self.initial_marvin_collections)

            self.updated_calibre_collections = self._get_calibre_collections()
            self.updated_marvin_collections = self._get_marvin_collections()
            self._log("updated calibre collections: %s" % self.updated_calibre_collections)
            self._log("updated marvin collections: %s" % self.updated_marvin_collections)

            self.accept()

        elif self.bb.buttonRole(button) == QDialogButtonBox.ActionRole:
            pass

        elif self.bb.buttonRole(button) == QDialogButtonBox.RejectRole:
            self._log("RejectRole")
            self.stored_command = None
            self.close()

    def esc(self, *args):
        self._log_location()
        self._clear_selected_rows()

    def export_to_marvin(self):
        self._log_location()

    def import_from_marvin(self):
        self._log_location()

    def initialize(self, parent, book_title, calibre_collections, marvin_collections, connected_device):
        '''
        __init__ is called on SizePersistedDialog()
        if calibre_collections is None, the book does not exist in calibre library
        '''
        self.setupUi(self)
        self.calibre_collections = calibre_collections
        self.calibre_selection = None
        self.marvin_collections = marvin_collections
        self.marvin_selection = None
        self.opts = parent.opts
        self.parent = parent
        self.stored_command = None
        self.verbose = parent.verbose

        self._log_location(book_title)
        self.setWindowTitle(book_title)

        # Subscribe to Marvin driver change events
        connected_device.marvin_device_signals.reader_app_status_changed.connect(
            self.marvin_status_changed)

        # ~~~~~~~~ Export to Marvin button ~~~~~~~~
        self.export_to_marvin_tb.setIcon(QIcon(os.path.join(self.opts.resources_path,
                                                   'icons',
                                                   'from_calibre.png')))
        self.export_to_marvin_tb.setToolTip("Export collection assignments to Marvin")
        self.export_to_marvin_tb.clicked.connect(self._export_to_marvin)

        # ~~~~~~~~ Import from Marvin button ~~~~~~~~
        self.import_from_marvin_tb.setIcon(QIcon(os.path.join(self.opts.resources_path,
                                                   'icons',
                                                   'from_marvin.png')))
        self.import_from_marvin_tb.setToolTip("Import collection assignments from Marvin")
        self.import_from_marvin_tb.clicked.connect(self._import_from_marvin)

        # ~~~~~~~~ Merge collections button ~~~~~~~~
        self.merge_collections_tb.setIcon(QIcon(os.path.join(self.opts.resources_path,
                                                          'icons',
                                                          'sync_collections.png')))
        self.merge_collections_tb.setToolTip("Merge collection assignments")
        self.merge_collections_tb.clicked.connect(self._merge_collections)

        # ~~~~~~~~Remove collection assignment button ~~~~~~~~
        self.remove_assignment_tb.setIcon(QIcon(I('trash.png')))
        self.remove_assignment_tb.setToolTip("Remove collection assignment")
        self.remove_assignment_tb.clicked.connect(self._remove_collection_assignment)

        # ~~~~~~~~ Clear all collections button ~~~~~~~~
        self.clear_all_collections_tb.setIcon(QIcon(os.path.join(self.opts.resources_path,
                                                          'icons',
                                                          'clear_all.png')))
        self.clear_all_collections_tb.setToolTip("Remove all collection assignments")
        self.clear_all_collections_tb.clicked.connect(self._clear_all_collections)

        # Populate collection models
        self._initialize_collections()

        if self.calibre_collections:
            # Save initial state
            self.initial_calibre_collections = list(self._get_calibre_collections())
            self.initial_marvin_collections = list(self._get_marvin_collections())
            # Remind the user of calibre's custom column, disable buttons if no calibre field
            calibre_cf = self.prefs.get('collection_field_comboBox', '')
            if calibre_cf:
                self.calibre_gb.setTitle("Calibre collections (%s)" % calibre_cf)
            else:
                self.calibre_gb.setTitle("Calibre (no collections field)")
                self.calibre_gb.setEnabled(False)
                # Disable import/export/sync
                self.export_to_marvin_tb.setEnabled(False)
                self.import_from_marvin_tb.setEnabled(False)
                self.merge_collections_tb.setEnabled(False)

            # If collections already equal, disable import/export/merge
            if self.calibre_collections == self.marvin_collections:
                self.export_to_marvin_tb.setEnabled(False)
                self.import_from_marvin_tb.setEnabled(False)
                self.merge_collections_tb.setEnabled(False)
        else:
            # Save initial state
            self.initial_marvin_collections = list(self._get_marvin_collections())
            # Hide the calibre panel, disable tool buttons
            self.calibre_gb.setVisible(False)
            self.export_to_marvin_tb.setEnabled(False)
            self.import_from_marvin_tb.setEnabled(False)
            self.merge_collections_tb.setEnabled(False)

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
        Save the requested operation
        '''
        self._log_location(command)
        self.stored_command = command
        self.close()

    # ~~~~~~~~ Helpers ~~~~~~~~
    def _clear_calibre_selection(self):
        '''
        '''
        self._log_location()
        self.calibre_lw.clearSelection()

    def _clear_marvin_selection(self):
        '''
        '''
        self._log_location()
        self.marvin_lw.clearSelection()

    def _clear_selected_rows(self):
        '''
        Clear any active selections
        '''
        self._log_location()
        self._clear_calibre_selection()
        self._clear_marvin_selection()

    def _clear_all_collections(self):
        '''
        '''
        self._log_location()
        self.stored_command = 'clear_all_collections'

        # Delete calibre collection assignments
        rows_to_delete = len(self.calibre_lw.model().listdata)
        for row in range(rows_to_delete - 1, -1, -1):
            self.calibre_lw.model().beginRemoveRows(QModelIndex(), row, row)
            del self.calibre_lw.model().listdata[row]
            self.calibre_lw.model().endRemoveRows()

        # Delete Marvin collection assignments
        rows_to_delete = len(self.marvin_lw.model().listdata)
        for row in range(rows_to_delete - 1, -1, -1):
            self.marvin_lw.model().beginRemoveRows(QModelIndex(), row, row)
            del self.marvin_lw.model().listdata[row]
            self.marvin_lw.model().endRemoveRows()

    def _export_to_marvin(self):
        '''
        '''
        self._log_location()
        scr = self._selected_calibre_rows()
        smr = self._selected_marvin_rows()
        self._log("scr: %s smr: %s" % (scr, smr))
        self.stored_command = 'export_to_marvin'

    def _import_from_marvin(self):
        '''
        '''
        self._log_location()
        scr = self._selected_calibre_rows()
        smr = self._selected_marvin_rows()
        self._log("scr: %s smr: %s" % (scr, smr))
        self.stored_command = 'import_from_marvin'

    def _get_calibre_collections(self):
        '''

        '''
        self._log_location()
        return self.calibre_lw.model().listdata

    def _get_marvin_collections(self):
        '''

        '''
        self._log_location()
        return self.marvin_lw.model().listdata

    def _initialize_collections(self):
        '''
        Populate the data model with current collection assignments
        Hook click, doubleClick events
        '''
        self._log_location()

        # Set the bg color of the description text fields to the dialog bg color
        if False:
            bgcolor = self.palette().color(QPalette.Background)
            palette = QPalette()
            palette.setColor(QPalette.Base, bgcolor)
            self.calibre_lw.setPalette(palette)
            self.marvin_lw.setPalette(palette)

        self.calibre_lw.setModel(MyListModel(self.calibre_collections))
        self.calibre_lw.setAlternatingRowColors(True)
        self.marvin_lw.setModel(MyListModel(self.marvin_collections))
        self.marvin_lw.setAlternatingRowColors(True)

        # Capture click events to clear selections in opposite list
        self.calibre_lw.clicked.connect(self._clear_marvin_selection)
        self.calibre_lw.doubleClicked.connect(self._clear_marvin_selection)
        self.marvin_lw.clicked.connect(self._clear_calibre_selection)
        self.marvin_lw.doubleClicked.connect(self._clear_calibre_selection)

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

    def _merge_collections(self):
        '''
        '''
        self._log_location()
        scr = self._selected_calibre_rows()
        smr = self._selected_marvin_rows()
        self._log("scr: %s smr: %s" % (scr, smr))
        self.stored_command = 'merge_collections'

    def _remove_collection_assignment(self):
        '''
        Only one selection can be active
        '''
        self._log_location()
        if self._selected_calibre_rows():
            row = self._selected_calibre_rows()[0]
            self.calibre_lw.model().beginRemoveRows(QModelIndex(), row, row)
            del self.calibre_lw.model().listdata[row]
            self.calibre_lw.model().endRemoveRows()
        elif self._selected_marvin_rows():
            row = self._selected_marvin_rows()[0]
            self.marvin_lw.model().beginRemoveRows(QModelIndex(), row, row)
            del self.marvin_lw.model().listdata[row]
            self.marvin_lw.model().endRemoveRows()

    def _selected_calibre_rows(self):
        '''
        Return a list of selected calibre rows
        '''
        srs = self.calibre_lw.selectionModel().selectedRows()
        return [sr.row() for sr in srs]

    def _selected_marvin_rows(self):
        '''
        Return a list of selected Marvin rows
        '''
        srs = self.marvin_lw.selectionModel().selectedRows()
        return [sr.row() for sr in srs]

