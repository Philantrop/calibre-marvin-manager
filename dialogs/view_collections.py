#!/usr/bin/env python
# coding: utf-8

from __future__ import (unicode_literals, division, absolute_import,
                        print_function)

__license__ = 'GPL v3'
__copyright__ = '2010, Gregory Riker'
__docformat__ = 'restructuredtext en'

import os, sys

from calibre.devices.usbms.driver import debug_print
from calibre.gui2 import error_dialog
from calibre.gui2.dialogs.device_category_editor import ListWidgetItem
from calibre.gui2.dialogs.message_box import MessageBox
from calibre.utils.icu import sort_key

from calibre_plugins.marvin_manager.book_status import dialog_resources_path
from calibre_plugins.marvin_manager.common_utils import Logger, SizePersistedDialog

from PyQt4.Qt import (Qt, QDialogButtonBox, QIcon, QPalette,
                      pyqtSignal)

# Import Ui_Form from form generated dynamically during initialization
if True:
    sys.path.insert(0, dialog_resources_path)
    from view_collections_ui import Ui_Dialog
    sys.path.remove(dialog_resources_path)


ADD_NEW_COLLECTION_ENABLED = True
RENAMING_ENABLED = False

class CollectionsViewerDialog(SizePersistedDialog, Ui_Dialog, Logger):
    marvin_device_status_changed = pyqtSignal(str)

    def accept(self):
        self._log_location()
        if self.calibre_collections is None:
            self.updated_calibre_collections = None
        else:
            self.updated_calibre_collections = self._get_calibre_collections()
        self.updated_marvin_collections = self._get_marvin_collections()
        self.results = {
            'updated_calibre_collections': self.updated_calibre_collections,
            'updated_marvin_collections': self.updated_marvin_collections
        }

        super(CollectionsViewerDialog, self).accept()

    def add_collection_assignment(self):
        '''
        Always add to Marvin, user can sync if they want
        '''
        self._log_location()
        ma = 'new collection assignment'
        self.marvin_lw.addItem(ma)
        item = self.marvin_lw.item(self.marvin_lw.count() - 1)
        item.setFlags(item.flags() | Qt.ItemIsEditable)
        self.marvin_lw.editItem(item)

    def close(self):
        self._log_location()
        super(CollectionsViewerDialog, self).close()

    def dispatch_button_click(self, button):
        '''
        BUTTON_ROLES = ['AcceptRole', 'RejectRole', 'DestructiveRole', 'ActionRole',
                        'HelpRole', 'YesRole', 'NoRole', 'ApplyRole', 'ResetRole']
        '''
        self._log_location(self.bb.buttonRole(button))
        if self.bb.buttonRole(button) == QDialogButtonBox.AcceptRole:
            #self._log("AcceptRole")
            self.updated_calibre_collections = self._get_calibre_collections()
            self.updated_marvin_collections = self._get_marvin_collections()
            self.accept()

        elif self.bb.buttonRole(button) == QDialogButtonBox.ActionRole:
            pass

        elif self.bb.buttonRole(button) == QDialogButtonBox.RejectRole:
            #self._log("RejectRole")
            self.updated_calibre_collections = self.calibre_collections
            self.updated_marvin_collections = self.marvin_collections
            self.close()

    def esc(self, *args):
        self._log_location()
        self._clear_selected_rows()

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
        self.setWindowTitle("Collection assignments")

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

        if ADD_NEW_COLLECTION_ENABLED:
            self.add_collection_tb.setIcon(QIcon(I('plus.png')))
            self.add_collection_tb.setToolTip("Add a collection assignment")
            self.add_collection_tb.clicked.connect(self.add_collection_assignment)
        else:
            self.add_collection_tb.setVisible(False)

        # ~~~~~~~~Remove collection assignment button ~~~~~~~~
        self.remove_assignment_tb.setIcon(QIcon(I('trash.png')))
        self.remove_assignment_tb.setToolTip("Remove a collection assignment")
        self.remove_assignment_tb.clicked.connect(self._remove_collection_assignment)

        if RENAMING_ENABLED:
            # ~~~~~~~~ Rename collection button ~~~~~~~~
            self.rename_collection_tb.setIcon(QIcon(I('edit_input.png')))
            self.rename_collection_tb.setToolTip("Rename collection")
            self.rename_collection_tb.clicked.connect(self._rename_collection)
        else:
            self.rename_collection_tb.setVisible(False)

        # ~~~~~~~~ Clear all collections button ~~~~~~~~
        self.remove_all_assignments_tb.setIcon(QIcon(os.path.join(self.opts.resources_path,
                                                     'icons',
                                                     'remove_all_collections.png')))
        self.remove_all_assignments_tb.setToolTip("Remove all collection assignments")
        self.remove_all_assignments_tb.clicked.connect(self._remove_all_assignments)

        # Populate collection models
        self._initialize_collections()

        # Save initial Marvin state
        self.initial_marvin_collections = list(self._get_marvin_collections())
        self.marvin_gb.setToolTip("Collections assigned in Marvin")

        if self.calibre_collections is not None:
            # Save initial state
            self.initial_calibre_collections = list(self._get_calibre_collections())
            # Remind the user of calibre's custom column, disable buttons if no calibre field
            calibre_cf = self.prefs.get('collection_field_comboBox', '')
            if calibre_cf:
                self.calibre_gb.setTitle("calibre ('%s')" % calibre_cf)
                self.calibre_gb.setToolTip("Collection assignments from '%s'" % calibre_cf)
            else:
                self.calibre_gb.setTitle("calibre")
                self.calibre_gb.setToolTip("No custom column selected for collection assignments")
                self.calibre_gb.setEnabled(False)
                # Disable import/export/sync
                self.export_to_marvin_tb.setEnabled(False)
                self.import_from_marvin_tb.setEnabled(False)
                self.merge_collections_tb.setEnabled(False)

            if False:
                # If collections already equal, disable import/export/merge
                if self.calibre_collections == self.marvin_collections:
                    self.export_to_marvin_tb.setEnabled(False)
                    self.import_from_marvin_tb.setEnabled(False)
                    self.merge_collections_tb.setEnabled(False)

        else:
            # No cid, this book is Marvin only
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

    def _export_to_marvin(self):
        '''
        Copy calibre assignments to Marvin
        '''
        self._log_location()
        self.marvin_lw.clear()
        for i in range(self.calibre_lw.count()):
            citem = self.calibre_lw.item(i).text()
            item = ListWidgetItem(citem)
            item.setData(Qt.UserRole, citem)
            item.setFlags(item.flags() | Qt.ItemIsEditable)
            self.marvin_lw.addItem(item)

    def _import_from_marvin(self):
        '''
        Copy Marvin collections to calibre
        '''
        self._log_location()
        self.calibre_lw.clear()
        for i in range(self.marvin_lw.count()):
            mitem = self.marvin_lw.item(i).text()
            item = ListWidgetItem(mitem)
            item.setData(Qt.UserRole, mitem)
            item.setFlags(item.flags() | Qt.ItemIsEditable)
            self.calibre_lw.addItem(item)

    def _get_calibre_collections(self):
        '''
        Return widget items as list
        '''
        cc = []
        for i in range(self.calibre_lw.count()):
            cc.append(unicode(self.calibre_lw.item(i).text()))
        return cc

    def _get_marvin_collections(self):
        '''
        Return widget items as list
        '''
        mc = []
        for i in range(self.marvin_lw.count()):
            mc.append(unicode(self.marvin_lw.item(i).text()))
        return mc

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

        if self.calibre_collections is not None:
            for ca in self.calibre_collections:
                item = ListWidgetItem(ca)
                item.setData(Qt.UserRole, ca)
                if RENAMING_ENABLED:
                    item.setFlags(item.flags() | Qt.ItemIsEditable)
                self.calibre_lw.addItem(item)

        for ma in self.marvin_collections:
            item = ListWidgetItem(ma)
            item.setData(Qt.UserRole, ma)
            if RENAMING_ENABLED:
                item.setFlags(item.flags() | Qt.ItemIsEditable)
            self.marvin_lw.addItem(item)

        # Capture click events to clear selections in opposite list
        self.calibre_lw.clicked.connect(self._clear_marvin_selection)
        self.marvin_lw.clicked.connect(self._clear_calibre_selection)

        # Hook double-click events
        if RENAMING_ENABLED:
            self.calibre_lw.doubleClicked.connect(self.rename_calibre_tag)
            self.marvin_lw.doubleClicked.connect(self.rename_marvin_tag)

        # Enable sorting
        if self.calibre_collections is not None:
            self.calibre_lw.setSortingEnabled(True)
        self.marvin_lw.setSortingEnabled(True)

    def _merge_collections(self):
        '''
        '''
        self._log_location()

        # Merge the two collection lists without dupes
        cl = set(self._get_calibre_collections())
        ml = set(self._get_marvin_collections())
        deltas = ml - cl
        merged_collections = sorted(self._get_calibre_collections() + list(deltas), key=sort_key)

        # Clear both
        self.calibre_lw.clear()
        self.marvin_lw.clear()

        # Assign to both
        for ca in merged_collections:
            item = ListWidgetItem(ca)
            item.setData(Qt.UserRole, ca)
            item.setFlags(item.flags() | Qt.ItemIsEditable)
            self.calibre_lw.addItem(item)

            item = ListWidgetItem(ca)
            item.setData(Qt.UserRole, ca)
            item.setFlags(item.flags() | Qt.ItemIsEditable)
            self.marvin_lw.addItem(item)

    def _remove_all_assignments(self):
        '''
        '''
        self._log_location()
        self.stored_command = 'clear_all_collections'

        # Confirm
        title = "Are you sure?"
        msg = ("<p>Delete all collection assignments from calibre and Marvin?</p>")
        d = MessageBox(MessageBox.QUESTION, title, msg,
                       show_copy_button=False)
        if d.exec_():
            self.calibre_lw.clear()
            self.marvin_lw.clear()

    def _remove_collection_assignment(self):
        '''
        Only one panel can have active selection
        '''
        def _remove_assignments(deletes, list_widget):
            row = list_widget.row(deletes[0])

            for item in deletes:
                list_widget.takeItem(list_widget.row(item))
            if row >= list_widget.count():
                row = list_widget.count() - 1
            if row >= 0:
                list_widget.scrollToItem(list_widget.item(row))

        self._log_location()

        if self.calibre_lw.selectedItems():
            deletes = self.calibre_lw.selectedItems()
            _remove_assignments(deletes, self.calibre_lw)

        elif self.marvin_lw.selectedItems():
            deletes = self.marvin_lw.selectedItems()
            _remove_assignments(deletes, self.marvin_lw)

        else:
            title = "No collection selected"
            msg = ("<p>Select a collection assignment to remove.</p>")
            MessageBox(MessageBox.INFO, title, msg,
                       show_copy_button=False).exec_()

    def _rename_collection(self):
        '''
        Only one panel can have active selection
        '''
        self._log_location()
        if self.calibre_lw.selectedItems():
            self.rename_calibre_tag()
        elif self.marvin_lw.selectedItems():
            self.rename_marvin_tag()
        else:
            title = "No collection selected"
            msg = ("<p>Select a collection to rename.</p>")
            MessageBox(MessageBox.INFO, title, msg,
                       show_copy_button=False).exec_()

    def rename_calibre_tag(self):
        item = self.calibre_lw.currentItem()
        self._rename_calibre_tag(item)

    def _rename_calibre_tag(self, item):
        if item is None:
            error_dialog(self, 'No item selected',
                         'Select a collection to rename.').exec_()
            return
        self.calibre_lw.editItem(item)

    def rename_marvin_tag(self):
        item = self.marvin_lw.currentItem()
        self._rename_marvin_tag(item)

    def _rename_marvin_tag(self, item):
        if item is None:
            error_dialog(self, 'No item selected',
                         'Select a collection to rename.').exec_()
            return
        self.marvin_lw.editItem(item)

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
