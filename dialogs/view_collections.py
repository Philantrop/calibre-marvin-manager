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
from calibre.gui2 import Application
from calibre.gui2.dialogs.device_category_editor import ListWidgetItem
from calibre.gui2.dialogs.message_box import MessageBox
from calibre.utils.icu import sort_key
from calibre.utils.magick.draw import add_borders_to_image, thumbnail

from calibre_plugins.marvin_manager.book_status import dialog_resources_path
from calibre_plugins.marvin_manager.common_utils import (MyAbstractItemModel,
    SizePersistedDialog)

from PyQt4.Qt import (Qt, QAbstractItemModel, QAbstractListModel, QColor,
                      QDialog, QDialogButtonBox, QIcon, QMimeData,
                      QModelIndex, QPalette, QPixmap, QSize, QSizePolicy, QVariant,
                      pyqtSignal, SIGNAL)

# Import Ui_Form from form generated dynamically during initialization
if True:
    sys.path.insert(0, dialog_resources_path)
    from view_collections_ui import Ui_Dialog
    sys.path.remove(dialog_resources_path)

class CollectionsViewerDialog(SizePersistedDialog, Ui_Dialog):
    LOCATION_TEMPLATE = "{cls}:{func}({arg1}) {arg2}"

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
            self.updated_calibre_collections = self.initial_calibre_collections
            self.updated_marvin_collections = self.initial_marvin_collections
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
        self.setWindowTitle("'%s' collection assignments" % book_title)

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
        self.remove_assignment_tb.setToolTip("Remove a collection assignment")
        self.remove_assignment_tb.clicked.connect(self._remove_collection_assignment)

        # ~~~~~~~~ Rename collection button ~~~~~~~~
        self.rename_collection_tb.setIcon(QIcon(I('edit_input.png')))
        self.rename_collection_tb.setToolTip("Rename collection")
        self.rename_collection_tb.clicked.connect(self._rename_collection)

        # ~~~~~~~~ Clear all collections button ~~~~~~~~
        self.remove_all_assignments_tb.setIcon(QIcon(os.path.join(self.opts.resources_path,
                                                          'icons',
                                                          'remove_all_collections.png')))
        self.remove_all_assignments_tb.setToolTip("Remove all collection assignments from calibre and Marvin")
        self.remove_all_assignments_tb.clicked.connect(self._remove_all_assignments)

        # Populate collection models
        self._initialize_collections()

        if self.calibre_collections:
            # Save initial state
            self.initial_calibre_collections = list(self._get_calibre_collections())
            self.initial_marvin_collections = list(self._get_marvin_collections())
            # Remind the user of calibre's custom column, disable buttons if no calibre field
            calibre_cf = self.prefs.get('collection_field_comboBox', '')
            if calibre_cf:
                self.calibre_gb.setTitle("Calibre (%s)" % calibre_cf)
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
                item.setFlags(item.flags() | Qt.ItemIsEditable)
                self.calibre_lw.addItem(item)

        for ma in self.marvin_collections:
            item = ListWidgetItem(ma)
            item.setData(Qt.UserRole, ma)
            item.setFlags(item.flags() | Qt.ItemIsEditable)
            self.marvin_lw.addItem(item)

        # Capture click events to clear selections in opposite list
        self.calibre_lw.clicked.connect(self._clear_marvin_selection)
        self.calibre_lw.doubleClicked.connect(self.rename_calibre_tag)

        self.marvin_lw.clicked.connect(self._clear_calibre_selection)
        self.marvin_lw.doubleClicked.connect(self.rename_marvin_tag)

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
        self._log_location()
        def _remove_assignments(deletes, list_widget):
            row = list_widget.row(deletes[0])

            for item in deletes:
                list_widget.takeItem(list_widget.row(item))
            if row >= list_widget.count():
                row = list_widget.count() - 1
            if row >= 0:
                list_widget.scrollToItem(list_widget.item(row))

        if self.calibre_lw.selectedItems():
            deletes = self.calibre_lw.selectedItems()
            _remove_assignments(deletes, self.calibre_lw)

        elif self.marvin_lw.selectedItems():
            deletes = self.marvin_lw.selectedItems()
            _remove_assignments(deletes, self.marvin_lw)

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

