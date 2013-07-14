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

from calibre.gui2.dialogs.device_category_editor import DeviceCategoryEditor, ListWidgetItem
from calibre.gui2.dialogs.device_category_editor_ui import Ui_DeviceCategoryEditor

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
    from manage_collections_ui import Ui_Dialog
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

class CollectionsManagementDialog(SizePersistedDialog, Ui_Dialog):
    LOCATION_TEMPLATE = "{cls}:{func}({arg1}) {arg2}"

    marvin_device_status_changed = pyqtSignal(str)

    def accept(self):
        self._log_location()
        super(CollectionsManagementDialog, self).accept()

    def close(self):
        self._log_location()
        super(CollectionsManagementDialog, self).close()

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
            if button.objectName() == 'rename_button':
                self._rename_collection()

        elif self.bb.buttonRole(button) == QDialogButtonBox.DestructiveRole:
            if button.objectName() == 'remove_button':
                self._remove_collection()

        elif self.bb.buttonRole(button) == QDialogButtonBox.RejectRole:
            self._log("RejectRole")
            self.close()

    def esc(self, *args):
        self._log_location()
        self._clear_selected_rows()

    def initialize(self, parent, original_collections, collection_ids, connected_device):
        '''
        __init__ is called on SizePersistedDialog()
        '''
        self.setupUi(self)
        self.collection_ids = collection_ids
        self.opts = parent.opts
        self.original_collections = original_collections
        self.parent = parent
        self.stored_command = None
        self.verbose = parent.verbose

        self._log_location()

        self.setWindowTitle("Collection Management")

        # Subscribe to Marvin driver change events
        connected_device.marvin_device_signals.reader_app_status_changed.connect(
            self.marvin_status_changed)

        # ~~~~~~~~ Rename button ~~~~~~~~
        self.rename_button = self.bb.addButton('Rename', QDialogButtonBox.ActionRole)
        self.rename_button.setObjectName('rename_button')
        self.rename_button.setIcon(QIcon(I('edit_input.png')))

        # ~~~~~~~~ Delete button ~~~~~~~~
        self.remove_button = self.bb.addButton('Remove', QDialogButtonBox.DestructiveRole)
        self.remove_button.setObjectName('remove_button')
        self.remove_button.setIcon(QIcon(I('trash.png')))

        # Populate collection model, save a copy of initial state
        self._initialize_collections()

        # Set the bg color of the description text fields to the dialog bg color
        bgcolor = self.palette().color(QPalette.Background)
        palette = QPalette()
        palette.setColor(QPalette.Base, bgcolor)
        self.collections_lw.setPalette(palette)

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
    def _clear_selection(self):
        '''
        '''
        self._log_location()
        self.collections_lw.clearSelection()

    def _clear_selected_rows(self):
        '''
        Clear any active selections
        '''
        self._log_location()
        self._clear_selection()

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
        rows_to_delete = len(self.collections_lw.model().listdata)
        for row in range(rows_to_delete - 1, -1, -1):
            self.collections_lw.model().beginRemoveRows(QModelIndex(), row, row)
            del self.collections_lw.model().listdata[row]
            self.collections_lw.model().endRemoveRows()

    def _get_collections(self):
        '''

        '''
        self._log_location()
        return self.collections_lw.model().listdata

    def _initialize_collections(self):
        '''
        Populate the data model with merged collection assignments
        '''
        self._log_location()
        self._log("original_collections: %s" % self.original_collections)
        cc = set(self.original_collections['calibre'])
        mc = set(self.original_collections['Marvin'])
        merged = sorted(list(cc.union(mc)), key=sort_key)

        self.collections_lw.setModel(MyListModel(merged))

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

    def _rename_collection(self):
        '''
        '''
        self._log_location()
        if self._selected_rows():
            row = self._selected_rows()[0]
            current_name = self.collections_lw.model().listdata[row]
            self._log("current_name: %s" % current_name)

    def _remove_collection(self):
        '''
        '''
        self._log_location()
        if self._selected_rows():
            row = self._selected_rows()[0]
            self.collections_lw.model().beginRemoveRows(QModelIndex(), row, row)
            del self.collections_lw.model().listdata[row]
            self.collections_lw.model().endRemoveRows()

    def _selected_rows(self):
        '''
        Return a list of selected rows
        '''
        srs = self.collections_lw.selectionModel().selectedRows()
        return [sr.row() for sr in srs]

class MyDeviceCategoryEditor(DeviceCategoryEditor):
    '''
    subclass of gui2.dialogs.device_category_editor
    '''
    def __init__(self, window, tag_to_match, data, key):
        QDialog.__init__(self, window)
        Ui_DeviceCategoryEditor.__init__(self)
        self.setupUi(self)
        # Remove help icon on title bar
        icon = self.windowIcon()
        self.setWindowFlags(self.windowFlags()&(~Qt.WindowContextHelpButtonHint))
        self.setWindowIcon(icon)
        self.setWindowTitle("Manage collections")
        self.label.setText("Active collections")

        self.to_rename = {}
        self.to_delete = set([])
        self.original_names = {}
        self.all_tags = {}

        """
        for k,v in data:
            self.all_tags[v] = k
            self.original_names[k] = v
        for tag in sorted(self.all_tags.keys(), key=key):
            item = ListWidgetItem(tag)
            item.setData(Qt.UserRole, self.all_tags[tag])
            item.setFlags (item.flags() | Qt.ItemIsEditable)
            self.available_tags.addItem(item)

        if tag_to_match is not None:
            items = self.available_tags.findItems(tag_to_match, Qt.MatchExactly)
            if len(items) == 1:
                self.available_tags.setCurrentItem(items[0])
        """

        cc = set(data['calibre'])
        mc = set(data['Marvin'])
        merged = list(cc.union(mc))
        for tag in sorted(merged, key=key):
            item = ListWidgetItem(tag)
            item.setData(Qt.UserRole, tag)
            item.setFlags (item.flags() | Qt.ItemIsEditable)
            self.available_tags.addItem(item)

        self.delete_button.clicked.connect(self.delete_tags)
        self.rename_button.clicked.connect(self.rename_tag)
        self.available_tags.itemDoubleClicked.connect(self._rename_tag)
        self.available_tags.itemChanged.connect(self.finish_editing)


