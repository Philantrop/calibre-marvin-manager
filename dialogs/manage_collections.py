#!/usr/bin/env python
# coding: utf-8

from __future__ import (unicode_literals, division, absolute_import,
                        print_function)

__license__ = 'GPL v3'
__copyright__ = '2010, Gregory Riker'
__docformat__ = 'restructuredtext en'

import os, sys

from calibre.devices.usbms.driver import debug_print
from calibre.gui2 import question_dialog, error_dialog
from calibre.gui2.dialogs.device_category_editor import DeviceCategoryEditor, ListWidgetItem
from calibre.gui2.dialogs.device_category_editor_ui import Ui_DeviceCategoryEditor

from PyQt4.Qt import (Qt, QDialog, QIcon,
                      pyqtSignal)


class MyDeviceCategoryEditor(DeviceCategoryEditor):
    '''
    subclass of gui2.dialogs.device_category_editor
    .available_tags is QListWidget
    .rename_button
    .delete_button
    '''
    LOCATION_TEMPLATE = "{cls}:{func}({arg1}) {arg2}"

    marvin_device_status_changed = pyqtSignal(str)

    def __init__(self, parent, tag_to_match, data, key, connected_device):
        QDialog.__init__(self, parent.opts.gui)
        Ui_DeviceCategoryEditor.__init__(self)
        self.setupUi(self)
        self.connected_device = connected_device
        self.verbose = parent.opts.verbose

        # Subscribe to Marvin driver change events
        connected_device.marvin_device_signals.reader_app_status_changed.connect(
            self.marvin_status_changed)

        # Remove help icon on title bar
        self.setWindowFlags(self.windowFlags() & (~Qt.WindowContextHelpButtonHint))
        self.setWindowIcon(QIcon(os.path.join(parent.opts.resources_path,
                                 'icons',
                                 'edit_collections.png')))
        self.setWindowTitle("Manage collections")
        self.label.setText("Active collections")

        self.to_rename = {}
        self.to_delete = set([])

        try:
            cc = set(data['calibre'])
        except:
            cc = set([])

        try:
            mc = set(data['Marvin'])
        except:
            mc = set([])

        merged = list(cc.union(mc))
        for tag in sorted(merged, key=key):
            item = ListWidgetItem(tag)
            item.setData(Qt.UserRole, tag)
            item.setFlags(item.flags() | Qt.ItemIsEditable)
            self.available_tags.addItem(item)

        self.delete_button.clicked.connect(self.delete_tags)
        self.rename_button.clicked.connect(self.rename_tag)
        self.available_tags.itemDoubleClicked.connect(self._rename_tag)
        self.available_tags.itemChanged.connect(self.finish_editing)

    def delete_tags(self):
        deletes = self.available_tags.selectedItems()
        if not deletes:
            error_dialog(self, 'No items selected',
                               'Select at least one collection to delete.',
                               show_copy_button=False).exec_()
            return
        ct = ', '.join([unicode(item.text()) for item in deletes])
        if not question_dialog(self, 'Are you sure?',
                               '<p>'+'Are you sure you want to delete the following collections?'+'<br>'+ct,
                               show_copy_button=False):
            return
        row = self.available_tags.row(deletes[0])
        for item in deletes:
            self.to_delete.add(unicode(item.text()))
            self.available_tags.takeItem(self.available_tags.row(item))

        if row >= self.available_tags.count():
            row = self.available_tags.count() - 1
        if row >= 0:
            self.available_tags.scrollToItem(self.available_tags.item(row))

    def finish_editing(self, item):
        if not item.text():
                error_dialog(self, 'Item is blank',
                             'An item cannot be set to nothing. Delete it instead.',
                             show_copy_button=False).exec_()
                item.setText(item.previous_text())
                return
        if item.text() != item.initial_text():
            self.to_rename[unicode(item.initial_text())] = unicode(item.text())

    def marvin_status_changed(self, command):
        '''
        '''
        self.marvin_device_status_changed.emit(command)

        self._log_location(command)

        if command in ['disconnected', 'yanked']:
            self._log("closing dialog: %s" % command)
            self.close()

    def rename_tag(self):
        item = self.available_tags.currentItem()
        self._rename_tag(item)

    def _rename_tag(self, item):
        if item is None:
            error_dialog(self, 'No item selected',
                         'Select a collection to rename.').exec_()
            return
        self.available_tags.editItem(item)

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
