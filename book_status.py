#!/usr/bin/env python
# coding: utf-8
from __future__ import (unicode_literals, division, absolute_import,
                        print_function)

__license__ = 'GPL v3'
__copyright__ = '2013, Greg Riker <griker@hotmail.com>'
__docformat__ = 'restructuredtext en'

import base64, cStringIO, hashlib, importlib, inspect, json
import locale, operator, os, cPickle as pickle, re, sqlite3, sys, time

from collections import OrderedDict
from datetime import datetime, timedelta
from dateutil import tz
from functools import partial
from lxml import etree
from threading import Timer
from xml.sax.saxutils import escape

from PyQt4 import QtCore
from PyQt4.Qt import (Qt, QAbstractTableModel,
                      QApplication, QBrush,
                      QColor, QCursor, QDialogButtonBox, QFont, QFontMetrics, QGridLayout,
                      QHeaderView, QHBoxLayout, QIcon,
                      QItemSelectionModel, QLabel, QLineEdit, QMenu, QModelIndex,
                      QPainter, QPixmap, QProgressDialog, QPushButton,
                      QSize, QSizePolicy, QSpacerItem, QString,
                      QTableView, QTableWidget, QTableWidgetItem, QTimer, QToolButton,
                      QVariant, QVBoxLayout, QWidget,
                      SIGNAL, pyqtSignal)

from calibre import strftime
from calibre.constants import islinux, isosx, iswindows
from calibre.devices.errors import UserFeedback
from calibre.devices.usbms.driver import debug_print
from calibre.ebooks.BeautifulSoup import BeautifulSoup, BeautifulStoneSoup, Tag, UnicodeDammit
from calibre.ebooks.oeb.iterator import EbookIterator
from calibre.gui2 import Application, Dispatcher, error_dialog, warning_dialog
from calibre.gui2.dialogs.message_box import MessageBox
from calibre.gui2.dialogs.progress import ProgressDialog
from calibre.gui2.progress_indicator import ProgressIndicator
from calibre.utils.config import config_dir, JSONConfig
from calibre.utils.date import strptime
from calibre.utils.icu import sort_key
from calibre.utils.magick.draw import thumbnail
from calibre.utils.wordcount import get_wordcount_obj
from calibre.utils.zipfile import ZipFile

from calibre_plugins.marvin_manager.annotations import merge_annotations

from calibre_plugins.marvin_manager.common_utils import (
    AbortRequestException, AnnotationStruct, Book, BookStruct, InventoryCollections,
    Logger, MyBlockingBusy, ProgressBar, RowFlasher, SizePersistedDialog,
    get_cc_mapping, get_icon, updateCalibreGUIView)

dialog_resources_path = os.path.join(config_dir, 'plugins', 'Marvin_XD_resources', 'dialogs')


class MyTableView(QTableView):
    def __init__(self, parent):
        super(MyTableView, self).__init__(parent)
        self.parent = parent

        # Hook header context menu events separately
        self.horizontalHeader().setContextMenuPolicy(Qt.CustomContextMenu)
        self.horizontalHeader().customContextMenuRequested.connect(self.header_event)

    def contextMenuEvent(self, event):

        index = self.indexAt(event.pos())
        col = index.column()
        row = index.row()
        selected_books = self.parent._selected_books()
        menu = QMenu(self)

        if self.parent.busy:
            # Don't show context menu if busy
            pass

        elif col == self.parent.ANNOTATIONS_COL:
            calibre_cids = False
            for row in selected_books:
                if selected_books[row]['cid'] is not None:
                    calibre_cids = True
                    break

            afn = get_cc_mapping('annotations', 'combobox', None)
            no_annotations = not selected_books[row]['has_annotations']

            ac = menu.addAction("View annotations")
            ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'annotations.png')))
            ac.triggered.connect(partial(self.parent.dispatch_context_menu_event, "show_highlights", row))
            if len(selected_books) > 1 or no_annotations:
                ac.setEnabled(False)

            # Fetch Annotations if custom field specified
            enabled = False
            if afn:
                # Do any of the selected books have annotations?
                if len(selected_books) > 1 and calibre_cids:
                    for sr in selected_books:
                        if selected_books[sr]['has_annotations']:
                            enabled = True
                            break
                elif len(selected_books) == 1 and selected_books[row]['has_annotations'] and calibre_cids:
                    enabled = True
                ac = menu.addAction("Add annotations to '{0}' column".format(afn))
                ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'annotations.png')))
                ac.triggered.connect(partial(self.parent.dispatch_context_menu_event, "fetch_annotations", row))
            else:
                ac = menu.addAction("No custom column specified for 'Annotations'")
                ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'annotations.png')))
            ac.setEnabled(enabled)

        elif col == self.parent.ARTICLES_COL:
            try:
                no_articles = not selected_books[row]['has_articles']
                ac = menu.addAction("View articles")
                ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'articles.png')))
                ac.triggered.connect(partial(self.parent.dispatch_context_menu_event, "show_deep_view_articles", row))
                if len(selected_books) > 1 or no_articles:
                    ac.setEnabled(False)
            except:
                pass

        elif col == self.parent.COLLECTIONS_COL:
            cfl = get_cc_mapping('collections', 'field', None)

            ac = menu.addAction("Add collection assignments")
            ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'star.png')))
            ac.triggered.connect(self.parent.show_add_collections_dialog)

            ac = menu.addAction("View collection assignments")
            ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'update_metadata.png')))
            ac.triggered.connect(partial(self.parent.dispatch_context_menu_event, "show_collections", row))
            if len(selected_books) > 1:
                ac.setEnabled(False)

            ac = menu.addAction("Export calibre collections to Marvin")
            ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'from_calibre.png')))
            if cfl:
                ac.triggered.connect(partial(self.parent.dispatch_context_menu_event, "export_collections", row))
            else:
                ac.setEnabled(False)

            ac = menu.addAction("Import Marvin collections to calibre")
            ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'from_marvin.png')))
            if cfl:
                ac.triggered.connect(partial(self.parent.dispatch_context_menu_event, "import_collections", row))
            else:
                ac.setEnabled(False)

            ac = menu.addAction("Merge collections")
            ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'sync_collections.png')))
            if cfl:
                ac.triggered.connect(partial(self.parent.dispatch_context_menu_event, "synchronize_collections", row))
            else:
                ac.setEnabled(False)

            ac = menu.addAction("Remove from all collections")
            ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'clear_all.png')))
            ac.triggered.connect(partial(self.parent.dispatch_context_menu_event, "clear_all_collections", row))

            menu.addSeparator()

            ac = menu.addAction("Manage collections")
            ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'edit_collections.png')))
            ac.triggered.connect(partial(self.parent.dispatch_context_menu_event, "manage_collections", row))

        elif col == self.parent.DEEP_VIEW_COL:
            try:
                no_dv_content = False
                for row in selected_books:
                    if not selected_books[row]['has_dv_content']:
                        no_dv_content = True
                        break

                ac = menu.addAction("Generate Deep View content")
                ac.setIcon(QIcon(I('exec.png')))
                ac.triggered.connect(partial(self.parent.dispatch_context_menu_event, "generate_deep_view", row))
                ac.setEnabled(no_dv_content)

                menu.addSeparator()

                ac = menu.addAction("Deep View articles")
                ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'deep_view.png')))
                ac.triggered.connect(partial(self.parent.dispatch_context_menu_event,
                                             "show_deep_view_articles", row))
                no_articles = not selected_books[row]['has_articles']
                if len(selected_books) > 1 or no_articles:
                    ac.setEnabled(False)

                ac = menu.addAction("Deep View items, sorted alphabetically")
                ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'deep_view.png')))
                ac.triggered.connect(partial(self.parent.dispatch_context_menu_event,
                                             "show_deep_view_alphabetically", row))
                if len(selected_books) > 1 or no_dv_content:
                    ac.setEnabled(False)

                ac = menu.addAction("Deep View items, sorted by importance")
                ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'deep_view.png')))
                ac.triggered.connect(partial(self.parent.dispatch_context_menu_event,
                                             "show_deep_view_by_importance", row))
                if len(selected_books) > 1 or no_dv_content:
                    ac.setEnabled(False)

                ac = menu.addAction("Deep View items, sorted by order of appearance")
                ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'deep_view.png')))
                ac.triggered.connect(partial(self.parent.dispatch_context_menu_event,
                                             "show_deep_view_by_appearance", row))
                if len(selected_books) > 1 or no_dv_content:
                    ac.setEnabled(False)

                ac = menu.addAction("Deep View items, notes and flags first")
                ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'deep_view.png')))
                ac.triggered.connect(partial(self.parent.dispatch_context_menu_event,
                                             "show_deep_view_by_annotations", row))
                if len(selected_books) > 1 or no_dv_content:
                    ac.setEnabled(False)
            except:
                pass

        elif col == self.parent.FLAGS_COL:
            ac = menu.addAction("Clear all")
            ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'clear_all.png')))
            ac.triggered.connect(partial(self.parent.dispatch_context_menu_event, "clear_all_flags", row))
            ac = menu.addAction("Clear New")
            ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'clear_new.png')))
            ac.triggered.connect(partial(self.parent.dispatch_context_menu_event, "clear_new_flag", row))
            ac = menu.addAction("Clear Reading list")
            ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'clear_reading.png')))
            ac.triggered.connect(partial(self.parent.dispatch_context_menu_event, "clear_reading_list_flag", row))
            ac = menu.addAction("Clear Read")
            ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'clear_read.png')))
            ac.triggered.connect(partial(self.parent.dispatch_context_menu_event, "clear_read_flag", row))
            menu.addSeparator()
            ac = menu.addAction("Set New")
            ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'set_new.png')))
            ac.triggered.connect(partial(self.parent.dispatch_context_menu_event, "set_new_flag", row))
            ac = menu.addAction("Set Reading list")
            ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'set_reading.png')))
            ac.triggered.connect(partial(self.parent.dispatch_context_menu_event, "set_reading_list_flag", row))
            ac = menu.addAction("Set Read")
            ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'set_read.png')))
            ac.triggered.connect(partial(self.parent.dispatch_context_menu_event, "set_read_flag", row))

            # Add Synchronize option
            calibre_cids = False
            for row in selected_books:
                if selected_books[row]['cid'] is not None:
                    calibre_cids = True
                    break

            read_field = get_cc_mapping('read', 'combobox', None)
            reading_list_field = get_cc_mapping('reading_list', 'combobox', None)
            if read_field or reading_list_field:
                menu.addSeparator()
                label = 'Synchronize Reading list, Read'
                if reading_list_field and not read_field:
                    label = 'Synchronize Reading list'
                elif read_field and not reading_list_field:
                    label = 'Synchronize Read'
                ac = menu.addAction(label)
                ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'sync_collections.png')))
                ac.triggered.connect(partial(self.parent.dispatch_context_menu_event, "synchronize_flags", row))
                ac.setEnabled(bool(calibre_cids))

        elif col == self.parent.LAST_OPENED_COL:
            date_read_field = get_cc_mapping('date_read', 'combobox', None)

            # Test for calibre cids
            calibre_cids = False
            for row in selected_books:
                if selected_books[row]['cid'] is not None:
                    calibre_cids = True
                    break

            # Test for active last_opened dates
            last_opened = False
            for row in selected_books:
                if selected_books[row]['last_opened'] > '':
                    last_opened = True
                    break

            title = "No custom column specified for 'Last read'"
            if date_read_field:
                title = "Apply to '%s' column" % date_read_field
            ac = menu.addAction(title)
            ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'from_marvin.png')))
            ac.triggered.connect(partial(self.parent.dispatch_context_menu_event, "apply_date_read", row))

            if (not date_read_field) or (not calibre_cids) or (not last_opened):
                ac.setEnabled(False)

        elif col == self.parent.LOCKED_COL:
            any_ids_locked = False
            for row in selected_books:
                if selected_books[row]['locked']:
                    any_ids_locked = True
                    break

            any_ids_unlocked = False
            for row in selected_books:
                if not selected_books[row]['locked']:
                    any_ids_unlocked = True
                    break

            ac = menu.addAction("Lock")
            ac.setEnabled(any_ids_unlocked)
            if any_ids_unlocked:
                icon = QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'lock_enabled.png'))
            else:
                icon = QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'lock_disabled.png'))
            ac.setIcon(icon)
            ac.triggered.connect(partial(self.parent.dispatch_context_menu_event, "set_locked", row))

            ac = menu.addAction("Unlock")
            ac.setEnabled(any_ids_locked)
            if any_ids_locked:
                icon = QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'unlock_enabled.png'))
            else:
                icon = QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'unlock_disabled.png'))
            ac.setIcon(icon)
            ac.triggered.connect(partial(self.parent.dispatch_context_menu_event, "set_unlocked", row))

        elif col == self.parent.PROGRESS_COL:
            progress_field = get_cc_mapping('progress', 'combobox', None)

            # Test for calibre cids
            calibre_cids = False
            for row in selected_books:
                if selected_books[row]['cid'] is not None:
                    calibre_cids = True
                    break

            # Test for active Progress
            progress = True
#             for row in selected_books:
#                 if selected_books[row]['progress'] > 0:
#                     progress = True
#                     break

            title = "No custom column specified for 'Progress'"
            if progress_field:
                title = "Apply to '{0}' column".format(progress_field)
            ac = menu.addAction(title)
            ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'from_marvin.png')))
            ac.triggered.connect(partial(self.parent.dispatch_context_menu_event, "apply_progress", row))

            if (not progress_field) or (not calibre_cids) or (not progress):
                ac.setEnabled(False)

        elif col in [self.parent.TITLE_COL, self.parent.AUTHOR_COL]:
            ac = menu.addAction("View metadata")
            ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'update_metadata.png')))
            ac.triggered.connect(partial(self.parent.dispatch_context_menu_event, "show_metadata", row))
            if len(selected_books) > 1:
                ac.setEnabled(False)

            # If match_quality < YELLOW, metadata updates disabled
            enable_metadata_updates = True
            if len(selected_books) == 1 and self.parent.tm.get_match_quality(row) < BookStatusDialog.MATCH_COLORS.index('YELLOW'):
                enable_metadata_updates = False

            ac = menu.addAction("Export metadata from calibre to Marvin")
            ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'from_calibre.png')))
            ac.triggered.connect(partial(self.parent.dispatch_context_menu_event, "export_metadata", row))
            ac.setEnabled(enable_metadata_updates)

            ac = menu.addAction("Import metadata from Marvin to calibre")
            ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'from_marvin.png')))
            ac.triggered.connect(partial(self.parent.dispatch_context_menu_event, "import_metadata", row))
            ac.setEnabled(enable_metadata_updates)

            menu.addSeparator()
            # Test for calibre cids
            in_library = True
            for row in selected_books:
                if selected_books[row]['cid'] is None:
                    in_library = False
                    break

            ac = menu.addAction("Add to calibre library")
            ac.setIcon(QIcon(I('plus.png')))
            ac.triggered.connect(self.parent._add_books_to_library)
            ac.setEnabled(not in_library)

            menu.addSeparator()
            ac = menu.addAction("Delete from Marvin library")
            ac.setIcon(QIcon(I('trash.png')))
            ac.triggered.connect(self.parent._delete_books)

        elif col == self.parent.VOCABULARY_COL:
            try:
                no_vocabulary = not selected_books[row]['has_vocabulary']

                ac = menu.addAction("View vocabulary for this book")
                ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'vocabulary.png')))
                ac.triggered.connect(partial(self.parent.dispatch_context_menu_event, "show_vocabulary", row))
                if len(selected_books) > 1 or no_vocabulary:
                    ac.setEnabled(False)

                ac = menu.addAction("View all vocabulary words")
                ac.setIcon(QIcon(I('books_in_series.png')))
                ac.triggered.connect(partial(self.parent.dispatch_context_menu_event, "show_global_vocabulary", row))
            except:
                pass

        elif col == self.parent.WORD_COUNT_COL:
            word_count_field = get_cc_mapping('word_count', 'combobox', None)

            # Test for calibre cids
            calibre_cids = False
            for row in selected_books:
                if selected_books[row]['cid'] is not None:
                    calibre_cids = True
                    break

            # Test for active word counts
            word_counts = False
            for row in selected_books:
                #print(repr(selected_books[row]['word_count']))
                if selected_books[row]['word_count']:
                    word_counts = True
                    break

            ac = menu.addAction("Calculate word count")
            ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'word_count.png')))
            ac.triggered.connect(partial(self.parent.dispatch_context_menu_event, "calculate_word_count", row))

            title = "No custom column specified for 'Word count'"
            if word_count_field:
                title = "Apply to '{0}' column".format(word_count_field)
            ac = menu.addAction(title)
            ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'from_marvin.png')))
            ac.triggered.connect(partial(self.parent.dispatch_context_menu_event, "apply_word_count", row))

            if (not word_count_field) or (not calibre_cids) or (not word_counts):
                ac.setEnabled(False)

        menu.exec_(event.globalPos())

    def header_event(self, pos):
        '''
        Context menu event handler for header
        Allow user to toggle column visibility
        '''

        menu = QMenu(self)

        for col, title in self.parent.USER_CONTROLLED_COLUMNS:
            visible = not self.isColumnHidden(col) and self.columnWidth(col) > 0
            ac = menu.addAction(title)
            ac.setCheckable(True)
            ac.setChecked(visible)
            ac.triggered.connect(partial(self.toggle_column_visibility, col))

        action = menu.exec_(self.mapToGlobal(pos))

    def toggle_column_visibility(self, col):
        '''
        Toggle visible state of col, resize to contents
        '''
        invisible = self.isColumnHidden(col) or self.columnWidth(col) == 0
        if invisible:
            self.showColumn(col)

            # Set width of shown column
            if col in [self.parent.AUTHOR_COL, self.parent.SUBJECTS_COL]:
                self.setColumnWidth(col, self.columnWidth(self.parent.TITLE_COL))
            elif col in [self.parent.WORD_COUNT_COL, self.parent.COLLECTIONS_COL]:
                width = self.columnWidth(self.parent.LAST_OPENED_COL)
                if not width:
                    width = 87
                self.setColumnWidth(col, width)
            elif col in [self.parent.ANNOTATIONS_COL, self.parent.VOCABULARY_COL,
                         self.parent.DEEP_VIEW_COL, self.parent.ARTICLES_COL]:
                width = self.columnWidth(self.parent.FLAGS_COL)
                if not width:
                    width = 53
                self.setColumnWidth(col, width)
            else:
                self.resizeColumnToContents(col)
        else:
            self.hideColumn(col)

        # Update Refresh button label based upon column visibility
        self.parent._update_refresh_button()


class SortableImageWidgetItem(QWidget):
    def __init__(self, path, sort_key):
        super(SortableImageWidgetItem, self).__init__()
        self.picture = QPixmap(path)
        self.sort_key = sort_key

    def __lt__(self, other):
        return self.sort_key < other.sort_key


class SortableTableWidgetItem(QTableWidgetItem):
    """
    Subclass widget sortable by sort_key
    """
    def __init__(self, text, sort_key):
        super(SortableTableWidgetItem, self).__init__(text)
        self.sort_key = sort_key

    def __lt__(self, other):
        return self.sort_key < other.sort_key


class MarkupTableModel(QAbstractTableModel):
    #http://www.saltycrane.com/blog/2007/12/pyqt-43-qtableview-qabstracttablemodel/

    SATURATION = 0.40
    HSVALUE = 1.0
    RED_HUE = 0.0           #   0/360
    ORANGE_HUE = 0.08325    #  30/360
    YELLOW_HUE = 0.1665     #  60/360
    GREEN_HUE = 0.333       # 120/360
    CYAN_HUE = 0.500        # 180/360
    MAGENTA_HUE = 0.875        # 315/360
    WHITE_HUE = 1.0

    def __init__(self, parent=None, centered_columns=[], right_aligned_columns=[], *args):
        """
        datain: a list of lists
        headerdata: a list of strings
        """
        QAbstractTableModel.__init__(self, parent, *args)
        self.parent = parent
        self.arraydata = parent.tabledata
        self.centered_columns = centered_columns
        self.right_aligned_columns = right_aligned_columns
        self.headerdata = parent.LIBRARY_HEADER
        self.show_match_colors = parent.show_match_colors

    def all_rows(self):
        return self.arraydata

    def columnCount(self, parent):
        return len(self.headerdata)

    def data(self, index, role):
        row, col = index.row(), index.column()
        if not index.isValid():
            return QVariant()

        elif role == Qt.ForegroundRole and self.show_match_colors:
            match_quality = self.get_match_quality(row)
            if match_quality == BookStatusDialog.MATCH_COLORS.index('DARK_GRAY'):
                return QVariant(QBrush(Qt.white))

        elif role == Qt.BackgroundRole and self.show_match_colors:
            match_quality = self.get_match_quality(row)
            if match_quality == BookStatusDialog.MATCH_COLORS.index('LIGHT_GRAY'):
                return QVariant(QBrush(QColor(0xD8, 0xD8,0xD8)))
            elif match_quality == BookStatusDialog.MATCH_COLORS.index('DARK_GRAY'):
                return QVariant(QBrush(QColor(0x98, 0x98,0x98)))
            elif match_quality == BookStatusDialog.MATCH_COLORS.index('GREEN'):
                return QVariant(QBrush(QColor.fromHsvF(self.GREEN_HUE, self.SATURATION, self.HSVALUE)))
            elif match_quality == BookStatusDialog.MATCH_COLORS.index('MAGENTA'):
                return QVariant(QBrush(QColor.fromHsvF(self.MAGENTA_HUE, self.SATURATION, self.HSVALUE)))
            elif match_quality == BookStatusDialog.MATCH_COLORS.index('ORANGE'):
                return QVariant(QBrush(QColor.fromHsvF(self.ORANGE_HUE, self.SATURATION, self.HSVALUE)))
            elif match_quality == BookStatusDialog.MATCH_COLORS.index('RED'):
                return QVariant(QBrush(QColor.fromHsvF(self.RED_HUE, self.SATURATION, self.HSVALUE)))
            elif match_quality == BookStatusDialog.MATCH_COLORS.index('YELLOW'):
                return QVariant(QBrush(QColor.fromHsvF(self.YELLOW_HUE, self.SATURATION, self.HSVALUE)))
            else:
                return QVariant(QBrush(QColor.fromHsvF(self.WHITE_HUE, 0.0, self.HSVALUE)))

        elif role == Qt.DecorationRole and col == self.parent.LOCKED_COL:
            return self.arraydata[row][self.parent.LOCKED_COL].picture

        elif role == Qt.DecorationRole and col == self.parent.FLAGS_COL:
            return self.arraydata[row][self.parent.FLAGS_COL].picture

        elif role == Qt.DecorationRole and col == self.parent.COLLECTIONS_COL:
            return self.arraydata[row][self.parent.COLLECTIONS_COL].picture

        elif (role == Qt.DisplayRole and
              col == self.parent.PROGRESS_COL
              and self.parent.prefs.get('show_progress_as_percentage', False)):
            return self.arraydata[row][self.parent.PROGRESS_COL].text()
        elif (role == Qt.DecorationRole and
              col == self.parent.PROGRESS_COL
              and not self.parent.prefs.get('show_progress_as_percentage', False)):
            return self.arraydata[row][self.parent.PROGRESS_COL].picture

        elif role == Qt.DisplayRole and col == self.parent.SERIES_COL:
            return self.arraydata[row][self.parent.SERIES_COL].text()

        elif role == Qt.DisplayRole and col == self.parent.WORD_COUNT_COL:
            return self.arraydata[row][self.parent.WORD_COUNT_COL].text()

        elif role == Qt.DisplayRole and col == self.parent.TITLE_COL:
            return self.arraydata[row][self.parent.TITLE_COL].text()
        elif role == Qt.DisplayRole and col == self.parent.AUTHOR_COL:
            return self.arraydata[row][self.parent.AUTHOR_COL].text()
        elif role == Qt.DisplayRole and col == self.parent.DATE_ADDED_COL:
            return self.arraydata[row][self.parent.DATE_ADDED_COL].text()
        elif role == Qt.DisplayRole and col == self.parent.LAST_OPENED_COL:
            return self.arraydata[row][self.parent.LAST_OPENED_COL].text()
        elif role == Qt.DisplayRole and col == self.parent.SUBJECTS_COL:
            return self.arraydata[row][self.parent.SUBJECTS_COL].text()
        elif role == Qt.DisplayRole and col == self.parent.ANNOTATIONS_COL:
            return self.arraydata[row][self.parent.ANNOTATIONS_COL].text()
        elif role == Qt.DisplayRole and col == self.parent.VOCABULARY_COL:
            return self.arraydata[row][self.parent.VOCABULARY_COL].text()
        elif role == Qt.DisplayRole and col == self.parent.ARTICLES_COL:
            return self.arraydata[row][self.parent.ARTICLES_COL].text()

        elif role == Qt.TextAlignmentRole and (col in self.centered_columns):
            return Qt.AlignHCenter
        elif role == Qt.TextAlignmentRole and (col in self.right_aligned_columns):
            return Qt.AlignRight

        elif role == Qt.ToolTipRole:
            if self.parent.busy:
                return "<p>Please wait until current operation completes</p>"
            else:
                match_quality = self.get_match_quality(row)
                tip = '<p>'
                if match_quality == BookStatusDialog.MATCH_COLORS.index('GREEN'):
                    tip += 'Matched in calibre library'
                elif match_quality == BookStatusDialog.MATCH_COLORS.index('YELLOW'):
                    tip += 'Matched in calibre library with differing metadata'
                elif match_quality == BookStatusDialog.MATCH_COLORS.index('ORANGE'):
                    tip += 'Duplicate of matched book in calibre library'
                elif match_quality == BookStatusDialog.MATCH_COLORS.index('LIGHT_GRAY'):
                    tip += 'Book updated in calibre library'
                elif match_quality == BookStatusDialog.MATCH_COLORS.index('DARK_GRAY'):
                    tip += 'Book updated in Marvin library'
                elif match_quality == BookStatusDialog.MATCH_COLORS.index('MAGENTA'):
                    tip += 'Multiple copies in calibre library'
                elif match_quality == BookStatusDialog.MATCH_COLORS.index('RED'):
                    tip += 'Duplicated in Marvin library'
                else:
                    tip += 'Book in Marvin library only'

                # Add the suffix based upon column
                if col in [self.parent.TITLE_COL, self.parent.AUTHOR_COL]:
                    return tip + "<br/>Double-click to view metadata<br/>Right-click for more options</p>"

                elif col in [self.parent.ANNOTATIONS_COL,
                             self.parent.ARTICLES_COL]:
                    has_content = bool(self.arraydata[row][col])
                    if has_content:
                        return tip + "<br/>Double-click to view details<br/>Right-click for more options</p>"
                    else:
                        return tip + '</p>'

                elif col == self.parent.COLLECTIONS_COL:
                    has_content = bool(self.arraydata[row][col].sort_key)
                    if has_content:
                        return tip + "<br/>Double-click to view details<br/>Right-click for more options</p>"
                    else:
                        return tip + '<br/>Right-click for more options</p>'

                elif col in [self.parent.VOCABULARY_COL]:
                    has_content = bool(self.arraydata[row][col])
                    if has_content:
                        return tip + "<br/>>Double-click to view Vocabulary words<br/>Right-click for more options</p>"
                    else:
                        return tip + '<br/>Right-click for options</p>'


                elif col in [self.parent.DEEP_VIEW_COL]:
                    has_content = bool(self.arraydata[row][col])
                    if has_content:
                        return tip + "<br/>Double-click to view Deep View content<br/>Right-click for more options</p>"
                    else:
                        return tip + '<br/>Double-click to generate Deep View content<br/>Right-click for more options</p>'

                elif col in [self.parent.FLAGS_COL]:
                    return tip + "<br/>Right-click for options</p>"

                elif col == self.parent.LOCKED_COL:
                    return ("<p>Double-click to toggle locked status" +
                            "<br/>Right-click for more options</p>")

                elif col in [self.parent.WORD_COUNT_COL]:
                    return (tip + "<br/>Double-click to generate word count" +
                                  "<br/>Right-click to generate word count for multiple books</p>")

                else:
                    return tip + '</p>'

        elif role != Qt.DisplayRole:
            return QVariant()

        return QVariant(self.arraydata[index.row()][index.column()])

    def headerData(self, col, orientation, role):
        if role == Qt.DisplayRole:
            if orientation == Qt.Horizontal:
                return QVariant(self.headerdata[col])

        if role == Qt.ToolTipRole:
            if orientation == Qt.Horizontal:
                if col == self.parent.ANNOTATIONS_COL:
                    tip = "<p>Annotations and Highlights.<br/>"
                elif col == self.parent.ARTICLES_COL:
                    tip = "<p>Pinned articles.<br/>"
                elif col == self.parent.AUTHOR_COL:
                    tip = "<p>Book author.<br/>"
                elif col == self.parent.COLLECTIONS_COL:
                    tip = "<p>Collection assignments.<br/>"
                elif col == self.parent.DATE_ADDED_COL:
                    tip = "<p>Date added to Marvin.<br/>"
                elif col == self.parent.DEEP_VIEW_COL:
                    tip = "<p>Deep View items.<br/>"
                elif col == self.parent.FLAGS_COL:
                    tip = "<p><i>New</i>, <i>Reading</i> and <i>Read</i> flags.<br/>"
                elif col == self.parent.LAST_OPENED_COL:
                    tip = "<p>Last opened in Marvin.<br/>"
                elif col == self.parent.LOCKED_COL:
                    tip = "<p>Locked status.<br/>"
                elif col == self.parent.PROGRESS_COL:
                    tip = "<p>Reading progress.<br/>"
                elif col == self.parent.SERIES_COL:
                    tip = "<p>Book series.<br/>"
                elif col == self.parent.SUBJECTS_COL:
                    tip = "<p>Book subjects.<br/>"
                elif col == self.parent.TITLE_COL:
                    tip = "<p>Book title.<br/>"
                elif col == self.parent.VOCABULARY_COL:
                    tip = "<p>Vocabulary words.<br/>"
                elif col == self.parent.WORD_COUNT_COL:
                    tip = "<p>Word count.<br/>"
                else:
                    tip = '<p>'

                suffix = "Right-click to show or hide columns.</p>"

                return QString(tip + suffix)

        return QVariant()

    def refresh(self, show_match_colors):
        self.show_match_colors = show_match_colors
        self.dataChanged.emit(self.createIndex(0, 0),
                              self.createIndex(self.rowCount(0), self.columnCount(0)))

    def rowCount(self, parent):
        return len(self.arraydata)

    def setData(self, index, value, role):
        row, col = index.row(), index.column()
        self.emit(SIGNAL("dataChanged(QModelIndex,QModelIndex)"), index, index)
        return True

    def sort(self, Ncol, order):
        """
        Sort table by given column number.
        """
        self.emit(SIGNAL("layoutAboutToBeChanged()"))
        self.arraydata = sorted(self.arraydata, key=operator.itemgetter(Ncol))
        if order == Qt.DescendingOrder:
            self.arraydata.reverse()
        self.emit(SIGNAL("layoutChanged()"))

    # ~~~~~~~~~~~ Getters and Setters ~~~~~~~~~~~
    def get_annotations(self, row):
        return self.arraydata[row][self.parent.ANNOTATIONS_COL]

    def get_articles(self, row):
        return self.arraydata[row][self.parent.ARTICLES_COL]

    def get_author(self, row):
        return self.arraydata[row][self.parent.AUTHOR_COL]

    def get_book_id(self, row):
        return self.arraydata[row][self.parent.BOOK_ID_COL]

    def get_calibre_id(self, row):
        return self.arraydata[row][self.parent.CALIBRE_ID_COL]

    def set_calibre_id(self, row, value):
        self.arraydata[row][self.parent.CALIBRE_ID_COL] = value

    def get_collections(self, row):
        return self.arraydata[row][self.parent.COLLECTIONS_COL]

    def set_collections(self, row, value):
        self.arraydata[row][self.parent.COLLECTIONS_COL] = value
        self.parent.repaint()

    def get_deep_view(self, row):
        return self.arraydata[row][self.parent.DEEP_VIEW_COL]

    def set_deep_view(self, row, value):
        self.arraydata[row][self.parent.DEEP_VIEW_COL] = value
        self.parent.repaint()

    def get_flags(self, row):
        return self.arraydata[row][self.parent.FLAGS_COL]

    def set_flags(self, row, value):
        self.arraydata[row][self.parent.FLAGS_COL] = value
        #self.parent.repaint()

    def get_last_opened(self, row):
        return self.arraydata[row][self.parent.LAST_OPENED_COL]

    def get_locked(self, row):
        return self.arraydata[row][self.parent.LOCKED_COL]

    def set_locked(self, row, value):
        self.arraydata[row][self.parent.LOCKED_COL] = value
        #self.parent.repaint()

    def get_match_quality(self, row):
        return self.arraydata[row][self.parent.MATCHED_COL]

    def set_match_quality(self, row, value):
        self.arraydata[row][self.parent.MATCHED_COL] = value
        self.parent.repaint()

    def get_path(self, row):
        return self.arraydata[row][self.parent.PATH_COL]

    def get_progress(self, row):
        return self.arraydata[row][self.parent.PROGRESS_COL]

    def set_progress(self, row, value):
        self.arraydata[row][self.parent.PROGRESS_COL] = value
        #self.parent.repaint()

    def get_series(self, row):
        return self.arraydata[row][self.parent.SERIES_COL]

    def get_subjects(self, row):
        return self.arraydata[row][self.parent.SUBJECTS_COL]

    def get_title(self, row):
        return self.arraydata[row][self.parent.TITLE_COL]

    def get_uuid(self, row):
        return self.arraydata[row][self.parent.UUID_COL]

    def get_vocabulary(self, row):
        return self.arraydata[row][self.parent.VOCABULARY_COL]

    def get_word_count(self, row):
        return self.arraydata[row][self.parent.WORD_COUNT_COL]

    def set_word_count(self, row, value):
        self.arraydata[row][self.parent.WORD_COUNT_COL] = value
        self.parent.repaint()


class BookStatusDialog(SizePersistedDialog, Logger):
    '''
    '''
#     CANCEL_NOT_REQUESTED = 0
#     CANCEL_REQUESTED = 1
#     CANCEL_ACKNOWLEDGED = 2
    CHECKMARK = u"\u2713"
    CIRCLE_SLASH = u"\u20E0"
    DEFAULT_REFRESH_TEXT = 'Refresh custom columns'
    DEFAULT_REFRESH_TOOLTIP = "<p>Refresh custom column content in calibre for the selected books.<br/>Assign custom column mappings in the <i>Customize plugin…</i> dialog.</p>"
    HASH_CACHE_FS = "content_hashes.db"
    HIGHLIGHT_COLORS = ['Pink', 'Yellow', 'Blue', 'Green', 'Purple']
    MATCH_COLORS = ['DARK_GRAY', 'LIGHT_GRAY', 'WHITE', 'RED', 'ORANGE', 'MAGENTA', 'YELLOW', 'GREEN']
    MATH_TIMES_CIRCLED = u" \u2297 "
    MATH_TIMES = u" \u00d7 "
    MAX_BOOKS_BEFORE_SPINNER = 4
    MAX_ELEMENT_DEPTH = 6
    REMOTE_CACHE_FOLDER = '/'.join(['/Library', 'calibre.mm'])
    UPDATING_MARVIN_MESSAGE = "Updating Marvin Library…"
    UTF_8_BOM = r'\xef\xbb\xbf'
    WATCHDOG_TIMEOUT = 10.0

    # Flag constants
    if True:
        FLAGS = {
            'new': 'NEW',
            'read': 'READ',
            'reading_list': 'READING LIST'
        }

        # Binary values for flag updates
        NEW_FLAG = 4
        READING_FLAG = 2
        READ_FLAG = 1

    # Column assignments. When changing order here, also change in _construct_table_data
    if True:
        LIBRARY_HEADER = [
                          'Title', 'Author', 'Series',
                          'Word count', 'Date added', 'Progress', 'Last read',
                          'Subjects', 'Collections', MATH_TIMES, 'Flags',
                          'Ann', 'Voc', 'DV', 'Art',
                          'Match Quality', 'uuid', 'cid', 'mid', 'path']
        ANNOTATIONS_COL = LIBRARY_HEADER.index('Ann')
        ARTICLES_COL = LIBRARY_HEADER.index('Art')
        AUTHOR_COL = LIBRARY_HEADER.index('Author')
        BOOK_ID_COL = LIBRARY_HEADER.index('mid')
        CALIBRE_ID_COL = LIBRARY_HEADER.index('cid')
        COLLECTIONS_COL = LIBRARY_HEADER.index('Collections')
        DATE_ADDED_COL = LIBRARY_HEADER.index('Date added')
        DEEP_VIEW_COL = LIBRARY_HEADER.index('DV')
        FLAGS_COL = LIBRARY_HEADER.index('Flags')
        LAST_OPENED_COL = LIBRARY_HEADER.index('Last read')
        LOCKED_COL = LIBRARY_HEADER.index(MATH_TIMES)
        MATCHED_COL = LIBRARY_HEADER.index('Match Quality')
        PATH_COL = LIBRARY_HEADER.index('path')
        PROGRESS_COL = LIBRARY_HEADER.index('Progress')
        TITLE_COL = LIBRARY_HEADER.index('Title')
        SERIES_COL = LIBRARY_HEADER.index('Series')
        SUBJECTS_COL = LIBRARY_HEADER.index('Subjects')
        UUID_COL = LIBRARY_HEADER.index('uuid')
        VOCABULARY_COL = LIBRARY_HEADER.index('Voc')
        WORD_COUNT_COL = LIBRARY_HEADER.index('Word count')

        HIDDEN_COLUMNS = [
            UUID_COL,
            CALIBRE_ID_COL,
            BOOK_ID_COL,
            PATH_COL,
            MATCHED_COL,
        ]
        CENTERED_COLUMNS = [
            ANNOTATIONS_COL,
            ARTICLES_COL,
            COLLECTIONS_COL,
            DEEP_VIEW_COL,
            LAST_OPENED_COL,
            VOCABULARY_COL,
        ]
        RIGHT_ALIGNED_COLUMNS = [
            PROGRESS_COL,
            WORD_COUNT_COL
        ]

    # User-controlled columns. Text is displayed in header context menu
    if True:
        USER_CONTROLLED_COLUMNS = [
            (AUTHOR_COL, 'Author'),
            (SERIES_COL, 'Series'),
            (WORD_COUNT_COL, 'Word count'),
            (DATE_ADDED_COL, 'Date added'),
            (PROGRESS_COL, 'Progress'),
            (LAST_OPENED_COL, 'Last read'),
            (SUBJECTS_COL, 'Subjects'),
            (COLLECTIONS_COL, 'Collections'),
            (FLAGS_COL, 'Flags'),
            (ANNOTATIONS_COL, 'Annotations'),
            (VOCABULARY_COL, 'Vocabulary'),
            (DEEP_VIEW_COL, 'Deep View'),
            (ARTICLES_COL, 'Articles'),
            ]

    # Marvin XML command template
    if True:
        METADATA_COMMAND_XML = b'''\xef\xbb\xbf<?xml version='1.0' encoding='utf-8'?>
        <{0} timestamp=\'{1}\'>
        <manifest>
        </manifest>
        </{0}>'''

        GENERAL_COMMAND_XML = b'''\xef\xbb\xbf<?xml version='1.0' encoding='utf-8'?>
        <command type=\'{0}\' timestamp=\'{1}\'>
        </command>'''

    marvin_device_status_changed = pyqtSignal(dict)

    def accept(self):
        self._log_location()
        self._save_column_widths()
        super(BookStatusDialog, self).accept()

    def busy_cancel_click(self):
        '''
        Capture the click, disable the button
        '''
        self._log_location()
        self.busy_cancel_requested = True
        self._busy_status_msg(msg="Cancelling, please wait…")
        self.busy_cancel_button.setEnabled(False)

    def capture_sort_column(self, sort_column):
        sort_order = self.tv.horizontalHeader().sortIndicatorOrder()
        self.opts.prefs.set('marvin_library_sort_column', sort_column)
        self.opts.prefs.set('marvin_library_sort_order', sort_order)

    def close(self):
        self._log_location()
        self._save_column_widths()
        super(BookStatusDialog, self).close()

    def dispatch_button_click(self, button):
        '''
        BUTTON_ROLES = ['AcceptRole', 'RejectRole', 'DestructiveRole', 'ActionRole',
                        'HelpRole', 'YesRole', 'NoRole', 'ApplyRole', 'ResetRole']
        '''
        self._log_location()
        if self.dialogButtonBox.buttonRole(button) == QDialogButtonBox.AcceptRole:
            self.accept()

        elif self.dialogButtonBox.buttonRole(button) == QDialogButtonBox.ActionRole:
            if button.objectName() == 'match_colors_button':
                self.toggle_match_colors()

            elif button.objectName() == 'manage_collections_button':
                self.show_manage_collections_dialog()

            elif button.objectName() == 'view_collections_button':
                selected_rows = self._selected_rows()
                if selected_rows:
                    self.show_view_collections_dialog(selected_rows[0])
                else:
                    title = "View collections"
                    msg = "<p>Select a book.</p>"
                    MessageBox(MessageBox.INFO, title, msg,
                               show_copy_button=False).exec_()

            elif button.objectName() == 'refresh_custom_columns_button':
                self.refresh_custom_columns()

            elif button.objectName() == 'view_global_vocabulary_button':
                self.show_html_dialog('show_global_vocabulary', 0)

            elif button.objectName() == 'view_metadata_button':
                selected_rows = self._selected_rows()
                if selected_rows:
                    self.show_view_metadata_dialog(selected_rows[0])
                else:
                    title = "View metadata"
                    msg = "<p>Select a book.</p>"
                    MessageBox(MessageBox.INFO, title, msg,
                               show_copy_button=False).exec_()

        elif self.dialogButtonBox.buttonRole(button) == QDialogButtonBox.DestructiveRole:
            self._delete_books()

        elif self.dialogButtonBox.buttonRole(button) == QDialogButtonBox.HelpRole:
            self.show_help()

        elif self.dialogButtonBox.buttonRole(button) == QDialogButtonBox.RejectRole:
            self.close()

    def dispatch_context_menu_event(self, action, row):
        '''
        '''
        self._log_location("%s row: %s" % (repr(action), row))

        if action == 'apply_date_read':
            self._apply_date_read()
        elif action == 'apply_progress':
            self._apply_progress()
        elif action == 'apply_word_count':
            self._apply_word_count()
        elif action == 'calculate_word_count':
            self._calculate_word_count()
        elif action in ['clear_all_collections', 'export_collections',
                        'import_collections', 'synchronize_collections']:
            self._update_collections(action)
        elif action in ['clear_new_flag', 'clear_reading_list_flag',
                        'clear_read_flag', 'clear_all_flags',
                        'set_new_flag', 'set_reading_list_flag', 'set_read_flag']:
            self._update_flags(action)
        elif action in ['export_metadata', 'import_metadata']:
            self._update_metadata(action)
        elif action == 'fetch_annotations':
            self._fetch_annotations(report_results=True)
        elif action == 'generate_deep_view':
            self._generate_deep_view()
        elif action == 'manage_collections':
            self.show_manage_collections_dialog()
        elif action in ['set_locked', 'set_unlocked']:
            self._update_locked_status(action)

        elif action in ['show_deep_view_articles',
                        'show_deep_view_alphabetically', 'show_deep_view_by_importance',
                        'show_deep_view_by_appearance', 'show_deep_view_by_annotations',
                        'show_vocabulary']:
            self.show_html_dialog(action, row)
        elif action == 'show_collections':
            self.show_view_collections_dialog(row)
        elif action == 'show_global_vocabulary':
            self.show_html_dialog('show_global_vocabulary', row)
        elif action == 'show_highlights':
            self.show_annotations(row)
        elif action == 'show_metadata':
            self.show_view_metadata_dialog(row)
        elif action == 'synchronize_flags':
            self._synchronize_flags()
        else:
            selected_books = self._selected_books()
            det_msg = ''
            for row in selected_books:
                det_msg += selected_books[row]['title'] + '\n'

            title = "Context menu event"
            msg = ("<p>{0}</p>".format(action) +
                   "<p>Click <b>Show details</b> for affected books</p>")

            MessageBox(MessageBox.INFO, title, msg, det_msg=det_msg,
                       show_copy_button=False).exec_()

    def dispatch_double_click(self, index):
        '''
        Display column data for selected book
        '''
        self._log_location()

        asset_actions = {
            self.ARTICLES_COL: 'show_deep_view_articles',
            self.DEEP_VIEW_COL: 'show_deep_view_by_importance',
            self.VOCABULARY_COL: 'show_vocabulary'
            }

        column = index.column()
        row = index.row()

        if column in [self.TITLE_COL, self.AUTHOR_COL]:
            self.show_view_metadata_dialog(row)

        elif column in [self.ARTICLES_COL, self.VOCABULARY_COL]:
            self.show_html_dialog(asset_actions[column], row)

        elif column == self.ANNOTATIONS_COL:
            self.show_annotations(row)

        elif column == self.DEEP_VIEW_COL:
            # If no DV content, generate DV content, else show it
            has_dv_content = self._selected_books()[row]['has_dv_content']
            if has_dv_content:
                self.show_html_dialog(asset_actions[column], row)
            else:
                self._generate_deep_view()

        elif column == self.COLLECTIONS_COL:
            self.show_view_collections_dialog(row)

        elif column in [self.FLAGS_COL]:
            title = "Flag options"
            msg = "<p>Right-click in the Flags column for flag management options.</p>"
            MessageBox(MessageBox.INFO, title, msg,
                       show_copy_button=False).exec_()

        elif column == self.LOCKED_COL:
            self._toggle_locked_status(row)

        elif column == self.WORD_COUNT_COL:
            self._calculate_word_count()

        else:
            self._log("no double-click handler for %s" % self.LIBRARY_HEADER[column])

    def esc(self, *args):
        '''
        Clear any active selections, filter
        '''
        self._log_location()
        self._clear_selected_rows()
        self.filter_clear()

    def filter_clear(self):
        '''
        Clear the filter, show all rows
        '''
        self._log_location()
        self.filter_le.clear()
        total_books = len(self.tm.all_rows())
        for i in range(total_books):
            self.tv.showRow(i)

        # Restore clickability
        self.tv.horizontalHeader().setClickable(True)

    def filter_table_rows(self, qstr):
        '''
        Hide rows not matching filter
        '''
        pattern = str(qstr).strip()
        if pattern == '':
            self.filter_clear()
            return

        self._log_location(pattern)
        total_books = len(self.tm.all_rows())
        for i in range(total_books):
            matched = False
            if re.search(pattern, self.tm.get_title(i).text(), re.IGNORECASE):
                matched = True
            elif re.search(pattern, self.tm.get_author(i).text(), re.IGNORECASE):
                matched = True
            elif re.search(pattern, self.tm.get_series(i).text(), re.IGNORECASE):
                matched = True
            elif re.search(pattern, self.tm.get_subjects(i).text(), re.IGNORECASE):
                matched = True

            if matched:
                self.tv.showRow(i)
            else:
                self.tv.hideRow(i)

        # Prevent sorting on cols, because we'll lose reference to the matched rows
        self.tv.horizontalHeader().setClickable(False)

    def initialize(self, parent):
        self.busy = False
        self.busy_cancel_requested = False
        self.busy_panel = None
        self.Dispatcher = partial(Dispatcher, parent=self)
        self.hash_cache = None
        self.icon = get_icon(parent.icon)
        self.ios = parent.ios
        self.installed_books = None
        self.opts = parent.opts
        self.parent = parent
        self.prefs = parent.opts.prefs
        self.library_scanner = parent.library_scanner
        self.library_title_map = None
        self.library_uuid_map = None
        self.local_cache_folder = self.parent.connected_device.temp_dir
        self.local_hash_cache = None
        self.marvin_cancellation_required = False
        self.remote_hash_cache = None
        self.show_match_colors = self.prefs.get('show_match_colors', False)
        self.soloed_books = set()
        self.updated_match_quality = None
        self.verbose = parent.verbose

        # Device-specific cover_hash cache
        device_cached_hashes = "plugins/Marvin_XD_resources/{0}_cover_hashes".format(
            re.sub('\W', '_', self.ios.device_name))
        self.archived_cover_hashes = JSONConfig(device_cached_hashes)

        # Subscribe to Marvin driver change events
        self.parent.connected_device.marvin_device_signals.reader_app_status_changed.connect(
            self.marvin_status_changed)

        self._log_location()

        self.installed_books = self._generate_booklist()

        self._busy_panel_setup("Preparing Marvin library view…")

        # ~~~~~~~~ Create the dialog ~~~~~~~~
        self.setWindowTitle(u'Marvin Library: %d books' % len(self.installed_books))
        self.setWindowIcon(self.icon)
        self.l = QVBoxLayout(self)
        self.setLayout(self.l)
        self.perfect_width = 0

        # ~~~~~~~~ Create the filter~~~~~~~~
        self.filter_hb = QHBoxLayout()

        # Line edit
        self.filter_le = QLineEdit()
        #self.filter_le.setFrame(False)
        self.filter_le.textEdited.connect(self.filter_table_rows)
        self.filter_le.setPlaceholderText("Filter by Title, Author, Series or Subject")
        self.filter_le.setToolTip("Filter books by Title, Author, Series or Subject")
        self.filter_hb.addWidget(self.filter_le)

        # Clear button
        self.filter_tb = QToolButton()
        self.filter_tb.setIcon(QIcon(I('clear_left.png')))
        self.filter_tb.setToolTip("Clear filter")
        self.filter_tb.clicked.connect(self.filter_clear)
        self.filter_hb.addWidget(self.filter_tb)

        # Spacer
        self.filter_spacer = QSpacerItem(16, 16, QSizePolicy.Expanding, QSizePolicy.Minimum)
        self.filter_hb.addItem(self.filter_spacer)

        # Busy status
        self.busy_status_label = QLabel('')
        self.filter_hb.addWidget(self.busy_status_label, 0, Qt.AlignRight)

        self.filter_hb.addSpacing(10)
        self.busy_cancel_button = QPushButton(QIcon(I('window-close.png')), 'Cancel')
        self.filter_hb.addWidget(self.busy_cancel_button, 0, Qt.AlignHCenter)
        self.busy_cancel_button.setVisible(False)
        self.busy_cancel_button.clicked.connect(self.busy_cancel_click)

        self.filter_hb.addSpacing(10)
        self.busy_status_pi = ProgressIndicator(self)
        self.busy_status_pi.setDisplaySize(24)
        self.filter_hb.addWidget(self.busy_status_pi, 0, Qt.AlignHCenter)
        self.busy_status_pi.setVisible(False)
        self.filter_hb.addSpacing(10)

        self.l.addLayout(self.filter_hb)

        # ~~~~~~~~ Create the Table ~~~~~~~~
        self.tv = MyTableView(self)
        self.l.addWidget(self.tv)

        self.tabledata = self._construct_table_data()
        self._construct_table_view()

        # Set the width of the filter control after we know the size of the other cols
        saved_column_widths = self.opts.prefs.get('marvin_library_column_widths')
        filter_width = 0
        for index in [self.TITLE_COL, self.AUTHOR_COL, self.SERIES_COL]:
            filter_width += saved_column_widths[index]
        self.filter_le.setFixedWidth(filter_width)

        # ~~~~~~~~ Create the ButtonBox ~~~~~~~~
        self.dialogButtonBox = QDialogButtonBox(QDialogButtonBox.Help)
        self.l.addWidget(self.dialogButtonBox)

        # Delete button
        if False:
            self.delete_button = self.dialogButtonBox.addButton(self.dialogButtonBox.Discard)
            self.delete_button.setText('Delete')

        # Done button
        self.done_button = self.dialogButtonBox.addButton(self.dialogButtonBox.Ok)
        self.done_button.setText('Close')

        self.dialogButtonBox.setOrientation(Qt.Horizontal)
        self.dialogButtonBox.setCenterButtons(False)

        # Show/Hide Match Quality
        self.show_match_colors_button = self.dialogButtonBox.addButton("undefined", QDialogButtonBox.ActionRole)
        self.show_match_colors_button.setObjectName('match_colors_button')
        self.show_match_colors = not self.show_match_colors
        self.toggle_match_colors()

        # Manage collections
        if True:
            self.mc_button = self.dialogButtonBox.addButton('Manage collections', QDialogButtonBox.ActionRole)
            self.mc_button.setObjectName('manage_collections_button')
            self.mc_button.setIcon(QIcon(os.path.join(self.parent.opts.resources_path,
                                                      'icons',
                                                      'edit_collections.png')))

        # Create the Refresh button
        button_title = self.DEFAULT_REFRESH_TEXT
        self.refresh_button = self.dialogButtonBox.addButton(button_title, QDialogButtonBox.ActionRole)
        self.refresh_button.setObjectName('refresh_custom_columns_button')
        self.refresh_button.setIcon(QIcon(os.path.join(self.parent.opts.resources_path,
                                                       'icons',
                                                       'sync_collections.png')))
        self.refresh_button.setToolTip(self.DEFAULT_REFRESH_TOOLTIP)
        self.refresh_button.setEnabled(False)

        # Apply proper text to Refresh button
        self._update_refresh_button()

        # View Global vocabulary
        if True:
            self.gv_button = self.dialogButtonBox.addButton('View all vocabulary words', QDialogButtonBox.ActionRole)
            self.gv_button.setObjectName('view_global_vocabulary_button')
            self.gv_button.setIcon(QIcon(I('books_in_series.png')))

        self.dialogButtonBox.clicked.connect(self.dispatch_button_click)

        # ~~~~~~~~ Connect signals ~~~~~~~~
        self.connect(self.tv, SIGNAL("doubleClicked(QModelIndex)"), self.dispatch_double_click)
        self.connect(self.tv.horizontalHeader(), SIGNAL("sectionClicked(int)"), self.capture_sort_column)

        self.resize_dialog()
        self.tv.setFocus()

        self._busy_panel_teardown()

        if self.parent.prefs.get('auto_refresh_at_startup', False):
            self._busy_panel_setup("Refreshing custom column content…")
            self.refresh_custom_columns(all_books=True, report_results=False)
            self._busy_panel_teardown()
            self._clear_selected_rows()

        # Report duplicates, updated, set temporary markers according to prefs
        self.soloed_books = set()
        self.parent.gui.library_view.model().db.set_marked_ids(self.soloed_books)
        self._report_calibre_duplicates()
        self._report_content_updates()

    def launch_collections_scanner(self):
        '''
        Invoke InventoryCollections to identify cids with collection assignments
        After indexing, self.library_collections.ids is list of cids
        '''
        self._log_location()
        self.library_collections = InventoryCollections(self)
        self.connect(self.library_collections, self.library_collections.signal, self.library_collections_complete)
        QTimer.singleShot(0, self.start_library_collections_inventory)

        if False:
            # Wait for scan to start
            while not self.library_collections.isRunning():
                Application.processEvents()

    def library_collections_complete(self):
        self._log_location()
        self._log("ids: %s" % repr(self.library_collections.ids))

    def marvin_status_changed(self, cmd_dict):
        '''

        '''
        self.marvin_device_status_changed.emit(cmd_dict)
        command = cmd_dict['cmd']

        self._log_location(command)

        if command in ['disconnected', 'yanked']:
            self._log("closing dialog: %s" % command)
            self.close()

    def refresh_custom_columns(self, all_books=False, report_results=True):
        '''
        Refresh enabled custom columns from Marvin content
        '''
        self._log_location()

        self.saved_selection_region = self.tv.visualRegionForSelection(self.tv.selectionModel().selection())

        enabled = []
        for cfn in ['annotations', 'date_read', 'progress', 'read',
            'reading_list','word_count']:
            cfv = get_cc_mapping(cfn, 'combobox', None)
            if cfv:
                enabled.append(cfv)
        cols_to_refresh = ', '.join(sorted(enabled, key=sort_key))


        if all_books:
            rows_to_refresh = list(reversed([i for i in range(len(self.tm.all_rows()))]))
        else:
            # Process selected books
            rows_to_refresh = sorted(self._selected_books())

        if rows_to_refresh:
            msg = "Refreshing %s for %s" % (cols_to_refresh,
                "1 book…" if len(rows_to_refresh) == 1 else
                "%d books…" % len(rows_to_refresh))
            self._busy_status_setup(msg=msg, show_cancel=len(rows_to_refresh) > 1)

            for row in rows_to_refresh:
                if self.busy_cancel_requested:
                    break

                self.tv.selectRow(row)
                self._fetch_annotations(update_gui=False)
                self._apply_date_read(update_gui=False)
                self._apply_flags(update_gui=False)
                self._apply_progress(update_gui=False)
                self._apply_word_count(update_gui=False)

            # _apply_flags may have updated Marvin mainDb
            self._localize_marvin_database()

            updateCalibreGUIView()
            self._busy_status_teardown()

            # Restore selection
            if self.saved_selection_region:
                for rect in self.saved_selection_region.rects():
                    self.tv.setSelection(rect, QItemSelectionModel.Select)
                self.saved_selection_region = None

            # Report results
            if report_results:
                title = 'Custom columns refreshed'
                refreshed = ''
                for col in enabled[0:-1]:
                    refreshed += '<b>%s</b>, ' % col
                refreshed += '<b>%s</b> ' % enabled[-1]
                msg = "<p>%s refreshed for %s.</p>" % (refreshed,
                                               "1 book" if len(rows_to_refresh) == 1 else
                                               "%d books" % len(rows_to_refresh))
                MessageBox(MessageBox.INFO, title, msg, det_msg='', show_copy_button=False).exec_()

        else:
            # No rows selected, inform user how the feature works
            title = 'No books selected'
            msg = ('No books selected.\n' +
                    'To refresh custom columns, select one or more books, ' +
                    "then click the 'Refresh custom columns' button.")
            MessageBox(MessageBox.WARNING, title, msg, det_msg='', show_copy_button=False).exec_()

    def show_add_collections_dialog(self):
        '''
        Get a new collection name(s), add to selected books in Marvin
        '''
        self._log_location()

        klass = os.path.join(dialog_resources_path, 'add_collections.py')
        if os.path.exists(klass):
            #self._log("importing metadata dialog from '%s'" % klass)
            sys.path.insert(0, dialog_resources_path)
            this_dc = importlib.import_module('add_collections')
            sys.path.remove(dialog_resources_path)
            dlg = this_dc.AddCollectionsDialog(self, self.parent.connected_device)
            dlg.initialize()
            dlg.exec_()
            if dlg.result() == dlg.Accepted:
                raw = str(dlg.new_collection_le.text())
                raw = raw.replace(',', ', ')
                added_collections = raw.split(', ')

                # Save the selection
                self.saved_selection_region = self.tv.visualRegionForSelection(self.tv.selectionModel().selection())

                selected_books = self._selected_books()
                for row in sorted(selected_books):
                    self.tv.selectRow(row)
                    book_id = selected_books[row]['book_id']
                    original_collections = self._get_marvin_collections(book_id)
                    updated_collections = sorted(original_collections + added_collections, key=sort_key)
                    self._update_marvin_collections(book_id, updated_collections)
                    self._update_collection_match(self.installed_books[book_id], row)

                # Restore the selection
                for rect in self.saved_selection_region.rects():
                    self.tv.setSelection(rect, QItemSelectionModel.Select)
                self.saved_selection_region = None

                # Update local_db for all changes
                #self._localize_marvin_database()

            else:
                self._log("User cancelled Add collections dialog")
        else:
            self._log("ERROR: Can't import from '%s'" % klass)

    def show_annotations(self, row):
        '''
        '''
        HTML_TEMPLATE = (
            '<?xml version=\'1.0\' encoding=\'utf-8\'?>' +
            '<html xmlns="http://www.w3.org/1999/xhtml">' +
            '<head>' +
            '<meta http-equiv="content-type" content="text/html; charset=utf-8"/>' +
            '<title>Annotations</title>' +
            '</head>' +
            '<body>{0}</body>' +
            '</html>'
            )

        book_id = self._selected_book_id(row)
        title = self.installed_books[book_id].title
        self._log_location(title)

        refresh = None

        if not self.installed_books[book_id].highlights:
            self._log("No annotations available for %s" % repr(title))
            return

        header = None
        group_box_title = 'Annotations'
        annotations = self._get_formatted_annotations(book_id)

        footer = None
        afn = get_cc_mapping('annotations', 'combobox', None)
        if afn:
            refresh = {
                'name': afn,
                'method': "_fetch_annotations"
                }

        content_dict = {
            'footer': footer,
            'group_box_title': group_box_title,
            'header': header,
            'html_content': HTML_TEMPLATE.format(annotations),
            'title': title,
            'toolTip': '<p>Annotations appearance may be fine-tuned in the <b>Customize plugin…</b> dialog</p>',
            'refresh': refresh
            }

        klass = os.path.join(dialog_resources_path, 'html_viewer.py')
        if os.path.exists(klass):
            #self._log("importing metadata dialog from '%s'" % klass)
            sys.path.insert(0, dialog_resources_path)
            this_dc = importlib.import_module('html_viewer')
            sys.path.remove(dialog_resources_path)
            dlg = this_dc.HTMLViewerDialog(self, 'html_viewer')
            dlg.initialize(self,
                           content_dict,
                           book_id,
                           self.installed_books[book_id],
                           self.parent.connected_device.local_db_path)
            dlg.exec_()

        else:
            self._log("ERROR: Can't import from '%s'" % klass)

    def show_help(self):
        '''
        Display help file
        '''
        self.parent.show_help()

    def show_html_dialog(self, action, row):
        '''
        Display assets associated with book
        Articles, Deep View, Vocabulary
        profile = {'title': <dlg title>,
                   'group_box_title':<gb title>,
                   'header': <header text>,
                   'content': <default content>,
                   'footer': <footer text>}
        '''
        self._log_location(action)

        book_id = self._selected_book_id(row)
        title = self.installed_books[book_id].title
        refresh = None

        if action == 'show_deep_view_articles':
            if not self.installed_books[book_id].articles:
                return

            command_name = "command"
            command_type = "GetDeepViewArticlesHTML"
            update_soup = BeautifulStoneSoup(self.GENERAL_COMMAND_XML.format(
                command_type, time.mktime(time.localtime())))
            parameters_tag = self._build_parameters(self.installed_books[book_id], update_soup)
            update_soup.command.insert(0, parameters_tag)

            header = None
            group_box_title = 'Deep View articles'
            default_content = ("<p>Deep View articles provided by Marvin…</p>")
            footer = None

        elif action in ('show_deep_view_alphabetically', 'show_deep_view_by_importance',
                        'show_deep_view_by_appearance', 'show_deep_view_by_annotations'):

            if not bool(self.tm.get_deep_view(row)):
                self._log("no DV content for %s" % title)
                return

            command_name = "command"
            command_type = "GetFirstOccurrenceHTML"
            update_soup = BeautifulStoneSoup(self.GENERAL_COMMAND_XML.format(
                command_type, time.mktime(time.localtime())))
            parameters_tag = Tag(update_soup, 'parameters')
            update_soup.command.insert(0, parameters_tag)

            # The bookID we know - entityID and hits we get from the user
            parameter_tag = Tag(update_soup, 'parameter')
            parameter_tag['name'] = "bookID"
            parameter_tag.insert(0, str(book_id))
            parameters_tag.insert(0, parameter_tag)

            header = None
            if action == 'show_deep_view_alphabetically':
                group_box_title = "Deep View items alphabetically"
                sort_order = "E.Name COLLATE NOCASE ASC"

            elif action == 'show_deep_view_by_importance':
                group_box_title = "Deep View items by importance"
                sort_order = "Cnt DESC, E.Name ASC"

            elif action == 'show_deep_view_by_appearance':
                group_box_title = "Deep View items by appearance"
                sort_order = "Loc ASC, E.Name ASC"

            elif action == 'show_deep_view_by_annotations':
                group_box_title = "Deep View items (notes and flags first)"
                sort_order = "NoteFlagOrder DESC, Cnt DESC, E.Name ASC"

            default_content = "{0} to be provided by Marvin.".format(group_box_title)
            footer = None

            # Get a list of DV items by querying mainDb
            entities = "Entities_%d" % book_id
            entity_locations = "EntityLocations_%d" % book_id
            con = sqlite3.connect(self.parent.connected_device.local_db_path)
            with con:
                con.row_factory = sqlite3.Row
                dv_names_cur = con.cursor()
                if action == "show_deep_view_by_annotations":
                    dv_names_cur.execute('''SELECT
                                             E.ID,
                                             E.Name,
                                             COUNT(L.SentenceIndex) AS Cnt,
                                             MIN(L.SectionIndex * 1000000 + L.SentenceIndex) AS Loc,
                                             Flag,
                                             Note,
                                             E.Confidence,
                                            CASE WHEN Note IS NULL THEN 0 ELSE 2 END + Flag AS NoteFlagOrder
                                            FROM {0} as E JOIN {1} AS L ON E.ID = L.EntityID
                                            GROUP BY E.ID
                                            ORDER BY {2}
                                         '''.format(entities, entity_locations, sort_order))
                else:
                    dv_names_cur.execute('''SELECT
                                             E.ID,
                                             E.Name,
                                             COUNT(L.SentenceIndex) AS Cnt,
                                             MIN(L.SectionIndex * 1000000 + L.SentenceIndex) as Loc,
                                             Flag,
                                             Note,
                                             E.Confidence
                                            FROM {0} as E JOIN {1} AS L ON E.ID = L.EntityID
                                            GROUP BY E.ID
                                            ORDER BY {2}
                                         '''.format(entities, entity_locations, sort_order))

                rows = dv_names_cur.fetchall()
                dv_names_cur.close()
            klass = os.path.join(dialog_resources_path, 'deep_view_items.py')
            if os.path.exists(klass):
                sys.path.insert(0, dialog_resources_path)
                this_dc = importlib.import_module('deep_view_items')
                sys.path.remove(dialog_resources_path)
                dlg = this_dc.DeepViewItems(self, group_box_title, rows)
                dlg.exec_()

                if dlg.result:
                    # entityID
                    parameter_tag = Tag(update_soup, 'parameter')
                    parameter_tag['name'] = "entityID"
                    parameter_tag.insert(0, dlg.result['ID'])
                    parameters_tag.insert(0, parameter_tag)

                    # hits
                    parameter_tag = Tag(update_soup, 'parameter')
                    parameter_tag['name'] = "hits"
                    parameter_tag.insert(0, dlg.result['hits'])
                    parameters_tag.insert(0, parameter_tag)

                    group_box_title = "Deep View hits for %s" % dlg.result['item']
                else:
                    return

        elif action == 'show_global_vocabulary':
            command_name = "command"
            command_type = "GetGlobalVocabularyHTML"
            update_soup = BeautifulStoneSoup(self.GENERAL_COMMAND_XML.format(
                command_type, time.mktime(time.localtime())))

            title = "All vocabulary words"
            header = None
            group_box_title = "Vocabulary words by book"
            default_content = "<p>No Global vocabulary list returned by Marvin.</p>"
            footer = None

        elif action == 'show_vocabulary':
            if not self.installed_books[book_id].vocabulary:
                return

            command_name = "command"
            command_type = "GetLocalVocabularyHTML"
            update_soup = BeautifulStoneSoup(self.GENERAL_COMMAND_XML.format(
                command_type, time.mktime(time.localtime())))
            parameters_tag = self._build_parameters(self.installed_books[book_id], update_soup)
            update_soup.command.insert(0, parameters_tag)

            header = None
            group_box_title = 'Vocabulary words'
            if self.installed_books[book_id].vocabulary:
                word_list = '<br/>'.join(sorted(self.installed_books[book_id].vocabulary, key=sort_key))
                default_content = "<p>{0}</p>".format(word_list)
            else:
                default_content = ("<p>No vocabulary words</p>")
            footer = None

        else:
            self._log("ERROR: unsupported action '%s'" % action)
            return

        self._busy_status_setup(msg="Retrieving %s…" % group_box_title)
        results = self._issue_command(command_name, update_soup,
                                      get_response="html_response.html",
                                      update_local_db=False)
        self._busy_status_teardown()

        if results['code']:
            return self._show_command_error(command_type, results)
        else:
            response = results['response']

        if response:
            # <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
            if re.match(self.UTF_8_BOM, response):
                u_response = UnicodeDammit(response).unicode
                response = self._inject_css(u_response).encode('utf-8')
            response = "<?xml version='1.0' encoding='utf-8'?>" + response
        else:
            response = default_content

        content_dict = {
            'footer': footer,
            'group_box_title': group_box_title,
            'header': header,
            'html_content': response,
            'title': title,
            'toolTip': None,
            'refresh': refresh
            }


        klass = os.path.join(dialog_resources_path, 'html_viewer.py')
        if os.path.exists(klass):
            #self._log("importing metadata dialog from '%s'" % klass)
            sys.path.insert(0, dialog_resources_path)
            this_dc = importlib.import_module('html_viewer')
            sys.path.remove(dialog_resources_path)
            dlg = this_dc.HTMLViewerDialog(self, 'html_viewer')
            dlg.initialize(self,
                           content_dict,
                           book_id,
                           self.installed_books[book_id],
                           self.parent.connected_device.local_db_path
                           )
            dlg.exec_()

        else:
            self._log("ERROR: Can't import from '%s'" % klass)

    def show_manage_collections_dialog(self):
        '''
        Present all active collection names, allow edit/deletion
        Marvin changes applied immediately to mainDb.
        Returns a dict of original, changed so we can update custom column assignments,
        connected device
        '''
        self._log_location()

        # Build an inventory of cids with collection assignments
        self.launch_collections_scanner()

        current_collections = {}

        # Get all Marvin collection names
        con = sqlite3.connect(self.parent.connected_device.local_db_path)
        with con:
            con.row_factory = sqlite3.Row

            collections_cur = con.cursor()
            collections_cur.execute('''SELECT
                                        Name
                                       FROM Collections
                                    ''')
            rows = collections_cur.fetchall()
            collections_cur.close()

        marvin_collection_list = []
        if len(rows):
            marvin_collection_list = [row[b'Name'] for row in rows]
            marvin_collection_list = sorted(marvin_collection_list, key=sort_key)
            current_collections['Marvin'] = marvin_collection_list

        # Get all calibre collection names
        calibre_collection_list = []
        cfl = get_cc_mapping('collections', 'field', None)
        if cfl:
            db = self.opts.gui.current_db
            calibre_collection_list = db.all_custom(db.field_metadata.key_to_label(cfl))
            current_collections['calibre'] = sorted(calibre_collection_list, key=sort_key)

        if current_collections:
            if True:
                klass = os.path.join(dialog_resources_path, 'manage_collections.py')
                if os.path.exists(klass):
                    sys.path.insert(0, dialog_resources_path)
                    this_dc = importlib.import_module('manage_collections')
                    sys.path.remove(dialog_resources_path)

                    dlg = this_dc.MyDeviceCategoryEditor(self, tag_to_match=None,
                                                         data=current_collections, key=sort_key,
                                                         connected_device=self.parent.connected_device)
                    dlg.exec_()
                    if dlg.result() == dlg.Accepted:
                        if self.library_collections.isRunning():
                            self.library_collections.wait()

                        details = {'rename': dlg.to_rename,
                                   'delete': dlg.to_delete,
                                   'active_cids': self.library_collections.ids,
                                   'locations': current_collections}

                        self._busy_status_setup(msg="Updating collections…")
                        self._update_global_collections(details)
                        self._busy_status_teardown()

                else:
                    self._log("ERROR: Can't import from '%s'" % klass)

            else:
                klass = os.path.join(dialog_resources_path, 'manage_collections.py')
                if os.path.exists(klass):
                    sys.path.insert(0, dialog_resources_path)
                    this_dc = importlib.import_module('manage_collections')
                    sys.path.remove(dialog_resources_path)
                    dlg = this_dc.CollectionsManagementDialog(self, 'collections_manager')

                    if self.library_collections.isRunning():
                        self.library_collections.wait()

                    dlg.initialize(self,
                                   current_collections,
                                   self.library_collections.ids,
                                   self.parent.connected_device)
                    dlg.exec_()
                else:
                    self._log("ERROR: Can't import from '%s'" % klass)
        else:
            title = "Manage collections"
            msg = "<p>No collections to manage.</p>"
            MessageBox(MessageBox.INFO, title, msg,
                       show_copy_button=False).exec_()

    def show_view_collections_dialog(self, row):
        '''
        Present collection assignments to user, get updates
        Updated calibre assignments need to be sent to custom column.
        Device model and cached_books updated with updated collections + current flags.
        Updated Marvin assignments need to be sent to Marvin
        '''
        self._log_location(row)
        cid = self._selected_cid(row)
        book_id = self._selected_book_id(row)

        original_calibre_collections = self._get_calibre_collections(cid)
        original_marvin_collections = self._get_marvin_collections(book_id)

        if original_calibre_collections == [] and original_marvin_collections == []:
            title = self.installed_books[book_id].title
            msg = "<p>This book is not assigned to any collections.</p>"
            MessageBox(MessageBox.INFO, title, msg,
                       show_copy_button=False).exec_()
        else:
            klass = os.path.join(dialog_resources_path, 'view_collections.py')
            if os.path.exists(klass):
                #self._log("importing metadata dialog from '%s'" % klass)
                sys.path.insert(0, dialog_resources_path)
                this_dc = importlib.import_module('view_collections')
                sys.path.remove(dialog_resources_path)
                dlg = this_dc.CollectionsViewerDialog(self, 'collections_viewer')
                cid = self._selected_cid(row)
                dlg.initialize(self,
                               self.installed_books[book_id].title,
                               original_calibre_collections,
                               original_marvin_collections,
                               self.parent.connected_device)
                dlg.exec_()
                if dlg.result() == dlg.Accepted:
                    updated_calibre_collections = dlg.results['updated_calibre_collections']
                    updated_marvin_collections = dlg.results['updated_marvin_collections']

                    if (original_calibre_collections == updated_calibre_collections and
                            original_marvin_collections == updated_marvin_collections):
                        self._log("no collection changes detected")
                    else:
                        if updated_calibre_collections != original_calibre_collections:
                            self._log("original_calibre_collections: %s" % original_calibre_collections)
                            self._log("updated_calibre_collections: %s" % updated_calibre_collections)
                            self._update_calibre_collections(book_id, cid, updated_calibre_collections)

                        if updated_marvin_collections != original_marvin_collections:
                            self._log("original_marvin_collections: %s" % original_marvin_collections)
                            self._log("updated_marvin_collections: %s" % updated_marvin_collections)
                            self._update_marvin_collections(book_id, updated_marvin_collections)

                        # Update collections match status
                        self._update_collection_match(self.installed_books[book_id], self._selected_rows()[0])
                else:
                    self._log("User cancelled Collections dialog")
            else:
                self._log("ERROR: Can't import from '%s'" % klass)

    def show_view_metadata_dialog(self, row):
        '''
        '''
        self._log_location(row)
        cid = self._selected_cid(row)
        klass = os.path.join(dialog_resources_path, 'view_metadata.py')
        if os.path.exists(klass):
            #self._log("importing metadata dialog from '%s'" % klass)
            sys.path.insert(0, dialog_resources_path)
            this_dc = importlib.import_module('view_metadata')
            sys.path.remove(dialog_resources_path)
            dlg = this_dc.MetadataComparisonDialog(self, 'metadata_comparison')
            book_id = self._selected_book_id(row)
            cid = self._selected_cid(row)
            mismatches = self.installed_books[book_id].metadata_mismatches
            enable_metadata_updates = self.tm.get_match_quality(row) >= self.MATCH_COLORS.index('YELLOW')

            dlg.initialize(self,
                           book_id,
                           cid,
                           self.installed_books[book_id],
                           enable_metadata_updates,
                           self.parent.connected_device.local_db_path)
            dlg.exec_()
            if dlg.result() == dlg.Accepted and mismatches:
                action = dlg.stored_command
                self._update_metadata(action)
            else:
                self._log("User cancelled Metadata dialog")

        else:
            self._log("ERROR: Can't import from '%s'" % klass)

    def size_hint(self):
        return QtCore.QSize(self.perfect_width, self.height())

    def start_library_collections_inventory(self):
        self.library_collections.start()

    def toggle_match_colors(self):
        self.show_match_colors = not self.show_match_colors
        self.opts.prefs.set('show_match_colors', self.show_match_colors)
        if self.show_match_colors:
            self.show_match_colors_button.setText("Hide match status")
            self.tv.sortByColumn(self.LIBRARY_HEADER.index('Match Quality'), Qt.DescendingOrder)
            self.capture_sort_column(self.LIBRARY_HEADER.index('Match Quality'))
            self.show_match_colors_button.setIcon(QIcon(os.path.join(self.parent.opts.resources_path,
                                                                     'icons',
                                                                     'matches_hide.png')))
        else:
            self.show_match_colors_button.setText("Show match status")
            self.show_match_colors_button.setIcon(QIcon(os.path.join(self.parent.opts.resources_path,
                                                                     'icons',
                                                                     'matches_show.png')))
        self.tv.setAlternatingRowColors(not self.show_match_colors)
        self.tm.refresh(self.show_match_colors)

    # Helpers
    def _add_books_to_library(self):
        '''
        Filter out books already in calibre
        Hook into gui.iactions['Add Books'].add_books_from_device()
        '''
        self._log_location()

        # Save the selection region for restoration
        self.saved_selection_region = self.tv.visualRegionForSelection(self.tv.selectionModel().selection())
        self.updated_match_quality = {}

        # Isolate the books to add, confirming no cid
        bta = self._selected_books()
        paths_to_add = []
        for b in bta:
            if bta[b]['cid'] is None:
                paths_to_add.append(bta[b]['path'])

        # Build map of added books from the model so we know which items to monitor for completion
        if paths_to_add:
            # Find the books in the Device model
            model = self.parent.gui.memory_view.model()
            added = {}
            for item in model.sorted_map:
                original_cids = self.parent.opts.gui.current_db.search_getting_ids('formats:EPUB', '')
                book = model.db[item]
                if book.path in paths_to_add:
                    # Tell calibre to add the paths
                    # gui2.actions.add #406
                    self.opts.gui.iactions['Add Books'].add_books_from_device(self.parent.gui.memory_view,
                                                                              paths=[book.path])
                    # Wait for add_books job
                    while not self.parent.gui.job_manager.unfinished_jobs():
                        Application.processEvents()
                    this_job = self.parent.gui.job_manager.unfinished_jobs()[0]

                    while not this_job.is_finished:
                        Application.processEvents()

                    # Wait for current_db to be updated with one new cid
                    while True:
                        Application.processEvents()
                        updated_cids = self.parent.opts.gui.current_db.search_getting_ids('formats:EPUB', '')
                        added_cids = list(set(updated_cids) - set(original_cids))
                        if added_cids:
                            break
                    added[item] = {'path': book.path, 'cid': added_cids[0], 'row': item}

            #self._log("added: %s" % added)

            # Update in-memory with newly minted cid, populate added
            for item in added:
                #self._log(model.db[item].all_field_keys())
                #cid = model.db[item].application_id
                cid = added[item]['cid']

                # Add book_id, cid to item dict, update installed_books with cid
                for book in self.installed_books.values():
                    if book.path == added[item]['path']:
                        #added[item]['cid'] = cid
                        added[item]['book_id'] = book.mid
                        book.cid = cid
                        break

                # Add model_row, update cid in spreadsheet
                for model_row in bta:
                    if bta[model_row]['book_id'] == added[item]['book_id']:
                        added[item]['model_row'] = model_row
                        self.tm.set_calibre_id(model_row, cid)

            total_books = len(added)
            self._busy_status_setup()

            db = self.opts.gui.current_db
            cached_books = self.parent.connected_device.cached_books

            # Update calibre metadata from Marvin metadata, bind uuid
            for i, item in enumerate(added):
                mismatches = {}
                this_book = cached_books[added[item]['path']]
                mismatches['authors'] = {'Marvin': this_book['authors']}
                mismatches['author_sort'] = {'Marvin': this_book['author_sort']}
                mismatches['cover_hash'] = {'Marvin': this_book['cover_hash']}
                mismatches['comments'] = {'Marvin': this_book['description']}
                mismatches['pubdate'] = {'Marvin': this_book['pubdate']}
                mismatches['publisher'] = {'Marvin': this_book['publisher']}
                mismatches['series'] = {'Marvin': this_book['series']}
                mismatches['series_index'] = {'Marvin': this_book['series_index']}
                mismatches['tags'] = {'Marvin': this_book['tags']}
                mismatches['title'] = {'Marvin': this_book['title']}
                mismatches['title_sort'] = {'Marvin': this_book['title_sort']}

                # Get the newly minted calibre uuid
                mi = db.get_metadata(added[item]['cid'], index_is_id=True)
                mismatches['uuid'] = {'calibre': mi.uuid,
                                      'Marvin': this_book['uuid']}

                # Update calibre metadata to match Marvin metadata
                # Book is added to updated_match_quality for flashing
                if total_books > 1:
                    msg = "Updating calibre metadata: {0} of {1}".format(i+1, total_books)
                else:
                    msg = "Updating calibre metadata"
                self._busy_status_msg(msg=msg)
                self._update_calibre_metadata(added[item]['book_id'],
                                              added[item]['cid'],
                                              mismatches,
                                              added[item]['model_row'],
                                              update_local_db=False)

                # Update .application_id in gui.booklists to match our new cid
                self._log("updating application_id in booklist")
                booklist = self.parent.gui.booklists()[0]
                for book in booklist:
                    if book.uuid == mi.uuid:
                        book.application_id = added[item]['cid']
                        break

            # Update local_db
            self._localize_marvin_database()
            self._busy_status_teardown()

            # Launch row flasher
            self._flash_affected_rows()

            # Refresh calibre view to reflect changed metadata
            updateCalibreGUIView()

            # Reset connected status in Library window
            self.parent.gui.book_on_device(None, reset=True)

    def _apply_date_read(self, update_gui=True):
        '''
        Fetch the LAST_OPENED date, convert to datetime, apply to custom field
        '''
        lookup = get_cc_mapping('date_read', 'field', None)
        if lookup:
            self._log_location()
            selected_books = self._selected_books()
            updated = False
            for row in selected_books:
                cid = selected_books[row]['cid']
                if cid is not None:
                    # Get the current value from the lookup field
                    db = self.opts.gui.current_db
                    mi = db.get_metadata(cid, index_is_id=True)
                    #old_date = mi.get_user_metadata(lookup, False)['#value#']
                    #self._log("Updating old date_read value: %s" % repr(old_date))

                    # Build a new datetime object from Last read
                    new_date = selected_books[row]['last_opened']
                    if new_date:
                        updated = True
                        um = mi.metadata_for_field(lookup)
                        ndo = strptime(new_date, "%Y-%m-%d %H:%M", as_utc=False, assume_utc=True)
                        try:
                            um['#value#'] = ndo
                            mi.set_user_metadata(lookup, um)
                            db.set_metadata(cid, mi, set_title=False, set_authors=False,
                                            commit=True)
                        except:
                            self._log("unable to set custom field value, calibre needs to be restarted")

                    else:
                        self._log("'%s' has no Last read date" % selected_books[row]['title'])
            if updated and update_gui:
                updateCalibreGUIView()

    def _apply_flags(self, update_gui=True):
        '''
        Synchronize Read and Reading list flags between calibre and connected iDevice
        Compare metadata timestamps to determine master
        '''

        def _get_marvin_last_modified(book_id):
            '''
            Get MetadataUpdated timestamp for book_id,
            '''
            UPDATE_FIELD = b'MetadataUpdated'
            arg2 = ''

            con = sqlite3.connect(self.parent.connected_device.local_db_path)
            with con:
                con.row_factory = sqlite3.Row

                lm_cur = con.cursor()
                lm_cur.execute('''SELECT
                                   *
                                  FROM Books
                                  WHERE ID = '{0}'
                               '''.format(book_id))
                row = lm_cur.fetchone()

                last_modified = datetime.now(tz.tzutc())
                if UPDATE_FIELD in row.keys():
                    try:
                        last_modified = datetime.utcfromtimestamp(row[UPDATE_FIELD]).replace(tzinfo=tz.tzutc())
                    except:
                        arg2 = "\n\t\t error retrieving {0}, returning now()".format(UPDATE_FIELD)
                        import traceback
                        exc_type, exc_value, exc_traceback = sys.exc_info()
                        self._log_location(traceback.format_exception_only(exc_type, exc_value)[0].strip())
                else:
                    arg2 = "\n\t\t {0} unavailable, returning now()".format(UPDATE_FIELD)

                last_modified = last_modified.astimezone(tz.tzlocal())
                lm_cur.close()

            self._log_location(last_modified, "{0}".format(arg2))
            return last_modified

        read_lookup = get_cc_mapping('read', 'field', None)
        reading_list_lookup = get_cc_mapping('reading_list', 'field', None)
        if read_lookup or reading_list_lookup:
            selected_books = self._selected_books()
            self._log_location(selected_books[selected_books.keys()[0]]['title'])
            for row in selected_books:
                cid = self._selected_books()[row]['cid']
                if cid is not None and (read_lookup or reading_list_lookup):
                    # Get the metadata object
                    db = self.opts.gui.current_db
                    mi = db.get_metadata(cid, index_is_id=True)

                    # Get the current Marvin flag values
                    flagbits = self.tm.get_flags(row).sort_key

                    # Get the Marvin LastModified date for this book
                    book_id = self._selected_books()[row]['book_id']
                    c_last_modified = mi.last_modified.astimezone(tz.tzlocal())
                    m_last_modified = _get_marvin_last_modified(book_id)

                    # ~~~~~~~~~ Process Read flag ~~~~~~~~~
                    c_read = False
                    if read_lookup:
                        c_read_um = mi.metadata_for_field(read_lookup)
                        c_read = bool(c_read_um['#value#'])
                    m_read = bool(flagbits & self.READ_FLAG)

                    #  If unequal, determine sync master and update
                    if read_lookup and (c_read != m_read):
                        self._log("Updating Read flag…")
                        self._log("calibre last_modified: %s" % c_last_modified)
                        self._log("Marvin last_modified: %s" % m_last_modified)
                        if c_last_modified > m_last_modified:
                            self._log("Using calibre as sync master. Read flag: %s" % c_read)
                            if c_read:
                                self._set_flags('set_read_flag', update_local_db=False)
                            else:
                                self._clear_flags('clear_read_flag', update_local_db=False)
                        else:
                            self._log("Using Marvin as sync master. Read flag: %s" % m_read)
                            if m_read:
                                self._set_flags('set_read_flag', update_local_db=False)
                            else:
                                self._clear_flags('clear_read_flag', update_local_db=False)
                    else:
                        if not read_lookup:
                            self._log("No custom column mapped for Read flag")
                        elif not c_read and not m_read:
                            self._log("No change: both flags already cleared")
                        elif c_read and m_read:
                            self._log("No change: both flags already set")

                    # ~~~~~~~~~ Process Reading list flag ~~~~~~~~~
                    c_reading_list = False
                    if reading_list_lookup:
                        c_reading_list_um = mi.metadata_for_field(reading_list_lookup)
                        c_reading_list = bool(c_reading_list_um['#value#'])
                    m_reading_list = bool(flagbits & self.READING_FLAG)

                    #  If unequal, determine sync master and update
                    if reading_list_lookup and (c_reading_list != m_reading_list):
                        self._log("Updating Reading list flag…")
                        self._log("calibre last_modified: %s" % c_last_modified)
                        self._log("Marvin last_modified: %s" % m_last_modified)
                        if c_last_modified > m_last_modified:
                            self._log("Using calibre as sync master. Reading list flag: %s" % c_reading_list)
                            if c_reading_list:
                                self._set_flags('set_read_flag', update_local_db=False)
                            else:
                                self._clear_flags('clear_read_flag', update_local_db=False)
                        else:
                            self._log("Using Marvin as sync master. Reading list flag: %s" % m_reading_list)
                            if m_reading_list:
                                self._set_flags('set_read_flag', update_local_db=False)
                            else:
                                self._clear_flags('clear_read_flag', update_local_db=False)
                    else:
                        if not reading_list_lookup:
                            self._log("No custom column mapped for Reading list flag")
                        elif not c_reading_list and not m_reading_list:
                            self._log("No change: both flags already cleared")
                        elif c_reading_list and m_reading_list:
                            self._log("No change: both flags already set")

            if update_gui:
                updateCalibreGUIView()

    def _apply_progress(self, update_gui=True):
        '''
        Fetch Progress, apply to custom field
        Need to assert force_changes for db to allow custom field to be set to None.
        '''
        lookup = get_cc_mapping('progress', 'field', None)
        if lookup:
            self._log_location()
            selected_books = self._selected_books()
            for row in selected_books:
                cid = selected_books[row]['cid']
                if cid is not None:
                    # Get the current value from the lookup field
                    db = self.opts.gui.current_db
                    mi = db.get_metadata(cid, index_is_id=True)
                    um = mi.metadata_for_field(lookup)

                    new_progress = self.tm.get_progress(row).sort_key
                    if new_progress is not None:
                        new_progress = new_progress * 100
                    um['#value#'] = new_progress

                    mi.set_user_metadata(lookup, um)
                    db.set_metadata(cid, mi, set_title=False, set_authors=False,
                                    commit=True, force_changes=True)
            if update_gui:
                updateCalibreGUIView()

    def _apply_word_count(self, update_gui=True):
        '''
        Fetch Progress, apply to custom field
        '''
        lookup = get_cc_mapping('word_count', 'field', None)
        if lookup:
            self._log_location()
            selected_books = self._selected_books()
            updated = False
            for row in selected_books:
                cid = selected_books[row]['cid']
                if cid is not None:
                    # Get the current value from the lookup field
                    db = self.opts.gui.current_db
                    mi = db.get_metadata(cid, index_is_id=True)
                    um = mi.metadata_for_field(lookup)

                    new_word_count = self.tm.get_word_count(row).sort_key
                    if new_word_count:
                        um['#value#'] = new_word_count

                        mi.set_user_metadata(lookup, um)
                        db.set_metadata(cid, mi, set_title=False, set_authors=False,
                                        commit=True, force_changes=True)
                        updated = True
            if updated and update_gui:
                updateCalibreGUIView()

    def _build_metadata_update(self, book_id, cid, book, mismatches):
        '''
        Build a metadata update command file for Marvin
        '''
        self._log_location()

        def _strftime(fmt, dt):
            '''
            Guarantee YYYY-MM-DD format for strftime results. Resolves problem when
            year < 1000
            '''
            result = strftime(fmt, t=dt)
            if not re.match("\d{4}-\d{2}-\d{2}", result):
                ans = re.match("(?P<year>\d+)-(?P<month>\d+)-(?P<day>\d+)", result)
                year = int(ans.group('year'))
                month = int(ans.group('month'))
                day = int(ans.group('day'))
                result = "{year:04d}-{month:02d}-{day:02d}".format(
                    year=year, month=month, day=day)
            return result

        cached_books = self.parent.connected_device.cached_books
        target_epub = self.installed_books[book_id].path

        # Init the update_metadata command file
        command_element = "updatemetadata"
        update_soup = BeautifulStoneSoup(self.METADATA_COMMAND_XML.format(
            command_element, time.mktime(time.localtime())))
        root = update_soup.find(command_element)
        root['cleanupcollections'] = 'yes'

        book_tag = Tag(update_soup, 'book')
        book_tag['author'] = escape(', '.join(book.authors))
        book_tag['authorsort'] = escape(book.author_sort)
        book_tag['filename'] = target_epub

        naive = book.pubdate.replace(hour=0, minute=0, second=0, tzinfo=None)
        book_tag['pubdate'] = _strftime('%Y-%m-%d', naive)
        book_tag['publisher'] = ''
        if book.publisher is not None:
            book_tag['publisher'] = escape(book.publisher)
        book_tag['series'] = ''
        if book.series:
            book_tag['series'] = escape(book.series)
        book_tag['seriesindex'] = ''
        if book.series_index:
            book_tag['seriesindex'] = book.series_index
        book_tag['title'] = escape(book.title)
        book_tag['titlesort'] = escape(book.title_sort)
        book_tag['uuid'] = book.uuid

        # Cover
        if 'cover_hash' in mismatches:
            desired_thumbnail_height = self.parent.connected_device.THUMBNAIL_HEIGHT
            try:
                cover = thumbnail(book.cover_data[1],
                                  desired_thumbnail_height,
                                  desired_thumbnail_height)
                cover_hash = hashlib.md5(cover[2]).hexdigest()

                cover_tag = Tag(update_soup, 'cover')
                cover_tag['hash'] = cover_hash
                cover_tag['encoding'] = 'base64'
                cover_tag.insert(0, base64.b64encode(cover[2]))
                book_tag.insert(0, cover_tag)
            except:
                self._log("error calculating cover_hash for %s (cid %d)" % (book.title, cid))
                import traceback
                self._log(traceback.format_exc())
        else:
            self._log(" '%s': cover is up to date" % book.title)

        # ~~~~~~ Subjects ~~~~~~
        subjects_tag = Tag(update_soup, 'subjects')
        for tag in sorted(book.tags, reverse=True):
            subject_tag = Tag(update_soup, 'subject')
            subject_tag.insert(0, escape(tag))
            subjects_tag.insert(0, subject_tag)
        book_tag.insert(0, subjects_tag)

        # ~~~~~~ Collections + Flags ~~~~~~
        ccas = self._get_calibre_collections(cid)
        if ccas is None:
            ccas = []
        flags = self.installed_books[book_id].flags
        collection_assignments = sorted(flags + ccas, key=sort_key)

        # Update the driver cache
        cached_books[target_epub]['device_collections'] = collection_assignments

        collections_tag = Tag(update_soup, 'collections')
        if collection_assignments:
            for tag in collection_assignments:
                c_tag = Tag(update_soup, 'collection')
                c_tag.insert(0, escape(tag))
                collections_tag.insert(0, c_tag)
        book_tag.insert(0, collections_tag)

        # Add the description
        try:
            description_tag = Tag(update_soup, 'description')
            description_tag.insert(0, escape(book.comments))
            book_tag.insert(0, description_tag)
        except:
            pass

        update_soup.manifest.insert(0, book_tag)

        return update_soup

    def _build_parameters(self, book, update_soup):
        parameters_tag = Tag(update_soup, 'parameters')

        parameter_tag = Tag(update_soup, 'parameter')
        parameter_tag['name'] = "filename"
        parameter_tag.insert(0, book.path)
        parameters_tag.insert(0, parameter_tag)

        parameter_tag = Tag(update_soup, 'parameter')
        parameter_tag['name'] = "uuid"
        parameter_tag.insert(0, book.uuid)
        parameters_tag.insert(0, parameter_tag)

        parameter_tag = Tag(update_soup, 'parameter')
        parameter_tag['name'] = "author"
        parameter_tag.insert(0, escape(', '.join(book.authors)))
        parameters_tag.insert(0, parameter_tag)

        parameter_tag = Tag(update_soup, 'parameter')
        parameter_tag['name'] = "title"
        parameter_tag.insert(0, book.title)
        parameters_tag.insert(0, parameter_tag)

        return parameters_tag

    def _busy_panel_setup(self, title, on_top=False, show_cancel=False):
        '''
        '''
        self._log_location(title)

        if self.busy_panel:
            self._log("busy_window is already active with '%s'" %
                      str(self.busy_panel.msg))
        else:
            QApplication.setOverrideCursor(QCursor(Qt.WaitCursor))
            self.busy_panel = MyBlockingBusy(self.parent.gui, title, size=60,
                                              on_top=on_top,
                                              show_cancel=show_cancel)
            self.busy_panel.start()
            self.busy_panel.show()
            Application.processEvents()

    def _busy_panel_teardown(self):
        '''
        '''
        self._log_location()
        if self.busy_panel:
            self.busy_panel.stop()
            self.busy_panel.accept()
            self.busy_panel = None
            QApplication.restoreOverrideCursor()
        else:
            self._log("no active busy_window")

    def _busy_status_msg(self, msg=''):
        self._log_location(msg)
        self.busy_status_label.setText(msg)
        Application.processEvents()

    def _busy_status_setup(self, msg='', show_cancel=False,
        marvin_cancellation_required=False):
        '''
        '''
        self._log_location(msg)
        self.busy = True
        self.tv.setEnabled(False)
        self.dialogButtonBox.setEnabled(False)
        self.filter_le.setEnabled(False)
        self.filter_tb.setEnabled(False)
        self.busy_status_pi.setVisible(True)
        self.busy_status_label.setText(msg)
        self.busy_status_pi.startAnimation()
        if show_cancel:
            self.busy_cancel_button.setVisible(True)
            self.busy_cancel_button.setEnabled(True)
            self.busy_cancel_requested = False
        self.marvin_cancellation_required = marvin_cancellation_required

    def _busy_status_teardown(self):
        '''
        '''
        self._log_location()
        self.busy_status_label.setText('')
        self.busy_cancel_button.setVisible(False)
        self.busy_cancel_button.setEnabled(True)
        self.busy_cancel_requested = False
        self.busy_status_pi.stopAnimation()
        self.busy_status_pi.setVisible(False)
        self.filter_le.setEnabled(True)
        self.filter_tb.setEnabled(True)
        self.dialogButtonBox.setEnabled(True)
        self.tv.setEnabled(True)
        self.busy = False

    def _calculate_word_count(self, silent=False):
        '''
        Calculate word count for each selected book
        selected_books: {row: {'book_id':, 'cid':, 'path':, 'title':}...}
        return stats {book_id: word_count}
        silent switch used when another method needs word count (Generate DV)
        Wait until completion to update local_db
        '''
        def _extract_body_text(data):
            '''
            Get the body text of this html content with any html tags stripped
            '''
            RE_HTML_BODY = re.compile(u'<body[^>]*>(.*)</body>', re.UNICODE | re.DOTALL | re.IGNORECASE)
            RE_STRIP_MARKUP = re.compile(u'<[^>]+>', re.UNICODE)

            body = RE_HTML_BODY.findall(data)
            if body:
                return RE_STRIP_MARKUP.sub('', body[0]).replace('.', '. ')
            return ''

        self._log_location()

        stats = {}
        db_update = False

        selected_books = self._selected_books()
        if selected_books:
            if not silent:
                msg = "Calculating word count"
                total_books = len(selected_books)
                self._busy_status_setup(show_cancel=len(selected_books) > 1)

            # Save the selection region for restoration
            self.saved_selection_region = self.tv.visualRegionForSelection(self.tv.selectionModel().selection())

            for i, row in enumerate(sorted(selected_books.keys())):
                if self.busy_cancel_requested:
                    break

                # Do we already know the word count?
                cwc = self.tm.get_word_count(row).sort_key
                if cwc:
                    stats[selected_books[row]['book_id']] = cwc
                    continue

                db_update = True

                # Highlight book we're working on
                self.tv.selectRow(row)

                if not silent:
                    if total_books > 1:
                        msg = "Calculating word count: {0} of {1}".format(i+1, total_books)
                    else:
                        msg = "Calculating word count"
                    self._busy_status_msg(msg=msg)

                # Copy the remote epub to local storage
                path = selected_books[row]['path']
                rbp = '/'.join(['/Documents', path])
                lbp = os.path.join(self.local_cache_folder, path)

                with open(lbp, 'wb') as out:
                    self.ios.copy_from_idevice(str(rbp), out)

                # Open the file
                iterator = EbookIterator(lbp)
                iterator.__enter__(only_input_plugin=True, run_char_count=True,
                                   read_anchor_map=False)
                book_files = []
                strip_html = True
                for path in iterator.spine:
                    with open(path, 'rb') as f:
                        html = f.read().decode('utf-8', 'replace')
                        if strip_html:
                            html = unicode(_extract_body_text(html)).strip()
                            #print('FOUND HTML:', html)
                    book_files.append(html)
                book_text = ''.join(book_files)

                wordcount = get_wordcount_obj(book_text)

                self._log("{0}: {1:,} words".format(
                    selected_books[row]['title'], wordcount.words))
                stats[selected_books[row]['book_id']] = wordcount.words

                # Delete the local copy
                os.remove(lbp)

                # Update the model
                wc = locale.format("%d", wordcount.words, grouping=True)
                if wc > "0":
                    word_count_item = SortableTableWidgetItem(
                        "{0} ".format(wc),
                        wordcount.words)
                else:
                    word_count_item = SortableTableWidgetItem('', 0)
                self.tm.set_word_count(row, word_count_item)

                # Update self.installed_books
                book_id = selected_books[row]['book_id']
                self.installed_books[book_id].word_count = wc

                # Tell Marvin about the updated word_count
                command_name = 'update_metadata_items'
                command_element = 'updatemetadataitems'
                update_soup = BeautifulStoneSoup(self.METADATA_COMMAND_XML.format(
                    command_element, time.mktime(time.localtime())))
                book_tag = Tag(update_soup, 'book')
                book_tag['author'] = escape(', '.join(self.installed_books[book_id].authors))
                book_tag['filename'] = self.installed_books[book_id].path
                book_tag['title'] = self.installed_books[book_id].title
                book_tag['uuid'] = self.installed_books[book_id].uuid

                book_tag['wordcount'] = wordcount.words
                update_soup.manifest.insert(0, book_tag)

                results = self._issue_command(command_name, update_soup, update_local_db=False)
                if results['code']:
                    if not silent:
                        #pb.hide()
                        self._busy_status_teardown()
                    self._show_command_error(command_name, results)
                    return stats

            # Update local_db for all changes
            if db_update:
                self._localize_marvin_database()

            if not silent:
                self._busy_status_teardown()

            # Restore selection
            if self.saved_selection_region:
                for rect in self.saved_selection_region.rects():
                    self.tv.setSelection(rect, QItemSelectionModel.Select)
                self.saved_selection_region = None

        else:
            self._log("No selected books")
            # Display a summary
            title = "Word count"
            msg = ("<p>Select one or more books to calculate word count.</p>")
            MessageBox(MessageBox.INFO, title, msg,
                       show_copy_button=False).exec_()

        #self._log(stats)
        return stats

    def _clear_flags(self, action, update_local_db=True):
        '''
        Clear specified flags for selected books
        sort_key is the bitfield representing current flag settings
        '''
        def _build_flag_list(flagbits):
            flags = []
            if flagbits & self.NEW_FLAG:
                flags.append(self.FLAGS['new'])
            if flagbits & self.READING_FLAG:
                flags.append(self.FLAGS['reading_list'])
            if flagbits & self.READ_FLAG:
                flags.append(self.FLAGS['read'])
            return flags

        def _update_in_memory(book_id, path):
            flags = self.installed_books[book_id].flags
            collections = self.installed_books[book_id].device_collections
            merged = sorted(flags + collections, key=sort_key)

            # Update driver (cached_books)
            cached_books = self.parent.connected_device.cached_books
            cached_books[path]['device_collections'] = merged

            # Update Device model
            for row in self.opts.gui.memory_view.model().map:
                book = self.opts.gui.memory_view.model().db[row]
                if book.path == path:
                    book.device_collections = merged
                    break

        self._log_location(action)
        if action == 'clear_new_flag':
            mask = self.NEW_FLAG
        elif action == 'clear_reading_list_flag':
            mask = self.READING_FLAG
        elif action == 'clear_read_flag':
            mask = self.READ_FLAG
        elif action == 'clear_all_flags':
            mask = 0

        local_update_required = False

        # Save the currently selected rows
        self.saved_selection_region = self.tv.visualRegionForSelection(self.tv.selectionModel().selection())

        selected_books = self._selected_books()
        for row in selected_books:
            self.tv.selectRow(row)
            book_id = selected_books[row]['book_id']
            flagbits = self.tm.get_flags(row).sort_key

#             self._log("%s: mask: %s  flagbits: %s" %
#                 (selected_books[row]['title'], mask, flagbits))
#
            path = selected_books[row]['path']
            if mask == 0 and flagbits:
                flagbits = 0
                basename = "flags0.png"
                new_flags_widget = SortableImageWidgetItem(os.path.join(self.parent.opts.resources_path,
                                                                        'icons', basename),
                                                           flagbits)
                # Update self.installed_books flags list
                self.installed_books[book_id].flags = []
                local_update_required = True

            elif flagbits & mask:
                # Clear the bit with XOR
                flagbits = flagbits ^ mask
                basename = "flags%d.png" % flagbits
                new_flags_widget = SortableImageWidgetItem(os.path.join(self.parent.opts.resources_path,
                                                                        'icons', basename),
                                                           flagbits)
                # Update self.installed_books flags list
                self.installed_books[book_id].flags = _build_flag_list(flagbits)
                local_update_required = True

            if local_update_required:
                # Update the model
                self.tm.set_flags(row, new_flags_widget)

                # Update reading progress based upon flag values
                self._update_reading_progress(self.installed_books[book_id], row)

                # Update in-memory caches
                _update_in_memory(book_id, path)

                # Update Marvin db
                self._inform_marvin_collections(book_id, update_local_db=False)
                self._update_device_flags(book_id, path, _build_flag_list(flagbits))
            else:
                self._log("Marvin flags already correct")

            self._inform_calibre_flags(book_id)

        # Restore selection
        if self.saved_selection_region:
            for rect in self.saved_selection_region.rects():
                self.tv.setSelection(rect, QItemSelectionModel.Select)
            self.saved_selection_region = None

        if update_local_db and local_update_required:
            self._localize_marvin_database()

        Application.processEvents()

    def _clear_selected_rows(self):
        '''
        Clear any active selections
        '''
        self._log_location()
        self.tv.clearSelection()
        self.repaint()

    def _compute_epub_hash(self, zipfile):
        '''
        Generate a hash of all text and css files in epub
        '''
        def _url_decode(s):
            subs = {
                    '%20': ' ',
                    '%21': '!',
                    '%22': '"',
                    '%23': '#',
                    '%25': '%'
                   }
            for k, v in subs.iteritems():
                s = s.replace(k, v)
            return s

        _local_debug = False
        if _local_debug:
            self._log_location(os.path.basename(zipfile))

        # Find the OPF file in the zipped ePub, extract a list of text files
        try:
            zf = ZipFile(zipfile, 'r')
            container = etree.fromstring(zf.read('META-INF/container.xml'))
            opf_tree = etree.fromstring(zf.read(container.xpath('.//*[local-name()="rootfile"]')[0].get('full-path')))

            text_hrefs = []
            manifest = opf_tree.xpath('.//*[local-name()="manifest"]')[0]
            for item in manifest.iterchildren():
                #self._log(etree.tostring(item, pretty_print=True))
                mt = item.get('media-type')
                if mt in ['application/xhtml+xml', 'text/css']:
                    thr = item.get('href').split('/')[-1]
                    text_hrefs.append(_url_decode(thr))
            zf.close()
        except:
            if _local_debug:
                import traceback
                self._log(traceback.format_exc())
            return None

        if _local_debug:
            self._log("{:-^80}".format(" text_hrefs[] "))
            for th in text_hrefs:
                self._log(th)
            self._log("{:-^80}".format(""))

        m = hashlib.md5()
        zfi = ZipFile(zipfile).infolist()
        for zi in zfi:
            base = zi.filename.split('/')[-1]
            if _local_debug:
                #self._log("evaluating %s" % zi.filename)
                self._log("evaluating %s" % repr(base))

            if base in text_hrefs:
                m.update(zi.filename)
                m.update(str(zi.file_size))
                for component in zi.date_time:
                    m.update(str(component))
                if _local_debug:
                    self._log(" adding filename %s" % (zi.filename))
                    self._log(" adding file_size %s" % (zi.file_size))
                    self._log(" adding date_time %s" % (repr(zi.date_time)))

        if _local_debug:
            self._log("computed hexdigest: %s" % m.hexdigest())

        return m.hexdigest()

    def _construct_table_data(self):
        '''
        Populate the table data from self.installed_books
        '''
        def _generate_articles(book_data):
            '''
            '''
            article_count = 0
            if 'Wiki' in book_data.articles:
                article_count += len(book_data.articles['Wiki'])
            if 'Pinned' in book_data.articles:
                article_count += len(book_data.articles['Pinned'])

            if article_count:
                articles = SortableTableWidgetItem(
                    "{0}".format(article_count),
                    article_count)
            else:
                articles = SortableTableWidgetItem('', 0)
            return articles

        def _generate_author(book_data):
            '''
            '''
            if not book_data.author_sort:
                book_data.author_sort = ', '.join(book_data.author)
            author = SortableTableWidgetItem(
                ', '.join(book_data.author),
                book_data.author_sort.upper())
            return author

        def _generate_date_added(book_data):
            '''
            Date added sorts by timestamp
            '''
            date_added_ts = ''
            date_added_sort = 0
            if book_data.date_added:
                date_added_ts = time.strftime("%Y-%m-%d",
                                               time.localtime(book_data.date_added))
                date_added_sort = book_data.date_added
            date_added = SortableTableWidgetItem(
                date_added_ts,
                date_added_sort)
            return date_added

        def _generate_flags_profile(book_data):
            '''
            Figure out which flags image to use, assign sort value
            NEW = 4
            READING LIST = 2
            READ = 1
            '''
            flag_list = book_data.flags
            flagbits = 0
            if 'NEW' in flag_list:
                flagbits += 4
            if 'READING LIST' in flag_list:
                flagbits += 2
            if 'READ' in flag_list:
                flagbits += 1
            base_name = "flags%d.png" % flagbits
            flags = SortableImageWidgetItem(os.path.join(self.parent.opts.resources_path,
                                                         'icons', base_name),
                                            flagbits)
            return flags

        def _generate_highlights(book_data):
            '''
            '''
            if len(book_data.highlights):
                highlights = SortableTableWidgetItem(
                    "{0}".format(len(book_data.highlights)),
                    len(book_data.highlights))
            else:
                highlights = SortableTableWidgetItem('', 0)
            return highlights

        def _generate_last_opened(book_data):
            '''
            last_opened sorts by timestamp
            '''
            last_opened_ts = ''
            last_opened_sort = 0
            if book_data.date_opened:
                last_opened_ts = time.strftime("%Y-%m-%d %H:%M",
                                               time.localtime(book_data.date_opened))
                last_opened_sort = book_data.date_opened
            last_opened = SortableTableWidgetItem(
                last_opened_ts,
                last_opened_sort)
            return last_opened

        def _generate_locked_status(book_data):
            '''
            Generate a SortableImageWidgetItem representing the Pin value
            '''
            #image_name = "unlock_enabled.png"
            image_name = "empty_16x16.png"
            if book_data.pin:
                image_name = "lock_enabled.png"
            locked = SortableImageWidgetItem(os.path.join(self.parent.opts.resources_path,
                                                          'icons', image_name),
                                             book_data.pin)
            return locked

        def _generate_match_quality(book_data):
            '''
            GREEN:          Marvin uuid matches calibre uuid (hard match)
            YELLOW:         Marvin hash matches calibre hash (soft match)
            MAGENTA:        Book has multiple UUIDs in calibre, one matched in Marvin
            ORANGE:         Calibre hash duplicates:
            RED:            Marvin hash duplicates
            WHITE:          Marvin only, single copy
            LIGHT_GRAY:     Book exists in Marvin and calibre, but no match identified
            DARK_GRAY:      Book updated in Marvin (different or non-existent UUID)
            '''

            if self.opts.prefs.get('development_mode', False):
                self._log_location("'{0}'".format(book_data.title))
                self._log("uuid ({1}): {0}".format(repr(book_data.uuid),
                    'local' if book_data.uuid in self.library_uuid_map else 'foreign'))
                self._log("matches: {0}".format(repr(book_data.matches)))
                self._log("on_device: {0}".format(repr(book_data.on_device)))
                self._log("hash: {0}".format(repr(book_data.hash)))
                self._log("metadata_mismatches: {0}".format(
                    '' if book_data.metadata_mismatches else '{}'))
                for k, v in book_data.metadata_mismatches.items():
                    self._log(" {0}: {1}".format(k, v))

            _main = _('Main')

            if book_data.on_device is not None:
                '''
                Book is in calibre library.
                Resolve to GREEN | YELLOW | ORANGE | MAGENTA | LIGHT_GRAY
                '''
                #match_quality = self.MATCH_COLORS.index('LIGHT_GRAY')

                if book_data.on_device.startswith("{0} (".format(_main)):
                    ''' ORANGE: Calibre detects multiple copies '''
                    match_quality = self.MATCH_COLORS.index('ORANGE')
                elif book_data.uuid:
                    if (book_data.uuid in book_data.matches and
                        len(book_data.matches) > 1):
                        ''' MAGENTA: Multiple calibre UUIDs resolving to hash '''
                        match_quality = self.MATCH_COLORS.index('MAGENTA')
                    elif ([book_data.uuid] == book_data.matches and
                        not book_data.metadata_mismatches):
                        ''' GREEN: Hard UUID match, no metadata mismatches '''
                        match_quality = self.MATCH_COLORS.index('GREEN')
                    elif ([book_data.uuid] == book_data.matches and
                        book_data.metadata_mismatches):
                        ''' YELLOW: Hard UUID match with metadata mismatches '''
                        match_quality = self.MATCH_COLORS.index('YELLOW')
                    elif (book_data.uuid not in self.library_uuid_map and
                        book_data.metadata_mismatches):
                        ''' YELLOW: Foreign UUID with metadata mismatches '''
                        match_quality = self.MATCH_COLORS.index('YELLOW')
                    elif book_data.uuid not in self.library_uuid_map:
                        ''' DARK_GRAY: Book has been updated in Marvin '''
                        match_quality = self.MATCH_COLORS.index('DARK_GRAY')
                    else:
                        ''' LIGHT_GRAY: Book has been updated in calibre '''
                        match_quality = self.MATCH_COLORS.index('LIGHT_GRAY')
                else:
                    # No UUID, but calibre recognizes as a match
                    ''' DARK_GRAY: Book has been updated in Marvin '''
                    match_quality = self.MATCH_COLORS.index('DARK_GRAY')
            else:
                '''
                Book is not in calibre library
                Resolve to WHITE | RED
                '''
                match_quality = self.MATCH_COLORS.index('WHITE')

                if (book_data.hash in self.marvin_hash_map and
                    len(self.marvin_hash_map[book_data.hash]) > 1):
                    match_quality = self.MATCH_COLORS.index('RED')

            if self.opts.prefs.get('development_mode', False):
                self._log("match_quality: {0}".format(self.MATCH_COLORS[match_quality]))

            return match_quality

        def _generate_series(book_data):
            '''
            Generate a sort key based on series index
            Force non-series to sort after series
            '''
            series_ts = ''
            series_sort = '~'
            if book_data.series:
                cs_index = book_data.series_index
                if book_data.series_index.endswith('.0'):
                    cs_index = book_data.series_index[:-2]
                series_ts = "%s [%s]" % (book_data.series, cs_index)
                try:
                    index = float(book_data.series_index)
                except:
                    index = 0.0
                integer = int(index)
                fraction = index - integer
                series_sort = '%s %04d%s' % (book_data.series,
                                             integer,
                                             str('%0.4f' % fraction).lstrip('0'))
            series = SortableTableWidgetItem(series_ts, series_sort)
            return series

        def _generate_subjects(book_data):
            '''
            '''
            subjects = SortableTableWidgetItem(
                ', '.join(book_data.tags),
                ', '.join(book_data.tags).lower())
            return subjects

        def _generate_title(book_data):
            '''
            '''
            # Title, Author sort by title_sort, author_sort
            if not book_data.title_sort:
                book_data.title_sort = book_data.title_sorter()
            title = SortableTableWidgetItem(
                book_data.title,
                book_data.title_sort.upper())
            return title

        def _generate_vocabulary(book_data):
            if len(book_data.vocabulary):
                vocabulary = SortableTableWidgetItem(
                    "{0}".format(len(book_data.vocabulary)),
                    len(book_data.vocabulary))
            else:
                vocabulary = SortableTableWidgetItem('', 0)
            return vocabulary

        def _generate_word_count(book_data):
            '''
            '''
            if book_data.word_count > "0":
                word_count = SortableTableWidgetItem(
                    "{0} ".format(book_data.word_count),
                    locale.atoi(book_data.word_count))
            else:
                word_count = SortableTableWidgetItem('', 0)
            return word_count

        self._log_location()

        tabledata = []

        for book in self.installed_books:
            book_data = self.installed_books[book]
            articles = _generate_articles(book_data)
            author = _generate_author(book_data)
            collection_match = self._generate_collection_match(book_data)
            date_added = _generate_date_added(book_data)
            flags = _generate_flags_profile(book_data)
            highlights = _generate_highlights(book_data)
            last_opened = _generate_last_opened(book_data)
            locked = _generate_locked_status(book_data)
            book_data.match_quality = _generate_match_quality(book_data)
            progress = self._generate_reading_progress(book_data)
            title = _generate_title(book_data)
            series = _generate_series(book_data)
            subjects = _generate_subjects(book_data)
            vocabulary = _generate_vocabulary(book_data)
            word_count = _generate_word_count(book_data)

            # List order matches self.LIBRARY_HEADER
            this_book = [
                title,
                author,
                series,
                word_count,
                date_added,
                progress,
                last_opened,
                subjects,
                collection_match,
                locked,
                flags,
                highlights,
                vocabulary,
                self.CHECKMARK if book_data.deep_view_prepared else '',
                articles,
                book_data.match_quality,
                book_data.uuid,
                book_data.cid,
                book_data.mid,
                book_data.path
                ]
            tabledata.append(this_book)
            Application.processEvents()

        return tabledata

    def _construct_table_view(self):
        '''
        '''
        self._log_location()
        self.tm = MarkupTableModel(self, centered_columns=self.CENTERED_COLUMNS,
                                   right_aligned_columns=self.RIGHT_ALIGNED_COLUMNS)

        self.tv.setModel(self.tm)
        self.tv.setShowGrid(False)
        if self.parent.prefs.get('use_monospace_font', False):
            if isosx:
                FONT = QFont('Monaco', 11)
            elif iswindows:
                FONT = QFont('Lucida Console', 9)
            elif islinux:
                FONT = QFont('Monospace', 9)
                FONT.setStyleHint(QFont.TypeWriter)
            self.tv.setFont(FONT)
        else:
            FONT = self.tv.font()

        # Set row height
        fm = QFontMetrics(FONT)
        nrows = len(self.tabledata)
        for row in xrange(nrows):
            self.tv.setRowHeight(row, fm.height() + 4)

        self.tvSelectionModel = self.tv.selectionModel()
        self.tv.setAlternatingRowColors(not self.show_match_colors)
        self.tv.setShowGrid(False)
        self.tv.setWordWrap(False)
        self.tv.setSelectionBehavior(self.tv.SelectRows)

        # Hide the vertical self.header
        self.tv.verticalHeader().setVisible(False)

        columns_to_hide = list(self.HIDDEN_COLUMNS)

        # Check whether we're showing LOCKED_COL
        if not self.parent.has_password:
            columns_to_hide.append(self.LOCKED_COL)

        # If initial run, hide DATE_ADDED_COL, LAST_OPENED_COL, and SUBJECTS_COL
        saved_column_widths = self.opts.prefs.get('marvin_library_column_widths', None)
        if not saved_column_widths or (len(saved_column_widths) != len(self.LIBRARY_HEADER)):
            columns_to_hide.append(self.DATE_ADDED_COL)
            columns_to_hide.append(self.SUBJECTS_COL)
            columns_to_hide.append(self.LAST_OPENED_COL)
        else:
            for col, name in self.USER_CONTROLLED_COLUMNS:
                #self._log("%s: %d" % (name, saved_column_widths[col]))
                if saved_column_widths[col] == 0:
                    columns_to_hide.append(col)

        # Set column width to fit contents
        self.tv.resizeColumnsToContents()

        # Hide hidden columns
        for index in sorted(columns_to_hide):
            #self._log("hiding %s" % index)
            self.tv.hideColumn(index)

        # Set horizontal self.header props
        #self.tv.horizontalHeader().setStretchLastSection(True)

        # Clip Author, Title to 250
        self.tv.setColumnWidth(self.TITLE_COL, 250)
        self.tv.setColumnWidth(self.AUTHOR_COL, 250)

        # Restore saved widths if available
        saved_column_widths = self.opts.prefs.get('marvin_library_column_widths', False)
        if saved_column_widths and (len(saved_column_widths) == len(self.LIBRARY_HEADER)):
            for i, width in enumerate(saved_column_widths):
                self.tv.setColumnWidth(i, width)
        else:
            # Set narrow cols to width of FLAGS
            fixed_width = self.tv.columnWidth(self.LAST_OPENED_COL)
            if not fixed_width:
                fixed_width = 87
            for col in [self.WORD_COUNT_COL, self.COLLECTIONS_COL]:
                self.tv.setColumnWidth(col, fixed_width)

            fixed_width = self.tv.columnWidth(self.FLAGS_COL)
            for col in [self.ANNOTATIONS_COL, self.VOCABULARY_COL,
                        self.DEEP_VIEW_COL, self.ARTICLES_COL]:
                self.tv.setColumnWidth(col, fixed_width)
            self._save_column_widths()

        # Show/hide the Locked column depending on restrictions
        if self.parent.has_password:
            self.tv.showColumn(self.LOCKED_COL)
            self.tv.setColumnWidth(self.LOCKED_COL, 28)
            #self.tv.horizontalHeader().setResizeMode(self.LOCKED_COL, QHeaderView.Fixed) # observed crash

        self.tv.setSortingEnabled(True)

        sort_column = self.opts.prefs.get('marvin_library_sort_column',
                                          self.LIBRARY_HEADER.index('Match Quality'))
        sort_order = self.opts.prefs.get('marvin_library_sort_order',
                                         Qt.DescendingOrder)
        self.tv.sortByColumn(sort_column, sort_order)

    def _delete_books(self):
        '''
        '''
        self._log_location()

        btd = self._selected_books()
        books_to_delete = sorted([btd[b]['title'] for b in btd], key=sort_key)

        self.saved_selection_region = None
        self.updated_match_quality = {}

        if books_to_delete:
            ''' Under the skirts approach '''
            title = "Delete %s" % ("%d books?" % len(books_to_delete)
                                   if len(books_to_delete) > 1 else "1 book?")
            msg = ("<p>Click <b>Show details</b> for a list of books that will be deleted " +
                   "from your Marvin library.</p>" +
                   '<p><b><font style="color:#FF0000; ">{0}</font></b></p>'.format(title))
            det_msg = '\n'.join(books_to_delete)
            d = MessageBox(MessageBox.QUESTION, title, msg, det_msg=det_msg,
                           show_copy_button=False)
            if d.exec_():
                model = self.parent.gui.memory_view.model()
                paths_to_delete = [btd[b]['path'] for b in btd]
                self._log("paths_to_delete: %s" % paths_to_delete)
                sorted_map = model.sorted_map
                delete_map = {}
                for item in sorted_map:
                    book = model.db[item]
                    if book.path in paths_to_delete:
                        delete_map[book.path] = item
                        continue

                # Delete the rows in MM spreadsheet
                rows_to_delete = self._selected_rows()
                for row in sorted(rows_to_delete, reverse=True):
                    self.tm.beginRemoveRows(QModelIndex(), row, row)
                    del self.tm.arraydata[row]
                    self.tm.endRemoveRows()

                # Delete the books on Device
                if self.prefs.get('execute_marvin_commands', True):
                    job = self.parent.gui.remove_paths(delete_map.keys())

                    # Delete books in the Device model
                    model.mark_for_deletion(job, delete_map.values(), rows_are_ids=True)
                    model.deletion_done(job, succeeded=True)
                    for rtd in delete_map.values():
                        del model.db[rtd]

                    # Put on a show while waiting for the delete job to finish
                    QApplication.setOverrideCursor(QCursor(Qt.WaitCursor))
                    blocking_busy = MyBlockingBusy(self.opts.gui, self.UPDATING_MARVIN_MESSAGE, size=60)
                    blocking_busy.start()
                    blocking_busy.show()
                    while not job.is_finished:
                        Application.processEvents()
                    blocking_busy.stop()
                    blocking_busy.accept()
                    QApplication.restoreOverrideCursor()

                    # Remove from cached_paths in driver
                    if True:
                        for ptd in paths_to_delete:
                            self.parent.connected_device.cached_books.pop(ptd)
                    else:
                        # The book is not in booklists, how/when removed?
                        self.parent.connected_device.remove_books_from_metadata(paths_to_delete,
                                                                                self.parent.gui.booklists())

                    # Update the visible Device model
                    model.paths_deleted(paths_to_delete)
                else:
                    self._log("~~~ execute_marvin_commands disabled in JSON ~~~")

                # Remove from self.installed_books
                book_ids_to_delete = [btd[b]['book_id'] for b in btd]
                for book_id in book_ids_to_delete:
                    deleted = self.installed_books.pop(book_id)
                    if False:
                        self._log("deleted: %s hash: %s matches: %s" % (deleted.title,
                                                                        deleted.hash,
                                                                        deleted.matches))
                        if deleted.hash in self.library_scanner.hash_map:
                            self._log("library_hash_map: %s" % self.library_scanner.hash_map[deleted.hash])

                    for book_id, book in self.installed_books.items():
                        if book.hash == deleted.hash:
                            row = self._find_book_id_in_model(book_id)
                            if row:
                                # Is this book in library or Marvin only?
                                if self.tm.get_calibre_id(row):
                                    new = self.MATCH_COLORS.index('GREEN')
                                else:
                                    new = self.MATCH_COLORS.index('WHITE')

                                old = self.tm.get_match_quality(row)
                                self.tm.set_match_quality(row, new)
                                self.updated_match_quality[row] = {'book_id': book_id,
                                                                   'old': old,
                                                                   'new': new}
                # Update the book count in the title bar
                self.setWindowTitle(u'Marvin Library: %d books' % len(self.installed_books))

                # Launch row flasher
                #self._flash_affected_rows()

            else:
                self._log("delete cancelled")

        else:
            self._log("no books selected")
            title = "No selected books"
            msg = "<p>Select one or more books to delete.</p>"
            MessageBox(MessageBox.INFO, title, msg,
                       show_copy_button=False).exec_()

    def _dump_hash_map(self, library_hash_map):
        '''
        '''
        self._log_location()
        self._log("{0:^32} {1:^32} {2:^42}".format("HASH", "TITLE", "UUID"))
        for hash in sorted(library_hash_map):
            uuid = library_hash_map[hash][0]
            title = self.library_scanner.uuid_map[uuid]['title']
            if len(title) > 30:
                title = title[0:30] + '…'
            self._log("{0:<32} {1:<32} {2}".format(hash, title[0:31], library_hash_map[hash]))

    def _fetch_annotations(self, update_gui=True, report_results=False):
        '''
        Retrieve formatted annotations
        '''
        lookup = get_cc_mapping('annotations', 'field', None)
        if lookup:
            self._log_location()
            updated = 0
            for row, book in self._selected_books().items():
                cid = book['cid']
                if cid is not None:
                    if book['has_annotations']:
                        self._log("%s (row %d): %d annotations" %
                                  (repr(book['title']), row, self.tm.get_annotations(row).sort_key))
                        book_id = book['book_id']
                        new_annotations = self._get_formatted_annotations(book_id)

                        # Apply to custom column
                        # Get the current value from the lookup field
                        db = self.opts.gui.current_db
                        mi = db.get_metadata(cid, index_is_id=True)
                        um = mi.metadata_for_field(lookup)
                        old_annotations = mi.get_user_metadata(lookup, False)['#value#']
                        if old_annotations is None:
                            self._log("adding new_annotations")
                            um['#value#'] = new_annotations
                        else:
                            self._log("merging old_annotations and new_annotations")
                            old_soup = BeautifulSoup(old_annotations)
                            new_soup = BeautifulSoup(new_annotations)
                            merged_soup = merge_annotations(self, cid, old_soup, new_soup)
                            um['#value#'] = unicode(merged_soup)
                        mi.set_user_metadata(lookup, um)
                        db.set_metadata(cid, mi, set_title=False, set_authors=False,
                                        commit=True)
                        updated += 1
                    else:
                        self._log("%s has no annotations" % repr(book['title']))
                else:
                    self._log("%s does not exist in calibre library" % repr(book['title']))

            if update_gui and updated:
                updateCalibreGUIView()

            if report_results:
                title = 'Annotations refreshed'
                msg = ("<p>Annotations refreshed for %s.</p>" %
                    ("1 book" if updated == 1 else "%d books" % updated))
                MessageBox(MessageBox.INFO, title, msg, det_msg='', show_copy_button=False).exec_()

    def _fetch_deep_view_status(self, book_ids):
        '''
        Get current status for book_ids
        '''
        self._log_location(book_ids)

        dvp_status = {}
        con = sqlite3.connect(self.parent.connected_device.local_db_path)
        with con:
            con.row_factory = sqlite3.Row
            # Get all the books
            cur = con.cursor()
            cur.execute('''SELECT
                            Books.ID as id_,
                            DeepViewPrepared
                           FROM Books
                        ''')

            rows = cur.fetchall()
            for i, row in enumerate(rows):
                book_id = row[b'id_']
                if book_id in book_ids:
                    dvp_status[book_id] = row[b'DeepViewPrepared']

        return dvp_status

    def _fetch_marvin_content_hash(self, path):
        '''
        Given a Marvin path, compute/fetch a hash of its contents (excluding OPF)
        self.hash_cache is current
        '''
        #self._log_location(path)

        # Try getting the hash from the cache
        if path in self.hash_cache:
            #self._log("returning hash from cache: %s" % self.hash_cache[path])
            return self.hash_cache[path]

        # Get a local copy of the book, generate hash
        rbp = '/'.join(['/Documents', path])
        lbp = os.path.join(self.local_cache_folder, path)

        try:
            with open(lbp, 'wb') as out:
                self.ios.copy_from_idevice(str(rbp), out)
        except:
            # We have an invalid filename, but we need to return a unique hash
            #self._log("ERROR: Unable to open %s for output" % repr(lbp))
            import traceback
            self._log(traceback.format_exc())
            m = hashlib.md5()
            m.update(lbp)
            return m.hexdigest()

        hash = self._compute_epub_hash(lbp)

        # Add it to the hash_cache
        self._log("adding hash to cache: %s" % hash)
        self.hash_cache[path] = hash

        # Delete the local copy
        os.remove(lbp)
        return hash

    def _fetch_marvin_cover(self, book_id):
        '''
        Retrieve large cover from cache
        '''
        cover_bytes = None
        self._log_location("fetching large cover from cache")
        con = sqlite3.connect(self.parent.connected_device.local_db_path)
        with con:
            con.row_factory = sqlite3.Row

            # Fetch Hash from mainDb
            cover_cur = con.cursor()
            cover_cur.execute('''SELECT
                                  Hash
                                 FROM Books
                                 WHERE ID = '{0}'
                              '''.format(book_id))
            row = cover_cur.fetchone()

        book_hash = row[b'Hash']
        large_covers_subpath = self.parent.connected_device._cover_subpath(size="large")
        cover_path = '/'.join([large_covers_subpath, '%s.jpg' % book_hash])
        stats = self.ios.exists(cover_path)
        if stats:
            self._log("cover size: {:,} bytes".format(int(stats['st_size'])))
            cover_bytes = self.ios.read(cover_path, mode='rb')
        return cover_bytes

    def _find_book_id_in_model(self, book_id):
        '''
        Given a book_id, find its row in the displayed model
        '''
        for row, item in enumerate(self.tm.arraydata):
            #if self.tm.get_book_id(row) == book_id:
            if item[self.BOOK_ID_COL] == book_id:
                self._log("found %s at row %d" % (book_id, row))
                break
        else:
            row = None
        return row

    def _find_cid_in_model(self, cid):
        '''
        Given a cid, return its book_id
        '''
        book_id = None
        for book in self.tm.arraydata:
            if book[self.CALIBRE_ID_COL] == cid:
                book_id = book[self.BOOK_ID_COL]
                break
        return book_id

    def _find_fuzzy_matches(self, library_scanner, installed_books):
        '''
        Compare computed hashes of installed books to library books.
        Look for potential dupes.
        Add .matches property to installed_books, a list of all Marvin uuids matching
        our hash
        '''
        self._log_location()

        library_hash_map = library_scanner.hash_map
        hard_matches = {}
        soft_matches = []
        for book in installed_books:
            mb = installed_books[book]
            #self._log("evaluating %s hash: %s uuid: %s" % (mb.title, mb.hash, mb.uuid))
            uuids = []
            if mb.hash in library_hash_map:
                if mb.uuid in library_hash_map[mb.hash]:
                    #self._log("%s matches hash + uuid" % mb.title)
                    hard_matches[mb.hash] = mb
                    uuids = library_hash_map[mb.hash]
                else:
                    #self._log("%s matches hash, but not uuid" % mb.title)
                    soft_matches.append(mb)
                    uuids = [mb.uuid]
            else:
                #self._log("%s not in library_hash_map" % mb.title)
                pass
            #self._log("storing %s" % repr(uuids))
            mb.matches = uuids

        # Review the soft matches against the hard matches
        if soft_matches:
            # Scan soft matches against hard matches for hash collision
            for mb in soft_matches:
                if mb.hash in hard_matches:
                    mb.matches += hard_matches[mb.hash].matches

    def _flash_affected_rows(self):
        '''
        '''
        if self.updated_match_quality:
            self._log_location(sorted(self.updated_match_quality.keys()))
            self.flasher = RowFlasher(self, self.tm, self.updated_match_quality)
            self.connect(self.flasher, self.flasher.signal, self._flasher_complete)
            self.flasher.start()

    def _flasher_complete(self):
        '''
        '''
        self._log_location()
        if self.saved_selection_region:
            for rect in self.saved_selection_region.rects():
                self.tv.setSelection(rect, QItemSelectionModel.Select)

    def _generate_booklist(self):
        '''
        '''
        self._log_location()

        if False:
            ''' Scan library books for hashes '''
            if self.library_scanner.isRunning():
                self._busy_panel_setup("Scanning calibre library…")
                self.library_scanner.wait()
                self._busy_panel_teardown()

        # Save a reference to the title, uuid map
        self.library_title_map = self.library_scanner.title_map
        self.library_uuid_map = self.library_scanner.uuid_map

        # Get the library hash_map
        library_hash_map = self.library_scanner.hash_map
        if library_hash_map is None:
            library_hash_map = self._scan_library_books(self.library_scanner)
        else:
            self._log("hash_map already generated")

        # Dump the hash_map
        if self.opts.prefs.get('development_mode', False):
            self._dump_hash_map(library_hash_map)

        # Scan Marvin
        installed_books = self._get_installed_books()

        # Generate a map of Marvin hashes to book_ids
        self.marvin_hash_map = self._generate_marvin_hash_map(installed_books)

        # Update installed_books with library matches
        self._find_fuzzy_matches(self.library_scanner, installed_books)

        return installed_books

    def _generate_collection_match(self, book_data):
        '''
        If no custom collections field assigned, always return sort_value 0
        '''
        if (book_data.calibre_collections is None and
                book_data.device_collections == []):
            base_name = 'collections_empty.png'
            sort_value = 0
        elif (book_data.device_collections == [] and
              book_data.calibre_collections == []):
            base_name = 'collections_empty.png'
            sort_value = 0
        elif (book_data.calibre_collections is None and
                book_data.device_collections > []):
            base_name = 'collections_info.png'
            sort_value = 1
        elif book_data.device_collections == book_data.calibre_collections:
            base_name = 'collections_equal.png'
            sort_value = 3
        else:
            base_name = 'collections_unequal.png'
            sort_value = 2
        collection_match = SortableImageWidgetItem(os.path.join(self.parent.opts.resources_path,
                                                                'icons', base_name),
                                                   sort_value)
        return collection_match

    def _generate_deep_view(self):
        '''
        '''
        #WORST_CASE_CONVERSION_RATE = 2800   # WPM iPad1
        WORST_CASE_CONVERSION_RATE = 2350   # GwR empirical including updates
        #BEST_CASE_CONVERSION_RATE = 6500    # WPM iPad4, iPhone5
        TIMEOUT_PADDING_FACTOR = 0.50

        self._log_location()
        selected_books = self._selected_books()
        if selected_books:

            # Estimate worst-case time required to generate DV, covering word count calculations
            self._busy_status_setup(msg="Estimating time…")
            word_counts = self._calculate_word_count(silent=True)
            self._busy_status_teardown()

            twc = sum(word_counts.itervalues())
            total_seconds = twc/WORST_CASE_CONVERSION_RATE + 1
            self._log("word_counts: %s" % word_counts)

            individual_times = [int(v/WORST_CASE_CONVERSION_RATE + 1) for v in word_counts.values()]
            total_seconds = sum(individual_times)
            self._log("individual_times: %s" % individual_times)

            longest = max(individual_times)
            self._log("longest: %d" % longest)

            timeout = int(longest + (longest * TIMEOUT_PADDING_FACTOR))
            self._log("timeout: %d" % timeout)

            m, s = divmod(total_seconds, 60)
            h, m = divmod(m, 60)
            if h:
                estimated_time = "%d:%02d:%02d" % (h, m, s)
            else:
                estimated_time = "%d:%02d" % (m, s)

            if timeout > self.WATCHDOG_TIMEOUT:
                # Confirm that user wants to proceed given estimated time to completion
                total_books = len(selected_books)
                book_descriptor = "books" if total_books > 1 else "book"
                title = "Estimated time to completion"
                msg = ("<p>Generating Deep View for " +
                       "selected {0} ".format(book_descriptor) +
                       "may take as long as {0}, depending on your iDevice.</p>".format(estimated_time) +
                       "<p>Proceed?</p>")
                dlg = MessageBox(MessageBox.QUESTION, title, msg,
                                 show_copy_button=False)
                if not dlg.exec_():
                    self._log("user declined to proceed with estimated_time of %s" % estimated_time)
                    return
            else:
                # Use method default timeout
                timeout = None

            command_name = "command"
            command_type = "GenerateDeepView"
            update_soup = BeautifulStoneSoup(self.GENERAL_COMMAND_XML.format(
                command_type, time.mktime(time.localtime())))

            # Build a manifest of selected books
            manifest_tag = Tag(update_soup, 'manifest')
            for row in sorted(selected_books.keys(), reverse=True):
                book_id = selected_books[row]['book_id']
                book_tag = Tag(update_soup, 'book')
                book_tag['author'] = escape(', '.join(self.installed_books[book_id].authors))
                book_tag['filename'] = self.installed_books[book_id].path
                book_tag['title'] = self.installed_books[book_id].title
                book_tag['uuid'] = self.installed_books[book_id].uuid
                manifest_tag.insert(0, book_tag)
            update_soup.command.insert(0, manifest_tag)

            busy_msg = ("Generating Deep View for %s" %
                ("1 book…" if len(selected_books) == 1 else
                 "%d books…" % len(selected_books)))

            self._busy_status_setup(msg=busy_msg, show_cancel=len(selected_books) > 1,
                marvin_cancellation_required=True)
            results = self._issue_command(command_name, update_soup,
                                          timeout_override=timeout,
                                          update_local_db=True)
            self._busy_status_teardown()

            if results['code']:
                return self._show_command_error(command_type, results)

            # Get the latest DeepViewPrepared status for selected books
            book_ids = [selected_books[row]['book_id'] for row in selected_books.keys()]
            dpv_status = self._fetch_deep_view_status(book_ids)

            # Update visible model, self.installed_books
            for row in sorted(selected_books.keys(), reverse=True):
                book_id = selected_books[row]['book_id']
                self.installed_books[book_id].deep_view_prepared = dpv_status[book_id]
                updated = self.CHECKMARK if dpv_status[book_id] else ''
                self.tm.set_deep_view(row, updated)

    def _generate_interior_location_sort(self, xpath):
        try:
            match = re.match(r'\/x:html\[1\]\/x:body\[1\]\/x:div\[1\]\/x:div\[1\]\/x:(.*)\/text.*$', xpath)
            steps = len(match.group(1).split('/x:'))
            full_ladder = []
            for item in match.group(1).split('/x:'):
                full_ladder.append(int(re.match(r'.*\[(\d+)\]', item).group(1)))
            if len(full_ladder) < self.MAX_ELEMENT_DEPTH:
                for x in range(steps, self.MAX_ELEMENT_DEPTH):
                    full_ladder.append(0)
            else:
                full_ladder = full_ladder[:self.MAX_ELEMENT_DEPTH]
            fmt_str = '.'.join(["%04d"] * self.MAX_ELEMENT_DEPTH)
            return fmt_str % tuple(full_ladder)
        except:
            return False

    def _generate_marvin_hash_map(self, installed_books):
        '''
        Generate a map of book_ids to hash values
        {hash: [book_id, book_id,...], ...}
        '''
        self._log_location()
        hash_map = {}
        for book_id in installed_books:
            hash = installed_books[book_id].hash
            if hash in hash_map:
                hash_map[hash].append(book_id)
            else:
                hash_map[hash] = [book_id]
        return hash_map

    def _generate_reading_progress(self, book_data):
        '''
        Special-case progress:
              0% if book is marked NEW
            100% if book is marked Read
        '''
        #self._log_location(book_data.title)

        percent_read = ''
        if self.opts.prefs.get('show_progress_as_percentage', False):
            pct_progress = book_data.progress
            if 'NEW' in book_data.flags:
                percent_read = ''
                pct_progress = None
            elif 'READ' in book_data.flags:
                percent_read = "100%   "
                pct_progress = 1.0
            else:
                # Pad the right side for visual comfort, since this col is
                # right-aligned
                percent_read = "{:3.0f}%   ".format(book_data.progress * 100)
            progress = SortableTableWidgetItem(percent_read, pct_progress)
        else:
            base_name = "progress000.png"
            #base_name = "progress_none.png"
            pct_progress = book_data.progress
            if 'NEW' in book_data.flags:
                base_name = "progress_none.png"
                pct_progress = None
            elif 'READ' in book_data.flags:
                base_name = "progress100.png"
                pct_progress = 1.0
            elif book_data.progress >= 0.01 and book_data.progress < 0.11:
                base_name = "progress010.png"
            elif book_data.progress >= 0.11 and book_data.progress < 0.22:
                base_name = "progress020.png"
            elif book_data.progress >= 0.22 and book_data.progress < 0.33:
                base_name = "progress030.png"
            elif book_data.progress >= 0.33 and book_data.progress < 0.44:
                base_name = "progress040.png"
            elif book_data.progress >= 0.44 and book_data.progress < 0.55:
                base_name = "progress050.png"
            elif book_data.progress >= 0.55 and book_data.progress < 0.66:
                base_name = "progress060.png"
            elif book_data.progress >= 0.66 and book_data.progress < 0.77:
                base_name = "progress070.png"
            elif book_data.progress >= 0.77 and book_data.progress < 0.88:
                base_name = "progress080.png"
            elif book_data.progress >= 0.88 and book_data.progress < 0.95:
                base_name = "progress090.png"
            elif book_data.progress >= 0.95:
                base_name = "progress100.png"

            progress = SortableImageWidgetItem(os.path.join(self.parent.opts.resources_path,
                                                            'icons', base_name),
                                               pct_progress)
        return progress

    def _get_calibre_collections(self, cid):
        '''
        Return a sorted list of current calibre collection assignments or
        None if no collection_field_lookup assigned or book does not exist in library
        '''
        cfl = get_cc_mapping('collections', 'field', None)
        if cfl is None or cid is None:
            return None
        else:
            lib_collections = []
            db = self.opts.gui.current_db
            mi = db.get_metadata(cid, index_is_id=True)
            lib_collections = mi.get(cfl)
            if lib_collections:
                if type(lib_collections) is not list:
                    lib_collections = [lib_collections]
            return sorted(lib_collections, key=sort_key)

    def _get_epub_toc(self, path, prepend_title=None):
        '''
        Given a Marvin path, return the epub TOC indexed by section
        '''
        toc = None
        fpath = path

        # Find the OPF file in the zipped ePub
        zfo = cStringIO.StringIO(self.ios.read(fpath, mode='rb'))
        try:
            zf = ZipFile(zfo, 'r')
            container = etree.fromstring(zf.read('META-INF/container.xml'))
            opf_tree = etree.fromstring(zf.read(container.xpath('.//*[local-name()="rootfile"]')[0].get('full-path')))

            spine = opf_tree.xpath('.//*[local-name()="spine"]')[0]
            ncx_fs = spine.get('toc')
            manifest = opf_tree.xpath('.//*[local-name()="manifest"]')[0]
            ncx = manifest.find('.//*[@id="%s"]' % ncx_fs).get('href')

            # Find the ncx file
            fnames = zf.namelist()
            _ncx = [x for x in fnames if ncx in x][0]
            ncx_tree = etree.fromstring(zf.read(_ncx))
        except:
            import traceback
            self._log_location()
            self._log(" unable to unzip '%s'" % fpath)
            self._log(traceback.format_exc())
            return toc

        # fpath points to epub (zipped or unzipped dir)
        # spine, ncx_tree populated
        try:
            toc = OrderedDict()
            # 1. capture idrefs from spine
            for i, el in enumerate(spine):
                toc[str(i)] = el.get('idref')

            # 2. Resolve <spine> idrefs to <manifest> hrefs
            for el in toc:
                toc[el] = manifest.find('.//*[@id="%s"]' % toc[el]).get('href')

            # 3. Build a dict of src:toc_entry
            src_map = OrderedDict()
            navMap = ncx_tree.xpath('.//*[local-name()="navMap"]')[0]
            for navPoint in navMap:
                # Get the first-level entry
                src = re.sub(r'#.*$', '', navPoint.xpath('.//*[local-name()="content"]')[0].get('src'))
                toc_entry = navPoint.xpath('.//*[local-name()="text"]')[0].text
                src_map[src] = toc_entry

                # Get any nested navPoints
                nested_navPts = navPoint.xpath('.//*[local-name()="navPoint"]')
                for nnp in nested_navPts:
                    src = re.sub(r'#.*$', '', nnp.xpath('.//*[local-name()="content"]')[0].get('src'))
                    toc_entry = nnp.xpath('.//*[local-name()="text"]')[0].text
                    src_map[src] = toc_entry

            # Resolve src paths to toc_entry
            for section in toc:
                if toc[section] in src_map:
                    if prepend_title:
                        toc[section] = "%s &middot; %s" % (prepend_title,  src_map[toc[section]])
                    else:
                        toc[section] = src_map[toc[section]]
                else:
                    toc[section] = None

            # 5. Fill in the gaps
            current_toc_entry = None
            for section in toc:
                if toc[section] is None:
                    toc[section] = current_toc_entry
                else:
                    current_toc_entry = toc[section]
        except:
            import traceback
            self._log_location()
            self._log("{:~^80}".format(" error parsing '%s' " % fpath))
            self._log(traceback.format_exc())
            self._log("{:~^80}".format(" end traceback "))

        return toc

    def _get_formatted_annotations(self, book_id):
        '''
        '''
        # ~~~~~~~~~~ Emulating get_installed_books() ~~~~~~~~~~
        local_db_path = getattr(self.parent.connected_device, "local_db_path")
        #self._log("local_db_path: %s" % local_db_path)

        template = "{0}_books"
        books_db = template.format(re.sub('\W', '_', self.ios.device_name))
        #self._log("books_db: %s" % books_db)

        # Create the books table as needed (#272)
        self.opts.db.create_books_table(books_db)

        # Populate a BookStuct
        b_mi = BookStruct()
        b_mi.active = True
        b_mi.author = ', '.join(self.installed_books[book_id].author)
        b_mi.author_sort = self.installed_books[book_id].author_sort
        b_mi.book_id = book_id
        b_mi.title = self.installed_books[book_id].title
        b_mi.title_sort = self.installed_books[book_id].title_sort
        b_mi.uuid = self.installed_books[book_id].uuid

        # Add to books_db (#330)
        self.opts.db.add_to_books_db(books_db, b_mi)

        # Get the toc_entries (#344)
        path = '/'.join(['/Documents', self.installed_books[book_id].path])
        self.tocs = {}
        self.tocs[book_id] = self._get_epub_toc(path)

        # Update the timestamp (#347)
        self.opts.db.update_timestamp(books_db)
        self.opts.db.commit()

        # ~~~~~~~~~~ Emulating get_active_annotations() ~~~~~~~~~~
        template = "{0}_annotations"
        cached_db = template.format(re.sub('\W', '_', self.ios.device_name))
        self._log("cached_db: %s" % cached_db)

        # Create annotations table as needed (#153)
        self.opts.db.create_annotations_table(cached_db)

        # Fetch the annotations (#158)
        con = sqlite3.connect(local_db_path)

        with con:
            con.row_factory = sqlite3.Row
            cur = con.cursor()
            cur.execute('''
                           SELECT * FROM Highlights
                           WHERE BookId = '{0}'
                           ORDER BY NoteDateTime
                        '''.format(book_id))
            rows = cur.fetchall()
            for row in rows:
                # Sanitize text, note to unicode
                highlight_text = re.sub('\xa0', ' ', row[b'Text'])
                highlight_text = UnicodeDammit(highlight_text).unicode
                highlight_text = highlight_text.rstrip('\n').split('\n')
                while highlight_text.count(''):
                    highlight_text.remove('')
                highlight_text = [line.strip() for line in highlight_text]

                note_text = None
                if row[b'Note']:
                    ntu = UnicodeDammit(row[b'Note']).unicode
                    note_text = ntu.rstrip('\n')

                # Populate an AnnotationStruct
                a_mi = AnnotationStruct()
                a_mi.annotation_id = row[b'UUID']
                a_mi.book_id = book_id
                a_mi.highlight_color = self.HIGHLIGHT_COLORS[row[b'Colour']]
                a_mi.highlight_text = '\n'.join(highlight_text)
                a_mi.last_modification = row[b'NoteDateTime']

                section = str(int(row[b'Section']) - 1)
                try:
                    a_mi.location = self.tocs[book_id][section]
                except:
                    a_mi.location = "Section %s" % row[b'Section']

                a_mi.note_text = note_text

                # If empty highlight_text and empty note_text, not a useful annotation
                if not highlight_text and not note_text:
                    continue

                # Generate location_sort
                interior = self._generate_interior_location_sort(row[b'StartXPath'])
                if not interior:
                    self._log("Marvin: unable to parse xpath:")
                    self._log(row[b'StartXPath'])
                    self._log(a_mi)
                    continue

                a_mi.location_sort = "%04d.%s.%04d" % (
                    int(row[b'Section']),
                    interior,
                    int(row[b'StartOffset']))

                # Add annotation
                self.opts.db.add_to_annotations_db(cached_db, a_mi)

                # Update last_annotation in books_db
                self.opts.db.update_book_last_annotation(books_db, row[b'NoteDateTime'], book_id)

            # Update the timestamp
            self.opts.db.update_timestamp(cached_db)
            self.opts.db.commit()

        book_mi = BookStruct()
        book_mi.book_id = book_id
        book_mi.reader_app = 'Marvin'
        book_mi.title = self.installed_books[book_id].title
        formatted_annotations = self.opts.db.annotations_to_html(cached_db, book_mi)

        return formatted_annotations

    def _get_marvin_collections(self, book_id):
        return sorted(self.installed_books[book_id].device_collections, key=sort_key)

    def _get_installed_books(self):
        '''
        Build a profile of all installed books for display
        On Device
        Pin
        Title
        Author
        CalibreSeries
        CalibreSeriesIndex
        Last read
        Bookmarks
        Highlights/annotations
        Deep View content
        Vocabulary
        Percent read
        Flags
        Collections
        Word count
        hard match: uuid - green
        soft match: author/title or md5 contents/size match excluding OPF - yellow
        PinnedArticles + Wiki articles

        {mid: Book, ...}

        Try to use previously generated installed_books if available
        '''
        def _get_articles(cur, book_id):
            '''
            Return True if PinnedArticles or Wiki entries exist for book_id
            '''
            articles = {}

            # Get PinnedArticles
            a_cur = con.cursor()
            a_cur.execute('''SELECT
                              BookID,
                              Title,
                              URL
                             FROM PinnedArticles
                             WHERE BookID = '{0}'
                          '''.format(book_id))
            pinned_article_rows = a_cur.fetchall()

            if len(pinned_article_rows):
                pinned_articles = {}
                for row in pinned_article_rows:
                    pinned_articles[row[b'Title']] = row[b'URL']
                articles['Pinned'] = pinned_articles

            # Get Wiki snippets
            a_cur.execute('''SELECT
                              BookID,
                              Title,
                              Snippet
                             FROM Wiki
                             WHERE BookID = '{0}'
                          '''.format(book_id))
            wiki_rows = a_cur.fetchall()

            if len(wiki_rows):
                wiki_snippets = {}
                for row in wiki_rows:
                    wiki_snippets[row[b'Title']] = row[b'Snippet']
                articles['Wiki'] = wiki_snippets

            return articles

        def _get_calibre_id(uuid, title, author):
            '''
            Find book in library, return cid, mi
            '''
            if self.opts.prefs.get('development_mode', False):
                self._log_location("%s %s" % (repr(title), repr(author)))
            cid = None
            mi = None
            db = self.opts.gui.current_db
            try:
                if uuid in self.library_uuid_map:
                    cid = self.library_uuid_map[uuid]['id']
                    mi = db.get_metadata(cid, index_is_id=True, get_cover=True, cover_as_data=True)
                    if self.opts.prefs.get('development_mode', False):
                        self._log("UUID match: %s" % uuid)
                elif title in self.library_title_map:
                    _cid = self.library_title_map[title]['id']
                    _mi = db.get_metadata(_cid, index_is_id=True, get_cover=True, cover_as_data=True)
                    authors = author.split(', ')
                    if authors == _mi.authors:
                        cid = _cid
                        mi = _mi
                        if self.opts.prefs.get('development_mode', False):
                            self._log("TITLE/AUTHOR match")
            except:
                # Book deleted since scan
                import traceback
                self._log_location(traceback.format_exc())

            # Confirm valid mi object
            if getattr(mi, 'uuid', None) == 'dummy':
                mi = None
            return cid, mi

        def _get_collections(cur, book_id):
            # Get the collection assignments
            ca_cur = con.cursor()
            ca_cur.execute('''SELECT
                                BookID,
                                CollectionID
                              FROM BookCollections
                              WHERE BookID = '{0}'
                           '''.format(book_id))
            collections = []
            collection_rows = ca_cur.fetchall()
            if collection_rows is not None:
                collection_assignments = [collection[b'CollectionID']
                                          for collection in collection_rows]
                collections += [collection_map[item] for item in collection_assignments]
                collections = sorted(collections, key=sort_key)
            ca_cur.close()
            return collections

        def _get_flags(cur, row):
            # Get the flag assignments
            flags = []
            if row[b'NewFlag']:
                flags.append(self.FLAGS['new'])
            if row[b'ReadingList']:
                flags.append(self.FLAGS['reading_list'])
            if row[b'IsRead']:
                flags.append(self.FLAGS['read'])
            return flags

        def _get_highlights(cur, book_id):
            '''
            Return highlight text/notes associated with book_id
            '''
            hl_cur = con.cursor()
            hl_cur.execute('''SELECT
                                Note,
                                Text
                              FROM Highlights
                              WHERE BookID = '{0}'
                           '''.format(book_id))
            hl_rows = hl_cur.fetchall()
            highlight_list = []
            if len(hl_rows):
                for row in hl_rows:
                    raw_text = row[b'Text']
                    text = "<p>{0}".format(raw_text)
                    if row[b'Note']:
                        raw_note = row[b'Note']
                        text += "<br/>&nbsp;<em>{0}</em></p>".format(raw_note)
                    else:
                        text += "</p>"
                    highlight_list.append(text)
            hl_cur.close()
            return highlight_list

        def _get_marvin_genres(book_id):
            # Return sorted genre(s) for this book
            genre_cur = con.cursor()
            genre_cur.execute('''SELECT
                                    Subject
                                 FROM BookSubjects
                                 WHERE BookID = '{0}'
                              '''.format(book_id))
            genres = []
            genre_rows = genre_cur.fetchall()
            if genre_rows is not None:
                genres = [genre[b'Subject'] for genre in genre_rows]
            genre_cur.close()
            genres = sorted(genres, key=sort_key)
            return genres

        def _get_metadata_mismatches(cur, book_id, row, mi, this_book):
            '''
            Return dict of metadata mismatches.
            author, author_sort, pubdate, publisher, series, series_index, title,
            title_sort, description, subjects, collections, cover
            '''
            def _get_cover_hash(mi, this_book):
                '''
                Retrieve cover_hash from archive, or create/store
                '''
                #self._log_location(this_book.title)
                ach = self.archived_cover_hashes.get(str(this_book.cid), {})
                cover_last_modified = self.opts.gui.current_db.cover_last_modified(this_book.cid, index_is_id=True)
                if ('cover_last_modified' in ach and
                        ach['cover_last_modified'] == cover_last_modified):
                    #self._log("returning cached cover_hash %s" % ach['cover_hash'])
                    return ach['cover_hash']

                # Generate calibre cover hash (same process used by driver when sending books)
                cover_hash = '0'
                desired_thumbnail_height = self.parent.connected_device.THUMBNAIL_HEIGHT
                try:
                    #self._log("mi.cover_data[0]: %s" % repr(mi.cover_data[0]))
                    sized_thumb = thumbnail(mi.cover_data[1],
                                            desired_thumbnail_height,
                                            desired_thumbnail_height)
                    cover_hash = hashlib.md5(sized_thumb[2]).hexdigest()
                    cover_last_modified = self.opts.gui.current_db.cover_last_modified(this_book.cid, index_is_id=True)
                    self.archived_cover_hashes.set(str(this_book.cid),
                                                   {'cover_hash': cover_hash,
                                                    'cover_last_modified': cover_last_modified})
                except:
                    if mi.cover_data[1]:
                        self._log_location("error calculating cover_hash for %s (cid %d)" %
                        (this_book.title, this_book.cid))
                    else:
                        self._log_location("no cover available for %s" % this_book.title)
                return cover_hash

            #self._log_location(row[b'Title'])
            mismatches = {}
            if mi is not None:
                # ~~~~~~~~ authors ~~~~~~~~
                if mi.authors != this_book.authors:
                    mismatches['authors'] = {'calibre': mi.authors,
                                             'Marvin': this_book.authors}

                # ~~~~~~~~ author_sort ~~~~~~~~
                if mi.author_sort != row[b'AuthorSort']:
                    mismatches['author_sort'] = {'calibre': mi.author_sort,
                                                 'Marvin': row[b'AuthorSort']}

                # ~~~~~~~~ cover_hash ~~~~~~~~
                cover_hash = _get_cover_hash(mi, this_book)
                if cover_hash != row[b'CalibreCoverHash']:
                    mismatches['cover_hash'] = {'calibre': cover_hash,
                                                'Marvin': row[b'CalibreCoverHash']}

                # ~~~~~~~~ pubdate ~~~~~~~~
                if (mi.pubdate.year == 101 and mi.pubdate.month == 1 and
                    not row[b'DatePublished']):
                    # Special case when calibre pubdate is unknown (101-01-01) and
                    # Marvin is None
                    pass
                else:
                    if bool(row[b'DatePublished']) or bool(mi.pubdate):
                        mb_pubdate = None
                        if row[b'DatePublished']:
                            try:
                                mb_pubdate = datetime.utcfromtimestamp(int(row[b'DatePublished']))
                                mb_pubdate = mb_pubdate.replace(hour=0, minute=0, second=0)
                            except:
                                if iswindows:
                                    ''' Windows doesn't like negative timestamps '''
                                    epoch = datetime(1970, 1, 1)
                                    mb_pubdate = epoch + timedelta(seconds=int(row[b'DatePublished']))
                                else:
                                    self._log("Error getting pubdate for %s" % repr(row[b'Title']))
                                    self._log("DatePublished: %s" % repr(row[b'DatePublished']))
                                    import traceback
                                    self._log(traceback.format_exc())
                                    mb_pubdate = None

                        naive = mi.pubdate.replace(hour=0, minute=0, second=0, tzinfo=None)

                        if naive and mb_pubdate:
                            td = naive - mb_pubdate
                            if abs(td.days) > 1:
                                mismatches['pubdate'] = {'calibre': naive,
                                                         'Marvin': mb_pubdate}
                        elif naive != mb_pubdate:
                            # One of them is None
                            mismatches['pubdate'] = {'calibre': naive,
                                                     'Marvin': mb_pubdate}

                # ~~~~~~~~ publisher ~~~~~~~~
                if mi.publisher != row[b'Publisher']:
                    if not (mi.publisher is None and row[b'Publisher'] == 'Unknown'):
                        mismatches['publisher'] = {'calibre': mi.publisher,
                                                   'Marvin': row[b'Publisher']}

                # ~~~~~~~~ series, series_index ~~~~~~~~
                # We only care about series_index if series is assigned
                if bool(mi.series) or bool(row[b'CalibreSeries']):
                    if mi.series != row[b'CalibreSeries']:
                        mismatches['series'] = {'calibre': mi.series,
                                                'Marvin': row[b'CalibreSeries']}

                    csi = row[b'CalibreSeriesIndex'] if row[b'CalibreSeriesIndex'] else 0.0
                    if bool(mi.series_index) or bool(csi):
                        if mi.series_index != float(csi):
                            mismatches['series_index'] = {'calibre': mi.series_index,
                                                          'Marvin': csi}

                # ~~~~~~~~ title ~~~~~~~~
                if mi.title != row[b'Title']:
                    mismatches['title'] = {'calibre': mi.title,
                                           'Marvin': row[b'Title']}

                # ~~~~~~~~ title_sort ~~~~~~~~
                if mi.title_sort != row[b'CalibreTitleSort']:
                    mismatches['title_sort'] = {'calibre': mi.title_sort,
                                                'Marvin': row[b'CalibreTitleSort']}

                # ~~~~~~~~ comments ~~~~~~~~
                if bool(mi.comments) or bool(row[b'Description']):
                    if mi.comments != row[b'Description']:
                        mismatches['comments'] = {'calibre': mi.comments,
                                                  'Marvin': row[b'Description']}

                # ~~~~~~~~ tags ~~~~~~~~
                if sorted(mi.tags, key=sort_key) != _get_marvin_genres(book_id):
                    mismatches['tags'] = {'calibre': sorted(mi.tags, key=sort_key),
                                          'Marvin': _get_marvin_genres(book_id)}

                # ~~~~~~~~ uuid ~~~~~~~~
                if mi.uuid != row[b'UUID']:
                    mismatches['uuid'] = {'calibre': mi.uuid,
                                          'Marvin': row[b'UUID']}

            else:
                #self._log("(no calibre metadata for %s)" % row[b'Title'])
                pass

            return mismatches

        def _get_on_device_status(cid):
            '''
            Given a uuid, return the on_device status of the book
            '''
            ans = None
            if cid:
                db = self.opts.gui.current_db
                mi = db.get_metadata(cid, index_is_id=True)
                try:
                    ans = mi._proxy_metadata.ondevice_col
                except:
                    self._log_location("ERROR: ondevice_col not available for '%s'" % mi.title)
                    #self._log(mi.all_field_keys())
                    ans = None
            return ans

        def _get_pubdate(row):
            pubdate = None
            if row[b'DatePublished'] != '' and row[b'DatePublished'] is not None:
                try:
                    pubdate = datetime.utcfromtimestamp(int(row[b'DatePublished']))
                except:
                    if iswindows:
                        ''' Windows doesn't like negative timestamps '''
                        epoch = datetime(1970, 1, 1)
                        pubdate = epoch + timedelta(seconds=int(row[b'DatePublished']))
                    else:
                        self._log("Error getting pubdate for %s" % repr(row[b'Title']))
                        self._log("DatePublished: %s" % repr(row[b'DatePublished']))
                        import traceback
                        self._log(traceback.format_exc())
            return pubdate

        def _get_publisher(row):
            publisher = row[b'Publisher']
            if publisher == 'Unknown':
                publisher = None
            return publisher

        def _get_vocabulary_list(cur, book_id):
            # Get the vocabulary content
            voc_cur = con.cursor()
            voc_cur.execute('''SELECT
                                BookID,
                                Word
                              FROM Vocabulary
                              WHERE BookID = '{0}'
                           '''.format(book_id))

            vocabulary_rows = voc_cur.fetchall()
            vocabulary_list = []
            if len(vocabulary_rows):
                vocabulary_list = [vocabulary_item[b'Word']
                                   for vocabulary_item in vocabulary_rows]
                vocabulary_list = sorted(vocabulary_list, key=sort_key)
            voc_cur.close()
            return vocabulary_list

        def _purge_cover_hash_orphans():
            '''
            Purge obsolete cover hashes
            '''
            self._log_location()
            # Get active cids
            active_cids = sorted([str(installed_books[book_id].cid) for book_id in installed_books])
            #self._log("active_cids: %s" % active_cids)

            # Get active cover_hash cids
            cover_hash_cids = sorted(self.archived_cover_hashes.keys())
            #self._log("cover_hash keys: %s" % cover_hash_cids)

            self._busy_panel_setup("Removing obsolete cover hashes")
            for ch_cid in cover_hash_cids:
                if ch_cid not in active_cids:
                    self._log("removing orphan cid %s from archived_cover_hashes" % ch_cid)
                    del self.archived_cover_hashes[ch_cid]
            self._busy_panel_teardown()

        # ~~~~~~~~~~~~~ Entry point ~~~~~~~~~~~~~~~~~~

        self._log_location()

        marvin_content_updated = getattr(self.parent, 'marvin_content_updated', False)
        installed_books = getattr(self.parent, 'installed_books', None)
        if installed_books is None or marvin_content_updated:
            if marvin_content_updated:
                setattr(self.parent, 'marvin_content_updated', False)

            installed_books = {}

            # Wait for device driver to complete initialization, but tell user what's happening
            if not hasattr(self.parent.connected_device, "cached_books"):
                self._busy_panel_setup("Waiting for driver to finish initialization…")

            while True:
                if not hasattr(self.parent.connected_device, "cached_books"):
                    Application.processEvents()
                else:
                    if self.busy_panel is not None:
                        self._busy_panel_teardown()
                    break

            # Is there a valid mainDb?
            local_db_path = getattr(self.parent.connected_device, "local_db_path")
            if local_db_path is not None:
                # Fetch/compute hashes
                cached_books = self.parent.connected_device.cached_books
                hashes = self._scan_marvin_books(cached_books)

                # Get the mainDb data
                con = sqlite3.connect(self.parent.connected_device.local_db_path)
                with con:
                    con.row_factory = sqlite3.Row

                    # Build a collection map
                    collections_cur = con.cursor()
                    collections_cur.execute('''SELECT
                                                ID,
                                                Name
                                               FROM Collections
                                            ''')
                    rows = collections_cur.fetchall()
                    collection_map = {}
                    for row in rows:
                        collection_map[row[b'ID']] = row[b'Name']
                    collections_cur.close()

                    # Get the books
                    cur = con.cursor()
                    cur.execute('''SELECT
                                    Author,
                                    AuthorSort,
                                    Books.ID as id_,
                                    CalibreCoverHash,
                                    CalibreSeries,
                                    CalibreSeriesIndex,
                                    CalibreTitleSort,
                                    CoverFile,
                                    DateAdded,
                                    DateOpened,
                                    DatePublished,
                                    DeepViewPrepared,
                                    Description,
                                    FileName,
                                    IsRead,
                                    NewFlag,
                                    Pin,
                                    Progress,
                                    Publisher,
                                    ReadingList,
                                    Title,
                                    UUID,
                                    WordCount
                                  FROM Books
                                ''')

                    rows = cur.fetchall()

                    pb = ProgressBar(parent=self.opts.gui, window_title="Scanning Marvin library: 2 of 2")
                    book_count = len(rows)
                    pb.set_maximum(book_count)
                    pb.set_value(0)
                    pb.set_label('{:^100}'.format("Performing metadata magic…"))
                    pb.show()

                    for i, row in enumerate(rows):
                        try:
                            cid, mi = _get_calibre_id(row[b'UUID'],
                                                      row[b'Title'],
                                                      row[b'Author'])

                            book_id = row[b'id_']
                            # Get the primary metadata from Books
                            this_book = Book(row[b'Title'], row[b'Author'].split(', '))
                            this_book.articles = _get_articles(cur, book_id)
                            this_book.author_sort = row[b'AuthorSort']
                            this_book.cid = cid
                            this_book.calibre_collections = self._get_calibre_collections(this_book.cid)
                            this_book.comments = row[b'Description']
                            this_book.cover_file = row[b'CoverFile']
                            this_book.date_added = row[b'DateAdded']
                            this_book.date_opened = row[b'DateOpened']
                            this_book.device_collections = _get_collections(cur, book_id)
                            this_book.deep_view_prepared = row[b'DeepViewPrepared']
                            this_book.flags = _get_flags(cur, row)
                            this_book.hash = hashes[row[b'FileName']]['hash']
                            this_book.highlights = _get_highlights(cur, book_id)
                            this_book.match_quality = None  # Added in _construct_table_data()
                            this_book.metadata_mismatches = _get_metadata_mismatches(cur, book_id, row, mi, this_book)
                            this_book.mid = book_id
                            this_book.on_device = _get_on_device_status(this_book.cid)
                            this_book.path = row[b'FileName']
                            this_book.pin = row[b'Pin']
                            this_book.progress = row[b'Progress']
                            this_book.pubdate = _get_pubdate(row)
                            this_book.series = row[b'CalibreSeries']
                            this_book.series_index = row[b'CalibreSeriesIndex']
                            this_book.tags = _get_marvin_genres(book_id)
                            this_book.title_sort = row[b'CalibreTitleSort']
                            this_book.uuid = row[b'UUID']
                            this_book.vocabulary = _get_vocabulary_list(cur, book_id)
                            this_book.word_count = locale.format("%d", row[b'WordCount'], grouping=True)
                            installed_books[book_id] = this_book
                        except:
                            self._log("ERROR adding to installed_books")
                            import traceback
                            self._log(traceback.format_exc())

                        pb.increment()

                    pb.hide()

                # Remove orphan cover_hashes, but only if we're dealing with entire library
                mdb = self.opts.gui.library_view.model().db
                current_vl = mdb.data.get_base_restriction_name()
                if current_vl == '':
                    _purge_cover_hash_orphans()

                if self.opts.prefs.get('development_mode', False):
                    self._log("%d cached books from Marvin:" % len(cached_books))
                    for book in installed_books:
                        self._log("%s %s %s" % (installed_books[book].title,
                                             repr(installed_books[book].authors),
                                             installed_books[book].hash))

            else:
                self._log("Marvin database is damaged")
                title = "Damaged database"
                msg = "<p>Marvin database is damaged. Unable to retrieve Marvin library.</p>"
                MessageBox(MessageBox.ERROR, title, msg,
                           show_copy_button=False).exec_()

        return installed_books

    def _inject_css(self, html):
        '''
        stick a <style> element into html
        Deep View content structured differently
        <html style=""><body style="">
        '''
        css = self.prefs.get('injected_css', None)
        if css:
            try:
                styled_soup = BeautifulSoup(html)
                head = styled_soup.find("head")
                style_tag = Tag(styled_soup, 'style')
                style_tag['type'] = "text/css"
                style_tag.insert(0, css)
                head.insert(0, style_tag)
                html = styled_soup.renderContents()
            except:
                return html
        return(html)

    def _inform_calibre_flags(self, book_id, update_gui=True):
        '''
        Update enabled custom columns to current flag settings
        '''
        read_lookup = get_cc_mapping('read', 'field', None)
        reading_list_lookup = get_cc_mapping('reading_list', 'field', None)

        if (read_lookup or reading_list_lookup):
            self._log_location()
            db = self.opts.gui.current_db
            cid = self.installed_books[book_id].cid
            mi = db.get_metadata(cid, index_is_id=True)
            flags = self.installed_books[book_id].flags
            db_requires_update = False

            if read_lookup:
                c_read_um = mi.metadata_for_field(read_lookup)
                if 'READ' in flags and not c_read_um['#value#']:
                    self._log("setting READ (%s)" % read_lookup)
                    c_read_um['#value#'] = 1
                    db_requires_update = True
                    mi.set_user_metadata(read_lookup, c_read_um)
                elif 'READ' not in flags and c_read_um['#value#']:
                    self._log("clearing READ (%s)" % read_lookup)
                    c_read_um['#value#'] = None
                    db_requires_update = True
                    mi.set_user_metadata(read_lookup, c_read_um)
                else:
                    self._log("calibre Read flag already correct")

            if reading_list_lookup:
                c_reading_list_um = mi.metadata_for_field(reading_list_lookup)
                if 'READING LIST' in flags and not c_reading_list_um['#value#']:
                    self._log("setting READING LIST (%s)" % reading_list_lookup)
                    c_reading_list_um['#value#'] = 1
                    db_requires_update = True
                    mi.set_user_metadata(reading_list_lookup, c_reading_list_um)
                elif 'READING LIST' not in flags and c_reading_list_um['#value#']:
                    self._log("clearing READING LIST (%s)" % reading_list_lookup)
                    c_reading_list_um['#value#'] = None
                    db_requires_update = True
                    mi.set_user_metadata(reading_list_lookup, c_reading_list_um)
                else:
                    self._log("calibre Reading list flag already matches")

            if db_requires_update:
                db.set_metadata(cid, mi, set_title=False, set_authors=False,
                    commit=True, force_changes=True)
            if update_gui:
                updateCalibreGUIView()

    def _inform_marvin_collections(self, book_id, update_local_db=True):
        '''
        Inform Marvin of updated flags + collections
        '''
        # ~~~~~~~~ Update Marvin with Flags + Collections ~~~~~~~~
        command_name = 'update_metadata_items'
        command_element = 'updatemetadataitems'
        update_soup = BeautifulStoneSoup(self.METADATA_COMMAND_XML.format(
            command_element, time.mktime(time.localtime())))
        book_tag = Tag(update_soup, 'book')
        book_tag['author'] = escape(', '.join(self.installed_books[book_id].authors))
        book_tag['filename'] = self.installed_books[book_id].path
        book_tag['title'] = self.installed_books[book_id].title
        book_tag['uuid'] = self.installed_books[book_id].uuid

        flags = self.installed_books[book_id].flags
        collections = self.installed_books[book_id].device_collections
        merged = sorted(flags + collections, key=sort_key)

        collections_tag = Tag(update_soup, 'collections')
        for tag in sorted(merged, key=sort_key):
            c_tag = Tag(update_soup, 'collection')
            c_tag.insert(0, escape(tag))
            collections_tag.insert(0, c_tag)
        book_tag.insert(0, collections_tag)

        update_soup.manifest.insert(0, book_tag)

        local_busy = False
        if self.busy:
            self._busy_status_msg(msg=self.UPDATING_MARVIN_MESSAGE)
        else:
            local_busy = True
            self._busy_status_setup(msg=self.UPDATING_MARVIN_MESSAGE)
        results = self._issue_command(command_name, update_soup,
                                      update_local_db=update_local_db)
        if local_busy:
            self._busy_status_teardown()

        if results['code']:
            return self._show_command_error(command_name, results)

    def _issue_command(self, command_name, update_soup,
                       get_response=None,
                       timeout_override=None,
                       update_local_db=True):
        '''
        Consolidated command handler
        '''
        self._log_location()

        QApplication.setOverrideCursor(QCursor(Qt.WaitCursor))

        # Wait for the driver to be silent
        while self.parent.connected_device.get_busy_flag():
            Application.processEvents()
        self.parent.connected_device.set_busy_flag(True)

        # Copy command file to staging folder
        self._stage_command_file(command_name, update_soup,
                                 show_command=self.prefs.get('show_staged_commands', False))

        # Wait for completion
        try:
            results = self._wait_for_command_completion(command_name,
                timeout_override=timeout_override,
                get_response=get_response,
                update_local_db=update_local_db)
        except:
            import traceback
            details = "An error occurred while executing '{0}'.\n\n".format(command_name)
            details += traceback.format_exc()
            results = {'code': '2',
                       'status': "Error communicating with Marvin",
                       'details': details}

        # Try to reset the busy flag, although it might fail
        try:
            self.parent.connected_device.set_busy_flag(False)
        except:
            pass

        QApplication.restoreOverrideCursor()
        return results

    def _localize_marvin_database(self):
        '''
        Copy remote_db_path from iOS to local storage using device pointers
        '''
        self._log_location("starting")
        msg = "Refreshing database"
        local_busy = False
        if self.busy:
            self._busy_status_msg(msg=msg)
        else:
            local_busy = True
            self._busy_status_setup(msg=msg)

        local_db_path = self.parent.connected_device.local_db_path
        remote_db_path = self.parent.connected_device.books_subpath

        # Report size of remote_db
        stats = self.ios.exists(remote_db_path)
        self._log("mainDb: {:,} bytes".format(int(stats['st_size'])))

        with open(local_db_path, 'wb') as out:
            self.ios.copy_from_idevice(remote_db_path, out)

        if local_busy:
            self._busy_status_teardown()
        self._log_location("finished")

    def _localize_hash_cache(self, cached_books):
        '''
        Check for existence of hash cache on iDevice. Confirm/create folder
        If existing cached, purge orphans
        '''
        self._log_location()

        # Existing hash cache?
        lhc = os.path.join(self.local_cache_folder, self.HASH_CACHE_FS)
        rhc = '/'.join([self.REMOTE_CACHE_FOLDER, self.HASH_CACHE_FS])

        cache_exists = (self.ios.exists(rhc) and
                        not self.opts.prefs.get('hash_caching_disabled'))
        if cache_exists:
            # Copy from existing remote cache to local cache
            with open(lhc, 'wb') as out:
                self.ios.copy_from_idevice(str(rhc), out)

            # Load hash_cache to memory
            with open(lhc, 'rb') as hcf:
                hash_cache = pickle.load(hcf)

            self._log("remote hash cache: v{0}, {1} books in cache".format(
                hash_cache['version'],
                len(hash_cache) - 1))

        else:
            # Confirm path to remote folder is valid store point
            folder_exists = self.ios.exists(self.REMOTE_CACHE_FOLDER)
            if not folder_exists:
                self._log("creating remote_cache_folder %s" % repr(self.REMOTE_CACHE_FOLDER))
                self.ios.mkdir(self.REMOTE_CACHE_FOLDER)

            # Create a local cache
            with open(lhc, 'wb') as hcf:
                hash_cache = {'version': 1}
                pickle.dump(hash_cache, hcf, pickle.HIGHEST_PROTOCOL)
            self._log("creating new local hash cache: version %d" %
                      hash_cache['version'])

            """
            # Clear the marvin_content_updated flag
            if getattr(self.parent, 'marvin_content_updated', False):
                self._log("clearing marvin_content_updated flag")
                setattr(self.parent, 'marvin_content_updated', False)
            """

        self.local_hash_cache = lhc
        self.remote_hash_cache = rhc

        # Purge cache orphans, but only if we're looking at entire library.
        mdb = self.opts.gui.library_view.model().db
        current_vl = mdb.data.get_base_restriction_name()

        if cache_exists and current_vl == '':
            hash_cache = self._purge_cached_orphans(cached_books)

        return hash_cache

    def _purge_cached_orphans(self, cached_books):
        '''

        '''
        self._log_location()

        # Find the orphans
        orphans = []
        with open(self.local_hash_cache, 'rb') as hcf:
            hash_cache = pickle.load(hcf)
            for key in hash_cache:
                if key not in cached_books and key != 'version':
                    self._log("removing %s from hash cache" % key)
                    orphans.append(key)

        # Remove the orphans
        for key in orphans:
            hash_cache.pop(key)

        # Write updated hash_cache
        with open(self.local_hash_cache, 'wb') as hcf:
            pickle.dump(hash_cache, hcf, pickle.HIGHEST_PROTOCOL)

        return hash_cache

    def _report_calibre_duplicates(self):
        '''
        Scan for multiple UUIDs matching single hash
        Displayed as MAGENTA in MXD
        '''
        apply_markers = self.prefs.get('apply_markers_to_duplicates', True)
        self._log_location("apply_markers: %s" % apply_markers)

        # Build a list of Marvin hashes
        marvin_hashes = [v.hash for v in self.installed_books.values()]

        library_hash_map = self.library_scanner.hash_map
        duplicates = []
        for hash in sorted(library_hash_map):
            if len(library_hash_map[hash]) > 1 and hash in marvin_hashes:
                titles = []
                for uuid in library_hash_map[hash]:
                    titles.append("'{0}' ({1})".format(
                        self.library_scanner.uuid_map[uuid]['title'],
                        self.library_scanner.uuid_map[uuid]['id']))
                    if apply_markers:
                        self.soloed_books.add(self.library_scanner.uuid_map[uuid]['id'])
                duplicates.append(titles)

        if duplicates:
            if self.soloed_books:
                 self.parent.gui.library_view.model().db.set_marked_ids(self.soloed_books)

            details = ''
            for duplicate_set in duplicates:
                details += '- ' + ', '.join(duplicate_set) + '\n'

            title = 'Duplicate content'
            if apply_markers:
                marker_msg = ('<p>Duplicates will be temporarily marked in the ' +
                              'Library window. Temporary markers for duplicate content ' +
                              'may be disabled in the Marvin XD configuration dialog.</p>')
            else:
                marker_msg = ('<p>Duplicate content may be temporarily marked in the ' +
                              'Library window by enabling the option in the ' +
                              'Marvin XD configuration dialog.</p>' )

            msg = ('<p>Duplicates were detected while scanning your calibre library.<p>' +
                   '<p>Marvin books matching multiple calibre books will be displayed ' +
                   'with a ' +
                   '<span style="background-color:#FF99E5">magenta background</span> ' +
                   'in the Marvin XD window.</p>' +
                   marker_msg +
                   '<p>Click <b>Show details</b> to display duplicates.</p>')
            MessageBox(MessageBox.WARNING, title, msg, det_msg=details,
                       show_copy_button=True).exec_()

    def _report_content_updates(self):
        '''
        Report books identified as being installed in Marvin without hash matches
        LIGHT_GRAY: Book has been updated in calibre. (UUIDs match, different hashes)
        DARK_GRAY:  Book has been updated in Marvin. (UUIDs do not match, different hashes)
        '''
        apply_markers = self.prefs.get('apply_markers_to_updated', True)
        self._log_location("apply_markers: %s" % apply_markers)
        calibre_updates = ''
        marvin_updates = ''
        for this_book in self.installed_books.values():
            if this_book.match_quality == self.MATCH_COLORS.index('LIGHT_GRAY'):
                if apply_markers:
                    self.soloed_books.add(this_book.cid)
                calibre_updates += "- {0}\n".format(this_book.title)
            if this_book.match_quality == self.MATCH_COLORS.index('DARK_GRAY'):
                if apply_markers:
                    self.soloed_books.add(this_book.cid)
                marvin_updates += "- {0}\n".format(this_book.title)

        if calibre_updates or marvin_updates:
            if self.soloed_books:
                 self.parent.gui.library_view.model().db.set_marked_ids(self.soloed_books)

            title = 'Updated content'
            if apply_markers:
                marker_msg = ('<p>Books with updated content will be temporarily marked in the ' +
                              'Library window. Temporary markers for updated content ' +
                              'may be disabled in the Marvin XD configuration dialog.</p>')
            else:
                marker_msg = ('<p>Books with updated content may be temporarily marked in the ' +
                              'Library window by enabling the option in the ' +
                              'Marvin XD configuration dialog.</p>' )

            msg = ('<p>Updated content was detected while comparing your calibre ' +
                   'library with your Marvin library.</p>' +
                   '<p>Books updated in calibre will be displayed with a ' +
                   '<span style="background-color:#D9D9D9">light gray background</span> ' +
                   'in the Marvin XD window.</p>' +
                   '<p>Books updated in Marvin will be displayed with a ' +
                   '<span style="color:#FFFFFF; background-color:#989898">' +
                   'dark gray background</span> ' +
                   'in the Marvin XD window.</p>' +
                   marker_msg +
                   '<p>Click <b>Show details</b> for a list of books with updated content.</p>')

            details = ''
            if calibre_updates:
                details += 'Books updated in calibre:\n' + calibre_updates
            if marvin_updates:
                details += 'Books updated in Marvin:\n' + marvin_updates

            MessageBox(MessageBox.WARNING, title, msg, det_msg=details,
                       show_copy_button=True).exec_()

    def _save_column_widths(self):
        '''
        '''
        self._log_location()
        try:
            widths = []
            for (i, c) in enumerate(self.LIBRARY_HEADER):
                widths.append(self.tv.columnWidth(i))
            self.opts.prefs.set('marvin_library_column_widths', widths)
            self.opts.prefs.commit()
        except:
            import traceback
            self._log(traceback.format_exc())

    def _scan_library_books(self, library_scanner):
        '''
        Generate hashes for library epubs
        '''
        def _get_hash(cid):
            path = db.format(cid, 'epub', index_is_id=True,
                             as_path=True, preserve_filename=True)
            hash = self._compute_epub_hash(path)
            os.remove(path)
            return hash

        pb = ProgressBar(parent=self.opts.gui, window_title="Scanning calibre library")
        pb.set_label('{:^100}'.format("Waiting for library scan to complete…"))
        pb.set_value(0)
        pb.show()

        # Scan library books for hashes
        if self.library_scanner.isRunning():
            #self._busy_panel_setup("Waiting for library scan to complete…")
            #Application.processEvents()
            self.library_scanner.wait()
            #self._busy_panel_teardown()

        uuid_map = library_scanner.uuid_map
        total_books = len(uuid_map)
        self._log_location("%d" % total_books)

        pb.set_maximum(total_books)
        pb.set_label('{:^100}'.format("Identifying %d books in calibre library…" % (total_books)))

        db = self.opts.gui.current_db

        if False:
            '''
            Determine if there have been any changes to this lib since we last scanned it
            last_modified appears to always change even when no user changes. Ask KG
            '''
            lib_name = os.path.dirname(db.dbpath).split(os.path.sep)[-1]
            last_modified = time.mktime(db.last_modified().timetuple())
            library_snapshots = self.opts.prefs.get('calibre_library_snapshots', {})
            rescan_required = True
            if lib_name in library_snapshots:
                self._log("calibre_library_snapshots: %s" % library_snapshots[lib_name])
                if library_snapshots[lib_name] == last_modified:
                    rescan_required = False
                    self._log("No changes detected since last scan")
                else:
                    self._log("lib appears to have changed")
                    self._log("last_modified: %s" % repr(last_modified))
                    self._log("library_snapshots[lib_name]: %s" % repr(library_snapshots[lib_name]))
            else:
                self._log("%s not found in calibre_library_snapshots" % lib_name)

            # Store last_modified to prefs for future reference
            library_snapshots[lib_name] = last_modified
            self.opts.prefs.set('calibre_library_snapshots', library_snapshots)

        close_requested = False

        all_cached_hashes = db.get_all_custom_book_data('epub_hash')
        for k, v in all_cached_hashes.items():
            all_cached_hashes[k] = json.loads(v)

        for i, uuid in enumerate(uuid_map):
            try:
                cid = uuid_map[uuid]['id']

                # Do we have a cached hash?
                cached_hash = all_cached_hashes.get(cid, None)
                if cached_hash is None:
                    # Generate the hash, save it to local hash map
                    #self._log("generating hash")
                    hash = _get_hash(cid)
                    uuid_map[uuid]['hash'] = hash

                    # Cache hash to db
                    #self._log("adding cached_hash to db")
                    mtime = db.format_last_modified(cid, 'epub')
                    cached_dict = {'mtime': time.mktime(mtime.timetuple()), 'hash': hash}
                    db.add_custom_book_data(cid, 'epub_hash', json.dumps(cached_dict))
                else:
                    mtime = db.format_last_modified(cid, 'epub')
                    if cached_hash['mtime'] == time.mktime(mtime.timetuple()):
                        hash = cached_hash['hash']
                        uuid_map[uuid]['hash'] = hash
                    else:
                        # Book has been modified since we generated the hash
                        hash = _get_hash(cid)
                        uuid_map[uuid]['hash'] = hash
                        if self.opts.prefs.get('development_mode', False):
                            self._log("generating new hash for '{0}': {1}".format(
                                uuid_map[uuid]['title'], hash))

                        # Update db
                        cached_dict = {'mtime': time.mktime(mtime.timetuple()), 'hash': hash}
                        db.add_custom_book_data(cid, 'epub_hash', json.dumps(cached_dict))

            except:
                # Book deleted since scan?

                if self.opts.prefs.get('development_mode', False):
                    import traceback
                    self._log(traceback.format_exc())

            pb.increment()

            if pb.close_requested:
                close_requested = True
                break
        else:
            # Only build the hash map if we completed without a close request
            hash_map = library_scanner.build_hash_map()

        pb.hide()

        if close_requested:
            raise AbortRequestException("user cancelled library scan")

        return hash_map

    def _scan_marvin_books(self, cached_books):
        '''
        Create the initial dict of installed books with hash values
        '''
        self._log_location("%d books" % len(cached_books))

        # Fetch pre-existing hash cache from device, purge orphans
        self.hash_cache = self._localize_hash_cache(cached_books)

        # Set up the progress bar
        pb = ProgressBar(parent=self.opts.gui, window_title="Scanning Marvin library: 1 of 2")
        total_books = len(cached_books)
        pb.set_maximum(total_books)
        pb.set_value(0)
        pb.set_label('{:^100}'.format("Identifying %d books in Marvin library…" % (total_books)))
        pb.show()

        close_requested = False
        installed_books = {}
        for i, path in enumerate(cached_books):
            this_book = {}
            #pb.set_label('{:^100}'.format("%d of %d" % (i+1, total_books)))
            this_book['hash'] = self._fetch_marvin_content_hash(path)

            installed_books[path] = this_book
            pb.increment()

            if pb.close_requested:
                close_requested = True
                break
        else:
            # Store the updated hash_cache if we finished
            with open(self.local_hash_cache, 'wb') as hcf:
                pickle.dump(self.hash_cache, hcf, pickle.HIGHEST_PROTOCOL)

            # Push the local hash to the iDevice
            self._update_remote_hash_cache()

        pb.hide()

        if close_requested:
            raise AbortRequestException("user cancelled Marvin scan")

        return installed_books

    def _selected_book_id(self, row):
        '''
        Return selected Marvin book_id
        '''
        return self.tm.get_book_id(row)

    def _selected_books(self):
        '''
        Generate a dict of books selected in the dialog
        '''
        selected_books = {}

        for row in self._selected_rows():
            author = str(self.tm.get_author(row).text())
            book_id = self.tm.get_book_id(row)
            cid = self.tm.get_calibre_id(row)
            has_annotations = self.tm.get_annotations(row).sort_key
            has_articles = self.tm.get_articles(row).sort_key
            has_dv_content = bool(self.tm.get_deep_view(row))
            has_vocabulary = self.tm.get_vocabulary(row).sort_key
            last_opened = str(self.tm.get_last_opened(row).text())
            locked = self.tm.get_locked(row).sort_key
            path = self.tm.get_path(row)
            progress = self.tm.get_progress(row).sort_key
            title = str(self.tm.get_title(row).text())
            uuid = self.tm.get_uuid(row)
            word_count = self.tm.get_word_count(row).sort_key
            selected_books[row] = {
                'author': author,
                'book_id': book_id,
                'cid': cid,
                'has_annotations': has_annotations,
                'has_articles': has_articles,
                'has_dv_content': has_dv_content,
                'has_vocabulary': has_vocabulary,
                'last_opened': last_opened,
                'locked': locked,
                'path': path,
                'progress': progress,
                'title': title,
                'uuid': uuid,
                'word_count': word_count}

        return selected_books

    def _selected_cid(self, row):
        '''
        Return selected calibre id
        '''
        return self.tm.get_calibre_id(row)

    def _selected_rows(self):
        '''
        Return a list of selected rows
        '''
        srs = self.tv.selectionModel().selectedRows()
        return [sr.row() for sr in srs]

    def _set_flags(self, action, update_local_db=True):
        '''
        Set specified flags for selected books
        '''
        def _build_flag_list(flagbits):
            flags = []
            if flagbits & self.NEW_FLAG:
                flags.append(self.FLAGS['new'])
            if flagbits & self.READING_FLAG:
                flags.append(self.FLAGS['reading_list'])
            if flagbits & self.READ_FLAG:
                flags.append(self.FLAGS['read'])
            return flags

        def _update_in_memory(book_id, path):
            flags = self.installed_books[book_id].flags
            collections = self.installed_books[book_id].device_collections
            merged = sorted(flags + collections, key=sort_key)

            # Update driver (cached_books)
            cached_books = self.parent.connected_device.cached_books
            cached_books[path]['device_collections'] = merged

            # Update Device model
            for row in self.opts.gui.memory_view.model().map:
                book = self.opts.gui.memory_view.model().db[row]
                if book.path == path:
                    book.device_collections = merged
                    break

        self._log_location(action)
        if action == 'set_new_flag':
            mask = self.NEW_FLAG
            inhibit = self.NEW_FLAG + self.READING_FLAG
        elif action == 'set_reading_list_flag':
            mask = self.READING_FLAG
            inhibit = self.NEW_FLAG + self.READING_FLAG + self.READ_FLAG
        elif action == 'set_read_flag':
            mask = self.READ_FLAG
            inhibit = self.READING_FLAG + self.READ_FLAG

        local_db_update_required = False

        # Save the currently selected rows
        self.saved_selection_region = self.tv.visualRegionForSelection(self.tv.selectionModel().selection())

        selected_books = self._selected_books()
        for row in selected_books:
            self.tv.selectRow(row)
            book_id = selected_books[row]['book_id']
            flagbits = self.tm.get_flags(row).sort_key

            if flagbits != mask:
                path = selected_books[row]['path']
                if not flagbits & mask:
                    # Set the bit with OR
                    flagbits = flagbits | mask
                    flagbits = flagbits & inhibit
                    basename = "flags%d.png" % flagbits
                    new_flags_widget = SortableImageWidgetItem(os.path.join(self.parent.opts.resources_path,
                                                                            'icons', basename),
                                                               flagbits)
                    # Update the spreadsheet
                    self.tm.set_flags(row, new_flags_widget)

                    # Update self.installed_books flags list
                    self.installed_books[book_id].flags = _build_flag_list(flagbits)

                    # Update reading progress based on flag values
                    self._update_reading_progress(self.installed_books[book_id], row)

                    # Update in-memory
                    _update_in_memory(book_id, path)

                    # Update Marvin db, calibre custom columns
                    self._inform_marvin_collections(book_id, update_local_db=False)
                    local_db_update_required = True
                    self._update_device_flags(book_id, path, _build_flag_list(flagbits))
            else:
                self._log("Marvin flags already correct")

            self._inform_calibre_flags(book_id)

        # Restore selection
        if self.saved_selection_region:
            for rect in self.saved_selection_region.rects():
                self.tv.setSelection(rect, QItemSelectionModel.Select)
            self.saved_selection_region = None

        if update_local_db and local_db_update_required:
            self._localize_marvin_database()

        Application.processEvents()

    def _show_command_error(self, command, results):
        '''
        Display contents of a non-successful result
        '''
        self._log_location(results)
        title = "Results"
        msg = ("<p>Error communicating with Marvin while executing <tt>{0}</tt> command.</p>".format(command) +
               "<p>Click <b>Show details</b> for more information.</p>")
        details = results['details']
        MessageBox(MessageBox.WARNING, title, msg, det_msg=details,
                   show_copy_button=False).exec_()

    def _stage_command_file(self, command_name, command_soup, show_command=False):

        self._log_location(command_name)

        if show_command:
            if command_name in ['update_metadata', 'update_metadata_items']:
                soup = BeautifulStoneSoup(command_soup.renderContents())
                # <descriptions>
                descriptions = soup.findAll('description')
                for description in descriptions:
                    d_tag = Tag(soup, 'description')
                    d_tag.insert(0, "(description removed for debug stream)")
                    description.replaceWith(d_tag)
                # <covers>
                covers = soup.findAll('cover')
                for cover in covers:
                    cover_tag = Tag(soup, 'cover')
                    cover_tag['hash'] = cover['hash']
                    cover_tag['encoding'] = cover['encoding']
                    cover_tag.insert(0, "(cover bytes removed for debug stream)")
                    cover.replaceWith(cover_tag)
                self._log(soup.prettify())
            else:
                self._log("command_name: %s" % command_name)
                self._log(command_soup.prettify())

        if self.prefs.get('execute_marvin_commands', True):

            self.ios.write(command_soup.renderContents(),
                           b'/'.join([self.parent.connected_device.staging_folder, b'%s.tmp' % command_name]))
            self.ios.rename(b'/'.join([self.parent.connected_device.staging_folder, b'%s.tmp' % command_name]),
                            b'/'.join([self.parent.connected_device.staging_folder, b'%s.xml' % command_name]))

        else:
            self._log("~~~ execute_marvin_commands disabled in JSON ~~~")

    def _synchronize_flags(self):
        '''
        Iteratively synchronize each selected row
        '''
        self._log_location()

        # Save the currently selected rows
        self.saved_selection_region = self.tv.visualRegionForSelection(self.tv.selectionModel().selection())

        selected_books = self._selected_books()
        for row in selected_books:
            self.tv.selectRow(row)
            self._apply_flags()

        # Restore selection
        if self.saved_selection_region:
            for rect in self.saved_selection_region.rects():
                self.tv.setSelection(rect, QItemSelectionModel.Select)
            self.saved_selection_region = None

        self._localize_marvin_database()

        Application.processEvents()

    def _toggle_locked_status(self, row):
        '''
        '''
        current_status = self.tm.get_locked(row).sort_key
        self._log_location("current_status: %s" %  current_status)

        if current_status == 1:
            action = "set_unlocked"
        else:
            action = "set_locked"

        self._update_locked_status(action)

    def _update_calibre_collections(self, book_id, cid, updated_calibre_collections):
        '''
        '''
        self._log_location()
        # Update collections custom column
        lookup = get_cc_mapping('collections', 'field', None)
        if lookup is not None and cid is not None:
            # Get the current value from the lookup field
            db = self.opts.gui.current_db
            mi = db.get_metadata(cid, index_is_id=True)
            #old_collections = mi.get_user_metadata(lookup, False)['#value#']
            #self._log("Updating old collections value: %s" % repr(old_collections))

            um = mi.metadata_for_field(lookup)
            um['#value#'] = updated_calibre_collections
            mi.set_user_metadata(lookup, um)
            db.set_metadata(cid, mi, set_title=False, set_authors=False,
                            commit=True)
            db.commit()

        # Update in-memory
        self.installed_books[book_id].calibre_collections = updated_calibre_collections

    def _update_calibre_metadata(self, book_id, cid, mismatches, model_row, update_local_db=True):
        '''
        Update calibre from Marvin metadata
        If uuids differ, we need to send an update_metadata command to Marvin
        pb is incremented twice per book.
        '''

        # Highlight the row we're working on
        self.tv.selectRow(model_row)

        # Get the current metadata
        db = self.opts.gui.current_db
        mi = db.get_metadata(cid, index_is_id=True, get_cover=True, cover_as_data=True)

        self._log_location("{0} cid:{1}".format(repr(mi.title), cid))
        if self.opts.prefs.get('development_mode', False):
            self._log("mismatches:\n%s" % mismatches)

        # We need these if uuid needs to be updated
        cached_books = self.parent.connected_device.cached_books
        path = self.installed_books[book_id].path

        # Find the book in Device map
        for device_view_row in self.opts.gui.memory_view.model().map:
            book = self.opts.gui.memory_view.model().db[device_view_row]
            if book.path == path:
                break
        else:
            self._log("ERROR: couldn't find '%s' in memory_view" % path)
            device_view_row = None

        '''
        Update calibre metadata from Marvin
        mismatch keys:
            authors, author_sort, comments, cover_hash, pubdate, publisher,
            series, series_index, tags, title, title_sort, uuid
        Process in alpha order so that cover change are processed before changing uuid
        cover_hash and uuid have special handling, as they require Marvin to be notified of changes
        '''
        for key in sorted(mismatches.keys()):
            if key == 'authors':
                authors = mismatches[key]['Marvin']
                db.set_authors(cid, authors, allow_case_change=True)

            if key == 'author_sort':
                author_sort = mismatches[key]['Marvin']
                db.set_author_sort(cid, author_sort)

            if key == 'comments':
                comments = mismatches[key]['Marvin']
                db.set_comment(cid, comments)

            if key == 'cover_hash':
                # If covers don't match, import Marvin cover, then send it back with new hash
                cover_hash = mismatches[key]['Marvin']
                # Get the Marvin cover, add it to calibre
                marvin_cover = self._fetch_marvin_cover(book_id)
                if marvin_cover is not None:
                    db.set_cover(cid, marvin_cover)

                    desired_thumbnail_height = self.parent.connected_device.THUMBNAIL_HEIGHT
                    cover = thumbnail(marvin_cover,
                                      desired_thumbnail_height,
                                      desired_thumbnail_height)
                    cover_hash = hashlib.md5(cover[2]).hexdigest()

                    # Tell Marvin about the updated cover_hash
                    command_name = 'update_metadata_items'
                    command_element = 'updatemetadataitems'
                    update_soup = BeautifulStoneSoup(self.METADATA_COMMAND_XML.format(
                        command_element, time.mktime(time.localtime())))
                    book_tag = Tag(update_soup, 'book')
                    book_tag['author'] = escape(', '.join(self.installed_books[book_id].authors))
                    book_tag['filename'] = self.installed_books[book_id].path
                    book_tag['title'] = self.installed_books[book_id].title
                    book_tag['uuid'] = mismatches[key]['Marvin']

                    cover_tag = Tag(update_soup, 'cover')
                    cover_tag['hash'] = cover_hash
                    cover_tag['encoding'] = 'base64'
                    cover_tag.insert(0, base64.b64encode(marvin_cover))
                    book_tag.insert(0, cover_tag)

                    update_soup.manifest.insert(0, book_tag)

                    results = self._issue_command(command_name, update_soup,
                                                  update_local_db=update_local_db)
                    if results['code']:
                        return self._show_command_error(command_name, results)

                    # Update cached_books
                    cached_books[path]['cover_hash'] = cover_hash
                else:
                    self._log("No cover data available from Marvin")

            if key == 'pubdate':
                pubdate = mismatches[key]['Marvin']
                db.set_pubdate(cid, pubdate)

            if key == 'publisher':
                publisher = mismatches[key]['Marvin']
                db.set_publisher(cid, publisher, allow_case_change=True)

            if key == 'series':
                series = mismatches[key]['Marvin']
                db.set_series(cid, series, allow_case_change=True)

            if key == 'series_index':
                series_index = mismatches[key]['Marvin']
                db.set_series_index(cid, series_index)

            if key == 'tags':
                tags = sorted(mismatches[key]['Marvin'], key=sort_key)
                db.set_tags(cid, tags, allow_case_change=True)

            if key == 'title':
                title = mismatches[key]['Marvin']
                db.set_title(cid, title)

            if key == 'title_sort':
                title_sort = mismatches[key]['Marvin']
                db.set_title_sort(cid, title_sort)

            if key == 'uuid':
                # Update Marvin's uuid to match calibre's
                uuid = mismatches[key]['calibre']
                cached_books[path]['uuid'] = uuid
                self.installed_books[book_id].matches = [uuid]

                self.installed_books[book_id].uuid = uuid
                self.opts.gui.memory_view.model().db[device_view_row].uuid = uuid
                self.opts.gui.memory_view.model().db[device_view_row].in_library = "UUID"

                # Add uuid to hash_map
                self.library_scanner.add_to_hash_map(self.installed_books[book_id].hash, uuid)

                # Tell Marvin about the updated uuid
                command_name = 'update_metadata_items'
                command_element = 'updatemetadataitems'
                update_soup = BeautifulStoneSoup(self.METADATA_COMMAND_XML.format(
                    command_element, time.mktime(time.localtime())))
                book_tag = Tag(update_soup, 'book')
                book_tag['author'] = escape(', '.join(self.installed_books[book_id].authors))
                book_tag['filename'] = self.installed_books[book_id].path
                book_tag['title'] = self.installed_books[book_id].title
                book_tag['uuid'] = mismatches[key]['Marvin']

                book_tag['newuuid'] = mismatches[key]['calibre']
                update_soup.manifest.insert(0, book_tag)

                results = self._issue_command(command_name, update_soup,
                                              update_local_db=update_local_db)
                if results['code']:
                    #return self._show_command_error(command_name, results)
                    return results

        self._clear_selected_rows()

        # Update metadata match quality in the visible model
        old = self.tm.get_match_quality(model_row)
        new = self.MATCH_COLORS.index('GREEN')
        self.tm.set_match_quality(model_row, new)
        self.updated_match_quality[model_row] = {'book_id': book_id,
                                                 'old': old,
                                                 'new': new}
        return None

    def _update_collection_match(self, book, row):
        '''
        Refresh the Collections column
        '''
        self._log_location(row)
        match_status = self._generate_collection_match(book)
        self.tm.set_collections(row, match_status)

    def _update_collections(self, action):
        '''
        Apply action to selected books.
            export_collections
            import_collections
            synchronize_collections
            clear_all_collections
        '''
        self._log_location(action)

        selected_books = self._selected_books()
        for row in selected_books:
            book_id = self._selected_book_id(row)
            cid = self._selected_cid(row)
            calibre_collections = self._get_calibre_collections(cid)
            marvin_collections = self._get_marvin_collections(book_id)
            if action == 'export_collections':
                # Apply calibre collections to Marvin
                self._log("export_collections: %s" % selected_books[row]['title'])
                self._update_marvin_collections(book_id, calibre_collections)
                self._update_collection_match(self.installed_books[book_id], row)

            elif action == 'import_collections':
                # Apply Marvin collections to calibre
                self._log("import_collections: %s" % selected_books[row]['title'])
                self._update_calibre_collections(book_id, cid, marvin_collections)
                self._update_collection_match(self.installed_books[book_id], row)

            elif action == 'synchronize_collections':
                # Merged collections to both calibre and Marvin
                self._log("synchronize_collections: %s" % selected_books[row]['title'])
                cl = set(calibre_collections)
                ml = set(marvin_collections)
                deltas = ml - cl
                merged_collections = sorted(calibre_collections + list(deltas), key=sort_key)
                self._update_calibre_collections(book_id, cid, merged_collections)
                self._update_marvin_collections(book_id, merged_collections)
                self._update_collection_match(self.installed_books[book_id], row)

            elif action == 'clear_all_collections':
                # Remove all collection assignments from both calibre and Marvin
                self._log("clear_all_collections: %s" % selected_books[row]['title'])
                self._update_calibre_collections(book_id, cid, [])
                self._update_marvin_collections(book_id, [])
                self._update_collection_match(self.installed_books[book_id], row)

            else:
                self._log("unsupported action: %s" % action)
                title = "Update collections"
                msg = ("<p>{0}: not implemented</p>".format(action))
                MessageBox(MessageBox.INFO, title, msg,
                           show_copy_button=False).exec_()

    def _update_device_flags(self, book_id, path, updated_flags):
        '''
        Given a set of updated flags for path, update local copies:
            cached_books[path]['device_collections']
            Device model book.device_collections
            (installed_books already updated in _clear_flags, _set_flags())
        '''
        self._log_location("%s: %s" % (self.installed_books[book_id].title, updated_flags))

        # Get current collection assignments
        marvin_collections = self.installed_books[book_id].device_collections
        updated_collections = sorted(updated_flags + marvin_collections, key=sort_key)

        # Update driver
        cached_books = self.parent.connected_device.cached_books
        cached_books[path]['device_collections'] = updated_collections

        # Update Device model
        for row in self.opts.gui.memory_view.model().map:
            book = self.opts.gui.memory_view.model().db[row]
            if book.path == path:
                book.device_collections = updated_collections
                break

    def _update_flags(self, action):
        '''
        Context menu entry point
        '''
        rows_to_refresh = len(self._selected_books())

        if action in ['clear_new_flag', 'clear_reading_list_flag',
                      'clear_read_flag', 'clear_all_flags']:
            self._clear_flags(action)

        elif action in ['set_new_flag', 'set_reading_list_flag', 'set_read_flag']:
            self._set_flags(action)

        else:
            self._log("unsupported action: %s" % action)
            title = "Update flags"
            msg = ("<p>{0}: not implemented</p>".format(action))
            MessageBox(MessageBox.INFO, title, msg,
                       show_copy_button=False).exec_()

        title = 'Flags updated'
        msg = ("<p>Flags updated for {0}.</p>".format(
            "1 book" if rows_to_refresh == 1 else "{0} books".format(rows_to_refresh)))
        MessageBox(MessageBox.INFO, title, msg, det_msg='', show_copy_button=False).exec_()

    def _update_global_collections(self, details):
        '''
        details = {'rename': {'<old_name>': '<new_name>', ...}
                   'delete': (['<collection_to_delete>', ...])
                   'active_cids': [cid, cid, cid ...],
                   'locations': {'calibre': [collection, collection, ...],
                                 'Marvin': [collection, collection, ...]}
                  }

        Deletions:
            - Calibre: if collection in locations['calibre']:
                test each active_cid for removal
                update self.installed_books[].calibre_collections
            - Marvin: if collection in locations['Marvin']:
                Delete collection from cached_books[].device_collections
                Delete collection from self.installed_books[].device_collections
                Delete collection from Device memory_view[].device_collections
                Tell Marvin about the deletion(s)

        Renaming:
            - Calibre: if collection in locations['calibre']:
                test each active_cid for renaming
                update self.installed_books[].calibre_collections
            - Marvin: if collection in locations['Marvin']:
                Rename collection from cached_books[].device_collections
                Rename collection from self.installed_books[].device_collections
                Rename collection from Device memory_view[].device_collections
                Tell Marvin about the renaming

        '''
        self._log_location(details)

        # ~~~~~~ calibre ~~~~~~
        lookup = get_cc_mapping('collections', 'field', None)
        if lookup is not None and details['active_cids']:

            deleted_in_calibre = []
            for ctd in details['delete']:
                previous_name = None
                for k, v in details['rename'].items():
                    if v == ctd:
                        previous_name = k
                        break

                if (ctd in details['locations']['calibre'] or
                        previous_name in details['locations']['calibre']):

                    if previous_name:
                        # Switch identity
                        ctd = previous_name
                        deleted_in_calibre.append(ctd)

                    for cid in details['active_cids']:
                        # Get the current value from the lookup field
                        db = self.opts.gui.current_db
                        mi = db.get_metadata(cid, index_is_id=True)
                        collections = mi.get_user_metadata(lookup, False)['#value#']
                        self._log("%s: old collections value: %s" % (mi.title, repr(collections)))
                        if ctd in collections:
                            collections.remove(ctd)
                            self._log("new collections value: %s" % (repr(collections)))
                            um = mi.metadata_for_field(lookup)
                            um['#value#'] = collections
                            mi.set_user_metadata(lookup, um)
                            db.set_metadata(cid, mi, set_title=False, set_authors=False,
                                            commit=True)

                            # Update in-memory if book exists in Marvin
                            book_id = self._find_cid_in_model(cid)
                            if book_id:
                                self._log("old calibre_collections: %s" % self.installed_books[book_id].calibre_collections)
                                self.installed_books[book_id].calibre_collections = collections

            for ctr in details['rename']:
                if ctr in details['locations']['calibre'] and ctr not in deleted_in_calibre:
                    for cid in details['active_cids']:
                        # Get the current value from the lookup field
                        db = self.opts.gui.current_db
                        mi = db.get_metadata(cid, index_is_id=True)
                        collections = mi.get_user_metadata(lookup, False)['#value#']
                        self._log("%s: old collections value: %s" % (mi.title, repr(collections)))
                        if ctr in collections:
                            collections.remove(ctr)
                            replacement = details['rename'][ctr]
                            collections.append(replacement)
                            collections.sort(key=sort_key)
                            self._log("new collections value: %s" % (repr(collections)))
                            um = mi.metadata_for_field(lookup)
                            um['#value#'] = collections
                            mi.set_user_metadata(lookup, um)
                            db.set_metadata(cid, mi, set_title=False, set_authors=False,
                                            commit=True)

                            # Update in-memory if book exists in Marvin
                            book_id = self._find_cid_in_model(cid)
                            if book_id:
                                self._log("old calibre_collections: %s" % self.installed_books[book_id].calibre_collections)
                                self.installed_books[book_id].calibre_collections = collections

        # ~~~~~~ Marvin ~~~~~~

        cached_books = self.parent.connected_device.cached_books

        deleted_in_marvin = []
        for ctd in details['delete']:
            previous_name = None
            for k, v in details['rename'].items():
                if v == ctd:
                    previous_name = k
                    break

            if (ctd in details['locations']['Marvin'] or
                    previous_name in details['locations']['Marvin']):

                if previous_name:
                    # Switch identity, delete from details['rename']
                    ctd = previous_name
                    deleted_in_marvin.append(ctd)

                # Issue one update per deleted collection
                for book_id, book in self.installed_books.items():

                    command_name = "command"
                    command_type = "UpdateGlobalCollections"
                    update_soup = BeautifulStoneSoup(self.GENERAL_COMMAND_XML.format(
                        command_type, time.mktime(time.localtime())))
                    parameters_tag = Tag(update_soup, 'parameters')
                    update_soup.command.insert(0, parameters_tag)

                    if ctd in book.device_collections:
                        self._log("%s: delete '%s'" % (book.title, book.device_collections))
                        book.device_collections.remove(ctd)

                        # Update Device model
                        for row in self.opts.gui.memory_view.model().map:
                            model_book = self.opts.gui.memory_view.model().db[row]
                            if model_book.path == book.path:
                                model_book.device_collections = book.device_collections
                                break

                        # Update driver
                        cached_books[book.path]['device_collections'] = book.device_collections

                        # Add a <parameter action="delete"> tag
                        parameter_tag = Tag(update_soup, 'parameter')
                        parameter_tag['name'] = "action"
                        parameter_tag.insert(0, "delete")
                        parameters_tag.insert(0, parameter_tag)

                        parameter_tag = Tag(update_soup, 'parameter')
                        parameter_tag['name'] = "name"
                        parameter_tag.insert(0, ctd)
                        parameters_tag.insert(0, parameter_tag)

                        results = self._issue_command(command_name, update_soup)
                        if results['code']:
                            return self._show_command_error(command_type, results)

        for ctr in details['rename']:
            if ctr in details['locations']['Marvin'] and ctr not in deleted_in_marvin:

                # Issue one update per renamed collection
                for book_id, book in self.installed_books.items():

                    command_name = "command"
                    command_type = "UpdateGlobalCollections"
                    update_soup = BeautifulStoneSoup(self.GENERAL_COMMAND_XML.format(
                        command_type, time.mktime(time.localtime())))
                    parameters_tag = Tag(update_soup, 'parameters')
                    update_soup.command.insert(0, parameters_tag)

                    if ctr in book.device_collections:
                        self._log("%s: rename '%s'" % (book.title, book.device_collections))
                        book.device_collections.remove(ctr)
                        replacement = details['rename'][ctr]
                        book.device_collections.append(replacement)
                        book.device_collections.sort(key=sort_key)

                        # Update Device model
                        for row in self.opts.gui.memory_view.model().map:
                            model_book = self.opts.gui.memory_view.model().db[row]
                            if model_book.path == book.path:
                                model_book.device_collections = book.device_collections
                                break

                        # Update driver
                        cached_books[book.path]['device_collections'] = book.device_collections

                        # Add the parameter tags
                        parameter_tag = Tag(update_soup, 'parameter')
                        parameter_tag['name'] = "action"
                        parameter_tag.insert(0, "rename")
                        parameters_tag.insert(0, parameter_tag)

                        parameter_tag = Tag(update_soup, 'parameter')
                        parameter_tag['name'] = "name"
                        parameter_tag.insert(0, ctr)
                        parameters_tag.insert(0, parameter_tag)

                        parameter_tag = Tag(update_soup, 'parameter')
                        parameter_tag['name'] = "newname"
                        parameter_tag.insert(0, replacement)
                        parameters_tag.insert(0, parameter_tag)

                        results = self._issue_command(command_name, update_soup)
                        if results['code']:
                            return self._show_command_error(command_type, results)

    def _update_locked_status(self, action):
        '''
        '''
        self._log_location(action)

        selected_books = self._selected_books()

        if action == 'set_locked':
            new_pin_value = 1
            new_image_name = "lock_enabled.png"
            command_type = "LockBooks"
        elif action == 'set_unlocked':
            new_pin_value = 0
            new_image_name = "empty_16x16.png"
            #new_image_name = "unlock_enabled.png"
            command_type = "UnlockBooks"

        # Build the command shell
        command_name = "command"
        update_soup = BeautifulStoneSoup(self.GENERAL_COMMAND_XML.format(
            command_type, time.mktime(time.localtime())))
        manifest_tag = Tag(update_soup, 'manifest')
        update_soup.command.insert(0, manifest_tag)

        new_locked_widget = SortableImageWidgetItem(os.path.join(self.parent.opts.resources_path,
                                                                 'icons', new_image_name),
                                                    new_pin_value)

        for row in selected_books:
            # Update the spreadsheet
            book_id = selected_books[row]['book_id']
            self.tm.set_locked(row, new_locked_widget)

            # Update self.installed_books
            self.installed_books[book_id].pin = new_pin_value

            # Add the book to the manifest
            book_tag = Tag(update_soup, 'book')
            book_tag['author'] = escape(', '.join(self.installed_books[book_id].authors))
            book_tag['filename'] = self.installed_books[book_id].path
            book_tag['title'] = self.installed_books[book_id].title
            book_tag['uuid'] = self.installed_books[book_id].uuid
            manifest_tag.insert(0, book_tag)

        show_spinner = bool(len(selected_books) > self.MAX_BOOKS_BEFORE_SPINNER)
        if show_spinner:
            self._busy_status_setup(msg=self.UPDATING_MARVIN_MESSAGE)
        results = self._issue_command(command_name, update_soup, update_local_db=True)

        if show_spinner:
            self._busy_status_teardown()

        if results['code']:
            return self._show_command_error('update_locked_status', results)

    def _update_marvin_collections(self, book_id, updated_marvin_collections):
        '''
        '''
        self._log_location()

        # Update installed_books
        self.installed_books[book_id].device_collections = updated_marvin_collections

        # Merge active flags with updated marvin collections
        cached_books = self.parent.connected_device.cached_books
        path = self.installed_books[book_id].path
        active_flags = self.installed_books[book_id].flags
        updated_collections = sorted(active_flags +
                                     updated_marvin_collections,
                                     key=sort_key)

        # Update driver (cached_books)
        cached_books[path]['device_collections'] = updated_collections

        # Update Device model
        for row in self.opts.gui.memory_view.model().map:
            book = self.opts.gui.memory_view.model().db[row]
            if book.path == path:
                book.device_collections = updated_collections
                break

        # Tell Marvin about the changes
        self._inform_marvin_collections(book_id)

    def _update_marvin_metadata(self, book_id, cid, mismatches, model_row):
        '''
        Update Marvin from calibre metadata
        This clones upload_books() in the iOS reader application driver
        All metadata is asserted, cover optional if changes

        Books in gui.memory_view.model().db are Metadata objects
        self._log("standard_field_keys: %s" % self.opts.gui.memory_view.model().db[0].standard_field_keys())
        '''

        # Highlight the row we're working on
        self.tv.selectRow(model_row)

        # Get the current metadata
        db = self.opts.gui.current_db
        mi = db.get_metadata(cid, index_is_id=True, get_cover=True, cover_as_data=True)

        self._log_location("'{0}' cid:{1}".format(mi.title, cid))
        #self._log("mismatches:\n%s" % mismatches)

        command_name = "update_metadata"
        update_soup = self._build_metadata_update(book_id, cid, mi, mismatches)
        results = self._issue_command(command_name, update_soup)
        if results['code']:
            #return self._show_command_error('_update_marvin_metadata', results)
            return results

        # Update in-memory caches
        cached_books = self.parent.connected_device.cached_books
        path = self.installed_books[book_id].path

        # Find the book in Device map
        for device_view_row in self.opts.gui.memory_view.model().map:
            try:
                book = self.opts.gui.memory_view.model().db[device_view_row]
                if book.path == path:
                    break
            except:
                import traceback
                self._log("ERROR: invalid device_view_row %s" % device_view_row)
                self._log(traceback.format_exc())

        else:
            # If we didn't find the path, then possibly the book was updated/replaced
            # If the book was originally downloaded via OPDS, we should have a uuid match
            if 'uuid' not in mismatches:
                self._log("path not found in memory_view, scanning by uuid")
                uuid = self.installed_books[book_id].uuid
                for device_view_row in self.opts.gui.memory_view.model().map:
                    book = self.opts.gui.memory_view.model().db[device_view_row]
                    if book.uuid == uuid:
                        break
                else:
                    self._log("ERROR: uuid '%s' not found in memory_view" % uuid)
                    device_view_row = None
            else:
                self._log("ERROR: path '%s' not found in memory_view, uuid mismatch" % path)
                self._log(" Device view will not be updated")
                device_view_row = None

        '''
        We need to tweak the in-memory versions of the Marvin library as if they had
        been loaded initially.
        mismatch keys:
            authors, author_sort, comments, cover_hash, pubdate, publisher,
            series, series_index, tags, title, title_sort, uuid
        visible/relevant memory_view properties (order of appearance):
            in_library, title, title_sort, authors, author_sort, device_collections
        '''

        for key in mismatches:
            if key == 'authors':
                authors = mismatches[key]['calibre']
                cached_books[path]['authors'] = authors
                cached_books[path]['author'] = ', '.join(authors)
                self.installed_books[book_id].authors = authors
                if device_view_row:
                    self.opts.gui.memory_view.model().db[device_view_row].authors = authors

            if key == 'author_sort':
                author_sort = mismatches[key]['calibre']
                cached_books[path]['author_sort'] = author_sort
                self.installed_books[book_id].author_sort = author_sort
                if device_view_row:
                    self.opts.gui.memory_view.model().db[device_view_row].author_sort = author_sort

            if key == 'comments':
                comments = mismatches[key]['calibre']
                cached_books[path]['description'] = comments
                self.installed_books[book_id].comments = comments

            if key == 'cover_hash':
                cover_hash = mismatches[key]['calibre']
                cached_books[path]['cover_hash'] = cover_hash

            if key == 'pubdate':
                pubdate = mismatches[key]['calibre']
                cached_books[path]['pubdate'] = pubdate
                self.installed_books[book_id].pubdate = pubdate

            if key == 'publisher':
                publisher = mismatches[key]['calibre']
                cached_books[path]['publisher'] = publisher

            if key == 'series':
                series = mismatches[key]['calibre']
                cached_books[path]['series'] = series

            if key == 'series_index':
                series_index = mismatches[key]['calibre']
                cached_books[path]['series_index'] = series_index

            if key == 'tags':
                tags = mismatches[key]['calibre']
                cached_books[path]['tags'] = tags
                self.installed_books[book_id].tags = tags

            if key == 'title':
                title = mismatches[key]['calibre']
                cached_books[path]['title'] = title
                self.installed_books[book_id].title = title
                if device_view_row:
                    self.opts.gui.memory_view.model().db[device_view_row].title = title

            if key == 'title_sort':
                title_sort = mismatches[key]['calibre']
                cached_books[path]['title_sort'] = title_sort
                self.installed_books[book_id].title_sort = title_sort
                if device_view_row:
                    self.opts.gui.memory_view.model().db[device_view_row].title_sort = title_sort

            if key == 'uuid':
                uuid = mismatches[key]['calibre']
                cached_books[path]['uuid'] = uuid
                self.installed_books[book_id].matches = [uuid]

                self.installed_books[book_id].uuid = uuid
                if device_view_row:
                    self.opts.gui.memory_view.model().db[device_view_row].uuid = uuid
                    self.opts.gui.memory_view.model().db[device_view_row].in_library = "UUID"

                # Add to hash_map
                self.library_scanner.add_to_hash_map(self.installed_books[book_id].hash, uuid)

            self._clear_selected_rows()

        # Update metadata match quality in the visible model
        old = self.tm.get_match_quality(model_row)
        self.tm.set_match_quality(model_row, self.MATCH_COLORS.index('GREEN'))
        self.updated_match_quality[model_row] = {'book_id': book_id,
                                                 'old': old,
                                                 'new': self.MATCH_COLORS.index('GREEN')}
        return None

    def _update_metadata(self, action):
        '''
        Dispatched method is responsible for updating progress bar twice per book
        '''
        self._log_location(action)

        # Save the selection region for restoration
        self.saved_selection_region = self.tv.visualRegionForSelection(self.tv.selectionModel().selection())

        selected_books = self._selected_books()

        for row in selected_books:
            book_id = self._selected_book_id(row)
            if self.tm.get_match_quality(row) == self.MATCH_COLORS.index('ORANGE'):
                title = "Duplicate book"
                msg = ("<p>'{0}' is a duplicate.</p>".format(self.installed_books[book_id].title) +
                       "<p>Remove duplicates before updating metadata.</p>")
                return MessageBox(MessageBox.WARNING, title, msg,
                                  show_copy_button=False).exec_()

        total_books = len(selected_books)
        self._busy_status_setup(show_cancel=total_books > 1)
        self.updated_match_quality = {}
        errors = []

        for i, row in enumerate(sorted(selected_books)):
            book_id = self._selected_book_id(row)
            cid = self._selected_cid(row)
            mismatches = self.installed_books[book_id].metadata_mismatches
            #self._busy_status_msg(msg="Updating '{0}'".format(self.installed_books[book_id].title))
            if total_books > 1:
                msg = "Updating metadata: {0} of {1}".format(i+1, total_books)
            else:
                msg = "Updating metadata"
            self._busy_status_msg(msg=msg)
            if action == 'export_metadata':
                # Apply calibre metadata to Marvin
                error = self._update_marvin_metadata(book_id, cid, mismatches, row)
                if error:
                    errors.append(error)

            elif action == 'import_metadata':
                # Apply Marvin metadata to calibre
                error = self._update_calibre_metadata(book_id, cid, mismatches, row)
                if error:
                    errors.append(error)

            # Clear the metadata_mismatch
            self.installed_books[book_id].metadata_mismatches = {}

            if self.busy_cancel_requested:
                break

        self._busy_status_teardown()

        # Launch row flasher
        self._flash_affected_rows()

        if errors:
            # Construct a compilation of the error details
            details = ''
            for error in errors:
                details += "{0}\n".format(error['details'])
            self._show_command_error('_update_marvin_metadata', {'details': details})

    def _update_reading_progress(self, book, row):
        '''
        Refresh Progress column
        '''
        self._log_location()
        progress = self._generate_reading_progress(book)
        self.tm.set_progress(row, progress)

    def _update_refresh_button(self):
        '''
        '''
        self._log_location()

        # Get a list of the active mapped custom columns
        enabled = []
        for cfn, col in [
                         ('annotations', self.ANNOTATIONS_COL),
                         ('date_read', self.LAST_OPENED_COL),
                         ('progress', self.PROGRESS_COL),
                         ('read', self.FLAGS_COL),
                         ('reading_list', self.FLAGS_COL),
                         ('word_count', self.WORD_COUNT_COL)
                        ]:
            cfv = get_cc_mapping(cfn, 'combobox', None)
            visible = not self.tv.isColumnHidden(col) and self.tv.columnWidth(col) > 0
            if cfv and visible:
                enabled.append(cfv)

        if False and enabled:
            button_title = 'Refresh %s' % ', '.join(sorted(enabled, key=sort_key))
            if len(button_title) > 40:
                button_title = button_title[0:39] + u"\u2026"
        else:
            button_title = self.DEFAULT_REFRESH_TEXT

        self.refresh_button.setText(button_title)
        self.refresh_button.setEnabled(bool(enabled))

    def _update_remote_hash_cache(self):
        '''
        Copy updated hash cache to iDevice
        self.local_hash_cache, self.remote_hash_cache initialized
        in _localize_hash_cache()
        '''
        self._log_location()

        if self.ios.exists(str(self.remote_hash_cache)):
            self.ios.remove(str(self.remote_hash_cache))

        if self.parent.prefs.get('hash_caching_disabled', False):
            self._log("hash_caching_disabled, deleting remote hash cache")
        else:
            # Copy local cache to iDevice
            self.ios.copy_to_idevice(self.local_hash_cache, str(self.remote_hash_cache))

    def _wait_for_command_completion(self, command_name, update_local_db=True,
            get_response=None, timeout_override=None):
        '''
        Wait for Marvin to issue progress reports via status.xml
        Marvin creates status.xml upon receiving command, increments <progress>
        from 0.0 to 1.0 as command progresses.
        '''
        import traceback

        # POLLING_DELAY affects the frequency with which the spinner is updated
        POLLING_DELAY = 0.25
        msg = ''
        if timeout_override:
            msg = "using timeout_override %d" % timeout_override
        self._log_location(msg)

        results = {'code': 0}

        if self.prefs.get('execute_marvin_commands', True):
            self._log("%s: waiting for '%s'" %
                      (datetime.now().strftime('%H:%M:%S.%f'),
                      self.parent.connected_device.status_fs))

            if not timeout_override:
                timeout_value = self.WATCHDOG_TIMEOUT
            else:
                timeout_value = timeout_override

            # Set initial watchdog timer for ACK with default timeout
            self.operation_timed_out = False
            self.watchdog = Timer(self.WATCHDOG_TIMEOUT, self._watchdog_timed_out)
            self.watchdog.start()

            while True:
                if not self.ios.exists(self.parent.connected_device.status_fs):
                    # status.xml not created yet
                    if self.operation_timed_out:
                        final_code = '-1'
                        self.ios.remove(self.parent.connected_device.status_fs)
                        results = {
                            'code': -1,
                            'status': 'timeout',
                            'response': None,
                            'details': 'timeout_value: %d' % timeout_value
                            }
                        break
                    Application.processEvents()
                    time.sleep(POLLING_DELAY)

                else:
                    # Start a new watchdog timer per iteration
                    self.watchdog.cancel()
                    self.watchdog = Timer(timeout_value, self._watchdog_timed_out)
                    self.operation_timed_out = False
                    self.watchdog.start()

                    self._log("%s: monitoring progress of %s" %
                              (datetime.now().strftime('%H:%M:%S.%f'),
                              command_name))

                    code = '-1'
                    current_timestamp = 0.0
                    while code == '-1':
                        try:
                            if self.operation_timed_out:
                                self.ios.remove(self.parent.connected_device.status_fs)
                                results = {
                                    'code': -1,
                                    'status': 'timeout',
                                    'response': None,
                                    'details': 'timeout_value: %d' % timeout_value
                                    }
                                break

                            # Cancel requested?
                            if self.busy_cancel_requested and self.marvin_cancellation_required:
                                self._log("user requesting cancellation")

                                # Create "cancel.command" in staging folder
                                ft = (b'/'.join([self.parent.connected_device.staging_folder,
                                                 b'cancel.tmp']))
                                fs = (b'/'.join([self.parent.connected_device.staging_folder,
                                                 b'cancel.command']))
                                self.ios.write("please stop", ft)
                                self.ios.rename(ft, fs)

                                # Update status
                                self._busy_status_msg(msg="Completing operation on current book…")

                                # Clear flags so we can complete processing
                                self.marvin_cancellation_required = False

                            status = etree.fromstring(self.ios.read(self.parent.connected_device.status_fs))
                            code = status.get('code')
                            timestamp = float(status.get('timestamp'))
                            if timestamp != current_timestamp:
                                current_timestamp = timestamp
                                d = datetime.now()
                                progress = float(status.find('progress').text)
                                self._log("{0}: {1:>2} {2:>3}%".format(
                                          d.strftime('%H:%M:%S.%f'),
                                          code,
                                          "%3.0f" % (progress * 100)))
                                """
                                # Report progress
                                if self.report_progress is not None:
                                    self.report_progress(0.5 + progress/2, '')
                                """

                                # Reset watchdog timer
                                self.watchdog.cancel()
                                self.watchdog = Timer(timeout_value, self._watchdog_timed_out)
                                self.watchdog.start()

                            Application.processEvents()
                            time.sleep(POLLING_DELAY)

                        except:
                            self.watchdog.cancel()

                            formatted_lines = traceback.format_exc().splitlines()
                            current_error = formatted_lines[-1]

                            time.sleep(POLLING_DELAY)
                            Application.processEvents()

                            self._log("{0}:  retry ({1})".format(
                                datetime.now().strftime('%H:%M:%S.%f'),
                                current_error))

                            self.watchdog = Timer(timeout_value, self._watchdog_timed_out)
                            self.watchdog.start()

                    # Command completed
                    self.watchdog.cancel()

                    # Construct the results
                    final_code = status.get('code')
                    if final_code == '-1':
                        final_status = "incomplete"
                    elif final_code == '0':
                        final_status = "completed successfully"
                    elif final_code == '1':
                        final_status = "completed with warnings"
                    elif final_code == '2':
                        final_status = "completed with errors"
                    elif final_code == '3':
                        final_status = "cancelled by user"
                    results = {'code': int(final_code), 'status': final_status}

                    '''
                    if True and command_name == 'update_metadata':
                        # *** Fake some errors to test ***
                        self._log("***falsifying error reporting***")
                        results = {'code': 2, 'status': 'completed with errors',
                            'details': "[Title - Author.epub] Cannot locate book to update metadata - skipping"}
                    '''

                    if final_code not in ['0']:
                        if final_code == '3':
                            msgs = ['operation cancelled by user']
                        else:
                            messages = status.find('messages')
                            msgs = [msg.text for msg in messages]
                        details = '\n'.join(["code: %s" % final_code, "status: %s" % final_status])
                        details += '\n'.join(msgs)
                        self._log(details)
                        results['details'] = '\n'.join(msgs)
                        self.ios.remove(self.parent.connected_device.status_fs)

                        self._log("%s: '%s' complete with errors" %
                                  (datetime.now().strftime('%H:%M:%S.%f'),
                                  command_name))

                    # Get the response file from the staging folder
                    if get_response:
                        rf = b'/'.join([self.parent.connected_device.staging_folder, get_response])
                        self._log("fetching response '%s'" % rf)
                        if not self.ios.exists(self.parent.connected_device.status_fs):
                            response = "%s not found" % rf
                        else:
                            response = self.ios.read(rf)
                            self.ios.remove(rf)
                        results['response'] = response

                    self.ios.remove(self.parent.connected_device.status_fs)

                    self._log("%s: '%s' complete" %
                              (datetime.now().strftime('%H:%M:%S.%f'),
                              command_name))
                    break

            # Update local copy of Marvin db
            if update_local_db and final_code == '0':
                self._localize_marvin_database()

        else:
            self._log("~~~ execute_marvin_commands disabled in JSON ~~~")
        return results

    def _watchdog_timed_out(self):
        '''
        Set flag if I/O operation times out
        '''
        self._log_location(datetime.now().strftime('%H:%M:%S.%f'))
        self.operation_timed_out = True
