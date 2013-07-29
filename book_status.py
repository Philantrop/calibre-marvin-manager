#!/usr/bin/env python
# coding: utf-8
from __future__ import (unicode_literals, division, absolute_import,
                        print_function)

__license__ = 'GPL v3'
__copyright__ = '2013, Greg Riker <griker@hotmail.com>'
__docformat__ = 'restructuredtext en'

import base64, hashlib, importlib, inspect, locale, operator, os, re, sqlite3, sys, time
from datetime import datetime
from functools import partial
from lxml import etree
from threading import Timer
from xml.sax.saxutils import escape

from PyQt4 import QtCore
from PyQt4.Qt import (Qt, QAbstractTableModel,
                      QApplication, QBrush,
                      QColor, QCursor, QDialogButtonBox, QFont, QIcon,
                      QItemSelectionModel, QLabel, QMenu, QModelIndex, QPainter, QPixmap,
                      QProgressDialog, QTableView, QTableWidgetItem, QTimer,
                      QVariant, QVBoxLayout, QWidget,
                      SIGNAL, pyqtSignal)

from calibre import strftime
from calibre.constants import islinux, isosx, iswindows
from calibre.devices.errors import UserFeedback
from calibre.devices.usbms.driver import debug_print
from calibre.ebooks.BeautifulSoup import BeautifulStoneSoup, Tag, UnicodeDammit
from calibre.ebooks.oeb.iterator import EbookIterator
from calibre.gui2 import Application, Dispatcher, error_dialog, warning_dialog
from calibre.gui2.dialogs.message_box import MessageBox
from calibre.gui2.dialogs.progress import ProgressDialog
from calibre.utils.config import config_dir, JSONConfig
from calibre.utils.icu import sort_key
from calibre.utils.magick.draw import thumbnail
from calibre.utils.wordcount import get_wordcount_obj
from calibre.utils.zipfile import ZipFile

from calibre_plugins.marvin_manager.common_utils import (
    AbortRequestException, Book, InventoryCollections,
    MyBlockingBusy, ProgressBar, RowFlasher, SizePersistedDialog,
    updateCalibreGUIView)

dialog_resources_path = os.path.join(config_dir, 'plugins', 'Marvin_Mangler_resources', 'dialogs')


class MyTableView(QTableView):
    def __init__(self, parent):
        super(MyTableView, self).__init__(parent)
        self.parent = parent

    def contextMenuEvent(self, event):

        index = self.indexAt(event.pos())
        col = index.column()
        row = index.row()
        selected_books = self.parent._selected_books()
        menu = QMenu(self)

        if col == self.parent.ANNOTATIONS_COL:
            calibre_cids = False
            for row in selected_books:
                if selected_books[row]['cid'] is not None:
                    calibre_cids = True
                    break

            afn = self.parent.prefs.get('annotations_field_comboBox', None)
            no_annotations = not selected_books[row]['has_annotations']

            ac = menu.addAction("View Highlights")
            ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'annotations.png')))
            ac.triggered.connect(partial(self.parent.dispatch_context_menu_event, "show_highlights", row))
            if len(selected_books) > 1 or no_annotations:
                ac.setEnabled(False)

            # Fetch Highlights if custom field specified
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
                ac = menu.addAction("Add Highlights to '{0}' column".format(afn))
                ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'annotations.png')))
                ac.triggered.connect(partial(self.parent.dispatch_context_menu_event, "fetch_annotations", row))
            else:
                ac = menu.addAction("No custom field specified for Highlights")
                ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'annotations.png')))
            ac.setEnabled(enabled)

        elif col == self.parent.ARTICLES_COL:
            no_articles = not selected_books[row]['has_articles']
            ac = menu.addAction("View articles")
            ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'articles.png')))
            ac.triggered.connect(partial(self.parent.dispatch_context_menu_event, "show_articles", row))
            if len(selected_books) > 1 or no_articles:
                ac.setEnabled(False)

        elif col == self.parent.COLLECTIONS_COL:
            cfl = self.parent.prefs.get('collection_field_lookup', '')
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
            no_dv_content = not selected_books[row]['has_dv_content']

            ac = menu.addAction("Generate Deep View content")
            ac.setIcon(QIcon(I('exec.png')))
            ac.triggered.connect(partial(self.parent.dispatch_context_menu_event, "generate_deep_view", row))
            ac.setEnabled(no_dv_content)

            menu.addSeparator()

            ac = menu.addAction("Deep View articles")
            ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'deep_view.png')))
            ac.triggered.connect(partial(self.parent.dispatch_context_menu_event,
                                         "show_deep_view_articles", row))
            if len(selected_books) > 1 or no_dv_content:
                ac.setEnabled(False)

            ac = menu.addAction("Deep View names sorted alphabetically")
            ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'deep_view.png')))
            ac.triggered.connect(partial(self.parent.dispatch_context_menu_event,
                                         "show_deep_view_alphabetically", row))
            if len(selected_books) > 1 or no_dv_content:
                ac.setEnabled(False)

            ac = menu.addAction("Deep View names sorted by importance")
            ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'deep_view.png')))
            ac.triggered.connect(partial(self.parent.dispatch_context_menu_event,
                                         "show_deep_view_by_importance", row))
            if len(selected_books) > 1 or no_dv_content:
                ac.setEnabled(False)

            ac = menu.addAction("Deep View names sorted by order of appearance")
            ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'deep_view.png')))
            ac.triggered.connect(partial(self.parent.dispatch_context_menu_event,
                                         "show_deep_view_by_appearance", row))
            if len(selected_books) > 1 or no_dv_content:
                ac.setEnabled(False)

            ac = menu.addAction("Deep View names with notes and flags first")
            ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'deep_view.png')))
            ac.triggered.connect(partial(self.parent.dispatch_context_menu_event,
                                         "show_deep_view_by_annotations", row))
            if len(selected_books) > 1 or no_dv_content:
                ac.setEnabled(False)

        elif col == self.parent.FLAGS_COL:
            ac = menu.addAction("Clear All")
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

        elif col == self.parent.LAST_OPENED_COL:
            date_read_field = self.parent.prefs.get('date_read_field_comboBox', None)

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

            title = "No 'Date read' custom field selected"
            if date_read_field:
                title = "Apply to '%s' column" % date_read_field
            ac = menu.addAction(title)
            ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'from_marvin.png')))
            ac.triggered.connect(partial(self.parent.dispatch_context_menu_event, "apply_date_read", row))

            if (not date_read_field) or (not calibre_cids) or (not last_opened):
                ac.setEnabled(False)

        elif col == self.parent.PROGRESS_COL:
            progress_field = self.parent.prefs.get('progress_field_comboBox', None)

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

            title = "No 'Progress' custom field selected"
            if progress_field:
                title = "Apply to '%s' column" % progress_field
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
            if len(selected_books) == 1 and self.parent.tm.get_match_quality(row) < self.parent.YELLOW:
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

            ac = menu.addAction("Add to library")
            ac.setIcon(QIcon(I('plus.png')))
            ac.triggered.connect(self.parent._add_books_to_library)
            ac.setEnabled(not in_library)

            menu.addSeparator()
            ac = menu.addAction("Delete")
            ac.setIcon(QIcon(I('trash.png')))
            ac.triggered.connect(self.parent._delete_books)

        elif col == self.parent.VOCABULARY_COL:
            no_vocabulary = not selected_books[row]['has_vocabulary']

            ac = menu.addAction("View vocabulary for this book")
            ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'vocabulary.png')))
            ac.triggered.connect(partial(self.parent.dispatch_context_menu_event, "show_vocabulary", row))
            if len(selected_books) > 1 or no_vocabulary:
                ac.setEnabled(False)

            ac = menu.addAction("View all vocabulary words")
            ac.setIcon(QIcon(I('books_in_series.png')))
            ac.triggered.connect(partial(self.parent.dispatch_context_menu_event, "show_global_vocabulary", row))

        elif col == self.parent.WORD_COUNT_COL:
            ac = menu.addAction("Calculate word count")
            ac.setIcon(QIcon(os.path.join(self.parent.opts.resources_path, 'icons', 'word_count.png')))
            ac.triggered.connect(partial(self.parent.dispatch_context_menu_event, "calculate_word_count", row))

        menu.exec_(event.globalPos())


class _SortableImageWidgetItem(QLabel):
    def __init__(self, parent, path, sort_key, column):
        super(SortableImageWidgetItem, self).__init__(parent=parent.tv)
        self.column = column
        self.parent = parent.tv
        self.picture = QPixmap(path)
        self.sort_key = sort_key
        self.width = self.picture.width()
        self.setAlignment(Qt.AlignCenter)
        self.setPixmap(self.picture)


class SortableImageWidgetItem(QWidget):
    def __init__(self, parent, path, sort_key, column):
        super(SortableImageWidgetItem, self).__init__(parent=parent.tv)
        self.column = column
        self.parent_tv = parent.tv
        self.picture = QPixmap(path)
        self.sort_key = sort_key
        self.width = self.picture.width()

    def __lt__(self, other):
        return self.sort_key < other.sort_key

    def _paintEvent(self, event):
        #print("column_width: %s" % (repr(self.parent_tv.columnWidth(self.column))))
        #print("picture_width: %s" % repr(self.width))
        #print("event: %s" % dir(event))
        #print("region: %s" % dir(event.region()))
        #print("region().boundingRect(): %s" % repr(event.region().boundingRect()))
        #print("boundingRect: %s" % dir(event.region().boundingRect()))
        #print("getCoords: %s" % repr(event.region().boundingRect().getCoords()))
        #print("getRect: %s" % repr(event.region().boundingRect().getRect()))
        #print("column_viewport_position: %d" % self.parent_tv.columnViewportPosition(self.column))
        #print("dir(self.parent_tv): %s" % dir(self.parent_tv))
        #cvp = self.parent_tv.columnViewportPosition(self.column)
        #painter = QPainter(self.parent_tv.viewport())
        painter = QPainter(self)
        #x_off = 0
        #col_width = self.parent_tv.columnWidth(self.column)
        #if col_width > self.width:
        #    x_off = int((col_width - self.width) / 2)
        #print("x_off: %d" % x_off)
        #painter.drawPixmap(x_off, 0, self.picture)
        painter.drawPixmap(event.region().boundingRect(), self.picture)
        painter.end()
        #QWidget.paintEvent(self, event)


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
    RED_HUE = 0.0
    ORANGE_HUE = 0.08325
    YELLOW_HUE = 0.1665
    GREEN_HUE = 0.333
    WHITE_HUE = 1.0

    # Match quality colors
    if True:
        GREEN = 4
        YELLOW = 3
        ORANGE = 2
        RED = 1
        WHITE = 0

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

        elif role == Qt.BackgroundRole and self.show_match_colors:
            match_quality = self.get_match_quality(row)
            if match_quality == 4:
                return QVariant(QBrush(QColor.fromHsvF(self.GREEN_HUE, self.SATURATION, self.HSVALUE)))
            elif match_quality == 3:
                return QVariant(QBrush(QColor.fromHsvF(self.YELLOW_HUE, self.SATURATION, self.HSVALUE)))
            elif match_quality == 2:
                return QVariant(QBrush(QColor.fromHsvF(self.ORANGE_HUE, self.SATURATION, self.HSVALUE)))
            elif match_quality == 1:
                return QVariant(QBrush(QColor.fromHsvF(self.RED_HUE, self.SATURATION, self.HSVALUE)))
            else:
                return QVariant(QBrush(QColor.fromHsvF(self.WHITE_HUE, 0.0, self.HSVALUE)))

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

        elif role == Qt.DisplayRole and col == self.parent.TITLE_COL:
            return self.arraydata[row][self.parent.TITLE_COL].text()
        elif role == Qt.DisplayRole and col == self.parent.AUTHOR_COL:
            return self.arraydata[row][self.parent.AUTHOR_COL].text()
        elif role == Qt.DisplayRole and col == self.parent.LAST_OPENED_COL:
            return self.arraydata[row][self.parent.LAST_OPENED_COL].text()
        elif role == Qt.TextAlignmentRole and (col in self.centered_columns):
            return Qt.AlignHCenter
        elif role == Qt.TextAlignmentRole and (col in self.right_aligned_columns):
            return Qt.AlignRight

        elif role == Qt.ToolTipRole:
            match_quality = self.get_match_quality(row)
            tip = '<p>'
            if match_quality == self.GREEN:
                tip += 'Matched in calibre library'
            elif match_quality == self.YELLOW:
                tip += 'Matched in calibre library with differing metadata'
            elif match_quality == self.ORANGE:
                tip += 'Duplicate of matched book in calibre library'
            elif match_quality == self.RED:
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

            elif col in [self.parent.DEEP_VIEW_COL, self.parent.VOCABULARY_COL]:
                has_content = bool(self.arraydata[row][col])
                if has_content:
                    return tip + "<br/>Double-click to view details<br/>Right-click for more options</p>"
                else:
                    return tip + '<br/>Right-click for more options</p>'

            elif col in [self.parent.FLAGS_COL]:
                return tip + "<br/>Right-click for options</p>"

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
#         if role == Qt.ToolTipRole:
#             if orientation == Qt.Horizontal:
#                 return QString("Tooltip for col %d" % col)

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

    def get_match_quality(self, row):
        return self.arraydata[row][self.parent.MATCHED_COL]

    def set_match_quality(self, row, value):
        self.arraydata[row][self.parent.MATCHED_COL] = value
        #self.parent.repaint()

    def get_path(self, row):
        return self.arraydata[row][self.parent.PATH_COL]

    def get_progress(self, row):
        return self.arraydata[row][self.parent.PROGRESS_COL]

    def set_progress(self, row, value):
        self.arraydata[row][self.parent.PROGRESS_COL] = value
        #self.parent.repaint()

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
        #self.parent.repaint()


class BookStatusDialog(SizePersistedDialog):
    '''
    '''
    CHECKMARK = u"\u2713"

    # Location reporting template
    LOCATION_TEMPLATE = "{cls}:{func}({arg1}) {arg2}"

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

    # Column assignments
    if True:
        LIBRARY_HEADER = ['uuid', 'cid', 'mid', 'path',
                          'Title', 'Author', 'Word Count', 'Progress', 'Last read',
                          'Collections', 'Flags',
                          'Annotations', 'Articles', 'Deep View', 'Vocabulary',
                          'Match Quality']
        ANNOTATIONS_COL = LIBRARY_HEADER.index('Annotations')
        ARTICLES_COL = LIBRARY_HEADER.index('Articles')
        AUTHOR_COL = LIBRARY_HEADER.index('Author')
        BOOK_ID_COL = LIBRARY_HEADER.index('mid')
        CALIBRE_ID_COL = LIBRARY_HEADER.index('cid')
        COLLECTIONS_COL = LIBRARY_HEADER.index('Collections')
        DEEP_VIEW_COL = LIBRARY_HEADER.index('Deep View')
        FLAGS_COL = LIBRARY_HEADER.index('Flags')
        LAST_OPENED_COL = LIBRARY_HEADER.index('Last read')
        MATCHED_COL = LIBRARY_HEADER.index('Match Quality')
        PATH_COL = LIBRARY_HEADER.index('path')
        PROGRESS_COL = LIBRARY_HEADER.index('Progress')
        TITLE_COL = LIBRARY_HEADER.index('Title')
        UUID_COL = LIBRARY_HEADER.index('uuid')
        VOCABULARY_COL = LIBRARY_HEADER.index('Vocabulary')
        WORD_COUNT_COL = LIBRARY_HEADER.index('Word Count')

        HIDDEN_COLUMNS = [
            UUID_COL,
            CALIBRE_ID_COL,
            BOOK_ID_COL,
            PATH_COL,
            MATCHED_COL,
        ]
        CENTERED_COLUMNS = [
            ANNOTATIONS_COL,
            COLLECTIONS_COL,
            DEEP_VIEW_COL,
            ARTICLES_COL,
            LAST_OPENED_COL,
            VOCABULARY_COL,
        ]
        RIGHT_ALIGNED_COLUMNS = [
            PROGRESS_COL,
            WORD_COUNT_COL
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

    # Match quality color constants
    if True:
        GREEN = 4
        YELLOW = 3
        ORANGE = 2
        RED = 1
        WHITE = 0

    marvin_device_status_changed = pyqtSignal(str)

    def accept(self):
        self._log_location()
        self._save_column_widths()
        super(BookStatusDialog, self).accept()

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
                self.show_assets_dialog('show_global_vocabulary', 0)

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
            self._fetch_annotations()
        elif action == 'generate_deep_view':
            self._generate_deep_view()
        elif action == 'manage_collections':
            self.show_manage_collections_dialog()
        elif action in ['show_articles', 'show_deep_view_articles',
                        'show_deep_view_alphabetically', 'show_deep_view_by_importance',
                        'show_deep_view_by_appearance', 'show_deep_view_by_annotations',
                        'show_highlights', 'show_vocabulary']:
            self.show_assets_dialog(action, row)
        elif action == 'show_collections':
            self.show_view_collections_dialog(row)
        elif action == 'show_global_vocabulary':
            self.show_assets_dialog('show_global_vocabulary', row)
        elif action == 'show_metadata':
            self.show_view_metadata_dialog(row)
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
            self.ANNOTATIONS_COL: 'show_highlights',
            self.ARTICLES_COL: 'show_deep_view_articles',
            self.DEEP_VIEW_COL: 'show_deep_view_articles',
            self.VOCABULARY_COL: 'show_vocabulary'
        }

        column = index.column()
        row = index.row()

        if column in [self.TITLE_COL, self.AUTHOR_COL]:
            self.show_view_metadata_dialog(row)
        elif column in [self.ANNOTATIONS_COL, self.DEEP_VIEW_COL,
                        self.ARTICLES_COL, self.VOCABULARY_COL]:
            self.show_assets_dialog(asset_actions[column], row)
        elif column == self.COLLECTIONS_COL:
            self.show_view_collections_dialog(row)
        elif column in [self.FLAGS_COL]:
            title = "Flag options"
            msg = "<p>Right-click in the Flags column for flag management options.</p>"
            MessageBox(MessageBox.INFO, title, msg,
                       show_copy_button=False).exec_()
        elif column == self.WORD_COUNT_COL:
            self._calculate_word_count()
        else:
            self._log("no double-click handler for %s" % self.LIBRARY_HEADER[column])

    def esc(self, *args):
        '''
        Clear any active selections
        '''
        self._log_location()
        self._clear_selected_rows()

    def initialize(self, parent):
        self.archived_cover_hashes = JSONConfig('plugins/Marvin_Mangler_resources/cover_hashes')
        self.busy_window = None
        self.Dispatcher = partial(Dispatcher, parent=self)
        self.hash_cache = 'content_hashes.zip'
        self.ios = parent.ios
        self.opts = parent.opts
        self.parent = parent
        self.prefs = parent.opts.prefs
        self.library_scanner = parent.library_scanner
        self.library_title_map = None
        self.library_uuid_map = None
        self.local_cache_folder = self.parent.connected_device.temp_dir
        self.local_hash_cache = None
        self.remote_cache_folder = '/'.join(['/Library', 'calibre.mm'])
        self.remote_hash_cache = None
        self.show_match_colors = self.prefs.get('show_match_colors', False)
        self.updated_match_quality = None
        self.verbose = parent.verbose

        # Subscribe to Marvin driver change events
        self.parent.connected_device.marvin_device_signals.reader_app_status_changed.connect(
            self.marvin_status_changed)

        self._log_location()

        self.installed_books = self._generate_booklist()

        # ~~~~~~~~ Create the dialog ~~~~~~~~
        self.setWindowTitle(u'Marvin Library: %d books' % len(self.installed_books))
        self.setWindowIcon(self.opts.icon)
        self.l = QVBoxLayout(self)
        self.setLayout(self.l)
        self.perfect_width = 0

        # ~~~~~~~~ Create the Table ~~~~~~~~
        self.tv = MyTableView(self)
        self.tabledata = self._construct_table_data()
        self._construct_table_view()

        # ~~~~~~~~ Create the ButtonBox ~~~~~~~~
        self.dialogButtonBox = QDialogButtonBox(QDialogButtonBox.Help)

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

        # Word count
        if False:
            self.wc_button = self.dialogButtonBox.addButton('Calculate word count', QDialogButtonBox.ActionRole)
            self.wc_button.setObjectName('calculate_word_count_button')
            self.wc_button.setIcon(QIcon(os.path.join(self.parent.opts.resources_path,
                                                      'icons',
                                                      'word_count.png')))

        # Generate DV content
        if False:
            self.gdv_button = self.dialogButtonBox.addButton('Generate Deep View', QDialogButtonBox.ActionRole)
            self.gdv_button.setObjectName('generate_deep_view_button')
            self.gdv_button.setIcon(QIcon(os.path.join(self.parent.opts.resources_path,
                                                       'icons',
                                                       'deep_view.png')))

        # View metadata
        if False:
            self.vm_button = self.dialogButtonBox.addButton('View metadata', QDialogButtonBox.ActionRole)
            self.vm_button.setObjectName('view_metadata_button')
            self.vm_button.setIcon(QIcon(os.path.join(self.parent.opts.resources_path,
                                                      'icons',
                                                      'update_metadata.png')))

        # View collections
        if False:
            self.vc_button = self.dialogButtonBox.addButton('View collection assignments', QDialogButtonBox.ActionRole)
            self.vc_button.setObjectName('view_collections_button')
            self.vc_button.setIcon(QIcon(os.path.join(self.parent.opts.resources_path,
                                                      'icons',
                                                      'update_metadata.png')))

        # Manage collections
        if True:
            self.mc_button = self.dialogButtonBox.addButton('Manage collections', QDialogButtonBox.ActionRole)
            self.mc_button.setObjectName('manage_collections_button')
            self.mc_button.setIcon(QIcon(os.path.join(self.parent.opts.resources_path,
                                                      'icons',
                                                      'edit_collections.png')))

        # Refresh custom columns
        if True:
            # Get a list of the active custom columns
            enabled = []
            for cfn in ['annotations_field_comboBox', 'date_read_field_comboBox',
                        'progress_field_comboBox']:
                cfv = self.parent.prefs.get(cfn, None)
                if cfv:
                    enabled.append(cfv)
            if enabled:
                button_title = 'Refresh %s' % ', '.join(sorted(enabled, key=sort_key))

                self.refresh_button = self.dialogButtonBox.addButton(button_title, QDialogButtonBox.ActionRole)
                self.refresh_button.setObjectName('refresh_custom_columns_button')
                self.refresh_button.setIcon(QIcon(os.path.join(self.parent.opts.resources_path,
                                                               'icons',
                                                               'from_marvin.png')))
                tooltip = "Refresh custom columns %s in calibre" % ', '.join(sorted(enabled, key=sort_key))
                self.refresh_button.setToolTip(tooltip)

        # View Global vocabulary
        if True:
            self.gv_button = self.dialogButtonBox.addButton('View all vocabulary words', QDialogButtonBox.ActionRole)
            self.gv_button.setObjectName('view_global_vocabulary_button')
            self.gv_button.setIcon(QIcon(I('books_in_series.png')))

        self.dialogButtonBox.clicked.connect(self.dispatch_button_click)

        self.l.addWidget(self.dialogButtonBox)

        # ~~~~~~~~ Connect signals ~~~~~~~~
        self.connect(self.tv, SIGNAL("doubleClicked(QModelIndex)"), self.dispatch_double_click)
        self.connect(self.tv.horizontalHeader(), SIGNAL("sectionClicked(int)"), self.capture_sort_column)

        self.resize_dialog()

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

    def marvin_status_changed(self, command):
        '''

        '''
        self.marvin_device_status_changed.emit(command)

        self._log_location(command)

        if command in ['disconnected', 'yanked']:
            self._log("closing dialog: %s" % command)
            self.close()

    def refresh_custom_columns(self):
        '''
        Refresh enabled custom columns from Marvin content
        '''
        self._log_location()

        enabled = []
        for cfn in ['annotations_field_comboBox', 'date_read_field_comboBox',
                    'progress_field_comboBox']:
            cfv = self.parent.prefs.get(cfn, None)
            if cfv:
                enabled.append(cfv)
        cols_to_refresh = ', '.join(sorted(enabled, key=sort_key))

        pb = ProgressBar(parent=self.opts.gui, window_title="Refreshing {0}".format(cols_to_refresh),
                         on_top=True)
        pb.set_label('{:^100}'.format(" label goes here "))
        pb.set_value(0)
        pb.show()

        if len(self._selected_rows()) < 2:
            total_books = len(self.tm.all_rows())
            pb.set_maximum(total_books)
            pb.set_value(0)
            pb.show()

            for i, row in enumerate(self.tm.all_rows()):
                self.tv.selectRow(row)
                pb.set_label('{:^100}'.format(self.tm.all_rows[row]['title']))
                self._fetch_annotations()
                self._apply_date_read()
                self._apply_progress()
                pb.increment()
        else:
            rows_to_refresh = sorted(self._selected_books())
            total_books = len(rows_to_refresh)
            pb.set_maximum(total_books)
            pb.set_value(0)
            pb.show()

            for row in rows_to_refresh:
                self.tv.selectRow(row)
                pb.set_label('{:^100}'.format(self._selected_books()[row]['title']))
                self._fetch_annotations()
                self._apply_date_read()
                self._apply_progress()
                pb.increment()

        pb.hide()
        updateCalibreGUIView()

    def show_assets_dialog(self, action, row):
        '''
        Display assets associated with book
        Articles, Annotations, Deep View, Vocabulary
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
            command_name = "command"
            command_type = "GetDeepViewNamesHTML"
            update_soup = BeautifulStoneSoup(self.GENERAL_COMMAND_XML.format(
                command_type, time.mktime(time.localtime())))
            parameters_tag = self._build_parameters(self.installed_books[book_id], update_soup)

            # <parameter> for order
            parameter_tag = Tag(update_soup, 'parameter')
            parameter_tag['name'] = "order"

            header = None
            if action == 'show_deep_view_alphabetically':
                group_box_title = "Deep View names alphabetically"
                parameter_tag.insert(0, "alphabetically")

            elif action == 'show_deep_view_by_importance':
                group_box_title = "Deep View names by importance"
                parameter_tag.insert(0, "importance")

            elif action == 'show_deep_view_by_appearance':
                group_box_title = "Deep View names by appearance"
                parameter_tag.insert(0, "appearance")

            elif action == 'show_deep_view_by_annotations':
                group_box_title = "Deep View names with notes and flags first"
                parameter_tag.insert(0, "annotated")

            parameters_tag.insert(0, parameter_tag)
            update_soup.command.insert(0, parameters_tag)

            default_content = "{0} to be provided by Marvin.".format(group_box_title)
            footer = None

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

        elif action == 'show_highlights':
            command_name = "command"
            command_type = "GetAnnotationsHTML"
            update_soup = BeautifulStoneSoup(self.GENERAL_COMMAND_XML.format(
                command_type, time.mktime(time.localtime())))
            parameters_tag = self._build_parameters(self.installed_books[book_id], update_soup)
            update_soup.command.insert(0, parameters_tag)

            header = None
            group_box_title = 'Highlights and Annotations'
            if self.installed_books[book_id].highlights:
                default_content = '\n'.join(self.installed_books[book_id].highlights)
            else:
                default_content = "<p>No highlights</p>"
            footer = (
                '<p>The <a href="http://www.mobileread.com/forums/showthread.php?t=205062" target="_blank">' +
                'Annotations plugin</a> imports highlights and annotations from Marvin.</p>')
            afn = self.parent.prefs.get('annotations_field_comboBox', None)
            if afn:
                refresh = {
                    'name': afn,
                    'method': "_fetch_annotations"
                    }

        elif action == 'show_vocabulary':
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
            header = None
            group_box_title = action
            default_content = "Default content"
            footer = None

        QApplication.setOverrideCursor(QCursor(Qt.WaitCursor))
        self.busy_window = MyBlockingBusy(self, "Retrieving %s…" % group_box_title, size=60)
        self.busy_window.start()
        self.busy_window.show()

        response = self._issue_command(command_name, update_soup,
                                       get_response="html_response.html",
                                       update_local_db=False)

        response = self._issue_command(command_name, update_soup,
                                       get_response="html_response.html",
                                       update_local_db=False)
        """
        QApplication.setOverrideCursor(QCursor(Qt.WaitCursor))

        # Copy command file to staging folder
        self._stage_command_file(command_name, update_soup,
                                 show_command=self.prefs.get('show_staged_commands', False))

        # Wait for completion
        content = self._wait_for_command_completion(command_name,
                                                    update_local_db=False,
                                                    get_response="html_response.html")
        QApplication.restoreOverrideCursor()
        """

        if response:
            # Convert to unicode
            response = UnicodeDammit(response).unicode

            # Strip the UTF-8 BOM
            BOM = '\xef\xbb\xbf'
            response = re.sub(BOM, '', response)
        else:
            response = default_content

        content_dict = {
            'footer': footer,
            'group_box_title': group_box_title,
            'header': header,
            'html_content': response,
            'title': title,
            'refresh': refresh
            }

        self.busy_window.stop()
        self.busy_window.accept()
        self.busy_window = None
        QApplication.restoreOverrideCursor()

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
        cfl = self.parent.prefs.get('collection_field_lookup', '')
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

                        self._update_global_collections(details)
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
            enable_metadata_updates = self.tm.get_match_quality(row) >= self.YELLOW

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
            self.show_match_colors_button.setText("Hide Matches")
            self.tv.sortByColumn(self.LIBRARY_HEADER.index('Match Quality'), Qt.DescendingOrder)
            self.capture_sort_column(self.LIBRARY_HEADER.index('Match Quality'))
            self.show_match_colors_button.setIcon(QIcon(os.path.join(self.parent.opts.resources_path,
                                                                     'icons',
                                                                     'matches_hide.png')))
        else:
            self.show_match_colors_button.setText("Show Matches")
            self.show_match_colors_button.setIcon(QIcon(os.path.join(self.parent.opts.resources_path,
                                                                     'icons',
                                                                     'matches_show.png')))
        self.tv.setAlternatingRowColors(not self.show_match_colors)
        self.tm.refresh(self.show_match_colors)

    # Helpers
    def _apply_date_read(self):
        '''
        Fetch the LAST_OPENED date, convert to datetime, apply to custom field
        '''
        self._log_location()
        lookup = self.parent.prefs.get('date_read_field_lookup', None)
        if lookup:
            selected_books = self._selected_books()
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
                        um = mi.metadata_for_field(lookup)
                        ndo = datetime.strptime(new_date, "%Y-%m-%d")
                        um['#value#'] = ndo.replace(hour=12)
                        mi.set_user_metadata(lookup, um)
                        db.set_metadata(cid, mi, set_title=False, set_authors=False,
                                        commit=True)
                    else:
                        self._log("'%s' has no Last read date" % selected_books[row]['title'])
            #updateCalibreGUIView()
        else:
            self._log("No date_read_field_lookup specified")

    def _apply_progress(self):
        '''
        Fetch Progress, apply to custom field
        Need to assert force_changes for db to allow custom field to be set to None.
        '''
        self._log_location()
        lookup = self.parent.prefs.get('progress_field_lookup', None)
        if lookup:
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
            updateCalibreGUIView()
        else:
            self._log("No progress_field_lookup specified")

    def _build_metadata_update(self, book_id, cid, book, mismatches):
        '''
        Build a metadata update command file for Marvin
        '''
        self._log_location()

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
        book_tag['pubdate'] = strftime('%Y-%m-%d', t=naive)
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
                self._log("error calculating cover_hash for cid %d (%s)" % (cid, book.title))
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

    def _calculate_word_count(self):
        '''
        Calculate word count for each selected book
        selected_books: {row: {'book_id':, 'cid':, 'path':, 'title':}...}
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

        selected_books = self._selected_books()
        if selected_books:
            stats = {}

            pb = ProgressBar(parent=self.opts.gui, window_title="Calculating word count",
                             on_top=True)
            total_books = len(selected_books)
            pb.set_maximum(total_books)
            pb.set_value(0)
            pb.set_label('{:^100}'.format("1 of %d" % (total_books)))
            pb.show()

            close_requested = False
            for i, row in enumerate(selected_books):
                if pb.close_requested:
                    close_requested = True
                    break

                pb.set_label('{:^100}'.format(selected_books[row]['title']))

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

                self._log("%s: %d words" % (selected_books[row]['title'], wordcount.words))
                stats[selected_books[row]['title']] = wordcount.words

                # Delete the local copy
                os.remove(lbp)

                # Update the model
                wc = locale.format("%d", wordcount.words, grouping=True)
                self.tm.set_word_count(row, "{0} ".format(wc))

                # Update the spreadsheet for those watching at home
                self.repaint()

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

                self._issue_command(command_name, update_soup)
                """
                # Copy command file to staging folder
                self._stage_command_file(command_name, update_soup,
                                         show_command=self.prefs.get('show_staged_commands', False))

                # Wait for completion
                self._wait_for_command_completion(command_name, update_local_db=True)
                """

                pb.increment()

            pb.hide()

            if close_requested:
                self._log("user cancelled, partial results delivered")

            if False:
                # Display a summary
                title = "Word count results"
                msg = ("<p>Calculated word count for {0} books.</p>".format(total_books) +
                       "<p>Click <b>Show details</b> for a summary.</p>")
                dl = ["%s: %s" % (stat, locale.format("%d", stats[stat], grouping=True))
                      for stat in stats]
                details = '\n'.join(dl)
                MessageBox(MessageBox.INFO, title, msg, det_msg=details,
                           show_copy_button=False).exec_()
        else:
            self._log("No selected books")
            # Display a summary
            title = "Word count"
            msg = ("<p>Select one or more books to calculate word count.</p>")
            MessageBox(MessageBox.INFO, title, msg,
                       show_copy_button=False).exec_()

    def _clear_flags(self, action):
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

        selected_books = self._selected_books()
        for row in selected_books:
            book_id = selected_books[row]['book_id']
            flagbits = self.tm.get_flags(row).sort_key
            path = selected_books[row]['path']
            if mask == 0:
                flagbits = 0
                basename = "flags0.png"
                new_flags_widget = SortableImageWidgetItem(self,
                                                           os.path.join(self.parent.opts.resources_path,
                                                           'icons', basename),
                                                           flagbits, self.FLAGS_COL)
                # Update self.installed_books flags list
                flags = []
                self.installed_books[book_id].flags = flags

                # Update the spreadsheet
                self.tm.set_flags(row, new_flags_widget)

                # Update reading progress based upon flag values
                self._update_reading_progress(self.installed_books[book_id], row)

                # Update in-memory caches
                _update_in_memory(book_id, path)

                # Update Marvin db
                self._inform_marvin_collections(book_id)

            elif flagbits & mask:
                # Clear the bit with XOR
                flagbits = flagbits ^ mask
                basename = "flags%d.png" % flagbits
                new_flags_widget = SortableImageWidgetItem(self,
                                                           os.path.join(self.parent.opts.resources_path,
                                                           'icons', basename),
                                                           flagbits, self.FLAGS_COL)
                # Update self.installed_books flags list
                self.installed_books[book_id].flags = _build_flag_list(flagbits)

                # Update the model
                self.tm.set_flags(row, new_flags_widget)

                # Update reading progress based upon flag values
                self._update_reading_progress(self.installed_books[book_id], row)

                # Update in-memory caches
                _update_in_memory(book_id, path)

                # Update Marvin db
                self._inform_marvin_collections(book_id)

            self._update_device_flags(book_id, path, _build_flag_list(flagbits))

        self.repaint()

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
        if self.opts.prefs.get('development_mode', False):
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
                    text_hrefs.append(item.get('href').split('/')[-1])
            zf.close()
        except:
            if self.opts.prefs.get('development_mode', False):
                import traceback
                self._log(traceback.format_exc())
            return None

        m = hashlib.md5()
        zfi = ZipFile(zipfile).infolist()
        for zi in zfi:
            if False and self.opts.prefs.get('development_mode', False):
                self._log("evaluating %s" % zi.filename)

            if zi.filename.split('/')[-1] in text_hrefs:
                m.update(zi.filename)
                m.update(str(zi.file_size))
                if False and self.opts.prefs.get('development_mode', False):
                    self._log("adding filename %s" % (zi.filename))
                    self._log("adding file_size %s" % (zi.file_size))

        return m.hexdigest()

    def _construct_table_data(self):
        '''
        Populate the table data from self.installed_books
        '''
        def _generate_author(book_data):
            '''
            '''
            if not book_data.author_sort:
                book_data.author_sort = ', '.join(book_data.author)
            author = SortableTableWidgetItem(
                ', '.join(book_data.author),
                book_data.author_sort.upper())
            return author

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
            flags = SortableImageWidgetItem(self,
                                            os.path.join(self.parent.opts.resources_path,
                                                         'icons', base_name),
                                            flagbits, self.FLAGS_COL)
            return flags

        def _generate_last_opened(book_data):
            '''
            last_opened sorts by timestamp
            '''
            last_opened_ts = ''
            last_opened_sort = 0
            if book_data.date_opened:
                last_opened_ts = time.strftime("%Y-%m-%d",
                                               time.localtime(book_data.date_opened))
                last_opened_sort = book_data.date_opened
            last_opened = SortableTableWidgetItem(
                last_opened_ts,
                last_opened_sort)
            return last_opened

        def _generate_match_quality(book_data):
            '''
            4: Marvin uuid matches calibre uuid (hard match): Green
            3: Marvin hash matches calibre hash (soft match): Yellow
            2: Marvin hash duplicates: Orange
            1: Calibre hash duplicates: Red
            0: Marvin only, single copy: White
            '''

            if self.opts.prefs.get('development_mode', False):
                self._log("%s uuid: %s matches: %s on_device: %s hash: %s" %
                          (book_data.title,
                           repr(book_data.uuid),
                           repr(book_data.matches),
                           repr(book_data.on_device),
                           repr(book_data.hash)))
                self._log("metadata_mismatches: %s" % repr(book_data.metadata_mismatches))
            match_quality = self.WHITE

            if (book_data.uuid > '' and
                    [book_data.uuid] == book_data.matches and
                    not book_data.metadata_mismatches):
                # GREEN: Hard match - uuid match, metadata match
                match_quality = self.GREEN

            elif ((book_data.on_device == 'Main' and
                   book_data.metadata_mismatches) or
                  ([book_data.uuid] == book_data.matches)):
                # YELLOW: Soft match - hash match,
                match_quality = self.YELLOW

            elif (book_data.uuid in book_data.matches):
                # ORANGE: Duplicate of calibre copy
                match_quality = self.ORANGE

            elif (book_data.hash in self.marvin_hash_map and
                  len(self.marvin_hash_map[book_data.hash]) > 1):
                # RED: Marvin-only duplicate
                match_quality = self.RED

            if self.opts.prefs.get('development_mode', False):
                self._log("%s match_quality: %s" % (book_data.title, match_quality))
            return match_quality

        def _generate_title(book_data):
            '''
            '''
            # Title, Author sort by title_sort, author_sort
            if not book_data.title_sort:
                book_data.title_sort = book_data.title_sort()
            title = SortableTableWidgetItem(
                book_data.title,
                book_data.title_sort.upper())
            return title

        self._log_location()

        tabledata = []

        for book in self.installed_books:
            book_data = self.installed_books[book]
            author = _generate_author(book_data)
            collection_match = self._generate_collection_match(book_data)
            flags = _generate_flags_profile(book_data)
            last_opened = _generate_last_opened(book_data)
            match_quality = _generate_match_quality(book_data)
            progress = self._generate_reading_progress(book_data)
            title = _generate_title(book_data)

            article_count = 0
            if 'Wiki' in book_data.articles:
                article_count += len(book_data.articles['Wiki'])
            if 'Pinned' in book_data.articles:
                article_count += len(book_data.articles['Pinned'])

            # List order matches self.LIBRARY_HEADER
            this_book = [
                book_data.uuid,
                book_data.cid,
                book_data.mid,
                book_data.path,
                title,
                author,
                "{0} ".format(book_data.word_count) if book_data.word_count > '0' else '',
                progress,
                last_opened,
                collection_match,
                flags,
                len(book_data.highlights) if len(book_data.highlights) else '',
                article_count if article_count else '',
                self.CHECKMARK if book_data.deep_view_prepared else '',
                len(book_data.vocabulary) if len(book_data.vocabulary) else '',
                match_quality]
            tabledata.append(this_book)
        return tabledata

    def _construct_table_view(self):
        '''
        '''
        #self.tv = QTableView(self)
        self._log_location()
        self.l.addWidget(self.tv)
        self.tm = MarkupTableModel(self, centered_columns=self.CENTERED_COLUMNS,
                                   right_aligned_columns=self.RIGHT_ALIGNED_COLUMNS)

        self.tv.setModel(self.tm)
        self.tv.setShowGrid(False)
        if self.parent.prefs.get('use_monospace_font', False):
            # Set row height
            nrows = len(self.tabledata)
            for row in xrange(nrows):
                self.tv.setRowHeight(row, 16)

            if isosx:
                FONT = QFont('Monaco', 11)
            elif iswindows:
                FONT = QFont('Lucida Console', 9)
            elif islinux:
                FONT = QFont('Monospace', 9)
                FONT.setStyleHint(QFont.TypeWriter)
            self.tv.setFont(FONT)
        else:
            # Set row height
            nrows = len(self.tabledata)
            for row in xrange(nrows):
                self.tv.setRowHeight(row, 18)

        self.tvSelectionModel = self.tv.selectionModel()
        self.tv.setAlternatingRowColors(not self.show_match_colors)
        self.tv.setShowGrid(False)
        self.tv.setWordWrap(False)
        self.tv.setSelectionBehavior(self.tv.SelectRows)

        # Hide the vertical self.header
        self.tv.verticalHeader().setVisible(False)

        # Hide hidden columns
        for index in self.HIDDEN_COLUMNS:
            self.tv.hideColumn(index)

        # Set horizontal self.header props
        self.tv.horizontalHeader().setStretchLastSection(True)

        # Set column width to fit contents
        self.tv.resizeColumnsToContents()

        # Restore saved widths if available
        saved_column_widths = self.opts.prefs.get('marvin_library_column_widths', False)
        if saved_column_widths:
            for i, width in enumerate(saved_column_widths):
                self.tv.setColumnWidth(i, width)

        self.tv.setSortingEnabled(True)

        sort_column = self.opts.prefs.get('marvin_library_sort_column',
                                          self.LIBRARY_HEADER.index('Match Quality'))
        sort_order = self.opts.prefs.get('marvin_library_sort_order',
                                         Qt.DescendingOrder)
        self.tv.sortByColumn(sort_column, sort_order)

    def _add_books_to_library(self):
        '''
        Filter out books already in calibre
        Hook into gui.iactions['Add Books'].add_books_from_device()
        gui2.actions.add #406
        '''
        self._log_location("not fully implemented")

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
            # Find the books in the model so we can monitor the in_library field
            model = self.parent.gui.memory_view.model()
            added = {}
            for item in model.sorted_map:
                book = model.db[item]
                if book.path in paths_to_add:
                    added[item] = {'path': book.path}

            # Tell calibre to add the paths
            self.opts.gui.iactions['Add Books'].add_books_from_device(self.parent.gui.memory_view,
                                                                      paths=paths_to_add)

            # Wait for added books to be updated in model.db
            # in_library property will be set to AUTHOR (or, less likely, UUID)
            incomplete = True
            while incomplete:
                Application.processEvents()
                for item in added:
                    if model.db[item].in_library is None:
                        break
                else:
                    incomplete = False

            # Update in-memory with newly minted cid, populate added
            for item in added:
                #self._log(model.db[item].all_field_keys())
                cid = model.db[item].application_id

                # Add book_id, cid to item dict, update installed_books with cid
                for book in self.installed_books.values():
                    if book.path == added[item]['path']:
                        added[item]['cid'] = cid
                        added[item]['book_id'] = book.mid
                        book.cid = cid
                        break

                # Add model_row
                for model_row in bta:
                    if bta[model_row]['book_id'] == added[item]['book_id']:
                        added[item]['model_row'] = model_row

            pb = ProgressBar(parent=self.opts.gui, window_title="Updating calibre metadata",
                             on_top=True)
            total_books = len(added)
            # Show progress in dispatched method - 2 times
            pb.set_maximum(total_books * 2)
            pb.set_value(0)
            pb.set_label('{:^100}'.format("1 of %d" % (total_books)))
            pb.show()

            db = self.opts.gui.current_db
            cached_books = self.parent.connected_device.cached_books

            # Update calibre metadata from Marvin metadata, bind uuid
            for item in added:
                mismatches = {}
                this_book = cached_books[added[item]['path']]
                mismatches['authors'] = {'Marvin': this_book['authors']}
                mismatches['author_sort'] = {'Marvin': this_book['author_sort']}
                mismatches['cover_hash'] = {'Marvin': this_book['cover_hash']}
                mismatches['comments'] = {'Marvin': this_book['description']}
                mismatches['pubdate'] = {'Marvin': this_book['pubdate']}
                mismatches['publisher'] = {'Marvin': this_book['publisher']}
                if this_book['series']:
                    mismatches['series'] = {'Marvin': this_book['series']}
                    mismatches['series_index'] = {'Marvin': this_book['series_index']}
                mismatches['tags'] = {'Marvin': this_book['tags']}
                mismatches['title'] = {'Marvin': this_book['title']}
                mismatches['title_sort'] = {'Marvin': this_book['title_sort']}

                # Get the newly minted calibre uuid
                mi = db.get_metadata(added[item]['cid'], index_is_id=True)
                mismatches['uuid'] = {'calibre': mi.uuid,
                                      'Marvin': this_book['uuid']}

                # Do the magic
                self._update_calibre_metadata(added[item]['book_id'],
                                              added[item]['cid'],
                                              mismatches,
                                              added[item]['model_row'],
                                              pb)
            pb.hide()

            # Launch row flasher
            self._flash_affected_rows()

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
                    blocking_busy = MyBlockingBusy(self.opts.gui, "Updating Marvin Library…", size=60)
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
                                    new = self.GREEN
                                else:
                                    new = self.WHITE

                                old = self.tm.get_match_quality(row)
                                self.tm.set_match_quality(row, new)
                                self.updated_match_quality[row] = {'book_id': book_id,
                                                                   'old': old,
                                                                   'new': new}

                # Launch row flasher
                self._flash_affected_rows()

            else:
                self._log("delete cancelled")

        else:
            self._log("no books selected")
            title = "No selected books"
            msg = "<p>Select one or more books to delete.</p>"
            MessageBox(MessageBox.INFO, title, msg,
                       show_copy_button=False).exec_()

    def _fetch_annotations(self):
        '''
        A lightweight version of fetch annotations
        Request HTML annotations from Marvin, add to custom column specified in config
        '''
        self._log_location()

        lookup = self.parent.prefs.get('annotations_field_lookup', None)
        if lookup:
            for row, book in self._selected_books().items():
                cid = book['cid']
                if cid is not None:
                    if book['has_annotations']:
                        self._log("row %d has annotations" % row)
                        book_id = book['book_id']

                        # Build the command
                        command_name = "command"
                        command_type = "GetAnnotationsHTML"
                        update_soup = BeautifulStoneSoup(self.GENERAL_COMMAND_XML.format(
                            command_type, time.mktime(time.localtime())))
                        parameters_tag = self._build_parameters(self.installed_books[book_id], update_soup)
                        update_soup.command.insert(0, parameters_tag)

                        response = self._issue_command(command_name, update_soup,
                                                       get_response="html_response.html",
                                                       update_local_db=False)
                        """
                        # Copy command file to staging folder
                        self._stage_command_file(command_name, update_soup,
                                                 show_command=self.prefs.get('show_staged_commands', False))

                        # Wait for completion
                        html_response = self._wait_for_command_completion(command_name,
                                                                          update_local_db=False,
                                                                          get_response="html_response.html")
                        """
                        if not response:
                            response = '\n'.join(self.installed_books[book_id].highlights)

                        # Apply to custom column
                        # Get the current value from the lookup field
                        db = self.opts.gui.current_db
                        mi = db.get_metadata(cid, index_is_id=True)
                        #old_value = mi.get_user_metadata(lookup, False)['#value#']
                        #self._log("Updating old value: %s" % repr(old_value))

                        um = mi.metadata_for_field(lookup)
                        um['#value#'] = response
                        mi.set_user_metadata(lookup, um)
                        db.set_metadata(cid, mi, set_title=False, set_authors=False,
                                        commit=True)

                    else:
                        self._log("'%s' has no annotations" % book['title'])

            updateCalibreGUIView()

    def _fetch_marvin_content_hash(self, path):
        '''
        Given a Marvin path, compute/fetch a hash of its contents (excluding OPF)
        '''
        if self.opts.prefs.get('development_mode', False):
            self._log_location(path)

        # Try getting the hash from the cache
        try:
            zfr = ZipFile(self.local_hash_cache)
            hash = zfr.read(path)
        except:
            if self.opts.prefs.get('development_mode', False):
                self._log("opening local hash cache for appending")
            zfw = ZipFile(self.local_hash_cache, mode='a')
        else:
            if self.opts.prefs.get('development_mode', False):
                self._log("returning hash from cache: %s" % hash)
            zfr.close()
            return hash

        # Get a local copy, generate hash
        rbp = '/'.join(['/Documents', path])
        lbp = os.path.join(self.local_cache_folder, path)

        with open(lbp, 'wb') as out:
            self.ios.copy_from_idevice(str(rbp), out)

        hash = self._compute_epub_hash(lbp)
        zfw.writestr(path, hash)
        zfw.close()

        # Delete the local copy
        os.remove(lbp)
        return hash

    def _fetch_marvin_cover(self, book_id):
        '''
        Retrieve Books:LargeCoverJpg if no cover_path
        '''
        marvin_cover = None
        if self.installed_books[book_id].cover_file:
            self._log_location("fetching cover from Marvin sandbox")
            self._log("*** NOT IMPLEMENTED ***")
            # Return cover file as bytes

        else:
            self._log_location("fetching cover from mainDb")
            con = sqlite3.connect(self.parent.connected_device.local_db_path)
            with con:
                con.row_factory = sqlite3.Row

                # Fetch LargeCoverJpg from mainDb
                cover_cur = con.cursor()
                cover_cur.execute('''SELECT
                                      LargeCoverJpg
                                     FROM Books
                                     WHERE ID = '{0}'
                                  '''.format(book_id))
                rows = cover_cur.fetchall()

            if len(rows):
                marvin_cover = rows[0][b'LargeCoverJpg']
            else:
                self._log_location("no cover data fetched from mainDb")

        return marvin_cover

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

        # Scan library books for hashes
        if self.library_scanner.isRunning():
            self.library_scanner.wait()

        # Save a reference to the title, uuid map
        self.library_title_map = self.library_scanner.title_map
        self.library_uuid_map = self.library_scanner.uuid_map

        # Get the library hash_map
        library_hash_map = self.library_scanner.hash_map
        if library_hash_map is None:
            library_hash_map = self._scan_library_books(self.library_scanner)
        else:
            self._log("hash_map already generated")

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
        elif (book_data.calibre_collections is None and
                book_data.device_collections > []):
            base_name = 'collections_info.png'
            sort_value = 0
        elif (book_data.device_collections == [] and
              book_data.calibre_collections == []):
            base_name = 'collections_empty.png'
            sort_value = 0
        elif book_data.device_collections == book_data.calibre_collections:
            base_name = 'collections_equal.png'
            sort_value = 2
        else:
            base_name = 'collections_unequal.png'
            sort_value = 1
        collection_match = SortableImageWidgetItem(self,
                                                   os.path.join(self.parent.opts.resources_path,
                                                                'icons', base_name),
                                                   sort_value, self.COLLECTIONS_COL)
        return collection_match

    def _generate_deep_view(self):
        '''
        '''
        self._log_location()

        command_name = "command"
        command_type = "GenerateDeepView"
        update_soup = BeautifulStoneSoup(self.GENERAL_COMMAND_XML.format(
            command_type, time.mktime(time.localtime())))

        selected_books = self._selected_books()
        if selected_books:
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

            self.busy_window = MyBlockingBusy(self,
                                              "Generating Deep View for %s" %
                                               ("1 book…" if len(selected_books) == 1 else
                                                "%d books…" % len(selected_books)),
                                              size=60,
                                              show_cancel=True)
            QTimer.singleShot(0, self._start_busy_window)

            # Wait for the window to show up
            while not self.busy_window.is_running:
                Application.processEvents()

            self._issue_command(command_name, update_soup,
                                ignore_timeouts=True)

            self.busy_window.stop()
            self.busy_window.accept()
            self.busy_window = None

            """
            # Copy command file to staging folder
            self._stage_command_file(command_name, update_soup,
                                     show_command=self.prefs.get('show_staged_commands', False))

            # Wait for completion
            self._wait_for_command_completion(command_name, update_local_db=True,
                                              ignore_timeouts=True)
            """

            # Update visible model, self.installed_books
            for row in sorted(selected_books.keys(), reverse=True):
                book_id = selected_books[row]['book_id']
                self.installed_books[book_id].deep_view_prepared = 1
                self.tm.set_deep_view(row, self.CHECKMARK)

    def _generate_marvin_hash_map(self, installed_books):
        '''
        Generate a map of book_ids to hash values
        {hash: [book_id, book_id,...], ...}
        '''
        self._log_location()
        hash_map = {}
        for book_id in installed_books:
            hash = installed_books[book_id].hash
            #title = installed_books[book_id].title
#             self._log("%s: %s" % (title, hash))
            if hash in hash_map:
                hash_map[hash].append(book_id)
            else:
                hash_map[hash] = [book_id]

#         for hash in hash_map:
#             self._log("%s: %s" % (hash, hash_map[hash]))

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
            elif book_data.progress < 0.01:
                percent_read = ''
            else:
                # Pad the right side for visual comfort, since this col is
                # right-aligned
                percent_read = "{:3.0f}%   ".format(book_data.progress * 100)
            progress = SortableTableWidgetItem(percent_read, pct_progress)
        else:
            #base_name = "progress000.png"
            base_name = "progress_none.png"
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

            progress = SortableImageWidgetItem(self,
                                               os.path.join(self.parent.opts.resources_path,
                                               'icons', base_name),
                                               pct_progress,
                                               self.PROGRESS_COL)
        return progress

    def _get_calibre_collections(self, cid):
        '''
        Return a sorted list of current calibre collection assignments or
        None if no collection_field_lookup assigned or book does not exist in library
        '''
        cfl = self.prefs.get('collection_field_lookup', '')
        if cfl == '' or cid is None:
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

    def _get_marvin_collections(self, book_id):
        return sorted(self.installed_books[book_id].device_collections, key=sort_key)

    def _get_installed_books(self):
        '''
        Build a profile of all installed books for display
        On Device
        Title
        Author
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
                        self._log("UUID match")
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
                pass
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
                ach = self.archived_cover_hashes.get(str(this_book.cid), {})
                cover_last_modified = self.opts.gui.current_db.cover_last_modified(this_book.cid, index_is_id=True)
                if ('cover_last_modified' in ach and
                        ach['cover_last_modified'] == cover_last_modified):
                    return ach['cover_hash']

                # Generate calibre cover hash (same process used by driver when sending books)
                cover_hash = 0
                desired_thumbnail_height = self.parent.connected_device.THUMBNAIL_HEIGHT
                try:
                    sized_thumb = thumbnail(mi.cover_data[1],
                                            desired_thumbnail_height,
                                            desired_thumbnail_height)
                    cover_hash = hashlib.md5(sized_thumb[2]).hexdigest()
                    cover_last_modified = self.opts.gui.current_db.cover_last_modified(this_book.cid, index_is_id=True)
                    self.archived_cover_hashes.set(str(this_book.cid),
                                                   {'cover_hash': cover_hash,
                                                    'cover_last_modified': cover_last_modified})
                except:
                    self._log("error calculating cover_hash for cid %d (%s)" % (this_book.cid, this_book.title))
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
                if bool(row[b'DatePublished']) or bool(mi.pubdate):
                    try:
                        mb_pubdate = datetime.utcfromtimestamp(int(row[b'DatePublished']))
                        mb_pubdate = mb_pubdate.replace(hour=0, minute=0, second=0)
                    except:
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
                    mismatches['publisher'] = {'calibre': mi.publisher,
                                               'Marvin': row[b'Publisher']}

                # ~~~~~~~~ series, series_index ~~~~~~~~
                if bool(mi.series) or bool(row[b'CalibreSeries']):
                    if mi.series != row[b'CalibreSeries']:
                        mismatches['series'] = {'calibre': mi.series,
                                                'Marvin': row[b'CalibreSeries']}

                if bool(mi.series_index) or bool(float(row[b'CalibreSeriesIndex'])):
                    if mi.series_index != float(row[b'CalibreSeriesIndex']):
                        mismatches['series_index'] = {'calibre': mi.series_index,
                                                      'Marvin': row[b'CalibreSeriesIndex']}

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
            if self.opts.prefs.get('development_mode', False):
                self._log_location()

            ans = None
            if cid:
                db = self.opts.gui.current_db
                mi = db.get_metadata(cid, index_is_id=True)
                ans = mi.ondevice_col
            return ans

        def _get_pubdate(row):
            try:
                pubdate = datetime.utcfromtimestamp(int(row[b'DatePublished']))
            except:
                pubdate = None
            return pubdate

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

            for ch_cid in cover_hash_cids:
                if ch_cid not in active_cids:
                    self._log("removing orphan cid %s from archived_cover_hashes" % ch_cid)
                    del self.archived_cover_hashes[ch_cid]

        self._log_location()

        if self.opts.prefs.get('development_mode', False):
            self._log("local_db_path: %s" % self.parent.connected_device.local_db_path)

        # Fetch/compute hashes
        cached_books = self.parent.connected_device.cached_books
        hashes = self._scan_marvin_books(cached_books)

        # Get the mainDb data
        installed_books = {}
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
                            DateOpened,
                            DatePublished,
                            DeepViewPrepared,
                            Description,
                            FileName,
                            IsRead,
                            NewFlag,
                            Progress,
                            Publisher,
                            ReadingList,
                            Title,
                            UUID,
                            WordCount
                          FROM Books
                        ''')

            rows = cur.fetchall()

            pb = ProgressBar(parent=self.opts.gui, window_title="Performing Marvin metadata magic", on_top=True)
            book_count = len(rows)
            pb.set_maximum(book_count)
            pb.set_value(0)
            pb.set_label('{:^100}'.format("1 of %d" % (book_count)))
            pb.show()

            for i, row in enumerate(rows):
                pb.set_label('{:^100}'.format("%d of %d" % (i+1, book_count)))

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
                this_book.device_collections = _get_collections(cur, book_id)
                this_book.date_opened = row[b'DateOpened']
                this_book.deep_view_prepared = row[b'DeepViewPrepared']
                this_book.flags = _get_flags(cur, row)
                this_book.hash = hashes[row[b'FileName']]['hash']
                this_book.highlights = _get_highlights(cur, book_id)
                this_book.metadata_mismatches = _get_metadata_mismatches(cur, book_id, row, mi, this_book)
                this_book.mid = book_id
                this_book.on_device = _get_on_device_status(this_book.cid)
                this_book.path = row[b'FileName']
                this_book.progress = row[b'Progress']
                this_book.pubdate = _get_pubdate(row)
                this_book.tags = _get_marvin_genres(book_id)
                this_book.title_sort = row[b'CalibreTitleSort']
                this_book.uuid = row[b'UUID']
                this_book.vocabulary = _get_vocabulary_list(cur, book_id)
                this_book.word_count = locale.format("%d", row[b'WordCount'], grouping=True)
                installed_books[book_id] = this_book

                pb.increment()

            pb.hide()

        # Remove orphan cover_hashes
        _purge_cover_hash_orphans()

        if self.opts.prefs.get('development_mode', False):
            self._log("%d cached books from Marvin:" % len(cached_books))
            for book in installed_books:
                self._log("%s word_count: %s" % (installed_books[book].title,
                                                 repr(installed_books[book].word_count)))
        return installed_books

    def _inform_marvin_collections(self, book_id):
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

        self._issue_command(command_name, update_soup)
        """
        # Copy command file to staging folder
        self._stage_command_file(command_name, update_soup,
                                 show_command=self.prefs.get('show_staged_commands', False))

        # Wait for completion
        self._wait_for_command_completion(command_name, update_local_db=True)
        """

    def _issue_command(self, command_name, update_soup,
                       get_response=None,
                       ignore_timeouts=False,
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
        response = self._wait_for_command_completion(command_name,
                                                     ignore_timeouts=ignore_timeouts,
                                                     get_response=get_response,
                                                     update_local_db=update_local_db)

        self.parent.connected_device.set_busy_flag(False)

        QApplication.restoreOverrideCursor()

        if get_response:
            return response

    def _localize_marvin_database(self):
        '''
        Copy remote_db_path from iOS to local storage using device pointers
        '''
        self._log_location()

        local_db_path = self.parent.connected_device.local_db_path
        remote_db_path = self.parent.connected_device.books_subpath
        with open(local_db_path, 'wb') as out:
            self.ios.copy_from_idevice(remote_db_path, out)

    def _localize_hash_cache(self, cached_books):
        '''
        Check for existence of hash cache on iDevice. Confirm/create folder
        If existing cached, purge orphans
        '''
        self._log_location()

        # Existing hash cache?
        lhc = os.path.join(self.local_cache_folder, self.hash_cache)
        rhc = '/'.join([self.remote_cache_folder, self.hash_cache])

        cache_exists = (self.ios.exists(rhc) and
                        not self.opts.prefs.get('hash_caching_disabled'))
        if cache_exists:
            # Copy from existing remote cache to local cache
            self._log("copying remote hash cache")
            with open(lhc, 'wb') as out:
                self.ios.copy_from_idevice(str(rhc), out)
        else:
            # Confirm path to remote folder is valid store point
            folder_exists = self.ios.exists(self.remote_cache_folder)
            if not folder_exists:
                self._log("creating remote_cache_folder %s" % repr(self.remote_cache_folder))
                self.ios.mkdir(self.remote_cache_folder)
            else:
                self._log("remote_cache_folder exists")

            # Create a local cache
            self._log("creating new local hash cache: %s" % repr(lhc))
            zfw = ZipFile(lhc, mode='w')
            zfw.writestr('Marvin hash cache', '')
            zfw.close()

            # Clear the marvin_content_updated flag
            if self.parent.marvin_content_updated:
                self.parent.marvin_content_updated = False

        self.local_hash_cache = lhc
        self.remote_hash_cache = rhc

        # Purge cache orphans
        if cache_exists:
            self._purge_cached_orphans(cached_books)

    def _log(self, msg=None):
        '''
        Print msg to console
        '''
        if not self.verbose:
            return

        if msg:
            debug_print(" %s" % str(msg))
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
            arg1 = str(args[0])
        if len(args) > 1:
            arg2 = str(args[1])

        debug_print(self.LOCATION_TEMPLATE.format(cls=self.__class__.__name__,
                    func=sys._getframe(1).f_code.co_name,
                    arg1=arg1, arg2=arg2))

    def _purge_cached_orphans(self, cached_books):
        '''

        '''
        self._log_location()
        zfa = ZipFile(self.local_hash_cache, mode='a')
        zfi = zfa.infolist()
        for zi in zfi:
            if zi.filename == 'Marvin hash cache':
                continue
            if zi.filename not in cached_books:
                self._log("removing %s from hash cache" % repr(zi.filename))
                zfa.delete(zi.filename)
        zfa.close()

    def _save_column_widths(self):
        '''
        '''
        self._log_location()
        try:
            widths = []
            for (i, c) in enumerate(self.LIBRARY_HEADER):
                widths.append(self.tv.columnWidth(i))
            self.opts.prefs.set('marvin_library_column_widths', widths)
        except:
            pass

    def _scan_library_books(self, library_scanner):
        '''
        Generate hashes for library epubs
        '''
        # Scan library books for hashes
        if self.library_scanner.isRunning():
            self.library_scanner.wait()

        uuid_map = library_scanner.uuid_map
        self._log_location("%d epubs" % len(uuid_map))

        pb = ProgressBar(parent=self.opts.gui, window_title="Scanning library", on_top=True)
        total_books = len(uuid_map)
        pb.set_maximum(total_books)
        pb.set_value(0)
        pb.set_label('{:^100}'.format("1 of %d" % (total_books)))
        pb.show()

        db = self.opts.gui.current_db
        close_requested = False
        for i, uuid in enumerate(uuid_map):
            try:
                pb.set_label('{:^100}'.format("%d of %d" % (i+1, total_books)))

                path = db.format(uuid_map[uuid]['id'], 'epub', index_is_id=True,
                                 as_path=True, preserve_filename=True)
                uuid_map[uuid]['hash'] = self._compute_epub_hash(path)
                os.remove(path)
            except:
                # Book deleted since scan
                pass

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
        self._log_location()

        # Fetch pre-existing hash cache from device
        self._localize_hash_cache(cached_books)

        # Set up the progress bar
        pb = ProgressBar(parent=self.opts.gui, window_title="Scanning Marvin", on_top=True)
        total_books = len(cached_books)
        pb.set_maximum(total_books)
        pb.set_value(0)
        pb.set_label('{:^100}'.format("1 of %d" % (total_books)))
        pb.show()

        close_requested = False
        installed_books = {}
        for i, path in enumerate(cached_books):
            this_book = {}
            pb.set_label('{:^100}'.format("%d of %d" % (i+1, total_books)))
            this_book['hash'] = self._fetch_marvin_content_hash(path)

            installed_books[path] = this_book
            pb.increment()

            if pb.close_requested:
                close_requested = True
                break
        else:
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
            has_annotations = bool(self.tm.get_annotations(row))
            has_articles = bool(self.tm.get_articles(row))
            has_dv_content = bool(self.tm.get_deep_view(row))
            has_vocabulary = bool(self.tm.get_vocabulary(row))
            last_opened = str(self.tm.get_last_opened(row).text())
            path = self.tm.get_path(row)
            progress = self.tm.get_progress(row).sort_key
            title = str(self.tm.get_title(row).text())
            uuid = self.tm.get_uuid(row)
            selected_books[row] = {
                'author': author,
                'book_id': book_id,
                'cid': cid,
                'has_annotations': has_annotations,
                'has_articles': has_articles,
                'has_dv_content': has_dv_content,
                'has_vocabulary': has_vocabulary,
                'last_opened': last_opened,
                'path': path,
                'progress': progress,
                'title': title,
                'uuid': uuid}

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

    def _set_flags(self, action):
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

        selected_books = self._selected_books()
        for row in selected_books:
            book_id = selected_books[row]['book_id']
            flagbits = self.tm.get_flags(row).sort_key
            path = selected_books[row]['path']
            if not flagbits & mask:
                # Set the bit with OR
                flagbits = flagbits | mask
                flagbits = flagbits & inhibit
                basename = "flags%d.png" % flagbits
                new_flags_widget = SortableImageWidgetItem(self,
                                                           os.path.join(self.parent.opts.resources_path,
                                                           'icons', basename),
                                                           flagbits, self.FLAGS_COL)
                # Update the spreadsheet
                self.tm.set_flags(row, new_flags_widget)

                # Update self.installed_books flags list
                self.installed_books[book_id].flags = _build_flag_list(flagbits)

                # Update reading progress based on flag values
                self._update_reading_progress(self.installed_books[book_id], row)

                # Update in-memory
                _update_in_memory(book_id, path)

                # Update Marvin db
                self._inform_marvin_collections(book_id)

            self._update_device_flags(book_id, path, _build_flag_list(flagbits))

        self.repaint()

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

    def _start_busy_window(self):
        '''
        '''
        self._log_location()
        self.busy_window.start()
        self.busy_window.show()

    def _update_calibre_collections(self, book_id, cid, updated_calibre_collections):
        '''
        '''
        self._log_location()
        # Update collections custom column
        lookup = self.parent.prefs.get('collection_field_lookup', None)
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

    def _update_calibre_metadata(self, book_id, cid, mismatches, model_row, pb):
        '''
        Update calibre from Marvin metadata
        If uuids differ, we need to send an update_metadata command to Marvin
        pb is incremented twice per book.
        '''

        pb.increment()

        # Highlight the row we're working on
        self.tv.selectRow(model_row)

        # Get the current metadata
        db = self.opts.gui.current_db
        mi = db.get_metadata(cid, index_is_id=True, get_cover=True, cover_as_data=True)

        self._log_location(mi.title)
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
                    cover_hash = hashlib.md5(marvin_cover).hexdigest()

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

                    self._issue_command(command_name, update_soup)
                    """
                    # Copy command file to staging folder
                    self._stage_command_file(command_name, update_soup,
                                             show_command=self.prefs.get('show_staged_commands', False))

                    # Wait for completion
                    self._wait_for_command_completion(command_name, update_local_db=True)
                    """

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

                self._issue_command(command_name, update_soup)
                """
                # Copy command file to staging folder
                self._stage_command_file(command_name, update_soup,
                                         show_command=self.prefs.get('show_staged_commands', False))

                # Wait for completion
                self._wait_for_command_completion(command_name, update_local_db=True)
                """

            self._clear_selected_rows()

        pb.increment()

        # Update metadata match quality in the visible model
        old = self.tm.get_match_quality(model_row)
        new = self.GREEN
        self.tm.set_match_quality(model_row, new)
        self.updated_match_quality[model_row] = {'book_id': book_id,
                                                 'old': old,
                                                 'new': new}

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
        '''
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
        lookup = self.parent.prefs.get('collection_field_lookup', None)
        if lookup is not None and details['active_cids']:
            for ctd in details['delete']:
                if ctd in details['locations']['calibre']:
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
                if ctr in details['locations']['calibre']:
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
        command_name = "command"
        command_type = "UpdateGlobalCollections"
        update_soup = BeautifulStoneSoup(self.GENERAL_COMMAND_XML.format(
            command_type, time.mktime(time.localtime())))
        parameters_tag = Tag(update_soup, 'parameters')
        update_soup.command.insert(0, parameters_tag)

        cached_books = self.parent.connected_device.cached_books
        for ctd in details['delete']:
            if ctd in details['locations']['Marvin']:
                # update self.installed_books, Device model
                for book_id, book in self.installed_books.items():
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
                        parameter_tag['action'] = "delete"
                        parameter_tag.insert(0, ctd)
                        parameters_tag.insert(0, parameter_tag)

        for ctr in details['rename']:
            if ctr in details['locations']['Marvin']:
                for book_id, book in self.installed_books.items():
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

                        # Add a <parameter action="rename"> tag
                        parameter_tag = Tag(update_soup, 'parameter')
                        parameter_tag['action'] = "rename"
                        parameter_tag['newname'] = replacement
                        parameter_tag.insert(0, ctr)
                        parameters_tag.insert(0, parameter_tag)

        # Tell Marvin
        if len(parameters_tag):

            self._issue_command(command_name, update_soup)
            """
            # Copy command file to staging folder
            self._stage_command_file(command_name, update_soup,
                                     show_command=self.prefs.get('show_staged_commands', False))

            # Wait for completion
            self._wait_for_command_completion(command_name, update_local_db=True)
            """

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

    def _update_marvin_metadata(self, book_id, cid, mismatches, model_row, pb):
        '''
        Update Marvin from calibre metadata
        This clones upload_books() in the iOS reader application driver
        All metadata is asserted, cover optional if changes
        '''

        # Highlight the row we're working on
        self.tv.selectRow(model_row)

        # Get the current metadata
        db = self.opts.gui.current_db
        mi = db.get_metadata(cid, index_is_id=True, get_cover=True, cover_as_data=True)

        self._log_location(mi.title)
        self._log("mismatches:\n%s" % mismatches)

        command_name = "update_metadata"
        update_soup = self._build_metadata_update(book_id, cid, mi, mismatches)
        self._issue_command(command_name, update_soup)
        """
        # Copy the command file to the staging folder
        self._stage_command_file("update_metadata", update_soup,
                                 show_command=self.prefs.get('show_staged_commands', False))

        # Wait for completion
        self._wait_for_command_completion("update_metadata", update_local_db=True)
        """

        # 2x progress on purpose
        pb.increment()
        pb.increment()

        # Update in-memory caches
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
                self.opts.gui.memory_view.model().db[device_view_row].authors = authors

            if key == 'author_sort':
                author_sort = mismatches[key]['calibre']
                cached_books[path]['author_sort'] = author_sort
                self.installed_books[book_id].author_sort = author_sort
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
                self.opts.gui.memory_view.model().db[device_view_row].title = title

            if key == 'title_sort':
                title_sort = mismatches[key]['calibre']
                cached_books[path]['title_sort'] = title_sort
                self.installed_books[book_id].title_sort = title_sort
                self.opts.gui.memory_view.model().db[device_view_row].title_sort = title_sort

            if key == 'uuid':
                uuid = mismatches[key]['calibre']
                cached_books[path]['uuid'] = uuid
                self.installed_books[book_id].matches = [uuid]

                self.installed_books[book_id].uuid = uuid
                self.opts.gui.memory_view.model().db[device_view_row].uuid = uuid
                self.opts.gui.memory_view.model().db[device_view_row].in_library = "UUID"

                # Add to hash_map
                self.library_scanner.add_to_hash_map(self.installed_books[book_id].hash, uuid)

            self._clear_selected_rows()

        # Update metadata match quality in the visible model
        old = self.tm.get_match_quality(model_row)
        self.tm.set_match_quality(model_row, self.GREEN)
        self.updated_match_quality[model_row] = {'book_id': book_id,
                                                 'old': old,
                                                 'new': self.GREEN}

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
            if self.tm.get_match_quality(row) == self.ORANGE:
                title = "Duplicate book"
                msg = ("<p>'{0}' is a duplicate.</p>".format(self.installed_books[book_id].title) +
                       "<p>Remove duplicates before updating metadata.</p>")
                return MessageBox(MessageBox.WARNING, title, msg,
                                  show_copy_button=False).exec_()

        pb = ProgressBar(parent=self.opts.gui, window_title="Updating metadata",
                         on_top=True)
        total_books = len(selected_books)
        # Show progress in dispatched method - 2 times
        pb.set_maximum(total_books * 2)
        pb.set_value(0)
        pb.set_label('{:^100}'.format("1 of %d" % (total_books)))
        pb.show()

        self.updated_match_quality = {}

        for row in sorted(selected_books):
            book_id = self._selected_book_id(row)
            cid = self._selected_cid(row)
            mismatches = self.installed_books[book_id].metadata_mismatches
            pb.set_label('{:^100}'.format(self.installed_books[book_id].title))
            if action == 'export_metadata':
                # Apply calibre metadata to Marvin
                self._update_marvin_metadata(book_id, cid, mismatches, row, pb)

            elif action == 'import_metadata':
                # Apply Marvin metadata to calibre
                self._update_calibre_metadata(book_id, cid, mismatches, row, pb)

            # Clear the metadata_mismatch
            self.installed_books[book_id].metadata_mismatches = {}

            if pb.close_requested:
                break

        pb.hide()

        # Launch row flasher
        self._flash_affected_rows()

    def _update_reading_progress(self, book, row):
        '''
        Refresh Progress column
        '''
        self._log_location()
        progress = self._generate_reading_progress(book)
        self.tm.set_progress(row, progress)

    def _update_remote_hash_cache(self):
        '''
        Copy updated hash cache to iDevice
        self.local_hash_cache, self.remote_hash_cache initialized
        in _localize_hash_cache()
        '''
        self._log_location()

        if self.parent.prefs.get('hash_caching_disabled', False):
            self._log("hash_caching_disabled, deleting remote hash cache")
            self.ios.remove(str(self.remote_hash_cache))
        else:
            # Copy local cache to iDevice
            self.ios.copy_to_idevice(self.local_hash_cache, str(self.remote_hash_cache))

    def _wait_for_command_completion(self, command_name, update_local_db=True,
            get_response=None, ignore_timeouts=False):
        '''
        Wait for Marvin to issue progress reports via status.xml
        Marvin creates status.xml upon receiving command, increments <progress>
        from 0.0 to 1.0 as command progresses.
        '''
        self._log_location(command_name)

        response = None

        if self.prefs.get('execute_marvin_commands', True):
            self._log("%s: waiting for '%s'" %
                      (datetime.now().strftime('%H:%M:%S.%f'),
                      self.parent.connected_device.status_fs))

            # Set initial watchdog timer for ACK

            self.watchdog = Timer(self.WATCHDOG_TIMEOUT, self._watchdog_timed_out)
            self.operation_timed_out = False
            self.ignore_timeouts = ignore_timeouts

            self.watchdog.start()

            while True:
                if not self.ios.exists(self.parent.connected_device.status_fs):
                    # status.xml not created yet
                    if self.operation_timed_out:
                        self.ios.remove(self.parent.connected_device.status_fs)
                        raise UserFeedback("Marvin operation timed out.",
                                           details=None, level=UserFeedback.WARN)
                        break
                    #time.sleep(0.10)
                    Application.processEvents()

                else:
                    self.watchdog.cancel()

                    self._log("%s: monitoring progress of %s" %
                              (datetime.now().strftime('%H:%M:%S.%f'),
                              command_name))

                    # Start a new watchdog timer per iteration
                    self.watchdog = Timer(self.WATCHDOG_TIMEOUT, self._watchdog_timed_out)
                    self.operation_timed_out = False
                    self.watchdog.start()

                    code = '-1'
                    current_timestamp = 0.0
                    while code == '-1':
                        try:
                            if self.operation_timed_out:
                                self.ios.remove(self.parent.connected_device.status_fs)
                                raise UserFeedback("Marvin operation timed out.",
                                                   details=None, level=UserFeedback.WARN)
                                break

                            # Cancel requested?
                            if self.busy_window is not None:
                                if self.busy_window.cancel_status == self.busy_window.REQUESTED:
                                    self._log("user requested cancel")

                                    # Create "cancel.command" in staging folder
                                    ft = (b'/'.join([self.parent.connected_device.staging_folder,
                                                     b'cancel.tmp']))
                                    fs = (b'/'.join([self.parent.connected_device.staging_folder,
                                                     b'cancel.command']))
                                    self.ios.write("please stop", ft)
                                    self.ios.rename(ft, fs)

                                    # Change dialog text
                                    self.busy_window.set_text("Completing current book…")

                                    self.busy_window.cancel_status = self.busy_window.ACKNOWLEDGED

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
                                self.watchdog = Timer(self.WATCHDOG_TIMEOUT, self._watchdog_timed_out)
                                self.watchdog.start()
                            #time.sleep(0.01)
                            Application.processEvents()

                        except:
                            #import traceback
                            #self._log(traceback.format_exc())
                            #time.sleep(1.0)
                            Application.processEvents()
                            self._log("%s:  retry" % datetime.now().strftime('%H:%M:%S.%f'))

                    # Command completed
                    self.watchdog.cancel()

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

                    if final_code not in ['0', '3']:
                        if final_code == '-1':
                            final_status = "in progress"
                        elif final_code == '1':
                            final_status = "warnings"
                        elif final_code == '2':
                            final_status = "errors"
                        elif final_code == '3':
                            final_status = "cancelled"

                        messages = status.find('messages')
                        msgs = [msg.text for msg in messages]
                        details = '\n'.join(["code: %s" % final_code, "status: %s" % final_status])
                        details += '\n'.join(msgs)
                        self._log(details)

                        self.ios.remove(self.parent.connected_device.status_fs)

                        self._log("%s: '%s' complete with errors" %
                                  (datetime.now().strftime('%H:%M:%S.%f'),
                                  command_name))

                        raise UserFeedback("Operation %s.\nClick 'Show details' for more information."
                                           % (final_status),
                                           details=details, level=UserFeedback.WARN)

                    # Get the response file from the staging folder
                    if get_response:
                        rf = b'/'.join([self.parent.connected_device.staging_folder, get_response])
                        self._log("fetching response '%s'" % rf)
                        if not self.ios.exists(self.parent.connected_device.status_fs):
                            response = "%s not found" % rf
                        else:
                            response = self.ios.read(rf)
                            self.ios.remove(rf)

                    self.ios.remove(self.parent.connected_device.status_fs)

                    self._log("%s: '%s' complete" %
                              (datetime.now().strftime('%H:%M:%S.%f'),
                              command_name))
                    break

            # Update local copy of Marvin db
            if update_local_db:
                self._localize_marvin_database()

        else:
            self._log("~~~ execute_marvin_commands disabled in JSON ~~~")

        return response

    def _watchdog_timed_out(self):
        '''
        Set flag if I/O operation times out
        '''
        self._log_location(datetime.now().strftime('%H:%M:%S.%f'))

        if self.ignore_timeouts:
            # Start a new watchdog timer per iteration
            self._log("timeouts ignored, resetting timer")
            self.watchdog = Timer(self.WATCHDOG_TIMEOUT, self._watchdog_timed_out)
            self.operation_timed_out = False
            self.watchdog.start()
        else:
            self.operation_timed_out = True
