#!/usr/bin/env python
# coding: utf-8

__license__ = 'GPL v3'
__copyright__ = '2013, Greg Riker <griker@hotmail.com>'
__docformat__ = 'restructuredtext en'

import hashlib, locale, operator, os, re, sqlite3, sys, time
from lxml import etree

from PyQt4 import QtCore, QtGui
from PyQt4.Qt import (Qt, QAbstractItemModel, QAbstractTableModel, QBrush,
                      QCheckBox, QColor, QDialog, QDialogButtonBox, QFont, QLabel,
                      QTableView, QTableWidgetItem,
                      QVariant, QVBoxLayout,
                      SIGNAL)
from PyQt4.QtWebKit import QWebView

from calibre.constants import islinux, isosx, iswindows
from calibre.devices.usbms.driver import debug_print
from calibre.ebooks.oeb.iterator import EbookIterator
from calibre.gui2.dialogs.message_box import MessageBox
from calibre.utils.icu import sort_key
from calibre.utils.wordcount import get_wordcount_obj
from calibre.utils.zipfile import ZipFile

from calibre_plugins.marvin_manager.common_utils import (
    Book, HelpView, ProgressBar, SizePersistedDialog)


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

    def __init__(self, parent=None, columns_to_center=[], right_aligned_columns=[], *args):
        """
        datain: a list of lists
        headerdata: a list of strings
        """
        QAbstractTableModel.__init__(self, parent, *args)
        self.parent = parent
        self.arraydata = parent.tabledata
        self.centered_columns = columns_to_center
        self.right_aligned_columns = right_aligned_columns
        self.headerdata = parent.library_header
        self.show_confidence_colors = parent.show_confidence_colors

    def rowCount(self, parent):
        return len(self.arraydata)

    def columnCount(self, parent):
        return len(self.headerdata)

    def data(self, index, role):
        row, col = index.row(), index.column()
        if not index.isValid():
            return QVariant()
        elif role == Qt.BackgroundRole and self.show_confidence_colors:
            match_quality = self.arraydata[row][self.parent.MATCHED_COL]

            saturation = 0.40
            value = 1.0
            red_hue = 0.0
            green_hue = 0.333
            yellow_hue = 0.1665
            white_hue = 1.0
            if match_quality == 3:
                return QVariant(QBrush(QColor.fromHsvF(green_hue, saturation, value)))
            elif match_quality == 2:
                return QVariant(QBrush(QColor.fromHsvF(yellow_hue, saturation, value)))
            elif match_quality == 1:
                return QVariant(QBrush(QColor.fromHsvF(red_hue, saturation, value)))
            else:
                return QVariant(QBrush(QColor.fromHsvF(white_hue, 0.0, value)))

        elif role == Qt.CheckStateRole and col == self.parent.ENABLED_COL:
            if self.arraydata[row][self.parent.ENABLED_COL].checkState():
                return QVariant(Qt.Checked)
            else:
                return QVariant(Qt.Unchecked)
        elif role == Qt.DisplayRole and col == self.parent.PROGRESS_COL:
            return self.arraydata[row][self.parent.PROGRESS_COL].text()
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
        elif role != Qt.DisplayRole:
            return QVariant()
        return QVariant(self.arraydata[index.row()][index.column()])

    def flags(self, index):
        if index.column() == self.parent.ENABLED_COL:
            return QAbstractItemModel.flags(self, index) | Qt.ItemIsUserCheckable
        else:
            return QAbstractItemModel.flags(self, index)

    def refresh(self, show_confidence_colors):
        self.show_confidence_colors = show_confidence_colors
        self.dataChanged.emit(self.createIndex(0,0),
                              self.createIndex(self.rowCount(0), self.columnCount(0)))

    def headerData(self, col, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return QVariant(self.headerdata[col])
        return QVariant()

    def setData(self, index, value, role):
        row, col = index.row(), index.column()
        if col == self.parent.ENABLED_COL:
            if self.arraydata[row][self.parent.ENABLED_COL].checkState():
                self.arraydata[row][self.parent.ENABLED_COL].setCheckState(False)
            else:
                self.arraydata[row][self.parent.ENABLED_COL].setCheckState(True)

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


class BookStatusDialog(SizePersistedDialog):
    '''
    '''
    # Location reporting template
    LOCATION_TEMPLATE = "{cls}:{func}({arg1}) {arg2}"

    CHECKMARK = u"\u2713"
    PROGRESS_READ = u"\u25AA"
    PROGRESS_UNREAD = u"\u25AB"

    if isosx:
        FONT = QFont('Monaco', 11)
    elif iswindows:
        FONT = QFont('Lucida Console', 9)
    elif islinux:
        FONT = QFont('Monospace', 9)
        FONT.setStyleHint(QFont.TypeWriter)

    #def __init__(self, parent):
    def initialize(self, parent):
        self.flags = {
            'new': 'NEW',
            'read': 'READ',
            'reading_list': 'READING LIST'
            }
        self.hash_cache = 'content_hashes.zip'
        self.opts = parent.opts
        self.parent = parent
        self.prefs = parent.opts.prefs
        self.library_title_map = None
        self.library_uuid_map = None
        self.local_cache_folder = self.parent.connected_device.temp_dir
        self.local_hash_cache = None
        self.remote_cache_folder = '/'.join(['/Library','calibre.mm'])
        self.remote_hash_cache = None
        self.show_confidence_colors = True
        self.verbose = parent.verbose
        self._log_location()

        self.installed_books = self._generate_booklist()
        self._construct_table_data()

        self.setWindowTitle(u'Marvin Library: %d books' % len(self.tabledata))
        self.setWindowIcon(self.opts.icon)
        self.l = QVBoxLayout(self)
        self.setLayout(self.l)
        self.perfect_width = 0

        self.tv = QTableView(self)
        self.l.addWidget(self.tv)
        self.library_header = [self.CHECKMARK, 'uuid', 'cid', 'mid', 'path',
                               'Title', 'Author', 'Progress',
                               'Last Opened', 'Word Count', 'Annotations',
                               'Collections', 'Deep View', 'Vocabulary',
                               'Match Quality']
        self.ENABLED_COL = 0
        self.UUID_COL = self.library_header.index('uuid')
        self.CALIBRE_ID_COL = self.library_header.index('cid')
        self.BOOK_ID_COL = self.library_header.index('mid')
        self.PATH_COL = self.library_header.index('path')
        self.TITLE_COL = self.library_header.index('Title')
        self.AUTHOR_COL = self.library_header.index('Author')
        self.PROGRESS_COL = self.library_header.index('Progress')
        self.LAST_OPENED_COL = self.library_header.index('Last Opened')
        self.WORD_COUNT_COL = self.library_header.index('Word Count')
        self.ANNOTATIONS_COL = self.library_header.index('Annotations')
        self.COLLECTIONS_COL = self.library_header.index('Collections')
        self.DEEP_VIEW_COL = self.library_header.index('Deep View')
        self.VOCABULARY_COL = self.library_header.index('Vocabulary')
        self.MATCHED_COL = self.library_header.index('Match Quality')

        hidden_columns =    [
                             self.UUID_COL,
                             self.CALIBRE_ID_COL,
                             self.BOOK_ID_COL,
                             self.PATH_COL,
                             self.MATCHED_COL,
                            ]
        centered_columns =  [
                             self.ANNOTATIONS_COL,
                             self.COLLECTIONS_COL,
                             self.DEEP_VIEW_COL,
                             self.LAST_OPENED_COL,
                             self.PROGRESS_COL,
                             self.VOCABULARY_COL,
                             ]
        right_aligned_columns = [
                             self.WORD_COUNT_COL
                             ]
        self.tm = MarkupTableModel(self, columns_to_center=centered_columns,
                                   right_aligned_columns=right_aligned_columns)

        self.tv.setModel(self.tm)
        self.tv.setShowGrid(False)
        self.tv.setFont(self.FONT)
        self.tvSelectionModel = self.tv.selectionModel()
        self.tv.setAlternatingRowColors(not self.show_confidence_colors)
        self.tv.setShowGrid(False)
        self.tv.setWordWrap(False)
        self.tv.setSelectionBehavior(self.tv.SelectRows)

        # Connect signals
        self.connect(self.tv, SIGNAL("doubleClicked(QModelIndex)"), self.double_clicked)
        self.connect(self.tv.horizontalHeader(), SIGNAL("sectionClicked(int)"), self.capture_sort_column)

        # Hide the vertical self.header
        self.tv.verticalHeader().setVisible(False)

        # Hide hidden columns
        for index in hidden_columns:
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

        # Set row height
        nrows = len(self.tabledata)
        for row in xrange(nrows):
            self.tv.setRowHeight(row, 16)

        self.tv.setSortingEnabled(True)

        sort_column = self.opts.prefs.get('marvin_library_sort_column',
                                          self.library_header.index('Match Quality'))
        sort_order = self.opts.prefs.get('marvin_library_sort_order',
                                         Qt.DescendingOrder)
        self.tv.sortByColumn(sort_column, sort_order)

        # ~~~~~~~~ Create the ButtonBox ~~~~~~~~
        self.dialogButtonBox = QDialogButtonBox(QDialogButtonBox.Help)

        self.delete_button = self.dialogButtonBox.addButton(self.dialogButtonBox.Discard)
        self.delete_button.setText('Delete')

        self.done_button = self.dialogButtonBox.addButton(self.dialogButtonBox.Ok)
        self.done_button.setText('Done')

        self.dialogButtonBox.setOrientation(Qt.Horizontal)
        self.dialogButtonBox.setCenterButtons(False)

        # Set All/Clear All
        self.toggle_checkmarks_button = self.dialogButtonBox.addButton(' Set All ', QDialogButtonBox.ActionRole)
        self.toggle_checkmarks_button.setObjectName('toggle_checkmarks_button')

        # Show/Hide Match Quality
        smq_text = 'Show Match Quality'
        if self.show_confidence_colors:
            smq_text = "Hide Match Quality"
        self.show_confidence_button = self.dialogButtonBox.addButton(smq_text, QDialogButtonBox.ActionRole)
        self.show_confidence_button.setObjectName('match_quality_button')

        # Word count
        self.wc_button = self.dialogButtonBox.addButton('Calculate word count', QDialogButtonBox.ActionRole)
        self.wc_button.setObjectName('calculate_word_count_button')

        # Synchronize collections
        self.sc_button = self.dialogButtonBox.addButton('Synchronize collections', QDialogButtonBox.ActionRole)
        self.sc_button.setObjectName('synchronize_collections_button')

        # Bind soft matches
        self.bsm_button = self.dialogButtonBox.addButton('Bind soft matches', QDialogButtonBox.ActionRole)
        self.bsm_button.setObjectName('bind_soft_matches_button')

        self.dialogButtonBox.clicked.connect(self.show_installed_books_dialog_clicked)
        self.l.addWidget(self.dialogButtonBox)

        self.resize_dialog()

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

    def double_clicked(self, index):
        '''
        Display column data for selected book
        '''
        row = index.row()
        col = index.column()
        column = self.library_header[col]
        cid = self.tm.arraydata[row][self.CALIBRE_ID_COL]
        path = self.tm.arraydata[row][self.PATH_COL]
        mid = self.tm.arraydata[row][self.BOOK_ID_COL]
        title = str(self.tm.arraydata[row][self.TITLE_COL].text())

        self._log_location("%s '%s' mid: %s cid: %s" % (title, column, mid, cid))

        if col == self.COLLECTIONS_COL:
            device_collections = self.installed_books[mid].collections
            if device_collections:
                msg = "Marvin: " + ', '.join(sorted(device_collections, key=sort_key))
            else:
                msg = "Marvin: No collections assigned"

            # Get calibre collection assignments
            library_collections = []
            if cid:
                cfl = self.prefs.get('collection_field_lookup', '')
                if cfl:
                    self._log("cfl: %s" % cfl)
                    db = self.opts.gui.current_db
                    mi = db.get_metadata(cid, index_is_id=True)
                    value = mi.get(cfl)
                    if value:
                        if type(value) is list:
                            self._log("value is list: %s" % repr(value))
                            msg += '\n' + "Calibre: " + ', '.join(value)
                        elif type(value) in [str, unicode]:
                            self._log("value is string/uni: %s" % repr(value))
                            msg += '\n' + "Calibre: " + value
                        else:
                            self._log("value is unexpected type: '%s'" % type(value))
                    else:
                        self._log("no value for '%s'" % cfl)
                        msg += '\n' + "Calibre: No collections assigned"
                else:
                    self._log("collection_field_lookup: %s" % repr(cfl))

            MessageBox(MessageBox.INFO, column, msg,
                           show_copy_button=False).exec_()

        elif col == self.VOCABULARY_COL:
            vocabulary = self.installed_books[mid].vocabulary
            if vocabulary:
                msg = ', '.join(sorted(vocabulary, key=sort_key))
            else:
                msg = ("<p>No vocabulary words.</p>")
            MessageBox(MessageBox.INFO, column, msg,
                           show_copy_button=False).exec_()

        elif col == self.WORD_COUNT_COL:
            selected_books = {}
            selected_books[cid] = {'title': title, 'path': path}
            self._calculate_word_count(selected_books=selected_books)

    def show_installed_books_dialog_clicked(self, button):
        '''
        BUTTON_ROLES = ['AcceptRole', 'RejectRole', 'DestructiveRole', 'ActionRole',
                        'HelpRole', 'YesRole', 'NoRole', 'ApplyRole', 'ResetRole']
        '''
        self._log_location()
        if self.dialogButtonBox.buttonRole(button) == QDialogButtonBox.AcceptRole:
            self._log("AcceptRole")
            self.accept()
        elif self.dialogButtonBox.buttonRole(button) == QDialogButtonBox.ActionRole:
            if button.objectName() == 'match_quality_button':
                self.toggle_confidence_colors()
            elif button.objectName() == 'toggle_checkmarks_button':
                self.toggle_checkmarks()
            elif button.objectName() == 'calculate_word_count_button':
                self._calculate_word_count()
            elif button.objectName() == 'synchronize_collections_button':
                self._synchronize_collections()
            elif button.objectName() == 'bind_soft_matches_button':
                self._bind_soft_matches()

        elif self.dialogButtonBox.buttonRole(button) == QDialogButtonBox.DestructiveRole:
            self._delete_books()
        elif self.dialogButtonBox.buttonRole(button) == QDialogButtonBox.HelpRole:
            self.show_help()
        elif self.dialogButtonBox.buttonRole(button) == QDialogButtonBox.RejectRole:
            self.close()

    def show_help(self):
        '''
        Display help file
        '''
        self.parent.show_help()

    def size_hint(self):
        return QtCore.QSize(self.perfect_width, self.height())

    def toggle_checkmarks(self):
        button_text = str(self.toggle_checkmarks_button.text())
        if button_text == 'Clear All':
            for i in range(len(self.tabledata)):
                self.tm.arraydata[i][self.ENABLED_COL].setCheckState(False)
            self.toggle_checkmarks_button.setText(' Set All ')
        else:
            for i in range(len(self.tabledata)):
                self.tm.arraydata[i][self.ENABLED_COL].setCheckState(True)
            self.toggle_checkmarks_button.setText('Clear All')
        self.repaint()

    def toggle_confidence_colors(self):
        self.show_confidence_colors = not self.show_confidence_colors
        self.opts.prefs.set('annotated_books_dialog_show_confidence_as_bg_color', self.show_confidence_colors)
        if self.show_confidence_colors:
            self.show_confidence_button.setText("Hide Match Quality")
            self.tv.sortByColumn(self.library_header.index('Match Quality'), Qt.DescendingOrder)
            self.capture_sort_column(self.library_header.index('Match Quality'))
        else:
            self.show_confidence_button.setText("Show Match Quality")
        self.tv.setAlternatingRowColors(not self.show_confidence_colors)
        self.tm.refresh(self.show_confidence_colors)

    # Helpers
    def _bind_soft_matches(self):
        '''
        '''
        self._log_location()
        title = "Bind soft matches"
        msg = ("<p>Not implemented</p>")
        MessageBox(MessageBox.INFO, title, msg,
                       show_copy_button=False).exec_()

    def _calculate_word_count(self, selected_books=[]):
        '''
        Calculate word count for each selected book
        '''
        def _extract_body_text(data):
            '''
            Get the body text of this html content with any html tags stripped
            '''
            RE_HTML_BODY = re.compile(u'<body[^>]*>(.*)</body>', re.UNICODE | re.DOTALL | re.IGNORECASE)
            RE_STRIP_MARKUP = re.compile(u'<[^>]+>', re.UNICODE)

            body = RE_HTML_BODY.findall(data)
            if body:
                return RE_STRIP_MARKUP.sub('', body[0]).replace('.','. ')
            return ''

        self._log_location()

        if not selected_books:
            selected_books = self._get_selected_books()

        if selected_books:
            stats = {}

            pb = ProgressBar(parent=self.opts.gui, window_title="Calculating word counts", on_top=True)
            total_books = len(selected_books)
            pb.set_maximum(total_books)
            pb.set_value(0)
            pb.set_label('{:^100}'.format("1 of %d" % (total_books)))
            pb.show()

            for i, cid in enumerate(selected_books):
                pb.set_label('{:^100}'.format(selected_books[cid]['title']))

                # Copy the remote epub to local storage
                path = selected_books[cid]['path']
                rbp = '/'.join(['/Documents', path])
                lbp = os.path.join(self.local_cache_folder, path)

                # Set the driver busy flag, copy the file
                self._wait_for_driver_not_busy()
                self.parent.connected_device.busy = True
                with open(lbp, 'wb') as out:
                    self.parent.ios.copy_from_idevice(str(rbp), out)
                self.parent.connected_device.busy = False

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

                self._log("%s: %d words" % (selected_books[cid]['title'], wordcount.words))
                stats[selected_books[cid]['title']] = wordcount.words

                # Delete the local copy
                os.remove(lbp)
                pb.increment()

                # Update Marvin db

            pb.hide()

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

    def _compute_epub_hash(self, zipfile):
        '''
        Generate a hash of all *.*html files names, sizes
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
                if item.get('media-type') == 'application/xhtml+xml':
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
        self._log_location()

        self.tabledata = []

        for book in self.installed_books:
            book_data = self.installed_books[book]
            enabled = QCheckBox()
            enabled.setChecked(False)

            # last_opened sorts by timestamp
            last_opened_ts = ''
            last_opened_sort = 0
            if book_data.date_opened:
                last_opened_ts = time.strftime("%Y-%m-%d",
                                               time.localtime(book_data.date_opened))
                last_opened_sort = book_data.date_opened
            last_opened = SortableTableWidgetItem(
                last_opened_ts,
                last_opened_sort)

            # title, author sort by title_sort, author_sort
            if not book_data.title_sort:
                book_data.title_sort = book_data.title_sort()
            title = SortableTableWidgetItem(
                book_data.title,
                book_data.title_sort.upper())

            if not book_data.author_sort:
                book_data.author_sort = ', '.join(book_data.author)
            author = SortableTableWidgetItem(
                ', '.join(book_data.author),
                book_data.author_sort.upper())

            # Reading progress
            percent_read = ''
            if book_data.progress > 0.01:
                if self.opts.prefs.get('show_progress_as_percentage', False):
                    percent_read = "{:3.0f}%".format(book_data.progress * 100)
                else:
                    if book_data.progress < 0.25:
                        percent_read = (1 * self.PROGRESS_READ) + (4 * self.PROGRESS_UNREAD)
                    elif book_data.progress >= 0.25 and book_data.progress < 0.50:
                        percent_read = (2 * self.PROGRESS_READ) + (3 * self.PROGRESS_UNREAD)
                    elif book_data.progress >= 0.50 and book_data.progress < 0.75:
                        percent_read = (3 * self.PROGRESS_READ) + (2 * self.PROGRESS_UNREAD)
                    elif book_data.progress >= 0.75 and book_data.progress < 0.95:
                        percent_read = (4 * self.PROGRESS_READ) + (1 * self.PROGRESS_UNREAD)
                    else:
                        percent_read = (5 * self.PROGRESS_READ)
            progress = SortableTableWidgetItem(
                percent_read,
                book_data.progress)

            # Match quality
            if self.opts.prefs.get('development_mode', False):
                self._log("%s uuid: %s matches: %s on_device: %s" %
                            (book_data.title,
                             repr(book_data.uuid),
                             repr(book_data.matches),
                             repr(book_data.on_device)))

            match_quality = 0
            if (book_data.uuid > '' and
                [book_data.uuid] == book_data.matches):
                # Exact uuid match
                match_quality = 3
            elif (book_data.uuid > '' and
                  book_data.uuid in book_data.matches):
                # Duplicates
                match_quality = 1
            elif book_data.on_device == 'Main':
                # Soft match
                match_quality = 2
            if self.opts.prefs.get('development_mode', False):
                self._log("match_quality: %s" % match_quality)

            # List order matches self.library_header
            this_book = [
                enabled,
                book_data.uuid,
                book_data.cid,
                book_data.mid,
                book_data.path,
                title,
                author,
                progress,
                last_opened,
                book_data.word_count if book_data.word_count > '0' else '',
                self.CHECKMARK if book_data.has_highlights else '',
                self.CHECKMARK if len(book_data.collections) else '',
                self.CHECKMARK if book_data.deep_view_prepared else '',
                self.CHECKMARK if len(book_data.vocabulary) else '',
                match_quality
                ]
            self.tabledata.append(this_book)

    def _delete_books(self):
        '''
        '''
        self._log_location()
        title = "Delete books"
        msg = ("<p>Not implemented</p>")
        MessageBox(MessageBox.INFO, title, msg,
                       show_copy_button=False).exec_()

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

        # Set the driver busy flag, copy the file
        self._wait_for_driver_not_busy()
        self.parent.connected_device.busy = True
        with open(lbp, 'wb') as out:
            self.parent.ios.copy_from_idevice(str(rbp), out)
        self.parent.connected_device.busy = False

        hash = self._compute_epub_hash(lbp)
        zfw.writestr(path, hash)
        zfw.close()

        # Delete the local copy
        os.remove(lbp)
        return hash

    def _find_fuzzy_matches(self, library_scanner, installed_books):
        '''
        Compare computed hashes of installed books to library books.
        Look for potential dupes
        '''
        self._log_location()

        library_hash_map = library_scanner.hash_map
        hard_matches = {}
        soft_matches = []
        for book in installed_books:
            #self._log("evaluating %s hash: %s uuid: %s" % (mb.title, mb.hash, mb.uuid))
            mb = installed_books[book]
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
            mb.matches = uuids

        # Review the soft matches against the hard matches
        if soft_matches:
            # Scan soft matches against hard matches for hash collision
            for mb in soft_matches:
                if mb.hash in hard_matches:
                    mb.matches += hard_matches[mb.hash].matches

    def _generate_booklist(self):
        '''
        '''
        self._log_location()

        # Scan library books for hashes
        if self.parent.library_scanner.isRunning():
            self.library_scanner.wait()

        # Save a reference to the title, uuid map
        self.library_title_map = self.parent.library_scanner.title_map
        self.library_uuid_map = self.parent.library_scanner.uuid_map

        # Get the library hash_map
        library_hash_map = self.parent.library_scanner.hash_map
        if library_hash_map is None:
            library_hash_map = self._scan_library_books(self.parent.library_scanner)
        else:
            self._log("hash_map already generated")

        # Scan Marvin
        installed_books = self._get_installed_books()

        # Update installed_books with library matches
        self._find_fuzzy_matches(self.parent.library_scanner, installed_books)

        return installed_books

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

        {mid: Book, ...}
        '''
        def _get_calibre_id(uuid, title, author):
            '''
            Find book in library, return cid
            '''
            if self.opts.prefs.get('development_mode', False):
                self._log_location("%s %s" % (repr(title), repr(author)))
            ans = None
            db = self.opts.gui.current_db
            if uuid in self.library_uuid_map:
                ans = self.library_uuid_map[uuid]['id']
                if self.opts.prefs.get('development_mode', False):
                    self._log("UUID match")
            elif title in self.library_title_map:
                cid = self.library_title_map[title]['id']
                mi = db.get_metadata(cid, index_is_id=True)
                authors = author.split(', ')
                if authors == mi.authors:
                    ans = cid
                    if self.opts.prefs.get('development_mode', False):
                        self._log("TITLE/AUTHOR match")
            return ans

        def _get_collections(cur, book_id, row):
            # Get the collection assignments
            ca_cur = con.cursor()
            ca_cur.execute('''SELECT
                                BookID,
                                CollectionID
                              FROM BookCollections
                              WHERE BookID = '{0}'
                           '''.format(book_id))
            collections = []
#             if row[b'NewFlag']:
#                 collections.append(self.flags['new'])
#             if row[b'ReadingList']:
#                 collections.append(self.flags['reading_list'])
#             if row[b'IsRead']:
#                 collections.append(self.flags['read'])

            collection_rows = ca_cur.fetchall()
            if collection_rows is not None:
                collection_assignments = [collection[b'CollectionID']
                                          for collection in collection_rows]
                collections += [collection_map[item] for item in collection_assignments]
                collections = sorted(collections, key=sort_key)
            ca_cur.close()
            return collections

        def _get_highlights(cur, book_id):
            '''
            Test for existing highlights in book_id
            '''
            hl_cur = con.cursor()
            hl_cur.execute('''SELECT
                                BookID
                              FROM Highlights
                              WHERE BookID = '{0}'
                           '''.format(book_id))
            hl_rows = hl_cur.fetchall()
            if len(hl_rows):
                return True
            else:
                return False

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
                            CalibreTitleSort,
                            DateOpened,
                            DeepViewPrepared,
                            FileName,
                            IsRead,
                            NewFlag,
                            Progress,
                            ReadingList,
                            Title,
                            UUID,
                            WordCount
                          FROM Books
                        ''')

            rows = cur.fetchall()
            book_count = len(rows)
            for i, row in enumerate(rows):
                book_id = row[b'id_']

                # Get the primary metadata from Books
                this_book = Book(row[b'Title'], row[b'Author'])
                this_book.author_sort = row[b'AuthorSort']
                this_book.cid = _get_calibre_id(row[b'UUID'],
                                                row[b'Title'],
                                                row[b'Author'])
                this_book.collections = _get_collections(cur, book_id, row)
                this_book.date_opened = row[b'DateOpened']
                this_book.deep_view_prepared = row[b'DeepViewPrepared']
                this_book.hash = hashes[row[b'FileName']]['hash']
                this_book.has_highlights = _get_highlights(cur, book_id)
                this_book.mid = book_id
                this_book.on_device = _get_on_device_status(this_book.cid)
                this_book.path = row[b'FileName']
                this_book.progress = row[b'Progress']
                this_book.title_sort = row[b'CalibreTitleSort']
                this_book.uuid = row[b'UUID']
                this_book.vocabulary = _get_vocabulary_list(cur, book_id)
                this_book.word_count = locale.format("%d", row[b'WordCount'], grouping=True)
                installed_books[book_id] = this_book

        if self.opts.prefs.get('development_mode', False):
            self._log("%d cached books from Marvin:" % len(cached_books))
            for book in installed_books:
                self._log("%s word_count: %s" % (installed_books[book].title,
                                                 repr(installed_books[book].word_count)))
        return installed_books

    def _get_selected_books(self):
        '''
        Generate a dict of books selected in the dialog
        '''
        selected_books = {}
        for i in range(len(self.tabledata)):
            self.tv.selectRow(i)
            enabled = bool(self.tm.arraydata[i][self.ENABLED_COL].checkState())
            if enabled:
                cid = self.tm.arraydata[i][self.library_header.index('cid')]
                path = self.tm.arraydata[i][self.library_header.index('path')]
                title = str(self.tm.arraydata[i][self.library_header.index('Title')].text())
                selected_books[cid] = {'title': title, 'path': path}
        return selected_books

    def _localize_hash_cache(self, cached_books):
        '''
        Check for existence of hash cache on iDevice. Confirm/create folder
        If existing cached, purge orphans
        '''
        self._log_location()

        # Set the driver busy flag
        self._wait_for_driver_not_busy()
        self.parent.connected_device.busy = True

        # Existing hash cache?
        lhc = os.path.join(self.local_cache_folder, self.hash_cache)
        rhc = '/'.join([self.remote_cache_folder, self.hash_cache])

        cache_exists = (self.parent.ios.exists(rhc) and
                        not self.parent.marvin_content_invalid and
                        not self.opts.prefs.get('hash_caching_disabled'))
        if cache_exists:
            # Copy from existing remote cache to local cache
            self._log("copying remote hash cache")
            with open(lhc, 'wb') as out:
                self.parent.ios.copy_from_idevice(str(rhc), out)
        else:
            # Confirm path to remote folder is valid store point
            folder_exists = self.parent.ios.exists(self.remote_cache_folder)
            if not folder_exists:
                self._log("creating remote_cache_folder %s" % repr(self.remote_cache_folder))
                self.parent.ios.mkdir(self.remote_cache_folder)
            else:
                self._log("remote_cache_folder exists")

            # Create a local cache
            self._log("creating new local hash cache: %s" % repr(lhc))
            zfw = ZipFile(lhc, mode='w')
            zfw.writestr('Marvin hash cache', '')
            zfw.close()

            # Clear the marvin_content_invalid flag
            if self.parent.marvin_content_invalid:
                self.parent.marvin_content_invalid = False

        self.local_hash_cache = lhc
        self.remote_hash_cache = rhc

        # Clear the driver busy flag
        self.parent.connected_device.busy = False

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
        widths = []
        for (i, c) in enumerate(self.library_header):
            widths.append(self.tv.columnWidth(i))
        self.opts.prefs.set('marvin_library_column_widths', widths)

    def _scan_library_books(self, library_scanner):
        '''
        Generate hashes for library epubs
        '''
        # Scan library books for hashes
        if self.parent.library_scanner.isRunning():
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

        for i, uuid in enumerate(uuid_map):
            pb.set_label('{:^100}'.format("%d of %d" % (i+1, total_books)))

            path = db.format(uuid_map[uuid]['id'], 'epub', index_is_id=True,
                             as_path=True, preserve_filename=True)
            uuid_map[uuid]['hash'] = self._compute_epub_hash(path)
            os.remove(path)
            pb.increment()

        hash_map = library_scanner.build_hash_map()
        pb.hide()

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

        installed_books = {}
        for i, path in enumerate(cached_books):
            this_book = {}
            pb.set_label('{:^100}'.format("%d of %d" % (i+1, total_books)))
            this_book['hash'] = self._fetch_marvin_content_hash(path)

            installed_books[path] = this_book
            pb.increment()
        pb.hide()

        # Push the local hash to the iDevice
        self._update_remote_hash_cache()

        return installed_books

    def _synchronize_collections(self):
        '''
        '''
        self._log_location()
        title = "Synchronize collections"
        msg = ("<p>Not implemented</p>")
        MessageBox(MessageBox.INFO, title, msg,
                       show_copy_button=False).exec_()

    def _update_remote_hash_cache(self):
        '''
        Copy updated hash cache to iDevice
        self.local_hash_cache, self.remote_hash_cache initialized
        in _localize_hash_cache()
        '''
        self._log_location()

        # Set the driver busy flag
        self._wait_for_driver_not_busy()
        self.parent.connected_device.busy = True

        if self.parent.prefs.get('hash_caching_disabled', False):
            self._log("hash_caching_disabled, deleting remote hash cache")
            self.parent.ios.remove(str(self.remote_hash_cache))
        else:
            # Copy local cache to iDevice
            self.parent.ios.copy_to_idevice(self.local_hash_cache, str(self.remote_hash_cache))

        # Clear the driver busy flag
        self.parent.connected_device.busy = False

    def _wait_for_driver_not_busy(self):
        '''
        Wait for driver to finish any existing I/O
        '''
        if self.opts.prefs.get('development_mode', False):
            self._log_location()
        if self.parent.connected_device.busy:
            if self.opts.prefs.get('development_mode', False):
                self._log("waiting for busy device")
            while True:
                time.sleep(0.05)
                if not self.parent.connected_device.busy:
                    break
