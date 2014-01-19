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
from calibre.utils.magick.draw import add_borders_to_image, thumbnail

from calibre_plugins.marvin_manager.book_status import dialog_resources_path
from calibre_plugins.marvin_manager.common_utils import Logger, SizePersistedDialog

from PyQt4.Qt import (Qt, QBrush, QColor, QDialogButtonBox, QIcon, QImage,
                      QPainter, QPalette, QPixmap, QPushButton, QSize, QSizePolicy,
                      pyqtSignal)

# Import Ui_Form from form generated dynamically during initialization
if True:
    sys.path.insert(0, dialog_resources_path)
    from view_metadata_ui import Ui_Dialog
    sys.path.remove(dialog_resources_path)


class MetadataComparisonDialog(SizePersistedDialog, Ui_Dialog, Logger):
    BORDER_COLOR = "#FDFF99"
    BORDER_WIDTH = 4
    COVER_ICON_SIZE = 200

    marvin_device_status_changed = pyqtSignal(dict)

    def accept(self):
        self._log_location()
        super(MetadataComparisonDialog, self).accept()

    def close(self):
        self._log_location()
        super(MetadataComparisonDialog, self).close()

    def dispatch_button_click(self, button):
        '''
        BUTTON_ROLES = ['AcceptRole', 'RejectRole', 'DestructiveRole', 'ActionRole',
                        'HelpRole', 'YesRole', 'NoRole', 'ApplyRole', 'ResetRole']
        '''
        self._log_location()
        if self.bb.buttonRole(button) == QDialogButtonBox.AcceptRole:
            self._log("AcceptRole")
            self.accept()
        elif self.bb.buttonRole(button) == QDialogButtonBox.RejectRole:
            self.close()

    def esc(self, *args):
        self.close()

    def initialize(self, parent, book_id, cid, installed_book, enable_metadata_updates, marvin_db_path):
        '''
        __init__ is called on SizePersistedDialog()
        shared attributes of interest:
            .authors
            .author_sort
            .cover_hash
            .pubdate
            .publisher
            .series
            .series_index
            .title
            .title_sort
            .comments
            .tags
            .uuid
        '''
        self.setupUi(self)
        self.book_id = book_id
        self.cid = cid
        self.connected_device = parent.opts.gui.device_manager.device
        self.installed_book = installed_book
        self.marvin_db_path = marvin_db_path
        self.opts = parent.opts
        self.parent = parent
        self.stored_command = None
        self.verbose = parent.verbose
        self.BORDER_LR = 4
        self.BORDER_TB = 8
        self.GREY_FG = '<font style="color:#A0A0A0">{0}</font>'
        self.YELLOW_BG = '<font style="background:#FDFF99">{0}</font>'

        self._log_location(installed_book.title)

        # Subscribe to Marvin driver change events
        self.connected_device.marvin_device_signals.reader_app_status_changed.connect(
            self.marvin_status_changed)

        #self._log("mismatches:\n%s" % repr(installed_book.metadata_mismatches))
        self.mismatches = installed_book.metadata_mismatches

        self._populate_title()
        self._populate_title_sort()
        self._populate_series()
        self._populate_authors()
        self._populate_author_sort()
        self._populate_uuid()
        self._populate_covers()
        self._populate_subjects()
        self._populate_publisher()
        self._populate_pubdate()
        self._populate_description()

        # ~~~~~~~~ Export to Marvin button ~~~~~~~~
        self.export_to_marvin_button.setIcon(QIcon(os.path.join(self.parent.opts.resources_path,
                                                   'icons',
                                                   'from_calibre.png')))
        self.export_to_marvin_button.clicked.connect(partial(self.store_command, 'export_metadata'))
        self.export_to_marvin_button.setEnabled(enable_metadata_updates)

        # ~~~~~~~~ Import from Marvin button ~~~~~~~~
        self.import_from_marvin_button.setIcon(QIcon(os.path.join(self.parent.opts.resources_path,
                                                     'icons',
                                                     'from_marvin.png')))
        self.import_from_marvin_button.clicked.connect(partial(self.store_command, 'import_metadata'))
        self.import_from_marvin_button.setEnabled(enable_metadata_updates)

        # If no calibre book, or no mismatches, adjust the display accordingly
        if not self.cid:
            #self._log("self.cid: %s" % repr(self.cid))
            #self._log("self.mismatches: %s" % repr(self.mismatches))
            self.calibre_gb.setVisible(False)
            self.import_from_marvin_button.setVisible(False)
            self.setWindowTitle(u'Marvin metadata')
        elif not self.mismatches:
            # Show both panels, but hide the transfer buttons
            self.export_to_marvin_button.setVisible(False)
            self.import_from_marvin_button.setVisible(False)
        else:
            self.setWindowTitle(u'Metadata Summary')

        if False:
            # Set the Marvin QGroupBox to Marvin red
            marvin_red = QColor()
            marvin_red.setRgb(189, 17, 20, alpha=255)
            palette = QPalette()
            palette.setColor(QPalette.Background, marvin_red)
            self.marvin_gb.setPalette(palette)

        # ~~~~~~~~ Add a Close or Cancel button ~~~~~~~~
        self.close_button = QPushButton(QIcon(I('window-close.png')), 'Close')
        if self.mismatches:
            self.close_button.setText('Cancel')
        self.bb.addButton(self.close_button, QDialogButtonBox.RejectRole)

        self.bb.clicked.connect(self.dispatch_button_click)

        # Restore position
        self.resize_dialog()

    def marvin_status_changed(self, cmd_dict):
        '''

        '''
        self.marvin_device_status_changed.emit(cmd_dict)
        command = cmd_dict['cmd']

        self._log_location(command)

        if command in ['disconnected', 'yanked']:
            self._log("closing dialog: %s" % command)
            self.close()

    def store_command(self, command):
        '''
        '''
        self._log_location(command)
        self.stored_command = command
        self.accept()

    def _populate_authors(self):
        if 'authors' in self.mismatches:
            cs_authors = ', '.join(self.mismatches['authors']['calibre'])
            self.calibre_authors.setText(self.YELLOW_BG.format(cs_authors))
            ms_authors = ', '.join(self.mismatches['authors']['Marvin'])
            self.marvin_authors.setText(self.YELLOW_BG.format(ms_authors))
        else:
            authors = ', '.join(self.installed_book.authors)
            self.calibre_authors.setText(authors)
            self.marvin_authors.setText(authors)

    def _populate_author_sort(self):
        if 'author_sort' in self.mismatches:
            cs_author_sort = self.mismatches['author_sort']['calibre']
            self.calibre_author_sort.setText(self.YELLOW_BG.format(cs_author_sort))
            ms_author_sort = self.mismatches['author_sort']['Marvin']
            self.marvin_author_sort.setText(self.YELLOW_BG.format(ms_author_sort))
        else:
            author_sort = self.installed_book.author_sort
            self.calibre_author_sort.setText(self.GREY_FG.format(author_sort))
            self.marvin_author_sort.setText(self.GREY_FG.format(author_sort))

    def _populate_covers(self):
        '''
        Display calibre cover for both unless mismatch
        '''
        def _fetch_marvin_cover(with_border=False):
            '''
            Retrieve LargeCoverJpg from cache
            '''
            self._log_location()
            con = sqlite3.connect(self.marvin_db_path)
            with con:
                con.row_factory = sqlite3.Row

                # Fetch Hash from mainDb
                cover_cur = con.cursor()
                cover_cur.execute('''SELECT
                                      Hash
                                     FROM Books
                                     WHERE ID = '{0}'
                                  '''.format(self.book_id))
                row = cover_cur.fetchone()

            book_hash = row[b'Hash']
            large_covers_subpath = self.connected_device._cover_subpath(size="large")
            cover_path = '/'.join([large_covers_subpath, '%s.jpg' % book_hash])
            stats = self.parent.ios.exists(cover_path)
            if stats:
                self._log("fetching large cover from cache")
                #self._log("cover size: {:,} bytes".format(int(stats['st_size'])))
                cover_bytes = self.parent.ios.read(cover_path, mode='rb')
                m_image = QImage()
                m_image.loadFromData(cover_bytes)

                if with_border:
                    m_image = m_image.scaledToHeight(
                        self.COVER_ICON_SIZE - self.BORDER_WIDTH * 2,
                        Qt.SmoothTransformation)

                    # Construct a QPixmap with yellow background
                    self.m_pixmap = QPixmap(
                        QSize(m_image.width() + self.BORDER_WIDTH * 2,
                              m_image.height() + self.BORDER_WIDTH * 2))

                    m_painter = QPainter(self.m_pixmap)
                    m_painter.setRenderHints(m_painter.Antialiasing)

                    m_painter.fillRect(self.m_pixmap.rect(), QColor(0xFD, 0xFF, 0x99))
                    m_painter.drawImage(self.BORDER_WIDTH, self.BORDER_WIDTH, m_image)
                else:
                    m_image = m_image.scaledToHeight(
                        self.COVER_ICON_SIZE,
                        Qt.SmoothTransformation)

                    self.m_pixmap = QPixmap(
                        QSize(m_image.width(),
                              m_image.height()))

                    m_painter = QPainter(self.m_pixmap)
                    m_painter.setRenderHints(m_painter.Antialiasing)

                    m_painter.drawImage(0, 0, m_image)

                self.marvin_cover.setPixmap(self.m_pixmap)
            else:
                # No cover available, use generic
                self._log("No cached cover, using generic")
                pixmap = QPixmap()
                pixmap.load(I('book.png'))
                pixmap = pixmap.scaled(self.COVER_ICON_SIZE,
                                       self.COVER_ICON_SIZE,
                                       aspectRatioMode=Qt.KeepAspectRatio,
                                       transformMode=Qt.SmoothTransformation)
                self.marvin_cover.setPixmap(pixmap)

        self.calibre_cover.setMaximumSize(QSize(self.COVER_ICON_SIZE, self.COVER_ICON_SIZE))
        self.calibre_cover.setText('')
        self.calibre_cover.setScaledContents(False)

        self.marvin_cover.setMaximumSize(QSize(self.COVER_ICON_SIZE, self.COVER_ICON_SIZE))
        self.marvin_cover.setText('')
        self.marvin_cover.setScaledContents(False)

        if self.cid:
            db = self.opts.gui.current_db
            if 'cover_hash' not in self.mismatches:
                mi = db.get_metadata(self.cid, index_is_id=True, get_cover=True, cover_as_data=True)

                c_image = QImage()
                if mi.has_cover:
                    c_image.loadFromData(mi.cover_data[1])
                    c_image = c_image.scaledToHeight(self.COVER_ICON_SIZE,
                                                     Qt.SmoothTransformation)
                    self.c_pixmap = QPixmap(QSize(c_image.width(),
                                                  c_image.height()))
                    c_painter = QPainter(self.c_pixmap)
                    c_painter.setRenderHints(c_painter.Antialiasing)
                    c_painter.drawImage(0, 0, c_image)
                else:
                    c_image.load(I('book.png'))
                    c_image = c_image.scaledToWidth(135,
                                                    Qt.SmoothTransformation)
                    # Construct a QPixmap with dialog background
                    self.c_pixmap = QPixmap(
                        QSize(c_image.width(),
                              c_image.height()))
                    c_painter = QPainter(self.c_pixmap)
                    c_painter.setRenderHints(c_painter.Antialiasing)
                    bgcolor = self.palette().color(QPalette.Background)
                    c_painter.fillRect(self.c_pixmap.rect(), bgcolor)
                    c_painter.drawImage(0, 0, c_image)

                # Set calibre cover
                self.calibre_cover.setPixmap(self.c_pixmap)

                if self.opts.prefs.get('development_mode', False):
                    # Show individual covers
                    _fetch_marvin_cover()
                else:
                    # Show calibre cover on both sides
                    self.marvin_cover.setPixmap(self.c_pixmap)

            else:
                # Covers don't match - render with border
                # Construct a QImage with the cover sized to fit inside border
                c_image = QImage()
                cdata = db.cover(self.cid, index_is_id=True)
                if cdata is None:
                    c_image.load(I('book.png'))
                else:
                    c_image.loadFromData(cdata)

                c_image = c_image.scaledToHeight(
                    self.COVER_ICON_SIZE - self.BORDER_WIDTH * 2,
                    Qt.SmoothTransformation)

                # Construct a QPixmap with yellow background
                self.c_pixmap = QPixmap(
                    QSize(c_image.width() + self.BORDER_WIDTH * 2,
                          c_image.height() + self.BORDER_WIDTH * 2))
                c_painter = QPainter(self.c_pixmap)
                c_painter.setRenderHints(c_painter.Antialiasing)
                c_painter.fillRect(self.c_pixmap.rect(), QColor(0xFD, 0xFF, 0x99))
                c_painter.drawImage(self.BORDER_WIDTH, self.BORDER_WIDTH, c_image)
                self.calibre_cover.setPixmap(self.c_pixmap)
                _fetch_marvin_cover(with_border=True)
        else:
            _fetch_marvin_cover()

    def _populate_description(self):
        # Set the bg color of the description text fields to the dialog bg color
        bgcolor = self.palette().color(QPalette.Background)
        palette = QPalette()
        palette.setColor(QPalette.Base, bgcolor)
        self.calibre_description.setPalette(palette)
        self.marvin_description.setPalette(palette)

        if 'comments' in self.mismatches:
            self.calibre_description_label.setText(self.YELLOW_BG.format("<b>Description</b>"))
            if self.mismatches['comments']['calibre']:
                self.calibre_description.setText(self.mismatches['comments']['calibre'])

            self.marvin_description_label.setText(self.YELLOW_BG.format("<b>Description</b>"))
            if self.mismatches['comments']['Marvin']:
                self.marvin_description.setText(self.mismatches['comments']['Marvin'])
        else:
            if self.installed_book.comments:
                self.calibre_description.setText(self.installed_book.comments)
                self.marvin_description.setText(self.installed_book.comments)

    def _populate_pubdate(self):
        if 'pubdate' in self.mismatches:
            if self.mismatches['pubdate']['calibre']:
                cs_pubdate = "<b>Published:</b> {0}".format(strftime("%d %B %Y", t=self.mismatches['pubdate']['calibre']))
            else:
                cs_pubdate = "<b>Published:</b> Date unknown"
            self.calibre_pubdate.setText(self.YELLOW_BG.format(cs_pubdate))

            if self.mismatches['pubdate']['Marvin']:
                ms_pubdate = "<b>Published:</b> {0}".format(strftime("%d %B %Y", t=self.mismatches['pubdate']['Marvin']))
            else:
                ms_pubdate = "<b>Published:</b> Date unknown"
            self.marvin_pubdate.setText(self.YELLOW_BG.format(ms_pubdate))
        elif self.installed_book.pubdate:
            pubdate = "<b>Published:</b> {0}".format(strftime("%d %B %Y", t=self.installed_book.pubdate))
            self.calibre_pubdate.setText(pubdate)
            self.marvin_pubdate.setText(pubdate)
        else:
            pubdate = "<b>Published:</b> Date unknown"
            self.calibre_pubdate.setText(pubdate)
            self.marvin_pubdate.setText(pubdate)

    def _populate_publisher(self):
        if 'publisher' in self.mismatches:
            csp = self.mismatches['publisher']['calibre']
            if not csp:
                cs_publisher = "<b>Publisher:</b> Unknown"
            else:
                cs_publisher = "<b>Publisher:</b> {0}".format(csp)
            self.calibre_publisher.setText(self.YELLOW_BG.format(cs_publisher))

            msp = self.mismatches['publisher']['Marvin']
            if not msp:
                ms_publisher = "<b>Publisher:</b> Unknown"
            else:
                ms_publisher = "<b>Publisher:</b> {0}".format(msp)
            self.marvin_publisher.setText(self.YELLOW_BG.format(ms_publisher))
        else:
            if not self.installed_book.publisher:
                publisher = "<b>Publisher:</b> Unknown"
            else:
                publisher = "<b>Publisher:</b> {0}".format(self.installed_book.publisher)
            self.calibre_publisher.setText(publisher)
            self.marvin_publisher.setText(publisher)

    def _populate_series(self):
        if 'series' in self.mismatches:
            cs_index = str(self.mismatches['series_index']['calibre'])
            if cs_index.endswith('.0'):
                cs_index = cs_index[:-2]
            cs = "%s (%s)" % (self.mismatches['series']['calibre'], cs_index)
            self.calibre_series.setText(self.YELLOW_BG.format(cs))
            ms_index = str(self.mismatches['series_index']['Marvin'])
            if ms_index.endswith('.0'):
                ms_index = ms_index[:-2]
            ms = "%s (%s)" % (self.mismatches['series']['Marvin'], ms_index)
            self.marvin_series.setText(self.YELLOW_BG.format(ms))
        elif self.installed_book.series:
            cs_index = str(self.installed_book.series_index)
            if cs_index.endswith('.0'):
                cs_index = cs_index[:-2]
            cs = "%s (%s)" % (self.installed_book.series, cs_index)
            self.calibre_series.setText(cs)
            self.marvin_series.setText(cs)
        else:
            self.calibre_series.setVisible(False)
            self.marvin_series.setVisible(False)

    def _populate_subjects(self):
        '''
        '''
        # Setting size policy allows us to match Subjects fields height
        sp = QSizePolicy()
        sp.setHorizontalStretch(True)
        sp.setVerticalStretch(False)
        sp.setHeightForWidth(False)
        self.calibre_subjects.setSizePolicy(sp)
        self.marvin_subjects.setSizePolicy(sp)

        if 'tags' in self.mismatches:
            cs = "<b>Subjects:</b> {0}".format(', '.join(self.mismatches['tags']['calibre']))
            self.calibre_subjects.setText(self.YELLOW_BG.format(cs))
            ms = "<b>Subjects:</b> {0}".format(', '.join(self.mismatches['tags']['Marvin']))
            self.marvin_subjects.setText(self.YELLOW_BG.format(ms))

            calibre_height = self.calibre_subjects.sizeHint().height()
            marvin_height = self.marvin_subjects.sizeHint().height()
            if calibre_height > marvin_height:
                self.marvin_subjects.setMinimumHeight(calibre_height)
                self.marvin_subjects.setMaximumHeight(calibre_height)
            elif marvin_height > calibre_height:
                self.calibre_subjects.setMinimumHeight(marvin_height)
                self.calibre_subjects.setMaximumHeight(marvin_height)
        else:
            #self._log(repr(self.installed_book.tags))
            cs = "<b>Subjects:</b> {0}".format(', '.join(self.installed_book.tags))
            #self._log("cs: %s" % repr(cs))
            self.calibre_subjects.setText(cs)
            self.marvin_subjects.setText(cs)

    def _populate_title(self):
        if 'title' in self.mismatches:
            ct = self.mismatches['title']['calibre']
            self.calibre_title.setText(self.YELLOW_BG.format(ct))
            mt = self.mismatches['title']['Marvin']
            self.marvin_title.setText(self.YELLOW_BG.format(mt))
        else:
            title = self.installed_book.title
            self.calibre_title.setText(title)
            self.marvin_title.setText(title)

    def _populate_title_sort(self):
        if 'title_sort' in self.mismatches:
            cts = self.mismatches['title_sort']['calibre']
            self.calibre_title_sort.setText(self.YELLOW_BG.format(cts))
            mts = self.mismatches['title_sort']['Marvin']
            self.marvin_title_sort.setText(self.YELLOW_BG.format(mts))
        else:
            title_sort = self.installed_book.title_sort
            self.calibre_title_sort.setText(self.GREY_FG.format(title_sort))
            self.marvin_title_sort.setText(self.GREY_FG.format(title_sort))

    def _populate_uuid(self):
        if 'uuid' in self.mismatches:
            if self.mismatches['uuid']['calibre']:
                self.calibre_uuid.setText(self.YELLOW_BG.format('uuid'))
            if self.mismatches['uuid']['Marvin']:
                self.marvin_uuid.setText(self.YELLOW_BG.format('uuid'))
            else:
                self.marvin_uuid.setText(self.YELLOW_BG.format('no uuid'))
        else:
            self.calibre_uuid.setVisible(False)
            self.marvin_uuid.setVisible(False)
