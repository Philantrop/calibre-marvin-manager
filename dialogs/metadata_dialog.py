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
from calibre_plugins.marvin_manager.common_utils import SizePersistedDialog

from PyQt4.Qt import (Qt, QColor, QDialog, QDialogButtonBox, QIcon, QPalette, QPixmap,
                      QSize, QSizePolicy)

# Import Ui_Form from form generated dynamically during initialization
if True:
    sys.path.insert(0, dialog_resources_path)
    from metadata_ui import Ui_Dialog
    sys.path.remove(dialog_resources_path)

class MetadataComparisonDialog(SizePersistedDialog, Ui_Dialog):
    COVER_ICON_SIZE = 200
    LOCATION_TEMPLATE = "{cls}:{func}({arg1}) {arg2}"

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

        elif self.bb.buttonRole(button) == QDialogButtonBox.ActionRole:
            if button.objectName() == 'export_to_marvin_button':
                self.export_to_marvin()
            elif button.objectName() == 'import_from_marvin_button':
                self.import_from_marvin()

        elif self.bb.buttonRole(button) == QDialogButtonBox.RejectRole:
            self.close()

    def esc(self, *args):
        self.close()

    def export_to_marvin(self):
        self._log_location()

    def import_from_marvin(self):
        self._log_location()

    def initialize(self, parent, book_id, cid, installed_book, marvin_db_path):
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
        self.installed_book = installed_book
        self.marvin_db_path = marvin_db_path
        self.opts = parent.opts
        self.parent = parent
        self.stored_command = None
        self.verbose = parent.verbose
        self.BORDER_COLOR = "#FDFF99"
        self.BORDER_LR = 4
        self.BORDER_TB = 8
        self.GREY_FG = '<font style="color:#A0A0A0">{0}</font>'
        self.YELLOW_BG = '<font style="background:#FDFF99">{0}</font>'

        self._log_location(installed_book.title)

        self._log("mismatches:\n%s" % repr(installed_book.metadata_mismatches))
        self.mismatches = installed_book.metadata_mismatches

        self._populate_title()
        self._populate_title_sort()
        self._populate_series()
        self._populate_authors()
        self._populate_author_sort()
        self._populate_covers()
        self._populate_subjects()
        self._populate_publisher()
        self._populate_pubdate()
        self._populate_description()

        # ~~~~~~~~ Export to Marvin button ~~~~~~~~
        self.export_to_marvin_button.setIcon(QIcon(os.path.join(self.parent.opts.resources_path,
                                                   'icons',
                                                   'from_calibre.png')))
        self.export_to_marvin_button.clicked.connect(partial(self.store_command, 'export_to_marvin'))

        # ~~~~~~~~ Import from Marvin button ~~~~~~~~
        self.import_from_marvin_button.setIcon(QIcon(os.path.join(self.parent.opts.resources_path,
                                                   'icons',
                                                   'from_marvin.png')))
        self.import_from_marvin_button.clicked.connect(partial(self.store_command, 'import_from_marvin'))

        # If no calibre book, or no mismatches, hide the Calibre group and action buttons
        if  not self.cid or not self.mismatches:
            self.calibre_gb.setVisible(False)
            self.import_from_marvin_button.setVisible(False)
            self.setWindowTitle(u'Metadata Summary')
        else:
            self.setWindowTitle(u'Metadata Comparison')

        # Set the Marvin QGroupBox to Marvin red
#         marvin_red = QColor()
#         marvin_red.setRgb(189, 17, 20, alpha=255)
#         palette = QPalette()
#         palette.setColor(QPalette.Background, marvin_red)
#         self.marvin_gb.setPalette(palette)

        self.bb.clicked.connect(self.dispatch_button_click)

        # Restore position
        self.resize_dialog()

    def store_command(self, command):
        '''
        '''
        self._log_location(command)
        self.stored_command = command
        self.close()

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
            # Retrieve Books:LargeCoverJpg if no cover_path
            if self.installed_book.cover_file:
                self._log("fetch cover from Marvin sandbox")
            else:
                self._log("fetch cover from mainDb")
                con = sqlite3.connect(self.marvin_db_path)
                with con:
                    con.row_factory = sqlite3.Row

                    # Fetch LargeCoverJpg from mainDb
                    cover_cur = con.cursor()
                    cover_cur.execute('''SELECT
                                          LargeCoverJpg
                                         FROM Books
                                         WHERE ID = '{0}'
                                      '''.format(self.book_id))
                    rows = cover_cur.fetchall()

                if len(rows):
                    try:
                        # Save Marvin cover in case we're importing to calibre
                        self.marvin_cover_jpg = rows[0][b'LargeCoverJpg']
                        marvin_thumb = thumbnail(self.marvin_cover_jpg,
                                                 self.COVER_ICON_SIZE,
                                                 self.COVER_ICON_SIZE)
                        pixmap = QPixmap()
                        if with_border:
                            bordered_thumb = add_borders_to_image(marvin_thumb[2],
                                                              left=self.BORDER_LR,
                                                              right=self.BORDER_LR,
                                                              top=self.BORDER_TB,
                                                              bottom=self.BORDER_TB,
                                                              border_color=self.BORDER_COLOR)
                            pixmap.loadFromData(bordered_thumb)
                        else:
                            pixmap.loadFromData(marvin_thumb[2])
                        self.marvin_cover.setPixmap(pixmap)
                    except:
                        # No cover available, use generic
                        import traceback
                        self._log(traceback.format_exc())

                        self._log("failed to fetch LargeCoverJpg for %s (%s)" %
                                  (self.installed_book.title, self.book_id))
                        pixmap = QPixmap()
                        pixmap.load(I('book.png'))
                        pixmap = pixmap.scaled(self.COVER_ICON_SIZE,
                                               self.COVER_ICON_SIZE,
                                               aspectRatioMode=Qt.KeepAspectRatio,
                                               transformMode=Qt.SmoothTransformation)
                        self.marvin_cover.setPixmap(pixmap)
                        self.marvin_cover_jpg = None

                else:
                    self._log("no cover data fetched from mainDb")

        self.calibre_cover.setMaximumSize(QSize(self.COVER_ICON_SIZE, self.COVER_ICON_SIZE))
        self.calibre_cover.setText('')
        self.calibre_cover.setScaledContents(False)

        self.marvin_cover.setMaximumSize(QSize(self.COVER_ICON_SIZE, self.COVER_ICON_SIZE))
        self.marvin_cover.setText('')
        self.marvin_cover.setScaledContents(False)

        if self.cid:
            if 'cover_hash' not in self.mismatches:
                db = self.opts.gui.current_db
                mi = db.get_metadata(self.cid, index_is_id=True, get_cover=True, cover_as_data=True)
                calibre_thumb = thumbnail(mi.cover_data[1],
                                          self.COVER_ICON_SIZE,
                                          self.COVER_ICON_SIZE)
                pixmap = QPixmap()
                pixmap.loadFromData(calibre_thumb[2])
                self.calibre_cover.setPixmap(pixmap)

                # Marvin cover matches calibre cover
                self.marvin_cover.setPixmap(pixmap)
            else:
                # Covers don't match - render with border
                db = self.opts.gui.current_db
                mi = db.get_metadata(self.cid, index_is_id=True, get_cover=True, cover_as_data=True)
                calibre_thumb = thumbnail(mi.cover_data[1],
                                          self.COVER_ICON_SIZE,
                                          self.COVER_ICON_SIZE)
                bordered_thumb = add_borders_to_image(calibre_thumb[2],
                                                      left=self.BORDER_LR,
                                                      right=self.BORDER_LR,
                                                      top=self.BORDER_TB,
                                                      bottom=self.BORDER_TB,
                                                      border_color=self.BORDER_COLOR)

                pixmap = QPixmap()
                pixmap.loadFromData(bordered_thumb)
                self.calibre_cover.setPixmap(pixmap)
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
            self.calibre_description_label.setText(self.YELLOW_BG.format("Description"))
            if not self.mismatches['comments']['calibre']:
                self.calibre_description.setVisible(False)
                self.calibre_description_label.setText(self.YELLOW_BG.format("No description available"))
            else:
                self.calibre_description.setText(self.mismatches['comments']['calibre'])

            self.marvin_description_label.setText(self.YELLOW_BG.format("Description"))
            if not self.mismatches['comments']['Marvin']:
                self.marvin_description.setVisible(False)
                self.marvin_description_label.setText(self.YELLOW_BG.format("No description available"))
            else:
                self.marvin_description.setText(self.mismatches['comments']['Marvin'])
        else:
            if self.installed_book.comments:
                self.calibre_description.setText(self.installed_book.comments)
                self.marvin_description.setText(self.installed_book.comments)
            else:
                self.calibre_description.setVisible(False)
                self.calibre_description_label.setText("No description available")
                self.marvin_description.setVisible(False)
                self.marvin_description_label.setText("No description available")

    def _populate_pubdate(self):
        if 'pubdate' in self.mismatches:
            if self.mismatches['pubdate']['calibre']:
                cs_pubdate = "Published %s" % strftime("%e %B %Y", t=self.mismatches['pubdate']['calibre'])
            else:
                cs_pubdate = "Unknown date of publication"
            self.calibre_pubdate.setText(self.YELLOW_BG.format(cs_pubdate))

            if self.mismatches['pubdate']['Marvin']:
                ms_pubdate = "Published %s" % strftime("%e %B %Y", t=self.mismatches['pubdate']['Marvin'])
            else:
                ms_pubdate = "Unknown date of publication"
            self.marvin_pubdate.setText(self.YELLOW_BG.format(ms_pubdate))
        elif self.installed_book.pubdate:
            pubdate = "Published %s" % strftime("%e %B %Y", t=self.installed_book.pubdate)
            self.calibre_pubdate.setText(pubdate)
            self.marvin_pubdate.setText(pubdate)
        else:
            pubdate = "Publication date not available"
            self.calibre_pubdate.setText(pubdate)
            self.marvin_pubdate.setText(pubdate)

    def _populate_publisher(self):
        if 'publisher' in self.mismatches:
            cs_publisher = self.mismatches['publisher']['calibre']
            if not cs_publisher:
                cs_publisher = "Unknown publisher"
            self.calibre_publisher.setText(self.YELLOW_BG.format(cs_publisher))

            ms_publisher = self.mismatches['publisher']['Marvin']
            if not ms_publisher:
                ms_publisher = "Unknown publisher"
            self.marvin_publisher.setText(self.YELLOW_BG.format(ms_publisher))
        else:
            publisher = self.installed_book.publisher
            if not publisher:
                publisher = "Unknown publisher"
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

        # Setting size policy allows us to set each Subjects fields to the same height
        sp = QSizePolicy()
        sp.setVerticalStretch(False)
        sp.setHeightForWidth(False)
        self.calibre_subjects.setSizePolicy(sp)
        self.marvin_subjects.setSizePolicy(sp)

        if 'tags' in self.mismatches:
            cs = "<b>Subjects:</b> %s" % ', '.join(self.mismatches['tags']['calibre'])
            self.calibre_subjects.setText(self.YELLOW_BG.format(cs))
            ms = "<b>Subjects:</b> %s" % ', '.join(self.mismatches['tags']['Marvin'])
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
            cs = "<b>Subjects:</b> %s" % ', '.join(self.installed_book.tags)
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

