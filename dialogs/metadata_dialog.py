#!/usr/bin/env python
# coding: utf-8

from __future__ import (unicode_literals, division, absolute_import,
                        print_function)

__license__ = 'GPL v3'
__copyright__ = '2010, Gregory Riker'
__docformat__ = 'restructuredtext en'

import os, sqlite3, sys

from calibre.devices.usbms.driver import debug_print
from calibre.utils.magick.draw import thumbnail

from calibre_plugins.marvin_manager.book_status import dialog_resources_path
from calibre_plugins.marvin_manager.common_utils import SizePersistedDialog

from PyQt4.Qt import (QDialog, QDialogButtonBox, QIcon, QPixmap, QSize)

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
        __init__ is called on SizePersistecDialog()
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
        self.verbose = parent.verbose
        self.YELLOW_BG = '<font style="background:#FDFF99">{0}</font>'

        self._log_location(installed_book.title)
        self._log("matches:\n%s" % repr(installed_book.metadata_matches))
        self._log("mismatches:\n%s" % repr(installed_book.metadata_mismatches))

        self.setWindowTitle(u'Metadata Comparison')
        self.matches = installed_book.metadata_matches
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
        #self.export_to_marvin_button = self.bb.addButton('Export to Marvin', QDialogButtonBox.ActionRole)
        #self.export_to_marvin_button.setObjectName('export_to_marvin_button')
        self.export_to_marvin_button.setIcon(QIcon(os.path.join(self.parent.opts.resources_path,
                                                   'icons',
                                                   'from_calibre.png')))
        # ~~~~~~~~ Import from Marvin button ~~~~~~~~
        #self.import_from_marvin_button = self.bb.addButton('Import from Marvin', QDialogButtonBox.ActionRole)
        #self.import_from_marvin_button.setObjectName('import_from_marvin_button')
        self.import_from_marvin_button.setIcon(QIcon(os.path.join(self.parent.opts.resources_path,
                                                   'icons',
                                                   'from_marvin.png')))

        self.bb.clicked.connect(self.dispatch_button_click)

        # Restore position
        self.resize_dialog()

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
            authors = ', '.join(self.matches['authors'])
            self.calibre_authors.setText(authors)
            self.marvin_authors.setText(authors)

    def _populate_author_sort(self):
        if 'author_sort' in self.mismatches:
            cs_author_sort = self.mismatches['author_sort']['calibre']
            self.calibre_author_sort.setText(self.YELLOW_BG.format(cs_author_sort))
            ms_author_sort = self.mismatches['author_sort']['Marvin']
            self.marvin_author_sort.setText(self.YELLOW_BG.format(ms_author_sort))
        else:
            author_sort = self.matches['author_sort']
            self.calibre_author_sort.setText(author_sort)
            self.marvin_author_sort.setText(author_sort)

    def _populate_covers(self):
        '''
        Display calibre cover for both unless mismatch
        '''
        self.calibre_cover.setMaximumSize(QSize(self.COVER_ICON_SIZE, self.COVER_ICON_SIZE))
        self.calibre_cover.setText('')
        self.calibre_cover.setScaledContents(False)

        self.marvin_cover.setMaximumSize(QSize(self.COVER_ICON_SIZE, self.COVER_ICON_SIZE))
        self.marvin_cover.setText('')
        self.marvin_cover.setScaledContents(False)

        # Calibre cover always set from library cover
        db = self.opts.gui.current_db
        mi = db.get_metadata(self.cid, index_is_id=True, get_cover=True, cover_as_data=True)
        calibre_thumb = thumbnail(mi.cover_data[1],
                                  self.COVER_ICON_SIZE,
                                  self.COVER_ICON_SIZE)
        pixmap = QPixmap()
        pixmap.loadFromData(calibre_thumb[2])
        self.calibre_cover.setPixmap(pixmap)

        if 'cover_hash' not in self.mismatches:
            # Marvin cover matches calibre cover
            self.marvin_cover.setPixmap(pixmap)
        else:
            # Retrieve Books:LargeCoverJpg if no cover_path
            self._log("cover mismatch")
            if self.installed_book.cover_file:
                self._log("fetch cover from Marvin sandbox")
            else:
                self._log("fetch cover from mainDb")
                con = sqlite3.connect(self.marvin_db_path)
                with con:
                    con.row_factory = sqlite3.Row

                    # Build a collection map
                    cover_cur = con.cursor()
                    cover_cur.execute('''SELECT
                                          LargeCoverJpg
                                         FROM Books
                                         WHERE ID = '{0}'
                                      '''.format(self.book_id))
                    rows = cover_cur.fetchall()

                if len(rows):
                    # Save Marvin cover in case we're importing to calibre
                    self.marvin_cover_jpg = rows[0][b'LargeCoverJpg']
                    marvin_thumb = thumbnail(self.marvin_cover_jpg,
                                             self.COVER_ICON_SIZE,
                                             self.COVER_ICON_SIZE)
                    pixmap = QPixmap()
                    pixmap.loadFromData(marvin_thumb[2])
                    self.marvin_cover.setPixmap(pixmap)
                else:
                    self._log("no cover data fetched from mainDb")

    def _populate_description(self):
        if 'comments' in self.mismatches:
            self.calibre_description.setText(self.mismatches['comments']['calibre'])
            self.marvin_description.setText(self.mismatches['comments']['Marvin'])
        elif 'comments' in self.matches:
            description = self.matches['comments']
            self.calibre_description.setText(description)
            self.marvin_description.setText(description)

    def _populate_pubdate(self):
        if 'pubdate' in self.mismatches:
            cs_pubdate = "Published %s-%s-%s" % self.mismatches['pubdate']['calibre']
            self.calibre_pubdate.setText(self.YELLOW_BG.format(cs_pubdate))
            ms_pubdate = "Published %s-%s-%s" % self.mismatches['pubdate']['Marvin']
            self.marvin_pubdate.setText(self.YELLOW_BG.format(ms_pubdate))
        else:
            pubdate = "Published %s-%s-%s" % self.matches['pubdate']
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
            publisher = self.matches['publisher']
            if not publisher:
                publisher = "Unknown publisher"
            self.calibre_publisher.setText(publisher)
            self.marvin_publisher.setText(publisher)

    def _populate_series(self):
        if 'series' in self.mismatches:
            cs_index = str(self.mismatches['series_index']['calibre'])
            if cs_index.endswith('.0'):
                cs_index = cs_index[:-2]
            cs = "%s [%s]" % (self.mismatches['series']['calibre'], cs_index)
            self.calibre_series.setText(self.YELLOW_BG.format(cs))
            ms_index = str(self.mismatches['series_index']['Marvin'])
            if ms_index.endswith('.0'):
                ms_index = ms_index[:-2]
            ms = "%s [%s]" % (self.mismatches['series']['Marvin'], ms_index)
            self.marvin_series.setText(self.YELLOW_BG.format(ms))
        elif 'series' in self.matches:
            cs_index = str(self.matches['series_index'])
            if cs_index.endswith('.0'):
                cs_index = cs_index[:-2]
            cs = "%s [%s]" % (self.matches['series'], cs_index)
            self.calibre_series.setText(cs)
            self.marvin_series.setText(cs)
        else:
            self.calibre_series.setVisible(False)
            self.marvin_series.setVisible(False)

    def _populate_subjects(self):
        if 'tags' in self.mismatches:
            cs = "Subjects: %s" % ', '.join(self.mismatches['tags']['calibre'])
            self.calibre_subjects.setText(cs)
            ms = "Subjects: %s" % ', '.join(self.mismatches['tags']['Marvin'])
            self.marvin_subjects.setText(ms)
        else:
            cs = "Subjects: %s" % ', '.join(self.matches['tags'])
            self.calibre_subjects.setText(cs)
            self.marvin_subjects.setText(cs)

    def _populate_title(self):
        if 'title' in self.mismatches:
            self.calibre_title.setText(self.mismatches['title']['calibre'])
            self.marvin_title.setText(self.mismatches['title']['Marvin'])
        else:
            title = self.matches['title']
            self.calibre_title.setText(title)
            self.marvin_title.setText(title)

    def _populate_title_sort(self):
        if 'title_sort' in self.mismatches:
            self.calibre_title_sort.setText(self.mismatches['title_sort']['calibre'])
            self.marvin_title_sort.setText(self.mismatches['title_sort']['Marvin'])
        else:
            title_sort = self.matches['title_sort']
            self.calibre_title_sort.setText(title_sort)
            self.marvin_title_sort.setText(title_sort)

