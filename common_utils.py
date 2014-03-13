#!/usr/bin/env python
# coding: utf-8

from __future__ import (unicode_literals, division, absolute_import,
                        print_function)

__license__ = 'GPL v3'
__copyright__ = '2013, Greg Riker <griker@hotmail.com>'
__docformat__ = 'restructuredtext en'

import base64, cStringIO, json, os, re, sys, time, traceback

from collections import defaultdict
from datetime import datetime
from lxml import etree
from threading import Timer
from time import sleep
#from zipfile import ZipFile

from calibre import sanitize_file_name
from calibre.constants import iswindows
from calibre.devices.usbms.driver import debug_print
from calibre.ebooks.BeautifulSoup import BeautifulSoup, BeautifulStoneSoup, Tag
from calibre.ebooks.metadata import title_sort
from calibre.ebooks.metadata.book.base import Metadata
from calibre.gui2 import Application
from calibre.gui2.dialogs.message_box import MessageBox
from calibre.gui2.progress_indicator import ProgressIndicator
from calibre.library import current_library_name
from calibre.utils.config import config_dir
from calibre.utils.ipc import RC
from calibre.utils.zipfile import ZipFile, ZIP_STORED

from PyQt4.Qt import (Qt, QAbstractItemModel, QAction, QApplication,
                      QCheckBox, QComboBox, QCursor, QDial, QDialog, QDialogButtonBox,
                      QDoubleSpinBox, QFont, QFrame, QIcon,
                      QKeySequence, QLabel, QLineEdit,
                      QPixmap, QProgressBar, QPushButton,
                      QRadioButton, QSizePolicy, QSlider, QSpinBox, QString,
                      QThread, QTimer, QUrl,
                      QVBoxLayout,
                      SIGNAL)
from PyQt4.QtWebKit import QWebView
from PyQt4.uic import compileUi

# Stateful controls: (<class>,<list_name>,<get_method>,<default>,<set_method(s)>)
# multiple set_methods are chained, i.e. the results of the first call are passed to the second
# Currently a max of two chained CONTROL_SET methods are implemented, explicity for comboBox
CONTROLS = [
    (QCheckBox, 'checkBox_controls', 'isChecked', False, 'setChecked'),
    (QComboBox, 'comboBox_controls', 'currentText', '', ('findText', 'setCurrentIndex')),
    (QDial, 'dial_controls', 'value', 0, 'setValue'),
    (QDoubleSpinBox, 'doubleSpinBox_controls', 'value', 0, 'setValue'),
    (QLineEdit, 'lineEdit_controls', 'text', '', 'setText'),
    (QRadioButton, 'radioButton_controls', 'isChecked', False, 'setChecked'),
    (QSlider, 'slider_controls', 'value', 0, 'setValue'),
    (QSpinBox, 'spinBox_controls', 'value', 0, 'setValue'),
]

CONTROL_CLASSES = [control[0] for control in CONTROLS]
CONTROL_TYPES = [control[1] for control in CONTROLS]
CONTROL_GET = [control[2] for control in CONTROLS]
CONTROL_DEFAULT = [control[3] for control in CONTROLS]
CONTROL_SET = [control[4] for control in CONTROLS]

plugin_tmpdir = 'calibre_annotations_plugin'

plugin_icon_resources = {}

'''     Constants       '''
EMPTY_STAR = u'\u2606'
FULL_STAR = u'\u2605'


'''     Base classes    '''

class Logger():
    LOCATION_TEMPLATE = "{cls}:{func}({arg1}) {arg2}"
    def _log(self, msg=None):
        '''
        Print msg to console
        '''
        from calibre_plugins.marvin_manager.config import plugin_prefs
        if not plugin_prefs.get('debug_plugin', False):
            return

        if msg:
            debug_print(" %s" % str(msg))
        else:
            debug_print()

    def _log_location(self, *args):
        '''
        Print location, args to console
        '''
        from calibre_plugins.marvin_manager.config import plugin_prefs
        if not plugin_prefs.get('debug_plugin', False):
            return

        arg1 = arg2 = ''

        if len(args) > 0:
            arg1 = str(args[0])
        if len(args) > 1:
            arg2 = str(args[1])

        debug_print(self.LOCATION_TEMPLATE.format(cls=self.__class__.__name__,
                    func=sys._getframe(1).f_code.co_name,
                    arg1=arg1, arg2=arg2))


class Book(Metadata):
    '''
    A simple class describing a book
    See ebooks.metadata.book.base #46
    '''
    # 13 standard field keys from Metadata
    mxd_standard_keys = ['author_sort', 'authors', 'comments', 'device_collections', 'pubdate',
                         'publisher', 'rating', 'series', 'series_index', 'tags', 'title', 'title_sort', 'uuid']
    # 19 private field keys
    mxd_custom_keys = ['articles', 'cid', 'calibre_collections', 'cover_file', 'date_added', 'date_opened',
                       'deep_view_prepared', 'flags', 'hash', 'highlights', 'match_quality',
                       'metadata_mismatches', 'mid', 'on_device', 'path', 'pin', 'progress', 'vocabulary',
                       'word_count']

    def __init__(self, title, author):
        if type(author) is list:
            Metadata.__init__(self, title, authors=author)
        else:
            Metadata.__init__(self, title, authors=[author])

    def __eq__(self, other):
        all_mxd_keys = self.mxd_standard_keys + self.mxd_custom_keys
        for attr in all_mxd_keys:
            v1, v2 = [getattr(obj, attr, object()) for obj in [self, other]]
            if v1 is object() or v2 is object():
                return False
            elif v1 != v2:
                return False
        return True

    def __ne__(self, other):
        all_mxd_keys = self.mxd_standard_keys + self.mxd_custom_keys
        for attr in all_mxd_keys:
            v1, v2 = [getattr(obj, attr, object()) for obj in [self, other]]
            if v1 is object() or v2 is object():
                return True
            elif v1 != v2:
                return True
        return False

    def title_sorter(self):
        return title_sort(self.title)


class MyAbstractItemModel(QAbstractItemModel):
    def __init__(self, *args):
        QAbstractItemModel.__init__(self, *args)


class Struct(dict):
    """
    Create an object with dot-referenced members or dictionary
    """
    def __init__(self, **kwds):
        dict.__init__(self, kwds)
        self.__dict__ = self

    def __repr__(self):
        return '\n'.join([" %s: %s" % (key, repr(self[key])) for key in sorted(self.keys())])


class AnnotationStruct(Struct):
    """
    Populate an empty annotation structure with fields for all possible values
    """
    def __init__(self):
        super(AnnotationStruct, self).__init__(
            annotation_id=None,
            book_id=None,
            epubcfi=None,
            genre=None,
            highlight_color=None,
            highlight_text=None,
            last_modification=None,
            location=None,
            location_sort=None,
            note_text=None,
            reader=None,
            )


class BookStruct(Struct):
    """
    Populate an empty book structure with fields for all possible values
    """
    def __init__(self):
        super(BookStruct, self).__init__(
            active=None,
            author=None,
            author_sort=None,
            book_id=None,
            genre='',
            last_annotation=None,
            path=None,
            title=None,
            title_sort=None,
            uuid=None
            )


class SizePersistedDialog(QDialog):
    '''
    This dialog is a base class for any dialogs that want their size/position
    restored when they are next opened.
    '''
    def __init__(self, parent, unique_pref_name, stays_on_top=False):
        if stays_on_top:
            QDialog.__init__(self, parent.opts.gui, Qt.WindowStaysOnTopHint)
        else:
            QDialog.__init__(self, parent.opts.gui)
        self.unique_pref_name = unique_pref_name
        self.prefs = parent.opts.prefs
        self.geom = self.prefs.get(unique_pref_name, None)
        self.finished.connect(self.dialog_closing)

        # Hook ESC key
        self.esc_action = a = QAction(self)
        self.addAction(a)
        a.triggered.connect(self.esc)
        a.setShortcuts([QKeySequence('Esc', QKeySequence.PortableText)])

    def dialog_closing(self, result):
        geom = bytearray(self.saveGeometry())
        self.prefs.set(self.unique_pref_name, geom)

    def esc(self, *args):
        pass

    def resize_dialog(self):
        if self.geom is None:
            self.resize(self.sizeHint())
        else:
            self.restoreGeometry(self.geom)


'''     Exceptions      '''


class AbortRequestException(Exception):
    '''
    '''
    pass


class DeviceNotMountedException(Exception):
    ''' '''
    pass


'''     Dialogs         '''


class HelpView(SizePersistedDialog):
    '''
    Modeless dialog for presenting HTML help content
    '''

    def __init__(self, parent, icon, prefs, html=None, page=None, title=''):
        self.prefs = prefs
        #QDialog.__init__(self, parent=parent)
        super(HelpView, self).__init__(parent, 'help_dialog')
        self.setWindowTitle(title)
        self.setWindowIcon(icon)
        self.l = QVBoxLayout(self)
        self.setLayout(self.l)

        self.wv = QWebView()
        if html is not None:
            self.wv.setHtml(html)
        elif page is not None:
            self.wv.load(QUrl(page))
        self.wv.setMinimumHeight(100)
        self.wv.setMaximumHeight(16777215)
        self.wv.setMinimumWidth(400)
        self.wv.setMaximumWidth(16777215)
        self.wv.setGeometry(0, 0, 400, 100)
        self.wv.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.l.addWidget(self.wv)

        # Sizing
        sizePolicy = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.sizePolicy().hasHeightForWidth())
        self.setSizePolicy(sizePolicy)
        self.resize_dialog()


class MyBlockingBusy(QDialog):

    NORMAL = 0
    REQUESTED = 1
    ACKNOWLEDGED = 2

    def __init__(self, gui, msg, size=100, window_title='Marvin XD', show_cancel=False,
                 on_top=False):
        flags = Qt.FramelessWindowHint
        if on_top:
            flags = Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint
        QDialog.__init__(self, gui, flags)

        self._layout = QVBoxLayout()
        self.setLayout(self._layout)
        self.cancel_status = 0
        self.is_running = False

        # Add the spinner
        self.pi = ProgressIndicator(self)
        self.pi.setDisplaySize(size)
        self._layout.addSpacing(15)
        self._layout.addWidget(self.pi, 0, Qt.AlignHCenter)
        self._layout.addSpacing(15)

        # Fiddle with the message
        self.msg = QLabel(msg)
        #self.msg.setWordWrap(True)
        self.font = QFont()
        self.font.setPointSize(self.font.pointSize() + 2)
        self.msg.setFont(self.font)
        self._layout.addWidget(self.msg, 0, Qt.AlignHCenter)
        sp = QSizePolicy()
        sp.setHorizontalStretch(True)
        sp.setVerticalStretch(False)
        sp.setHeightForWidth(False)
        self.msg.setSizePolicy(sp)
        self.msg.setMinimumHeight(self.font.pointSize() + 8)
        #self.msg.setFrameStyle(QFrame.Panel | QFrame.Sunken)

        self._layout.addSpacing(15)

        if show_cancel:
            self.bb = QDialogButtonBox()
            self.cancel_button = QPushButton(QIcon(I('window-close.png')), 'Cancel')
            self.bb.addButton(self.cancel_button, self.bb.RejectRole)
            self.bb.clicked.connect(self.button_handler)
            self._layout.addWidget(self.bb)

        self.setWindowTitle(window_title)
        self.resize(self.sizeHint())

    def accept(self):
        self.stop()
        return QDialog.accept(self)

    def button_handler(self, button):
        '''
        Only change cancel_status from NORMAL to REQUESTED
        '''
        if self.bb.buttonRole(button) == QDialogButtonBox.RejectRole:
            if self.cancel_status == self.NORMAL:
                self.cancel_status = self.REQUESTED
                self.cancel_button.setEnabled(False)

    def reject(self):
        '''
        Cannot cancel this dialog manually
        '''
        pass

    def set_text(self, text):
        self.msg.setText(text)

    def start(self):
        self.is_running = True
        self.pi.startAnimation()

    def stop(self):
        self.is_running = False
        self.pi.stopAnimation()


class ProgressBar(QDialog, Logger):
    def __init__(self, parent=None, max_items=100, window_title='Progress Bar',
                 label='Label goes here', frameless=True, on_top=False,
                 alignment=Qt.AlignHCenter):
        if on_top:
            _flags = Qt.WindowStaysOnTopHint
            if frameless:
                _flags |= Qt.FramelessWindowHint
            QDialog.__init__(self, parent=parent,
                             flags=_flags)
        else:
            _flags = Qt.Dialog
            if frameless:
                _flags |= Qt.FramelessWindowHint
            QDialog.__init__(self, parent=parent,
                             flags=_flags)
        self.application = Application
        self.setWindowTitle(window_title)
        self.l = QVBoxLayout(self)
        self.setLayout(self.l)

        self.l.addSpacing(15)

        self.label = QLabel(label)
        self.label.setAlignment(alignment)
        self.font = QFont()
        self.font.setPointSize(self.font.pointSize() + 2)
        self.label.setFont(self.font)
        self.l.addWidget(self.label)

        self.l.addSpacing(15)

        self.progressBar = QProgressBar(self)
        self.progressBar.setRange(0, max_items)
        self.progressBar.setMinimum(0)
        self.progressBar.setMaximum(0)
        self.progressBar.setValue(0)
        self.l.addWidget(self.progressBar)

        self.l.addSpacing(15)

        self.close_requested = False

    def closeEvent(self, event):
        self._log_location()
        self.close_requested = True

    def increment(self):
        try:
            if self.progressBar.value() < self.progressBar.maximum():
                self.progressBar.setValue(self.progressBar.value() + 1)
                self.refresh()
        except:
            self._log_location()
            import traceback
            self._log(traceback.format_exc())

    def refresh(self):
        self.application.processEvents()

    def set_label(self, value):
        self.label.setText(value)
        self.label.repaint()
        self.refresh()

    def set_maximum(self, value):
        self.progressBar.setMaximum(value)
        self.refresh()

    def set_range(self, min, max):
        self.progressBar.setRange(min, max)
        self.refresh()

    def set_value(self, value):
        self.progressBar.setValue(value)
        self.progressBar.repaint()
        self.refresh()


'''     Threads         '''

class IndexLibrary(QThread):
    '''
    Build indexes of library:
    title_map: {title: {'authors':…, 'id':…, 'uuid:'…}, …}
    uuid_map:  {uuid:  {'author's:…, 'id':…, 'title':…, 'path':…}, …}
    id_map:    {id:    {'uuid':…, 'author':…}, …}
    '''

    def __init__(self, parent):
        QThread.__init__(self, parent)
        self.signal = SIGNAL("library_index_complete")
        self.cdb = parent.opts.gui.current_db
        self.id_map = None
        self.hash_map = None
        self.active_virtual_library = None

    def run(self):
        self.title_map = self.index_by_title()
        self.uuid_map = self.index_by_uuid()
        self.emit(self.signal)

    def add_to_hash_map(self, hash, uuid):
        '''
        When a book has been bound to a calibre uuid, we need to add it to the hash map
        '''
        if hash not in self.hash_map:
            self.hash_map[hash] = [uuid]
        else:
            self.hash_map[hash].append(uuid)

    def build_hash_map(self):
        '''
        Generate a reverse dict of hash:[uuid] from self.uuid_map
        Allow for multiple uuids with same hash (dupes)
        Hashes are added to uuid_map in book_status:_scan_library_books()
        '''
        hash_map = {}
        for uuid, v in self.uuid_map.items():
            try:
                if v['hash'] not in hash_map:
                    hash_map[v['hash']] = [uuid]
                else:
                    hash_map[v['hash']].append(uuid)
            except:
                # Book deleted since scan
                pass
        self.hash_map = hash_map
        return hash_map

    def index_by_title(self):
        '''
        By default, any search restrictions or virtual libraries are applied
        calibre.db.view:search_getting_ids()
        '''
        by_title = {}

        cids = self.cdb.search_getting_ids('formats:EPUB', '')
        for cid in cids:
            title = self.cdb.title(cid, index_is_id=True)
            by_title[title] = {
                'authors': self.cdb.authors(cid, index_is_id=True).split(','),
                'id': cid,
                'uuid': self.cdb.uuid(cid, index_is_id=True)
                }
        return by_title

    def index_by_uuid(self):
        '''
        By default, any search restrictions or virtual libraries are applied
        calibre.db.view:search_getting_ids()
        '''
        by_uuid = {}

        cids = self.cdb.search_getting_ids('formats:EPUB', '')
        for cid in cids:
            uuid = self.cdb.uuid(cid, index_is_id=True)
            by_uuid[uuid] = {
                'authors': self.cdb.authors(cid, index_is_id=True).split(','),
                'id': cid,
                'title': self.cdb.title(cid, index_is_id=True),
                }

        return by_uuid


class InventoryCollections(QThread):
    '''
    Build a list of books with collection assignments
    '''

    def __init__(self, parent):
        QThread.__init__(self, parent)
        self.signal = SIGNAL("collection_inventory_complete")
        self.cdb = parent.opts.gui.current_db
        self.cfl = get_cc_mapping('collections', 'field', None)
        self.ids = []
        #self.heatmap = {}

    def run(self):
        self.inventory_collections()
        self.emit(self.signal)

    def inventory_collections(self):
        id = self.cdb.FIELD_MAP['id']
        if self.cfl is not None:
            for record in self.cdb.data.iterall():
                mi = self.cdb.get_metadata(record[id], index_is_id=True)
                collection_list = mi.get_user_metadata(self.cfl, False)['#value#']
                if collection_list:
                    # Add this cid to list of library books with active collection assignments
                    self.ids.append(record[id])

                    if False:
                        # Update the heatmap
                        for ca in collection_list:
                            if ca not in self.heatmap:
                                self.heatmap[ca] = 1
                            else:
                                self.heatmap[ca] += 1


class MoveBackup(QThread, Logger):
    '''
    Move a (potentially large) backup file from connected device to local fs
    '''
    IOS_TRANSFER_RATE = 7000000  # ~7 MB/second
    TIMER_TICK = 0.25

    def __init__(self, **kwargs):
        '''
        kwargs: {'backup_folder', 'destination_folder', 'ios', 'parent', 'pb',
                 'storage_name', 'stats', total_seconds}
        '''
        self._log_location()
        try:
            for key in kwargs:
                setattr(self, key, kwargs.get(key))
            QThread.__init__(self, self.parent)

            for prop in ['iosra_booklist', 'mxd_mainDb_profile', 'mxd_device_cached_hashes',
                         'mxd_installed_books', 'mxd_remote_content_hashes', 'success',
                         'timer']:
                setattr(self, prop, None)

            self.dst = os.path.join(self.destination_folder,
                                    sanitize_file_name(self.storage_name))
            self.src = "{0}/marvin.backup".format(self.backup_folder)
            self.total_seconds *= 1.25   # allow for MXD component processing
            self._init_pb()

        except:
            import traceback
            self._log(traceback.format_exc())


    def run(self):
        try:
            backup_size = "{:,} MB".format(int(int(self.src_stats['st_size'])/(1024*1024)))
            self._log_location()
            self._log("moving {0} to '{1}'".format(backup_size, self.destination_folder))
            # Remove any older file of the same name at destination
            if os.path.isfile(self.dst):
                os.remove(self.dst)

            # Copy from the iDevice to destination
            with open(self.dst, 'wb') as out:
                self.ios.copy_from_idevice(self.src, out)

            # Validate transferred file sizes, do cleanup
            self._verify()

            # Append MXD components
            self._append_mxd_components()

            self.pb.set_value(self.total_seconds)
            self._cleanup()

        except:
            import traceback
            self._log(traceback.format_exc())

    def _append_mxd_components(self):
        self._log_location()
        if (self.iosra_booklist or
            self.mxd_mainDb_profile or
            self.mxd_device_cached_hashes or
            self.mxd_installed_books or
            self.mxd_remote_content_hashes):

            with ZipFile(self.dst, mode='a') as zfa:
                if self.iosra_booklist:
                    zfa.write(self.iosra_booklist, arcname="iosra_booklist.zip")

                if self.mxd_mainDb_profile:
                    zfa.writestr("mxd_mainDb_profile.json",
                                 json.dumps(self.mxd_mainDb_profile, sort_keys=True))

                if self.mxd_device_cached_hashes:
                    base_name = "mxd_cover_hashes.json"
                    zfa.write(self.mxd_device_cached_hashes, arcname=base_name)

                if self.mxd_installed_books:
                    base_name = "mxd_installed_books.json"
                    zfa.writestr(base_name, self.mxd_installed_books)

                if self.mxd_remote_content_hashes:
                    from calibre_plugins.marvin_manager.book_status import BookStatusDialog
                    base_name = "mxd_{0}".format(BookStatusDialog.HASH_CACHE_FS)
                    zfa.write(self.mxd_remote_content_hashes, arcname=base_name)

    def _cleanup(self):
        self._log_location()
        try:
            self.ios.remove(self.src)
            self.ios.remove(self.backup_folder)
            self.timer.cancel()
            self.pb.hide()
        except:
            import traceback
            self._log(traceback.format_exc())

    def _init_pb(self):
        self._log_location()
        try:
            max = int(self.total_seconds/self.TIMER_TICK) + 1
            self.pb.set_maximum(max)
            self.pb.set_range(0, max)
            self.timer = Timer(self.TIMER_TICK, self._ticked)
            self.timer.start()
        except:
            import traceback
            self._log(traceback.format_exc())

    def _ticked(self):
        '''
        Increment the progress bar, restart the timer
        '''
        try:
            self.pb.increment()
            self.timer = Timer(self.TIMER_TICK, self._ticked)
            self.timer.start()
        except:
            import traceback
            self._log(traceback.format_exc())

    def _verify(self):
        '''
        Confirm that the file was properly transferred
        '''
        src_size = int(self.src_stats['st_size'])
        dst_size = os.stat(self.dst).st_size
        self.success = (src_size == dst_size)
        self._log_location('backup verified' if self.success else '')
        if not self.success:
            self._log("file sizes did not match:")
            self._log("src_size: {0}".format(src_size))
            self._log("dst_size: {0}".format(dst_size))


class RestoreBackup(QThread, Logger):
    '''
    Copy a (potentially large) backup file from local fs to connected device
    ProgressBar needs to be created from main GUI thread
    '''
    IOS_TRANSFER_RATE = 7000000  # ~7 MB/second
    TIMER_TICK = 0.25

    def __init__(self, **kwargs):
        '''
        kwargs: {'backup_image', 'ios', 'msg', 'parent', 'pb', 'total_seconds'}
        '''
        self._log_location()
        try:
            for key in kwargs:
                setattr(self, key, kwargs.get(key))
            QThread.__init__(self, self.parent)
            self.src_size = os.stat(self.backup_image).st_size
            self.success = None
            self.timer = None
            self._init_pb()
        except:
            import traceback
            self._log(traceback.format_exc())

    def run(self):
        self._log_location()
        try:
            self._log("moving {0:,} bytes to '/Documents'".format(self.src_size))
            tmp = b'/'.join(['/Documents', 'restore_image.tmp'])
            self.dst = b'/'.join(['/Documents', 'marvin.backup'])
            self.ios.copy_to_idevice(self.backup_image, tmp)
            self.ios.rename(tmp, self.dst)
            self._verify()
            self.pb.set_value(self.total_seconds)
            self._cleanup()
        except:
            import traceback
            self._log(traceback.format_exc())

    def _cleanup(self):
        self._log_location()
        try:
            self.timer.cancel()
            self.pb.hide()
        except:
            import traceback
            self._log(traceback.format_exc())

    def _init_pb(self):
        self._log_location()
        try:
            max = int(self.total_seconds/self.TIMER_TICK)
            self.pb.set_maximum(max)
            self.pb.set_range(0, max)
            self.timer = Timer(self.TIMER_TICK, self._ticked)
            self.timer.start()
        except:
            import traceback
            self._log(traceback.format_exc())

    def _ticked(self):
        '''
        Increment the progress bar, restart the timer
        '''
        try:
            self.pb.increment()
            self.timer = Timer(self.TIMER_TICK, self._ticked)
            self.timer.start()
        except:
            import traceback
            self._log(traceback.format_exc())

    def _verify(self):
        '''
        Confirm source size == dest size
        '''
        self._log_location()
        try:
            self.dst_size = int(self.ios.exists(self.dst)['st_size'])
        except:
            self.dst_size = -1

        if self.src_size != self.dst_size:
            self.success = False
            self.ios.remove(self.dst)
        else:
            self.success = True

class RowFlasher(QThread):
    '''
    Flash rows_to_flash to show where ops occurred
    '''

    def __init__(self, parent, model, rows_to_flash):
        QThread.__init__(self)
        self.signal = SIGNAL("flasher_complete")
        self.model = model
        self.parent = parent
        self.rows_to_flash = rows_to_flash
        self.mode = 'old'

        self.cycles = self.parent.prefs.get('flasher_cycles', 3) + 1
        self.new_time = self.parent.prefs.get('flasher_new_time', 300)
        self.old_time = self.parent.prefs.get('flasher_old_time', 100)

    def run(self):
        QTimer.singleShot(self.old_time, self.update)
        while self.cycles:
            QApplication.processEvents()
        self.emit(self.signal)

    def toggle_values(self, mode):
        for row, item in self.rows_to_flash.items():
            self.model.set_match_quality(row, item[mode])

    def update(self):
        if self.mode == 'new':
            self.toggle_values('old')
            self.mode = 'old'
            QTimer.singleShot(self.old_time, self.update)
        elif self.mode == 'old':
            self.toggle_values('new')
            self.mode = 'new'
            self.cycles -= 1
            if self.cycles:
                QTimer.singleShot(self.new_time, self.update)

'''     Helper Classes  '''


class CommandHandler(Logger):
    '''
    Consolidated class for handling Marvin commands
    Construct two types of commands:
    METADATA_COMMAND_XML: specific
    GENERAL_COMMAND_XML: general
    '''
    POLLING_DELAY = 0.25        # Spinner frequency
    WATCHDOG_TIMEOUT = 10.0

    GENERAL_COMMAND_XML = b'''\xef\xbb\xbf<?xml version='1.0' encoding='utf-8'?>
    <command type=\'{0}\' timestamp=\'{1}\'>
    </command>'''

    METADATA_COMMAND_XML = b'''\xef\xbb\xbf<?xml version='1.0' encoding='utf-8'?>
    <{0} timestamp=\'{1}\'>
    <manifest>
    </manifest>
    </{0}>'''

    def __init__(self, parent):
        self._log_location()
        self.busy_cancel_requested = False
        self.command_name = None
        self.command_soup = None
        self.connected_device = parent.connected_device
        self.get_response = None
        self.ios = parent.ios
        self.marvin_cancellation_required = False
        self.operation_timed_out = False
        self.pb = None
        self.prefs = parent.prefs
        self.results = None
        self.timeout_override = None

    def construct_general_command(self, cmd_type):
        '''
        Create GENERAL_COMMAND_XML soup
        '''
        self._log_location("type={0}".format(cmd_type))
        self.command_name = 'command'
        self.command_type = cmd_type
        self.command_soup = BeautifulStoneSoup(self.GENERAL_COMMAND_XML.format(
            cmd_type, time.mktime(time.localtime())))

    def construct_metadata_command(self, cmd_element=None, cmd_name=None):
        '''
        Create METADATA_COMMAND_XML soup
        '''
        self.command_name = cmd_name
        self.command_soup = BeautifulStoneSoup(self.METADATA_COMMAND_XML.format(
            cmd_element, time.mktime(time.localtime())))

    def init_pb(self, total_seconds):
        self._log_location()
        try:
            max = int(total_seconds/self.POLLING_DELAY)
            self.pb.set_maximum(max)
            self.pb.set_range(0, max)
        except:
            import traceback
            self._log(traceback.format_exc())

    def issue_command(self, get_response=None, timeout_override=None):
        '''
        Consolidated command handler
        '''
        self._log_location()

        self.get_response = get_response
        self.timeout_override = timeout_override

        QApplication.setOverrideCursor(QCursor(Qt.WaitCursor))

        # Wait for the driver to be silent
        while self.connected_device.get_busy_flag():
            Application.processEvents()
        self.connected_device.set_busy_flag(True)

        # Copy command file to staging folder
        self._stage_command_file()

        # Wait for completion
        try:
            results = self._wait_for_command_completion()
        except:
            import traceback
            details = "An error occurred while executing '{0}'.\n\n".format(self.command_name)
            details += traceback.format_exc()
            results = {'code': '2',
                       'status': "Error communicating with Marvin",
                       'details': details}

        # Try to reset the busy flag, although it might fail
        try:
            self.connected_device.set_busy_flag(False)
        except:
            pass

        QApplication.restoreOverrideCursor()
        self.results = results

    def _stage_command_file(self):

        self._log_location()

        if self.prefs.get('show_staged_commands', False):
            if self.command_name in ['update_metadata', 'update_metadata_items']:
                soup = BeautifulStoneSoup(self.command_soup.renderContents())
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
                self._log(self.command_soup.prettify())

        if self.prefs.get('execute_marvin_commands', True):
            # Make sure there is no orphan status.xml from a previous timeout
            if self.ios.exists(self.connected_device.status_fs):
                self.ios.remove(self.connected_device.status_fs)

            tmp = b'/'.join([self.connected_device.staging_folder, b'%s.tmp' % self.command_name])
            final = b'/'.join([self.connected_device.staging_folder, b'%s.xml' % self.command_name])
            self.ios.write(self.command_soup.renderContents(), tmp)
            self.ios.rename(tmp, final)

        else:
            self._log("~~~ execute_marvin_commands disabled in JSON ~~~")

    def _wait_for_command_completion(self):
        '''
        Wait for Marvin to issue progress reports via status.xml
        Marvin creates status.xml upon receiving command, increments <progress>
        from 0.0 to 1.0 as command progresses.
        '''

        msg = "timeout: {0}".format(self.WATCHDOG_TIMEOUT)
        if self.timeout_override:
            msg = "timeout_override: {0}".format(self.timeout_override)
        self._log_location(msg)

        results = {'code': 0}

        if self.prefs.get('execute_marvin_commands', True):
            self._log("%s: waiting for '%s'" %
                      (datetime.now().strftime('%H:%M:%S.%f'),
                      self.connected_device.status_fs))

            if not self.timeout_override:
                timeout_value = self.WATCHDOG_TIMEOUT
            else:
                timeout_value = self.timeout_override

            # Set initial watchdog timer for ACK with default timeout
            self.operation_timed_out = False
            self.watchdog = Timer(self.WATCHDOG_TIMEOUT, self._watchdog_timed_out)
            self.watchdog.start()

            while True:
                if not self.ios.exists(self.connected_device.status_fs):
                    # status.xml not created yet
                    if self.operation_timed_out:
                        final_code = '-1'
                        self.ios.remove(self.connected_device.status_fs)
                        results = {
                            'code': -1,
                            'status': 'timeout',
                            'response': None,
                            'details': 'timeout_value: %d' % timeout_value
                            }
                        break
                    Application.processEvents()
                    if self.pb:
                        self.pb.increment()
                    time.sleep(self.POLLING_DELAY)

                else:
                    # Start a new watchdog timer per iteration
                    self.watchdog.cancel()
                    self.watchdog = Timer(timeout_value, self._watchdog_timed_out)
                    self.operation_timed_out = False
                    self.watchdog.start()

                    self._log("%s: monitoring progress of %s" %
                              (datetime.now().strftime('%H:%M:%S.%f'),
                              self.command_name))

                    code = '-1'
                    current_timestamp = 0.0
                    while code == '-1':
                        try:
                            if self.operation_timed_out:
                                self.ios.remove(self.connected_device.status_fs)
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
                                ft = (b'/'.join([self.connected_device.staging_folder,
                                                 b'cancel.tmp']))
                                fs = (b'/'.join([self.connected_device.staging_folder,
                                                 b'cancel.command']))
                                self.ios.write("please stop", ft)
                                self.ios.rename(ft, fs)

                                # Update status
                                self._busy_status_msg(msg="Completing operation on current book…")

                                # Clear flags so we can complete processing
                                self.marvin_cancellation_required = False

                            status = etree.fromstring(self.ios.read(self.connected_device.status_fs))
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
                            if self.pb:
                                self.pb.increment()
                            time.sleep(self.POLLING_DELAY)

                        except:
                            self.watchdog.cancel()

                            formatted_lines = traceback.format_exc().splitlines()
                            current_error = formatted_lines[-1]

                            time.sleep(self.POLLING_DELAY)
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
                    if True:
                        # *** Fake some errors to test ***
                        if self.command_name == 'update_metadata':
                            self._log("***falsifying update_metadata error***")
                            results = {'code': 2, 'status': 'completed with errors',
                                'details': "[Title - Author.epub] Cannot locate book to update metadata - skipping"}
                        elif self.command_name == 'command' and self.command_type == 'backup':
                            self._log("***falsifying backup error***")
                            results = {'code': 2, 'status': 'completed with errors',
                                'details': "Insufficient space available to create backup"}
                    '''

                    # Get the response file from the staging folder
                    if self.get_response:
                        rf = b'/'.join([self.connected_device.staging_folder, self.get_response])
                        self._log("fetching response '%s'" % rf)
                        if not self.ios.exists(rf):
                            response = "%s not found" % rf
                        else:
                            response = self.ios.read(rf)
                            self.ios.remove(rf)
                        results['response'] = response

                    self.ios.remove(self.connected_device.status_fs)

                    if final_code not in ['0']:
                        if final_code == '-1' and self.operation_timed_out:
                            msgs = ['Operation timed out',
                                    'timeout_value: {0} seconds'.format(timeout_value)]
                        elif final_code == '3':
                            msgs = ['operation cancelled by user']
                        else:
                            messages = status.find('messages')
                            msgs = [msg.text for msg in messages]

                        details = '\n'.join(["code: %s" % final_code, "status: %s" % final_status])
                        details += '\n'.join(msgs)
                        results['details'] = '\n'.join(msgs)

                        self._log("%s: <%s> complete with errors" %
                                  (datetime.now().strftime('%H:%M:%S.%f'),
                                  self.command_name))
                    else:
                        self._log("%s: <%s> complete" %
                                  (datetime.now().strftime('%H:%M:%S.%f'),
                                  self.command_name))
                    break

        else:
            self._log("~~~ execute_marvin_commands disabled in JSON ~~~")

        return results

    def _watchdog_timed_out(self):
        '''
        Set flag if I/O operation times out
        '''
        self._log_location(datetime.now().strftime('%H:%M:%S.%f'))
        self.operation_timed_out = True


class CompileUI():
    '''
    Compile Qt Creator .ui files at runtime
    '''
    def __init__(self, parent):
        self.compiled_forms = {}
        self.help_file = None
        self._log = parent._log
        self._log_location = parent._log_location
        self.parent = parent
        self.verbose = parent.verbose
        self.compiled_forms = self.compile_ui()

    def compile_ui(self):
        pat = re.compile(r'''(['"]):/images/([^'"]+)\1''')

        def sub(match):
            ans = 'I(%s%s%s)' % (match.group(1), match.group(2), match.group(1))
            return ans

        # >>> Entry point
        self._log_location()

        compiled_forms = {}
        self._find_forms()

        # Cribbed from gui2.__init__:build_forms()
        for form in self.forms:
            with open(form) as form_file:
                soup = BeautifulStoneSoup(form_file.read())
                property = soup.find('property', attrs={'name': 'windowTitle'})
                string = property.find('string')
                window_title = string.renderContents()

            compiled_form = self._form_to_compiled_form(form)
            if (not os.path.exists(compiled_form) or
                    os.stat(form).st_mtime > os.stat(compiled_form).st_mtime):

                if not os.path.exists(compiled_form):
                    if self.verbose:
                        self._log(' compiling %s' % form)
                else:
                    if self.verbose:
                        self._log(' recompiling %s' % form)
                    os.remove(compiled_form)
                buf = cStringIO.StringIO()
                compileUi(form, buf)
                dat = buf.getvalue()
                dat = dat.replace('__appname__', 'calibre')
                dat = dat.replace('import images_rc', '')
                dat = re.compile(r'(?:QtGui.QApplication.translate|(?<!def )_translate)\(.+?,\s+"(.+?)(?<!\\)",.+?\)').sub(r'_("\1")', dat)
                dat = dat.replace('_("MMM yyyy")', '"MMM yyyy"')
                dat = pat.sub(sub, dat)
                with open(compiled_form, 'wb') as cf:
                    cf.write(dat)

            compiled_forms[window_title] = compiled_form.rpartition(os.sep)[2].partition('.')[0]
        return compiled_forms

    def _find_forms(self):
        forms = []
        for root, _, files in os.walk(self.parent.resources_path):
            for name in files:
                if name.endswith('.ui'):
                    forms.append(os.path.abspath(os.path.join(root, name)))
        self.forms = forms

    def _form_to_compiled_form(self, form):
        compiled_form = form.rpartition('.')[0]+'_ui.py'
        return compiled_form


'''     Helper functions   '''

def _log(msg=None):
    '''
    Print msg to console
    '''
    from calibre_plugins.marvin_manager.config import plugin_prefs
    if not plugin_prefs.get('debug_plugin', False):
        return

    if msg:
        debug_print(" %s" % str(msg))
    else:
        debug_print()


def _log_location(*args):
    LOCATION_TEMPLATE = "{cls}:{func}({arg1}) {arg2}"

    from calibre_plugins.marvin_manager.config import plugin_prefs
    if not plugin_prefs.get('debug_plugin', False):
        return

    arg1 = arg2 = ''

    if len(args) > 0:
        arg1 = str(args[0])
    if len(args) > 1:
        arg2 = str(args[1])

    debug_print(LOCATION_TEMPLATE.format(cls='common_utils',
                func=sys._getframe(1).f_code.co_name,
                arg1=arg1, arg2=arg2))


def existing_annotations(parent, field, return_all=False):
    '''
    Return count of existing annotations, or existence of any
    '''
    #import calibre_plugins.marvin_manager.config as cfg
    _log_location(field)
    annotation_map = []
    if field:
        db = parent.opts.gui.current_db
        id = db.FIELD_MAP['id']
        for i, record in enumerate(db.data.iterall()):
            mi = db.get_metadata(record[id], index_is_id=True)
            if field == 'Comments':
                if mi.comments:
                    soup = BeautifulSoup(mi.comments)
                else:
                    continue
            else:
                soup = BeautifulSoup(mi.get_user_metadata(field, False)['#value#'])
            if soup.find('div', 'user_annotations') is not None:
                annotation_map.append(mi.id)
                if not return_all:
                    break
        if return_all:
            _log("Identified %d annotated books of %d total books" %
                (len(annotation_map), len(db.data)))

        _log("annotation_map: %s" % repr(annotation_map))
    else:
       _log("no active field")

    return annotation_map


def from_json(obj):
    '''
    Models calibre.utils.config:from_json
    uses local parse_date()
    '''
    if '__class__' in obj:
        if obj['__class__'] == 'bytearray':
            return bytearray(base64.standard_b64decode(obj['__value__']))
        if obj['__class__'] == 'datetime.datetime':
            return parse_date(obj['__value__'])
    return obj


def get_cc_mapping(cc_name, element, default=None):
    '''
    Return the element mapped to cc_name in prefs
    '''
    from calibre_plugins.marvin_manager.config import plugin_prefs

    if element not in ['field', 'combobox']:
        raise ValueError("invalid element '{0}' requested for custom column '{1}'".format(
            element, cc_name))

    ans = default
    cc_mappings = plugin_prefs.get('cc_mappings', {})
    current_library = current_library_name()
    if (current_library in cc_mappings and
        cc_name in cc_mappings[current_library] and
        element in cc_mappings[current_library][cc_name]):
        ans = cc_mappings[current_library][cc_name][element]
    return ans


def get_icon(icon_name):
    '''
    Retrieve a QIcon for the named image from the zip file if it exists,
    or if not then from Calibre's image cache.
    '''
    if icon_name:
        pixmap = get_pixmap(icon_name)
        if pixmap is None:
            # Look in Calibre's cache for the icon
            return QIcon(I(icon_name))
        else:
            return QIcon(pixmap)
    return QIcon()


def get_local_images_dir(subfolder=None):
    '''
    Returns a path to the user's local resources/images folder
    If a subfolder name parameter is specified, appends this to the path
    '''
    images_dir = os.path.join(config_dir, 'resources/images')
    if subfolder:
        images_dir = os.path.join(images_dir, subfolder)
    if iswindows:
        images_dir = os.path.normpath(images_dir)
    return images_dir


def get_pixmap(icon_name):
    '''
    Retrieve a QPixmap for the named image
    Any zipped icons belonging to the plugin must be prefixed with 'images/'
    '''
    global plugin_icon_resources

    if not icon_name.startswith('images/'):
        # We know this is definitely not an icon belonging to this plugin
        pixmap = QPixmap()
        pixmap.load(I(icon_name))
        return pixmap

    # As we did not find an icon elsewhere, look within our zip resources
    if icon_name in plugin_icon_resources:
        pixmap = QPixmap()
        pixmap.loadFromData(plugin_icon_resources[icon_name])
        return pixmap
    return None


def isoformat(date_time, sep='T'):
    '''
    Mocks calibre.utils.date:isoformat()
    '''
    return unicode(date_time.isoformat(str(sep)))


def move_annotations(parent, annotation_map, old_destination_field, new_destination_field,
                     window_title="Moving annotations"):
    '''
    Move annotations from old_destination_field to new_destination_field
    annotation_map precalculated in thread in config.py
    '''
    import calibre_plugins.marvin_manager.config as cfg
    from calibre_plugins.marvin_manager.annotations import BookNotes, BookmarkNotes

    _log_location(annotation_map)
    _log(" %s -> %s" % (repr(old_destination_field), repr(new_destination_field)))

    db = parent.opts.gui.current_db
    id = db.FIELD_MAP['id']

    # Show progress
    pb = ProgressBar(parent=parent, window_title=window_title)
    total_books = len(annotation_map)
    pb.set_maximum(total_books)
    pb.set_value(1)
    pb.set_label('{:^100}'.format('Moving annotations for %d books' % total_books))
    pb.show()

    transient_db = 'transient'

    # Prepare a new COMMENTS_DIVIDER
    comments_divider = '<div class="comments_divider"><p style="text-align:center;margin:1em 0 1em 0">{0}</p></div>'.format(
        cfg.plugin_prefs.get('COMMENTS_DIVIDER', '&middot;  &middot;  &bull;  &middot;  &#x2726;  &middot;  &bull;  &middot; &middot;'))

    for cid in annotation_map:
        mi = db.get_metadata(cid, index_is_id=True)

        # Comments -> custom
        if old_destination_field == 'Comments' and new_destination_field.startswith('#'):
            if mi.comments:
                old_soup = BeautifulSoup(mi.comments)
                uas = old_soup.find('div', 'user_annotations')
                if uas:
                    # Remove user_annotations from Comments
                    uas.extract()

                    # Remove comments_divider from Comments
                    cd = old_soup.find('div', 'comments_divider')
                    if cd:
                        cd.extract()

                    # Save stripped Comments
                    mi.comments = unicode(old_soup)

                    # Capture content
                    parent.opts.db.capture_content(uas, cid, transient_db)

                    # Regurgitate content with current CSS style
                    new_soup = parent.opts.db.rerender_to_html(transient_db, cid)

                    # Add user_annotations to destination
                    um = mi.metadata_for_field(new_destination_field)
                    um['#value#'] = unicode(new_soup)
                    mi.set_user_metadata(new_destination_field, um)

                    # Update the record with stripped Comments, populated custom field
                    db.set_metadata(cid, mi, set_title=False, set_authors=False,
                                    commit=True, force_changes=True, notify=True)
                    pb.increment()

        # custom -> Comments
        elif old_destination_field.startswith('#') and new_destination_field == 'Comments':
            if mi.get_user_metadata(old_destination_field, False)['#value#'] is not None:
                old_soup = BeautifulSoup(mi.get_user_metadata(old_destination_field, False)['#value#'])
                uas = old_soup.find('div', 'user_annotations')
                if uas:
                    # Remove user_annotations from custom field
                    uas.extract()

                    # Capture content
                    parent.opts.db.capture_content(uas, cid, transient_db)

                    # Regurgitate content with current CSS style
                    new_soup = parent.opts.db.rerender_to_html(transient_db, cid)

                    # Save stripped custom field data
                    um = mi.metadata_for_field(old_destination_field)
                    um['#value#'] = unicode(old_soup)
                    mi.set_user_metadata(old_destination_field, um)

                    # Add user_annotations to Comments
                    if mi.comments is None:
                        mi.comments = unicode(new_soup)
                    else:
                        mi.comments = mi.comments + \
                                      unicode(comments_divider) + \
                                      unicode(new_soup)

                    # Update the record with stripped custom field, updated Comments
                    db.set_metadata(cid, mi, set_title=False, set_authors=False,
                                    commit=True, force_changes=True, notify=True)
                    pb.increment()

        # same field -> same field - called from config:configure_appearance()
        elif (old_destination_field == new_destination_field):
            pb.set_label('{:^100}'.format('Updating annotations for %d books' % total_books))

            if new_destination_field == 'Comments':
                if mi.comments:
                    old_soup = BeautifulSoup(mi.comments)
                    uas = old_soup.find('div', 'user_annotations')
                    if uas:
                        # Remove user_annotations from Comments
                        uas.extract()

                        # Remove comments_divider from Comments
                        cd = old_soup.find('div', 'comments_divider')
                        if cd:
                            cd.extract()

                        # Save stripped Comments
                        mi.comments = unicode(old_soup)

                        # Capture content
                        parent.opts.db.capture_content(uas, cid, transient_db)

                        # Regurgitate content with current CSS style
                        new_soup = parent.opts.db.rerender_to_html(transient_db, cid)

                        # Add user_annotations to Comments
                        if mi.comments is None:
                            mi.comments = unicode(new_soup)
                        else:
                            mi.comments = mi.comments + \
                                          unicode(comments_divider) + \
                                          unicode(new_soup)

                        # Update the record with stripped custom field, updated Comments
                        db.set_metadata(cid, mi, set_title=False, set_authors=False,
                                        commit=True, force_changes=True, notify=True)
                        pb.increment()

            else:
                # Update custom field
                old_soup = BeautifulSoup(mi.get_user_metadata(old_destination_field, False)['#value#'])

                # Rerender book notes div
                bnd = old_soup.find('div', 'book_note')
                if bnd:
                    bnd.replaceWith(BookNotes().reconstruct(bnd))

                # Rerender bookmark notes div
                bmnd = old_soup.find('div', 'bookmark_notes')
                if bmnd:
                    bmnd.replaceWith(BookmarkNotes().reconstruct(bmnd))

                # Rerender annotations
                uas = old_soup.find('div', 'user_annotations')
                if uas:
                    # Capture content
                    parent.opts.db.capture_content(uas, cid, transient_db)

                    # Regurgitate annotations with current CSS style
                    #new_soup = parent.opts.db.rerender_to_html(transient_db, cid)
                    rerendered_annotations = BeautifulSoup(
                        parent.opts.db.rerender_to_html(transient_db, cid))
                    uas.replaceWith(rerendered_annotations)

                # Add stripped old_soup plus new_soup to destination field
                um = mi.metadata_for_field(new_destination_field)
                um['#value#'] = unicode(old_soup)
                mi.set_user_metadata(new_destination_field, um)

                # Update the record
                db.set_metadata(cid, mi, set_title=False, set_authors=False,
                                commit=True, force_changes=True, notify=True)


                pb.increment()

        # custom -> custom
        elif old_destination_field.startswith('#') and new_destination_field.startswith('#'):

            if mi.get_user_metadata(old_destination_field, False)['#value#'] is not None:
                old_soup = BeautifulSoup(mi.get_user_metadata(old_destination_field, False)['#value#'])

                # Rerender book notes div
                bnd = old_soup.find('div', 'book_note')
                if bnd:
                    bnd.replaceWith(BookNotes().reconstruct(bnd))

                # Rerender bookmark notes div
                bmnd = old_soup.find('div', 'bookmark_notes')
                if bmnd:
                    bmnd.replaceWith(BookmarkNotes().reconstruct(bmnd))

                # Rerender annotations
                uas = old_soup.find('div', 'user_annotations')
                if uas:
                    # Capture content
                    parent.opts.db.capture_content(uas, cid, transient_db)

                    # Regurgitate content with current CSS style
                    #new_soup = parent.opts.db.rerender_to_html(transient_db, cid)
                    rerendered_annotations = BeautifulSoup(
                        parent.opts.db.rerender_to_html(transient_db, cid))
                    uas.replaceWith(rerendered_annotations)

                    # Save stripped custom field data
                    um = mi.metadata_for_field(old_destination_field)
                    #um['#value#'] = unicode(old_soup)
                    um['#value#'] = None
                    mi.set_user_metadata(old_destination_field, um)

                    # Add updated soup to destination field
                    um = mi.metadata_for_field(new_destination_field)
                    um['#value#'] = unicode(old_soup)
                    mi.set_user_metadata(new_destination_field, um)

                # Update the record
                db.set_metadata(cid, mi, set_title=False, set_authors=False,
                                commit=True, force_changes=True, notify=True)
                pb.increment()

    # Hide the progress bar
    pb.hide()

    # Get the eligible custom fields
    all_custom_fields = db.custom_field_keys()
    custom_fields = {}
    for cf in all_custom_fields:
        field_md = db.metadata_for_field(cf)
        if field_md['datatype'] in ['comments']:
            custom_fields[field_md['name']] = {'field': cf,
                                               'datatype': field_md['datatype']}

    # Change field value to friendly name
    if old_destination_field.startswith('#'):
        for cf in custom_fields:
            if custom_fields[cf]['field'] == old_destination_field:
                old_destination_field = cf
                break
    if new_destination_field.startswith('#'):
        for cf in custom_fields:
            if custom_fields[cf]['field'] == new_destination_field:
                new_destination_field = cf
                break

    # Report what happened
    if old_destination_field == new_destination_field:
        msg = "<p>Annotations updated to new appearance settings for %d {0}.</p>" % len(annotation_map)
    else:
        msg = ("<p>Annotations for %d {0} moved from <b>%s</b> to <b>%s</b>.</p>" %
                (len(annotation_map), old_destination_field, new_destination_field))
    if len(annotation_map) == 1:
        msg = msg.format('book')
    else:
        msg = msg.format('books')
    MessageBox(MessageBox.INFO,
               '',
               msg=msg,
               show_copy_button=False,
               parent=parent.gui).exec_()
    _log("INFO: %s" % msg)

    # Update the UI
    updateCalibreGUIView()


def parse_date(date_string):
    '''
    Mocks calibre.utils.date:parse_date()
    https://labix.org/python-dateutil#head-42a94eedcff96da7fb1f77096b5a3b519c859ba9
    '''
    UNDEFINED_DATE = datetime(101,1,1, tzinfo=None)
    from dateutil.parser import parse
    if not date_string:
        return UNDEFINED_DATE
    return parse(date_string, ignoretz=True)


def inventory_controls(ui, dump_controls=False):
    '''
     Build an inventory of stateful controls
    '''
    controls = {'owner': ui.__class__.__name__}
    control_dict = defaultdict(list)
    for control_type in CONTROL_TYPES:
        control_dict[control_type] = []

    # Inventory existing controls
    for item in ui.__dict__:
        if type(ui.__dict__[item]) in CONTROL_CLASSES:
            index = CONTROL_CLASSES.index(type(ui.__dict__[item]))
            control_dict[CONTROL_TYPES[index]].append(str(ui.__dict__[item].objectName()))

    for control_list in CONTROL_TYPES:
        if control_dict[control_list]:
            controls[control_list] = control_dict[control_list]

    if dump_controls:
        _log_location()
        _log("Inventoried controls:")
        for control_type in CONTROL_TYPES:
            if control_type in controls:
                _log(" %s: %s" % (control_type, controls[control_type]))

    return controls


def restore_state(ui, prefs, restore_position=False):
    def _restore_ui_position(ui, owner):
        parent_loc = ui.iap.gui.pos()
        if True:
            last_x = prefs.get('%s_last_x' % owner, parent_loc.x())
            last_y = prefs.get('%s_last_y' % owner, parent_loc.y())
        else:
            last_x = parent_loc.x()
            last_y = parent_loc.y()
        ui.move(last_x, last_y)

    if restore_position:
        _restore_ui_position(ui, ui.controls['owner'])

    # Restore stateful controls
    for control_list in ui.controls:
        if control_list == 'owner':
            continue
        index = CONTROL_TYPES.index(control_list)
        for control in ui.controls[control_list]:
            control_ref = getattr(ui, control, None)
            if control_ref is not None:
                if isinstance(CONTROL_SET[index], unicode):
                    setter_ref = getattr(control_ref, CONTROL_SET[index], None)
                    if setter_ref is not None:
                        if callable(setter_ref):
                            setter_ref(prefs.get(control, CONTROL_DEFAULT[index]))
                elif isinstance(CONTROL_SET[index], tuple) and len(CONTROL_SET[index]) == 2:
                    # Special case for comboBox - first findText, then setCurrentIndex
                    setter_ref = getattr(control_ref, CONTROL_SET[index][0], None)
                    if setter_ref is not None:
                        if callable(setter_ref):
                            result = setter_ref(prefs.get(control, CONTROL_DEFAULT[index]))
                            setter_ref = getattr(control_ref, CONTROL_SET[index][1], None)
                            if setter_ref is not None:
                                if callable(setter_ref):
                                    setter_ref(result)
                else:
                    _log_location()
                    _log("invalid CONTROL_SET tuple for '%s'" % control)
                    _log("maximum of two chained methods")


def save_state(ui, prefs, save_position=False):
    def _save_ui_position(ui, owner):
        prefs.set('%s_last_x' % owner, ui.pos().x())
        prefs.set('%s_last_y' % owner, ui.pos().y())

    if save_position:
        _save_ui_position(ui, ui.controls['owner'])

    # Save stateful controls
    for control_list in ui.controls:
        if control_list == 'owner':
            continue
        index = CONTROL_TYPES.index(control_list)

        for control in ui.controls[control_list]:
            # Intercept QString objects, coerce to unicode
            qt_type = getattr(getattr(ui, control), CONTROL_GET[index])()
            if type(qt_type) is QString:
                qt_type = unicode(qt_type)
            prefs.set(control, qt_type)


def set_cc_mapping(cc_name, field=None, combobox=None):
    '''
    Store element to cc_name in prefs:cc_mappings
    '''
    from calibre_plugins.marvin_manager.config import plugin_prefs

    cc_mappings = plugin_prefs.get('cc_mappings', {})
    current_library = current_library_name()
    if current_library in cc_mappings:
        cc_mappings[current_library][cc_name] = {'field': field, 'combobox': combobox}
    else:
        cc_mappings[current_library] = {cc_name: {'field': field, 'combobox': combobox}}
    plugin_prefs.set('cc_mappings', cc_mappings)


def set_plugin_icon_resources(name, resources):
    '''
    Set our global store of plugin name and icon resources for sharing between
    the InterfaceAction class which reads them and the ConfigWidget
    if needed for use on the customization dialog for this plugin.
    '''
    global plugin_icon_resources, plugin_name
    plugin_name = name
    plugin_icon_resources = resources


def to_json(obj):
    '''
    Models calibre.utils.config:to_json
    Uses local isoformat()
    '''
    if isinstance(obj, bytearray):
        return {'__class__': 'bytearray',
                '__value__': base64.standard_b64encode(bytes(obj))}
    if isinstance(obj, datetime):
        return {'__class__': 'datetime.datetime',
                '__value__': isoformat(obj)}
    raise TypeError(repr(obj) + ' is not JSON serializable')


def updateCalibreGUIView():
    '''
    Refresh the GUI view
    '''
    t = RC(print_error=False)
    t.start()
    sleep(0.5)
    while True:
        if t.done:
            t.conn.send('refreshdb:')
            t.conn.close()
            break
        sleep(0.5)

