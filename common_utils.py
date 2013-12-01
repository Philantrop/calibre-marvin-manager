#!/usr/bin/env python
# coding: utf-8

from __future__ import (unicode_literals, division, absolute_import,
                        print_function)

__license__ = 'GPL v3'
__copyright__ = '2013, Greg Riker <griker@hotmail.com>'
__docformat__ = 'restructuredtext en'

import cStringIO, os, re, sys

from collections import defaultdict
from time import sleep

from calibre.constants import iswindows
from calibre.devices.usbms.driver import debug_print
from calibre.ebooks.BeautifulSoup import BeautifulSoup, BeautifulStoneSoup
from calibre.gui2 import Application
from calibre.gui2.dialogs.message_box import MessageBox
from calibre.gui2.progress_indicator import ProgressIndicator
from calibre.ebooks.metadata.book.base import Metadata
from calibre.utils.config import config_dir
from calibre.utils.ipc import RC

from PyQt4.Qt import (Qt, QAbstractItemModel, QAction, QApplication,
                      QCheckBox, QComboBox, QDial, QDialog, QDialogButtonBox,
                      QDoubleSpinBox, QFont, QIcon,
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
    def __init__(self, title, author):
        if type(author) is list:
            Metadata.__init__(self, title, authors=author)
        else:
            Metadata.__init__(self, title, authors=[author])

    @property
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


class ProgressBar(QDialog):
    def __init__(self, parent=None, max_items=100, window_title='Progress Bar',
                 label='Label goes here', frameless=True, on_top=False):
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

        self.label = QLabel(label)
        self.label.setAlignment(Qt.AlignHCenter)
        self.l.addWidget(self.label)

        self.progressBar = QProgressBar(self)
        self.progressBar.setRange(0, max_items)
        self.progressBar.setMinimum(0)
        self.progressBar.setMaximum(0)
        self.progressBar.setValue(0)
        self.l.addWidget(self.progressBar)

        self.close_requested = False

    def closeEvent(self, event):
        debug_print("ProgressBar:closeEvent()")
        self.close_requested = True

    def increment(self):
        self.progressBar.setValue(self.progressBar.value() + 1)
        self.refresh()

    def refresh(self):
        self.application.processEvents()

    def set_label(self, value):
        self.label.setText(value)
        self.label.repaint()
        self.refresh()

    def set_maximum(self, value):
        self.progressBar.setMaximum(value)
        self.refresh()

    def set_value(self, value):
        self.progressBar.setValue(value)
        self.progressBar.repaint()
        self.refresh()


'''     Threads         '''


class IndexLibrary(QThread):
    '''
    Build indexes of library:
    {title: {'authors':…, 'id':…, 'uuid:'…}, …}
    {uuid:  {'author's:…, 'id':…, 'title':…, 'path':…}, …}
    {id:    {'uuid':…, 'author':…}, …}
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
        self.cfl = parent.prefs.get('collection_field_lookup', None)
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


def move_annotations(parent, annotation_map, old_destination_field, new_destination_field,
                     window_title="Moving annotations"):
    '''
    Move annotations from old_destination_field to new_destination_field
    annotation_map precalculated in thread in config.py
    '''
    import calibre_plugins.marvin_manager.config as cfg

    _log_location(annotation_map)
    _log(" %s -> %s" % (old_destination_field, new_destination_field))

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

        # custom -> custom
        elif old_destination_field.startswith('#') and new_destination_field.startswith('#'):

            if mi.get_user_metadata(old_destination_field, False)['#value#'] is not None:
                old_soup = BeautifulSoup(mi.get_user_metadata(old_destination_field, False)['#value#'])
                uas = old_soup.find('div', 'user_annotations')
                if uas:
                    # Remove user_annotations from originating custom field
                    uas.extract()

                    # Capture content
                    parent.opts.db.capture_content(uas, cid, transient_db)

                    # Regurgitate content with current CSS style
                    new_soup = parent.opts.db.rerender_to_html(transient_db, cid)

                    # Save stripped custom field data
                    um = mi.metadata_for_field(old_destination_field)
                    um['#value#'] = unicode(old_soup)
                    mi.set_user_metadata(old_destination_field, um)

                    # Add new_soup to destination field
                    um = mi.metadata_for_field(new_destination_field)
                    um['#value#'] = unicode(new_soup)
                    mi.set_user_metadata(new_destination_field, um)

                    # Update the record
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
                uas = old_soup.find('div', 'user_annotations')
                if uas:
                    # Remove user_annotations from originating custom field
                    uas.extract()

                    # Capture content
                    parent.opts.db.capture_content(uas, cid, transient_db)

                    # Regurgitate content with current CSS style
                    new_soup = parent.opts.db.rerender_to_html(transient_db, cid)

                    # Add stripped old_soup plus new_soup to destination field
                    um = mi.metadata_for_field(new_destination_field)
                    um['#value#'] = unicode(old_soup) + unicode(new_soup)
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


def set_plugin_icon_resources(name, resources):
    '''
    Set our global store of plugin name and icon resources for sharing between
    the InterfaceAction class which reads them and the ConfigWidget
    if needed for use on the customization dialog for this plugin.
    '''
    global plugin_icon_resources, plugin_name
    plugin_name = name
    plugin_icon_resources = resources


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
