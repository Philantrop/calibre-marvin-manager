#!/usr/bin/env python
# coding: utf-8

from __future__ import (unicode_literals, division, absolute_import,
                        print_function)

__license__ = 'GPL v3'
__copyright__ = '2013, Greg Riker <griker@hotmail.com>'
__docformat__ = 'restructuredtext en'

import atexit, cPickle as pickle, os, re, sqlite3, sys, threading

from datetime import datetime
from functools import partial
from lxml import etree, html
from zipfile import ZipFile

from PyQt4.Qt import (Qt, QApplication, QCursor, QFileDialog, QIcon,
                      QMenu, QTimer, QUrl,
                      pyqtSignal)

from calibre.constants import DEBUG
from calibre.customize.ui import device_plugins, disabled_device_plugins
from calibre.devices.idevice.libimobiledevice import libiMobileDevice
from calibre.devices.usbms.driver import debug_print
from calibre.ebooks.BeautifulSoup import BeautifulSoup
from calibre.gui2 import Application, open_url
from calibre.gui2.actions import InterfaceAction
from calibre.gui2.device import device_signals
from calibre.gui2.dialogs.message_box import MessageBox
from calibre.library import current_library_name
from calibre.utils.config import config_dir

from calibre_plugins.marvin_manager import MarvinManagerPlugin
from calibre_plugins.marvin_manager.annotations_db import AnnotationsDB
from calibre_plugins.marvin_manager.book_status import BookStatusDialog
from calibre_plugins.marvin_manager.common_utils import (AbortRequestException,
    CommandHandler, CompileUI, IndexLibrary, Logger, MoveBackup, MyBlockingBusy,
    ProgressBar, RestoreBackup, Struct,
    get_icon, set_plugin_icon_resources, updateCalibreGUIView)
import calibre_plugins.marvin_manager.config as cfg
#from calibre_plugins.marvin_manager.dropbox import PullDropboxUpdates

# The first icon is the plugin icon, referenced by position.
# The rest of the icons are referenced by name
PLUGIN_ICONS = ['images/connected.png', 'images/disconnected.png']

class MarvinManagerAction(InterfaceAction, Logger):

    # Location reporting template
    LOCATION_TEMPLATE = "{cls}:{func}({arg1}) {arg2}"

    icon = PLUGIN_ICONS[0]
    minimum_ios_driver_version = (1, 3, 5)
    name = 'Marvin XD'
    prefs = cfg.plugin_prefs
    verbose = prefs.get('debug_plugin', False)

    # Declare the main action associated with this plugin
    action_spec = ('Marvin XD', None, None, None)
    #popup_type = QToolButton.InstantPopup
    action_add_menu = True
    action_menu_clone_qaction = True

    marvin_device_status_changed = pyqtSignal(dict)
    plugin_device_connection_changed = pyqtSignal(object)

    def about_to_show_menu(self):
        self.rebuild_menus()

    def create_backup(self):
        '''
        iPad1:      500 books in 90 seconds - 5.5 books/second
        iPad Mini:  500 books in 64 seconds - 7.8 books/second
        '''
        WORST_CASE_ARCHIVE_RATE = 4.0  # Books/second
        TIMEOUT_PADDING_FACTOR = 1.5
        backup_folder = b'/'.join(['/Documents', 'Backup'])
        backup_target = backup_folder + '/marvin.backup'
        last_backup_folder = self.prefs.get('backup_folder', os.path.expanduser("~"))

        def _confirm_overwrite():
            '''
            Check for existing backup before overwriting
            stats['st_mtime'], stats['st_size']
            Return True: continue
            Return False: cancel
            '''
            stats = self.ios.exists(backup_target)
            if stats:
                d = datetime.fromtimestamp(float(stats['st_mtime']))
                friendly_date = d.strftime("%A, %B %d, %Y")
                friendly_time = d.strftime("%I:%M%p")

                title = "A backup already exists!"
                msg = ('<p>There is an existing backup created ' +
                       '{0} at {1}.</p>'
                       '<p>Proceeding with this backup will '
                       'overwrite the existing backup.</p>'
                       '<p>Proceed?</p>'.format(friendly_date, friendly_time))
                dlg = MessageBox(MessageBox.QUESTION, title, msg,
                                 show_copy_button=False)
                return dlg.exec_()
            return True

        def _confirm_lengthy_backup():
            '''
            If this is going to take some time, warn the user
            '''
            self._log("estimated time to backup {0} books: {1}".format(total_books, estimated_time))
            # Confirm that user wants to proceed given estimated time to completion
            book_descriptor = "books" if total_books > 1 else "book"
            title = "Estimated time to create backup"
            msg = ("<p>Creating a backup of " +
                   "{0} books ".format(total_books) +
                   "may take as long as {0}, depending on your iDevice.</p>".format(estimated_time) +
                   "<p>Proceed?</p>")
            dlg = MessageBox(MessageBox.QUESTION, title, msg,
                             show_copy_button=False)
            return dlg.exec_()

        def _count_books():
            # Get a count of the books
            con = sqlite3.connect(self.connected_device.local_db_path)
            with con:
                con.row_factory = sqlite3.Row
                cur = con.cursor()
                cur.execute('''SELECT
                                title
                               FROM Books
                            ''')
                rows = cur.fetchall()
            return(len(rows))

        def _estimate_time():
            # Estimate worst-case time required to create backup
            m, s = divmod(total_seconds, 60)
            h, m = divmod(m, 60)
            if h:
                estimated_time = "%d:%02d:%02d" % (h, m, s)
            else:
                estimated_time = "%d:%02d" % (m, s)
            return estimated_time

        # ~~~ Entry point ~~~
        self._log_location()
        total_books = _count_books()
        total_seconds = int(total_books/WORST_CASE_ARCHIVE_RATE) + 1
        timeout = int(total_seconds * TIMEOUT_PADDING_FACTOR)
        estimated_time = _estimate_time()

        if timeout > CommandHandler.WATCHDOG_TIMEOUT:
            if not _confirm_lengthy_backup():
                return
        else:
            timeout = CommandHandler.WATCHDOG_TIMEOUT

        if not _confirm_overwrite():
            self._log("user declined to overwrite existing backup")
            return

        # Issue the command
        self._busy_panel_setup("Backing up {0:,} books from {1}…".format(
            total_books, self.ios.device_name))
        ch = CommandHandler(self)
        ch.construct_general_command('backup')
        ch.issue_command(timeout_override=timeout)
        self._busy_panel_teardown()

        if ch.results['code']:
            self._log("results: %s" % ch.results)
            title = "Backup unsuccessful"
            msg = ('<p>Unable to create backup of {0}.</p>'
                   '<p>Click <b>Show details</b> for more information.</p>').format(
                   self.ios.device_name)
            det_msg = ch.results['details']
            MessageBox(MessageBox.WARNING, title, msg, det_msg=det_msg).exec_()
            return

        # Move backup to the specified location
        stats = self.ios.exists(backup_target)
        if stats:
            dn = self.ios.device_name
            d = datetime.fromtimestamp(float(stats['st_mtime']))
            storage_name = "{0} {1}.backup".format(
                self.ios.device_name, d.strftime("%Y-%m-%d"))
            destination_folder = str(QFileDialog.getExistingDirectory(
                self.gui,
                "Select destination folder for backup",
                last_backup_folder,
                QFileDialog.ShowDirsOnly | QFileDialog.DontResolveSymlinks))

            if destination_folder:
                # Qt apparently sometimes returns a file within the selected directory,
                # rather than the directory itself. Validate destination_folder
                if not os.path.isdir(destination_folder):
                    destination_folder = os.path.dirname(destination_folder)

                # Move from iDevice to destination_folder
                move_operation = MoveBackup(self, backup_folder, destination_folder, storage_name, stats)
                msg = '<p>Moving marvin.backup ({0:,} bytes)…</p>'.format(
                       int(stats['st_size']))
                self._busy_panel_setup(msg)
                move_operation.start()
                while not move_operation.isFinished():
                    Application.processEvents()
                self._busy_panel_teardown()

                # Inform user backup operation is complete
                title = "Backup operation complete"
                msg = '<p>Marvin library has been backed up to {0}.</p>'.format(destination_folder)
                MessageBox(MessageBox.INFO, title, msg).exec_()

                # Save the backup folder
                self.prefs.set('backup_folder', destination_folder)
            else:
                # Inform user backup operation cancelled
                title = "Backup cancelled"
                msg = '<p>Backup of {0} cancelled.</p>'.format(self.ios.device_name)
                MessageBox(MessageBox.WARNING, title, msg, show_copy_button=False).exec_()
        else:
            self._log("No backup file found at {0}".format(backup_target))

    def create_menu_item(self, m, menu_text, image=None, tooltip=None, shortcut=None):
        ac = self.create_action(spec=(menu_text, None, tooltip, shortcut), attr=menu_text)
        if image:
            ac.setIcon(QIcon(image))
        m.addAction(ac)
        return ac

    def developer_utilities(self, action):
        '''
        'Delete calibre hashes', 'Delete Marvin hashes'
        remote_cache_folder = '/'.join(['/Library', 'calibre.mm'])
        '''
        self._log_location(action)
        if action in ['Create backup', 'Delete calibre hashes', 'Delete Marvin hashes',
                      'Nuke annotations', 'Reset column widths',
                      'Restore from backup']:
            if action == 'Delete Marvin hashes':
                remote_cache_folder = '/'.join(['/Library', 'calibre.mm'])
                rhc = b'/'.join([remote_cache_folder, BookStatusDialog.HASH_CACHE_FS])

                if self.ios.exists(rhc):
                    self.ios.remove(rhc)
                    self._log("remote hash cache at %s deleted" % rhc)

                # Remove cover hashes for connected device
                device_cached_hashes = "{0}_cover_hashes.json".format(
                    re.sub('\W', '_', self.ios.device_name))
                dch = os.path.join(self.resources_path, device_cached_hashes)
                if os.path.exists(dch):
                    os.remove(dch)
                    self._log("cover hashes at {0} deleted".format(dch))
                else:
                    self._log("no cover hashes found at {0}".format(dch))

            elif action == 'Delete calibre hashes':
                self.gui.current_db.delete_all_custom_book_data('epub_hash')
                self._log("cached epub hashes deleted")
                # Invalidate the library hash map, as library contents may change before reconnection
                if hasattr(self, 'library_scanner'):
                    if hasattr(self.library_scanner, 'hash_map'):
                        self.library_scanner.hash_map = None
            elif action == 'Nuke annotations':
                self.nuke_annotations()
            elif action == 'Reset column widths':
                self._log("deleting marvin_library_column_widths")
                self.prefs.pop('marvin_library_column_widths')
                self.prefs.commit()
            elif action == 'Create backup':
                self.create_backup()
            elif action == 'Restore from backup':
                self.restore_from_backup()

        else:
            self._log("unrecognized action")

    def discover_iosra_status(self):
        '''
        Confirm that iOSRA is installed and not disabled
        '''
        IOSRA = 'iOS reader applications'
        # Confirm that iOSRA is installed
        installed = False
        disabled = False
        status = "Marvin not connected"
        for dp in device_plugins(include_disabled=True):
            if dp.name == IOSRA:
                installed = True
                for ddp in disabled_device_plugins():
                    if ddp.name == IOSRA:
                        disabled = True
                break

        msg = None
        if not installed:
            status = "iOSRA plugin not installed"
            msg = ('<p>Marvin XD requires the iOS reader applications plugin to be installed.</p>' +
                   '<p>Install the plugin, configure it with Marvin ' +
                   'as the preferred reader application, then restart calibre.</p>' +
                   '<p><a href="http://www.mobileread.com/forums/showthread.php?t=215624">' +
                   'iOS reader applications support</a><br/>'
                   '<a href="http://www.mobileread.com/forums/showthread.php?t=221357">' +
                   'Marvin XD support</a></p>')
        elif installed and disabled:
            status = "iOSRA plugin disabled"
            msg = ('<p>Marvin XD requires the iOS reader applications plugin to be enabled.</p>' +
                   '<p>Enable the plugin in <i>Preferences|Advanced|Plugins</i>, ' +
                   'configure it with Marvin as the preferred reader application, ' +
                   'then restart calibre.</p>' +
                   '<p><a href="http://www.mobileread.com/forums/showthread.php?t=215624">' +
                   'iOS reader applications support</a><br/>'
                   '<a href="http://www.mobileread.com/forums/showthread.php?t=221357">' +
                   'Marvin XD support</a></p>')
        if msg:
            MessageBox(MessageBox.WARNING, status, msg, det_msg='', show_copy_button=False).exec_()

        return status

    # subclass override
    def genesis(self):
        self._log_location("v%d.%d.%d" % MarvinManagerPlugin.version)

        # General initialization, occurs when calibre launches
        self.book_status_dialog = None
        self.blocking_busy = MyBlockingBusy(self.gui, "Updating Marvin Library…", size=50)
        self.connected_device = None
        self.current_location = 'library'
        self.dialog_active = False
        self.dropbox_processed = False
        self.ios = None
        self.installed_books = None
        self.marvin_content_updated = False
        self.menus_lock = threading.RLock()
        self.sync_lock = threading.RLock()
        self.indexed_library = None
        self.library_indexed = False
        self.library_last_modified = None
        self.marvin_connected = False
        self.resources_path = os.path.join(config_dir, 'plugins', "%s_resources" % self.name.replace(' ', '_'))
        if not os.path.exists(self.resources_path):
            os.makedirs(self.resources_path)
        self.virtual_library = None

        # Build a current opts object
        self.opts = self.init_options()

        # Read the plugin icons and store for potential sharing with the config widget
        icon_resources = self.load_resources(PLUGIN_ICONS)
        set_plugin_icon_resources(self.name, icon_resources)

        # Assign our menu to this action and an icon
        self.menu = QMenu(self.gui)
        self.qaction.setMenu(self.menu)
        self.qaction.setIcon(get_icon("images/disconnected.png"))
        self.qaction.triggered.connect(self.main_menu_button_clicked)
        self.menu.aboutToShow.connect(self.about_to_show_menu)

        # Init the prefs file
        self.init_prefs()

        # Populate CSS resources
        self.inflate_css_resources()

        # Populate dialog resources
        self.inflate_dialog_resources()

        # Populate the help resources
        self.inflate_help_resources()

        # Populate icon resources
        self.inflate_icon_resources()

        # Compile .ui files as needed
        CompileUI(self)

        '''
        # Hook exit in case we need to do cleanup
        atexit.register(self.onexit)
        '''

    def inflate_css_resources(self):
        '''
        Extract CSS resources from the plugin. If the file already exists,
        don't replace it as user may have edited.
        '''
        css = []
        with ZipFile(self.plugin_path, 'r') as zf:
            for candidate in zf.namelist():
                if candidate.endswith('/'):
                    continue
                if candidate.startswith('css/'):
                    css.append(candidate)
        ir = self.load_resources(css)
        for css_file in css:
            if not css_file in ir:
                continue
            fs = os.path.join(self.resources_path, css_file)
            if not os.path.exists(fs):
                if not os.path.exists(os.path.dirname(fs)):
                    os.makedirs(os.path.dirname(fs))
                with open(fs, 'wb') as f:
                    f.write(ir[css_file])

    def inflate_dialog_resources(self):
        '''
        Copy the dialog files to our resource directory
        '''
        self._log_location()

        dialogs = []
        with ZipFile(self.plugin_path, 'r') as zf:
            for candidate in zf.namelist():
                # Qt UI files
                if candidate.startswith('dialogs/') and candidate.endswith('.ui'):
                    dialogs.append(candidate)
                # Corresponding class definitions
                if candidate.startswith('dialogs/') and candidate.endswith('.py'):
                    dialogs.append(candidate)
        dr = self.load_resources(dialogs)
        for dialog in dialogs:
            if not dialog in dr:
                continue
            fs = os.path.join(self.resources_path, dialog)
            if not os.path.exists(fs):
                # If the file doesn't exist in the resources dir, add it
                if not os.path.exists(os.path.dirname(fs)):
                    os.makedirs(os.path.dirname(fs))
                with open(fs, 'wb') as f:
                    f.write(dr[dialog])
            else:
                # Is the .ui file current?
                update_needed = False
                with open(fs, 'r') as f:
                    if f.read() != dr[dialog]:
                        update_needed = True
                if update_needed:
                    with open(fs, 'wb') as f:
                        f.write(dr[dialog])

    def inflate_help_resources(self):
        '''
        Extract the help resources from the plugin
        '''
        help_resources = []
        with ZipFile(self.plugin_path, 'r') as zf:
            for candidate in zf.namelist():
                if (candidate.startswith('help/') and candidate.endswith('.html') or
                    candidate.startswith('help/images/')):
                    help_resources.append(candidate)

        rd = self.load_resources(help_resources)
        for resource in help_resources:
            if not resource in rd:
                continue
            fs = os.path.join(self.resources_path, resource)
            if os.path.isdir(fs) or fs.endswith('/'):
                continue
            if not os.path.exists(os.path.dirname(fs)):
                os.makedirs(os.path.dirname(fs))
            with open(fs, 'wb') as f:
                f.write(rd[resource])

    def inflate_icon_resources(self):
        '''
        Extract the icon resources from the plugin
        '''
        icons = []
        with ZipFile(self.plugin_path, 'r') as zf:
            for candidate in zf.namelist():
                if candidate.endswith('/'):
                    continue
                if candidate.startswith('icons/'):
                    icons.append(candidate)
        ir = self.load_resources(icons)
        for icon in icons:
            if not icon in ir:
                continue
            fs = os.path.join(self.resources_path, icon)
            if not os.path.exists(fs):
                if not os.path.exists(os.path.dirname(fs)):
                    os.makedirs(os.path.dirname(fs))
                with open(fs, 'wb') as f:
                    f.write(ir[icon])

    def init_options(self, disable_caching=False):
        """
        Build an opts object with a ProgressBar, Annotations db
        """
        opts = Struct(
            gui=self.gui,
            #icon=get_icon(PLUGIN_ICONS[0]),
            prefs=self.prefs,
            resources_path=self.resources_path,
            verbose=DEBUG)

        self._log_location()

        # Attach a Progress bar
        opts.pb = ProgressBar(parent=self.gui, window_title=self.name)

        # Instantiate the Annotations database
        opts.db = AnnotationsDB(opts, path=os.path.join(self.resources_path, 'annotations.db'))
        opts.conn = opts.db.connect()

        return opts

    def init_prefs(self):
        '''
        Set the initial default values as needed, do any needed maintenance
        '''
        pref_map = {
            'plugin_version': "%d.%d.%d" % self.interface_action_base_plugin.version,
            'injected_css': "h1\t{font-size: 1.5em;}\nh2\t{font-size: 1.25em;}\nh3\t{font-size: 1em;}"
            }

        for pm in pref_map:
            if not self.prefs.get(pm, None):
                self.prefs.set(pm, pref_map[pm])

    # subclass override
    def initialization_complete(self):
        self.rebuild_menus()

        # Subscribe to device connection events
        device_signals.device_connection_changed.connect(self.on_device_connection_changed)

    def launch_library_scanner(self):
        '''
        Call IndexLibrary() to index current_db by uuid, title
        Need a test to see if db has been updated since last run. Until then,
        optimization disabled.
        After indexing, self.library_scanner.uuid_map and .title_map are populated
        '''

        mdb = self.gui.library_view.model().db
        current_vl = mdb.data.get_base_restriction_name()

        if (self.library_last_modified == self.gui.current_db.last_modified() and
                self.indexed_library is self.gui.current_db and
                self.library_indexed and
                self.library_scanner is not None and
                self.virtual_library == current_vl):
            self._log_location("library index current for virtual library %s" % repr(current_vl))
        else:
            self._log_location("updating library index for virtual library %s" % repr(current_vl))
            self.library_scanner = IndexLibrary(self)

            if False:
                self.connect(self.library_scanner, self.library_scanner.signal, self.library_index_complete)
                QTimer.singleShot(1, self.start_library_indexing)

                # Wait for indexing to complete
                while not self.library_scanner.isFinished():
                    Application.processEvents()
            else:
                self.start_library_indexing()
                while not self.library_scanner.isFinished():
                    Application.processEvents()
                self.library_index_complete()

    # subclass override
    def library_changed(self, db):
        self._log_location(current_library_name())
        self.indexed_library = None
        self.library_indexed = False
        self.library_scanner = None
        self.library_last_modified = None

    def library_index_complete(self):
        self._log_location()
        self.library_indexed = True
        self.indexed_library = self.gui.current_db
        self.library_last_modified = self.gui.current_db.last_modified()

        # Save the virtual library name we ran the indexing against
        mdb = self.gui.library_view.model().db
        current_vl = mdb.data.get_base_restriction_name()
        self.virtual_library = self.library_scanner.active_virtual_library = current_vl

        # Reset the hash_map in case we had a prior instance from a different vl
        self.library_scanner.hash_map = None

        # Reset self.installed_books
        self.installed_books = None

        self._busy_panel_teardown()

    # subclass override
    def location_selected(self, loc):
        self._log_location(loc)
        self.current_location = loc

    def main_menu_button_clicked(self):
        '''
        Primary click on menu button
        '''
        self._log_location()
        if self.connected_device:
            if not self.dialog_active:
                self.dialog_active = True
                try:
                    self.show_installed_books()
                except AbortRequestException, e:
                    self._log(e)
                    self.book_status_dialog = None
                self.dialog_active = False
        else:
            self.show_help()

    def marvin_status_changed(self, cmd_dict):
        '''
        The Marvin driver emits a signal after completion of protocol commands.
        This method receives the notification. If the content on Marvin changed
        as a result of the operation, we need to invalidate our cache of Marvin's
        installed books.
        '''
        self.marvin_device_status_changed.emit(cmd_dict)
        command = cmd_dict['cmd']

        self._log_location(cmd_dict)
        if command in ['delete_books', 'upload_books']:
            self.marvin_content_updated = True

        if command == 'remove_books':
            self.remove_paths_from_hash_cache(cmd_dict['paths'])

    def nuke_annotations(self):
        db = self.gui.current_db
        id = db.FIELD_MAP['id']

        # Get all eligible custom fields
        all_custom_fields = db.custom_field_keys()
        custom_fields = {}
        for custom_field in all_custom_fields:
            field_md = db.metadata_for_field(custom_field)
            if field_md['datatype'] in ['comments']:
                custom_fields[field_md['name']] = {'field': custom_field,
                                                        'datatype': field_md['datatype']}

        fields = ['Comments']
        for cfn in custom_fields:
            fields.append(cfn)
        fields.sort()

        # Warn the user that we're going to do it
        title = 'Remove annotations?'
        msg = ("<p>All existing annotations will be removed from %s.</p>" %
               ', '.join(fields) +
               "<p>Proceed?</p>")
        d = MessageBox(MessageBox.QUESTION,
                       title, msg,
                       show_copy_button=False)
        if not d.exec_():
            return
        self._log_location("QUESTION: %s" % msg)

        # Show progress
        pb = ProgressBar(parent=self.gui, window_title="Removing annotations")
        total_books = len(db.data)
        pb.set_maximum(total_books)
        pb.set_value(0)
        pb.set_label('{:^100}'.format("Scanning 0 of %d" % (total_books)))
        pb.show()

        for i, record in enumerate(db.data.iterall()):
            mi = db.get_metadata(record[id], index_is_id=True)
            pb.set_label('{:^100}'.format("Scanning %d of %d" % (i, total_books)))

            # Remove user_annotations from Comments
            if mi.comments:
                soup = BeautifulSoup(mi.comments)
                uas = soup.find('div', 'user_annotations')
                if uas:
                    uas.extract()

                # Remove comments_divider from Comments
                cd = soup.find('div', 'comments_divider')
                if cd:
                    cd.extract()

                # Save stripped Comments
                mi.comments = unicode(soup)

                # Update the record
                db.set_metadata(record[id], mi, set_title=False, set_authors=False,
                                commit=True, force_changes=True, notify=True)

            # Removed user_annotations from custom fields
            for cfn in custom_fields:
                cf = custom_fields[cfn]['field']
                if True:
                    soup = BeautifulSoup(mi.get_user_metadata(cf, False)['#value#'])
                    uas = soup.findAll('div', 'user_annotations')
                    if uas:
                        # Remove user_annotations from originating custom field
                        for ua in uas:
                            ua.extract()

                        # Save stripped custom field data
                        um = mi.metadata_for_field(cf)
                        stripped = unicode(soup)
                        if stripped == u'':
                            stripped = None
                        um['#value#'] = stripped
                        mi.set_user_metadata(cf, um)

                        # Update the record
                        db.set_metadata(record[id], mi, set_title=False, set_authors=False,
                                        commit=True, force_changes=True, notify=True)
                else:
                    um = mi.metadata_for_field(cf)
                    um['#value#'] = None
                    mi.set_user_metadata(cf, um)
                    # Update the record
                    db.set_metadata(record[id], mi, set_title=False, set_authors=False,
                                    commit=True, force_changes=True, notify=True)

            pb.increment()

        # Hide the progress bar
        pb.hide()

        # Update the UI
        updateCalibreGUIView()

    def onexit(self):
        '''
        Called as calibre is exiting.
        '''
        self._log_location()

    def on_device_connection_changed(self, is_connected):
        '''
        self.connected_device is the handle to the driver.
        '''
        self.plugin_device_connection_changed.emit(is_connected)

        if is_connected:
            self.connected_device = self.gui.device_manager.device
            self.marvin_connected = (hasattr(self.connected_device, 'ios_reader_app') and
                                     self.connected_device.ios_reader_app == 'Marvin')
            if self.marvin_connected:

                self._log_location(self.connected_device.gui_name)

                # Init libiMobileDevice
                self.ios = libiMobileDevice(verbose=self.prefs.get('debug_libimobiledevice', False))
                self._log("mounting %s" % self.connected_device.app_id)
                self.ios.mount_ios_app(app_id=self.connected_device.app_id)

                # Change our icon
                self.qaction.setIcon(get_icon("images/connected.png"))

                # Subscribe to Marvin driver change events
                self.connected_device.marvin_device_signals.reader_app_status_changed.connect(
                    self.marvin_status_changed)

                # Explore connected.xml for <has_password>
                connected_fs = getattr(self.connected_device, 'connected_fs', None)
                if connected_fs and self.ios.exists(connected_fs):

                    # Wait for the driver to be silent before exploring connected.xml
                    while self.connected_device.get_busy_flag():
                        Application.processEvents()
                    self.connected_device.set_busy_flag(True)

                    # connection.keys(): ['timestamp', 'marvin', 'device', 'system']
                    connection = etree.fromstring(self.ios.read(connected_fs))
                    #self._log(etree.tostring(connection, pretty_print=True))
                    self._log("%s running iOS %s" % (connection.get('device'), connection.get('system')))

                    self.has_password = False
                    chp = connection.find('has_password')
                    if chp is not None:
                        self.has_password = bool(chp.text == "true")
                    self._log("has_password: %s" % self.has_password)

                    self.process_updates()

                    self.connected_device.set_busy_flag(False)
            else:
                self._log("Marvin not connected")

        else:
            if self.marvin_connected:
                self._log_location("device disconnected")

                # Change our icon
                self.qaction.setIcon(get_icon("images/disconnected.png"))

                # Close libiMobileDevice connection, reset references to mounted device
                self.ios.disconnect_idevice()
                self.ios = None
                self.connected_device.marvin_device_signals.reader_app_status_changed.disconnect()
                self.connected_device = None

                # Invalidate the library hash map, as library contents may change before reconnection
                #self.library_scanner.hash_map = None

                # Clear has_password
                self.has_password = None

                # Dump our saved copy of installed_books
                self.installed_books = None

                self.marvin_connected = False

        self.rebuild_menus()

    """
    def process_dropbox_sync_records(self):
        '''
        Scan local Dropbox folder for metadata update records
        Show progress bar in dialog box reporting titles
        '''
        self._log_location()

        self.launch_library_scanner()
        foo = PullDropboxUpdates(self)
    """

    def process_updates(self):
        '''
        Handle any version-related maintenance here
        '''
        def _log_update():
            self._log_location("updating prefs from %s to %s" %
                (prefs_version, "%d.%d.%d" % self.interface_action_base_plugin.version))

        prefs_version = self.prefs.get("plugin_version", "0.0.0")
        updated = False

        # Clean up JSON file < v1.1.0
        if prefs_version < "1.1.0":
            _log_update()
            for obsolete_setting in [
                'annotations_field_comboBox', 'annotations_field_lookup',
                'collection_field_comboBox', 'collection_field_lookup',
                'date_read_field_comboBox', 'date_read_field_lookup',
                'progress_field_comboBox', 'progress_field_lookup',
                'read_field_comboBox', 'read_field_lookup',
                'reading_list_field_comboBox', 'reading_list_field_lookup',
                'word_count_field_comboBox', 'word_count_field_lookup']:
                if self.prefs.get(obsolete_setting, None) is not None:
                    self._log("removing obsolete entry '{0}'".format(obsolete_setting))
                    self.prefs.__delitem__(obsolete_setting)
            updated = True

        # Delete obsolete cached hashes
        if prefs_version < "1.2.0":
            _log_update()
            self._log("Deleting calibre and Marvin hashes")
            self.developer_utilities('Delete calibre hashes')
            self.developer_utilities('Delete Marvin hashes')
            updated = True

        # Change CSS prefs 'Timestamp' to 'Location'
        if prefs_version <= "1.2.0":
            appearance_css = self.prefs.get('appearance_css', None)
            if appearance_css is not None:
                for element in appearance_css:
                    if element['name'] == 'Timestamp':
                        element['name'] = 'Location'
                        self.prefs.set('appearance_css', appearance_css)
                        updated = True
                        _log_update()
                        self._log("changing appearance_css 'Timestamp' to 'Location'")
                    break

        if updated:
            self.prefs.set('plugin_version', "%d.%d.%d" % self.interface_action_base_plugin.version)

    def rebuild_menus(self):
        self._log_location()
        with self.menus_lock:
            m = self.menu
            m.clear()

            # Add 'About…'
            ac = self.create_menu_item(m, 'About' + '…')
            ac.triggered.connect(self.show_about)
            m.addSeparator()

            # Add menu options for connected Marvin, Dropbox syncing when no connection
            marvin_connected = False

            dropbox_syncing_enabled = self.prefs.get('dropbox_syncing', False)
            process_dropbox = False

            if self.connected_device and hasattr(self.connected_device, 'ios_reader_app'):
                if (self.connected_device.ios_reader_app == 'Marvin' and
                        self.connected_device.ios_connection['connected'] is True):
                    self._log("Marvin connected")
                    marvin_connected = True
                    ac = self.create_menu_item(m, 'Explore Marvin Library', image=I("dialog_information.png"))
                    ac.triggered.connect(self.show_installed_books)

                else:
                    self._log("Marvin not connected")
                    ac = self.create_menu_item(m, 'Marvin not connected')
                    ac.setEnabled(False)

            elif False and not self.connected_device:
                ac = self.create_menu_item(m, 'Update metadata via Dropbox')
                ac.triggered.connect(self.process_dropbox_sync_records)

                # If syncing enabled in Config dialog, automatically process 1x
                if dropbox_syncing_enabled and not self.dropbox_processed:
                    process_dropbox = True
            else:
                iosra_status = self.discover_iosra_status()
                self._log(iosra_status)
                ac = self.create_menu_item(m, iosra_status)
                ac.setEnabled(False)

            m.addSeparator()

            # Add 'Customize plugin…'
            ac = self.create_menu_item(m, 'Customize plugin' + '…', image=I("config.png"))
            ac.triggered.connect(self.show_configuration)

            m.addSeparator()

            # Add 'Help'
            ac = self.create_menu_item(m, 'Help', image=I('help.png'))
            ac.triggered.connect(self.show_help)

            # If Alt/Option key pressed, show Developer submenu
            modifiers = Application.keyboardModifiers()
            if bool(modifiers & Qt.AltModifier):
                m.addSeparator()
                self.developer_menu = m.addMenu(QIcon(I('config.png')),
                                                "Developer…")
                action = 'Delete calibre hashes'
                ac = self.create_menu_item(self.developer_menu, action, image=I('trash.png'))
                ac.triggered.connect(partial(self.developer_utilities, action))

                action = 'Delete Marvin hashes'
                ac = self.create_menu_item(self.developer_menu, action, image=I('trash.png'))
                ac.triggered.connect(partial(self.developer_utilities, action))
                ac.setEnabled(marvin_connected)

                action = 'Nuke annotations'
                ac = self.create_menu_item(self.developer_menu, action, image=I('trash.png'))
                ac.triggered.connect(partial(self.developer_utilities, action))

                action = 'Reset column widths'
                ac = self.create_menu_item(self.developer_menu, action, image=I('trash.png'))
                ac.triggered.connect(partial(self.developer_utilities, action))

                # Backup/restore
                if hasattr(self.connected_device, 'local_db_path'):
                    m.addSeparator()
                    #if self.connected_device.marvin_version >= (2, 7, 0):
                    action = 'Create backup'
                    icon = QIcon(os.path.join(self.resources_path, 'icons', 'sync_collections.png'))
                    ac = self.create_menu_item(m, action, image=icon)
                    ac.triggered.connect(partial(self.developer_utilities, action))

                    action = 'Restore from backup'
                    icon = QIcon(os.path.join(self.resources_path, 'icons', 'sync_collections.png'))
                    ac = self.create_menu_item(m, action, image=icon)
                    ac.triggered.connect(partial(self.developer_utilities, action))

            # Process Dropbox sync records automatically once only.
            if process_dropbox:
                self.process_dropbox_sync_records()
                self.dropbox_processed = True

    def remove_paths_from_hash_cache(self, paths):
        '''
        Remove cached hashes when iOSRA deletes books
        '''
        self._log_location()
        rhc = '/'.join([BookStatusDialog.REMOTE_CACHE_FOLDER,
            BookStatusDialog.HASH_CACHE_FS])

        if self.ios.exists(rhc):
            # Copy remote hash_cache to local file
            lhc = os.path.join(self.connected_device.temp_dir,
                BookStatusDialog.HASH_CACHE_FS)
            with open(lhc, 'wb') as out:
                self.ios.copy_from_idevice(str(rhc), out)

            # Load hash_cache
            with open(lhc, 'rb') as hcf:
                hash_cache = pickle.load(hcf)

            # Scan the cached hashes
            updated_hash_cache = {}
            updated = False
            for key, value in hash_cache.items():
                if key in paths:
                    self._log("%s removed from hash_cache" % key)
                    updated = True
                else:
                    updated_hash_cache[key] = value

            if updated:
                # Write the edited hash_cache locally
                with open(lhc, 'wb') as hcf:
                    pickle.dump(updated_hash_cache, hcf, pickle.HIGHEST_PROTOCOL)

                # Copy to iDevice
                self.ios.remove(str(rhc))
                self.ios.copy_to_idevice(lhc, str(rhc))

    def reset_marvin_library(self):
        self._log_location("not implemented")

    def restore_from_backup(self):
        '''
        Invoke RestoreBackup() in a separate thread
        Display a dialog telling user how to complete restore
        '''
        self._log_location()

        # Get the backup file
        source = QFileDialog.getOpenFileName(
            self.gui,
            "Select Marvin backup file to restore",
            self.prefs.get('backup_folder', os.path.expanduser("~")),
            "*.backup")
        if source:
            # Save the selected backup folder
            self.prefs.set('backup_folder', os.path.dirname(str(source)))

            copy_operation = RestoreBackup(self, source)
            self._busy_panel_setup("<p>Copying {0} ({1:,} bytes)<br/>to {2}…</p>".format(
                os.path.basename(str(source)),
                copy_operation.src_size,
                self.ios.device_name))
            copy_operation.start()
            while not copy_operation.isFinished():
                Application.processEvents()
            self._busy_panel_teardown()

            s_size = copy_operation.src_size
            d_size = copy_operation.dst_size

            # Verify transferred size
            if s_size == d_size:
                # Delete cached Marvin data
                self.developer_utilities('Delete Marvin hashes')

                self._log("Backup verifies: {0:,} bytes".format(s_size))
                # Display dialog detailing how to complete restore
                title = "Restore Marvin from backup"
                msg = ('<p>To complete the restore process on {0}:</p>'
                       '<ul>'
                       '<li>Touch <b>Disconnect</b> on the calibre connector</li>'
                       '<li>Press the Home button to return to the Home screen</li>'
                       '<li>Force-quit Marvin</li>'
                       '<li>Restart Marvin</li>'
                       '</ul>'
                       ).format(self.ios.device_name)
                d = MessageBox(MessageBox.INFO, title, msg, det_msg='', show_copy_button=False).exec_()

            else:
                self._log("Backup does not verify: source {0:,} dest {1:,}".format(
                    s_size, d_size))
                title = "Restore unsuccessful"
                msg = ('<p>Backup does not verify.</p>'
                       '<p>'
                       '<tt>Src: {0:,}</tt><br/>'
                       '<tt>Dst: {1:,}</tt></p>'
                       '<p>Restore cancelled.</p>'
                      ).format(s_size, d_size)
                d = MessageBox(MessageBox.WARNING, title, msg, det_msg='', show_copy_button=False).exec_()
                self._log("Restore cancelled")

    def show_configuration(self):
        self.interface_action_base_plugin.do_user_config(self.gui)

    def show_about(self):
        version = self.interface_action_base_plugin.version
        title = "%s v %d.%d.%d" % (self.name, version[0], version[1], version[2])
        msg = ('<p>To learn more about this plugin, visit the '
               '<a href="http://www.mobileread.com/forums/showthread.php?t=221357">Marvin XD</a> '
               'support thread at MobileRead’s Calibre forum.</p>')
        text = get_resources('about.txt')
        text = text.decode('utf-8')
        d = MessageBox(MessageBox.INFO, title, msg, det_msg=text, show_copy_button=False)
        d.exec_()

    def show_help(self):
        path = os.path.join(self.resources_path, 'help/help.html')
        open_url(QUrl.fromLocalFile(path))

    def show_installed_books(self):
        '''
        Show Marvin Library spreadsheet
        '''
        self._log_location()

        if self.connected_device.version < self.minimum_ios_driver_version:
            title = "Update required"
            msg = "<p>{0} requires v{1}.{2}.{3} (or later) of the iOS reader applications device driver.</p>".format(
                self.name,
                self.minimum_ios_driver_version[0],
                self.minimum_ios_driver_version[1],
                self.minimum_ios_driver_version[2])
            MessageBox(MessageBox.INFO, title, msg, det_msg='', show_copy_button=False).exec_()
        else:
            self.launch_library_scanner()

            # Assure that Library is active view. Avoids problems with _delete_books
            restore_to = None
            if self.current_location != 'library':
                restore_to = self.current_location
                self.gui.location_selected('library')

            self.book_status_dialog = BookStatusDialog(self, 'marvin_library')
            self.book_status_dialog.initialize(self)
            self._log_location("{0} books".format(len(self.book_status_dialog.installed_books)))
            self.book_status_dialog.exec_()

            # Keep a copy of installed_books in case user reopens w/o disconnect
            self.installed_books = self.book_status_dialog.installed_books

            # Restore the Device view if active before MXD window launched
            if restore_to:
                self.gui.location_selected(restore_to)

            self.book_status_dialog = None

    # subclass override
    def shutting_down(self):
        self._log_location()

    def start_library_indexing(self):
        self._log_location()
        self._busy_panel_setup("Indexing calibre library…")
        self.library_scanner.start()

    def _busy_panel_setup(self, title, show_cancel=False):
        '''
        '''
        self._log_location()
        Application.setOverrideCursor(QCursor(Qt.WaitCursor))
        self.busy_window = MyBlockingBusy(self.gui, title, size=60, show_cancel=show_cancel)
        self.busy_window.start()
        self.busy_window.show()

    def _busy_panel_teardown(self):
        '''
        '''
        self._log_location()
        if self.busy_window is not None:
            self.busy_window.stop()
            self.busy_window.accept()
            self.busy_window = None
        try:
            Application.restoreOverrideCursor()
        except:
            pass