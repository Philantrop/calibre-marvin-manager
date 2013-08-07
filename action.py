#!/usr/bin/env python
# coding: utf-8

from __future__ import (unicode_literals, division, absolute_import,
                        print_function)

__license__ = 'GPL v3'
__copyright__ = '2013, Greg Riker <griker@hotmail.com>'
__docformat__ = 'restructuredtext en'

import os, sys, threading

from functools import partial
from lxml import etree, html
from zipfile import ZipFile

from PyQt4.Qt import (Qt, QCursor, QIcon, QMenu, QTimer, QUrl,
                      pyqtSignal)

from calibre.constants import DEBUG
from calibre.devices.idevice.libimobiledevice import libiMobileDevice
from calibre.gui2 import Application, open_url
from calibre.gui2.actions import InterfaceAction
from calibre.gui2.device import device_signals
from calibre.gui2.dialogs.message_box import MessageBox
from calibre.devices.usbms.driver import debug_print
from calibre.utils.config import config_dir

from calibre_plugins.marvin_manager import MarvinManagerPlugin
from calibre_plugins.marvin_manager.book_status import BookStatusDialog
from calibre_plugins.marvin_manager.common_utils import (AbortRequestException,
    CompileUI, IndexLibrary, MyBlockingBusy, ProgressBar, Struct,
    get_icon, set_plugin_icon_resources)
import calibre_plugins.marvin_manager.config as cfg

# The first icon is the plugin icon, referenced by position.
# The rest of the icons are referenced by name
PLUGIN_ICONS = ['images/connected.png', 'images/disconnected.png']

class MarvinManagerAction(InterfaceAction):

    # Location reporting template
    LOCATION_TEMPLATE = "{cls}:{func}({arg1}) {arg2}"

    icon = PLUGIN_ICONS[0]
    minimum_ios_driver_version = (1, 0, 5)
    name = 'Marvin XD'
    prefs = cfg.plugin_prefs
    verbose = prefs.get('debug_plugin', False)

    # Declare the main action associated with this plugin
    action_spec = ('Marvin XD', None, None, None)
    #popup_type = QToolButton.InstantPopup
    action_add_menu = True
    action_menu_clone_qaction = True

    marvin_device_status_changed = pyqtSignal(str)
    plugin_device_connection_changed = pyqtSignal(object)

    def about_to_show_menu(self):
        self.rebuild_menus()

    def backup_restore(self):
        self._log_location("not implemented")

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
        if action in ['Delete calibre hashes', 'Delete Marvin hashes']:
            if action == 'Delete Marvin hashes':
                hash_cache = 'content_hashes.zip'
                remote_cache_folder = '/'.join(['/Library', 'calibre.mm'])
                rhc = b'/'.join([remote_cache_folder, hash_cache])

                if self.ios.exists(rhc):
                    self.ios.remove(rhc)
                    self._log("remote hash cache at %s deleted" % rhc)
            elif action == 'Delete calibre hashes':
                self.gui.current_db.delete_all_custom_book_data('epub_hash')
                self._log("cached epub hashes deleted")
                # Invalidate the library hash map, as library contents may change before reconnection
                self.library_scanner.hash_map = None

        else:
            self._log("unrecognized action")

    # subclass override
    def genesis(self):
        self._log_location("v%d.%d.%d" % MarvinManagerPlugin.version)

        # General initialization, occurs when calibre launches
        self.book_status_dialog = None
        self.blocking_busy = MyBlockingBusy(self.gui, "Updating Marvin Library…", size=50)
        self.connected_device = None
        self.ios = None
        self.installed_books = None
        self.marvin_content_updated = False
        self.menus_lock = threading.RLock()
        self.sync_lock = threading.RLock()
        self.indexed_library = None
        self.library_indexed = False
        self.library_last_modified = None
        self.resources_path = os.path.join(config_dir, 'plugins', "%s_resources" % self.name.replace(' ', '_'))

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

        # Populate dialog resources
        self.inflate_dialog_resources()

        # Populate the help resources
        self.inflate_help_resources()

        # Populate icon resources
        self.inflate_icon_resources()

        # Compile .ui files as needed
        CompileUI(self)

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
                if candidate == 'help/help.html' or candidate.startswith('help/images/'):
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
        Build an opts object with a ProgressBar
        """
        opts = Struct(
            gui=self.gui,
            #icon=get_icon(PLUGIN_ICONS[0]),
            prefs=self.prefs,
            resources_path=self.resources_path,
            verbose=DEBUG)

        opts['pb'] = ProgressBar(parent=self.gui, window_title=self.name)
        self._log_location()
        return opts

    def init_prefs(self):
        '''
        Set the initial default values as needed
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
        if (self.library_last_modified == self.gui.current_db.last_modified() and
                self.indexed_library is self.gui.current_db and
                self.library_indexed and
                self.library_scanner is not None):
            self._log_location("library index current")
        else:
            self._log_location("updating library index")
            self.library_scanner = IndexLibrary(self)
            self.connect(self.library_scanner, self.library_scanner.signal, self.library_index_complete)
            QTimer.singleShot(1, self.start_library_indexing)

    # subclass override
    def library_changed(self, db):
        self._log_location()
        self.indexed_library = None
        self.library_indexed = False
        self.library_scanner = None
        self.library_last_modified = None

    def library_index_complete(self):
        self._log_location()
        self.library_indexed = True
        self.indexed_library = self.gui.current_db
        self.library_last_modified = self.gui.current_db.last_modified()
        self._busy_operation_teardown()

    def main_menu_button_clicked(self):
        '''
        Primary click on menu button
        '''
        self._log_location()
        if self.connected_device:
            if not self.book_status_dialog:
                try:
                    self.show_installed_books()
                except AbortRequestException, e:
                    self._log(e)
                    self.book_status_dialog = None
        else:
            self.show_help()

    def marvin_status_changed(self, command):
        '''
        The Marvin driver emits a signal after completion of protocol commands.
        This method receives the notification. If the content on Marvin changed
        as a result of the operation, we need to invalidate our cache of Marvin's
        installed books.
        '''
        self.marvin_device_status_changed.emit(command)

        self._log_location(command)
        if command in ['delete_books', 'upload_books']:
            self.marvin_content_updated = True

    def on_device_connection_changed(self, is_connected):
        '''
        self.connected_device is the handle to the driver.
        '''
        self.plugin_device_connection_changed.emit(is_connected)
        if is_connected:
            self.connected_device = self.gui.device_manager.device

            self._log_location(self.connected_device.gui_name)

            # Init libiMobileDevice
            self.ios = libiMobileDevice(log=self._log,
                                        verbose=self.prefs.get('debug_libimobiledevice', False))
            self._log("mounting %s" % self.connected_device.app_id)
            self.ios.mount_ios_app(app_id=self.connected_device.app_id)

            # Change our icon
            self.qaction.setIcon(get_icon("images/connected.png"))

            # Subscribe to Marvin driver change events
            self.connected_device.marvin_device_signals.reader_app_status_changed.connect(
                self.marvin_status_changed)

            # If we've already built the hash map and the library hasn't changed, don't rescan
            if (hasattr(self.connected_device, 'ios_reader_app') and
                    self.connected_device.ios_reader_app == 'Marvin'):
                if self.indexed_library is not None and self.library_indexed:
                    self._log("library already indexed")
                else:
                    self.launch_library_scanner()

            # Explore connected.xml for <has_password>
            connected_fs = getattr(self.connected_device, 'connected_fs', None)
            if connected_fs and self.ios.exists(connected_fs):

                # Wait for the driver to be silent to explore connected.xml
                while self.connected_device.get_busy_flag():
                    Application.processEvents()
                self.connected_device.set_busy_flag(True)

                # connection.keys(): ['timestamp', 'marvin', 'device', 'system']
                connection = etree.fromstring(self.ios.read(connected_fs))
                #self._log(etree.tostring(connection, pretty_print=True))
                self._log_location("%s running iOS %s" % (connection.get('device'), connection.get('system')))

                has_password = connection.find('has_password')
                if has_password is not None:
                    self.has_password = bool(has_password.text == "true")
                    self._log("self.has_password: %s" % self.has_password)

                self.connected_device.set_busy_flag(False)

        else:
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

        self.rebuild_menus()

    def rebuild_menus(self):
        self._log_location()
        with self.menus_lock:
            m = self.menu
            m.clear()

            # Add 'About…'
            ac = self.create_menu_item(m, 'About' + '…')
            ac.triggered.connect(self.show_about)
            m.addSeparator()

            # Add menu options for connected Marvin
            marvin_connected = False
            if self.connected_device:
                if (self.connected_device.ios_reader_app == 'Marvin' and
                        self.connected_device.ios_connection['connected'] is True):
                    self._log("Marvin connected")
                    marvin_connected = True
                    ac = self.create_menu_item(m, 'Explore Marvin Library', image=I("dialog_information.png"))
                    ac.triggered.connect(self.show_installed_books)

                    if False:
                        ac = self.create_menu_item(m, 'Backup or Restore Library', image=I("swap.png"))
                        ac.triggered.connect(self.backup_restore)
                        ac.setEnabled(False)

                        ac = self.create_menu_item(m, 'Reset Marvin Library', image=I("trash.png"))
                        ac.triggered.connect(self.reset_marvin_library)
                        ac.setEnabled(False)

                else:
                    self._log("Marvin not connected")
                    ac = self.create_menu_item(m, 'Marvin not connected')
                    ac.setEnabled(False)
            else:
                self._log("Marvin not connected")
                ac = self.create_menu_item(m, 'Marvin not connected')
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

    def reset_marvin_library(self):
        self._log_location("not implemented")

    def show_configuration(self):
        self.interface_action_base_plugin.do_user_config(self.gui)

    def show_about(self):
        version = self.interface_action_base_plugin.version
        title = "%s v %d.%d.%d" % (self.name, version[0], version[1], version[2])
        msg = ('<p>To learn more about this plugin, visit the '
               '<a href="http://www.mobileread.com/forums/showthread.php?t=205062">THIS NEEDS TO BE ADDED</a> '
               'at MobileRead’s Calibre forum.</p>')
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
            self.book_status_dialog = BookStatusDialog(self, 'marvin_library')
            self.book_status_dialog.initialize(self)
            self._log_location("BookStatus initialized")
            self.book_status_dialog.exec_()

            # Keep a copy of installed_books in case user reopens w/o disconnect
            self.installed_books = self.book_status_dialog.installed_books

            self.book_status_dialog = None

    # subclass override
    def shutting_down(self):
        self._log_location()

    def start_library_indexing(self):
        self._log_location()
        self._busy_operation_setup("Indexing calibre library…")
        self.library_scanner.start()

    def _busy_operation_setup(self, title, show_cancel=False):
        '''
        '''
        self._log_location()
        Application.setOverrideCursor(QCursor(Qt.WaitCursor))
        self.busy_window = MyBlockingBusy(self.gui, title, size=60, show_cancel=show_cancel)
        self.busy_window.start()
        self.busy_window.show()

    def _busy_operation_teardown(self):
        '''
        '''
        self._log_location()
        self.busy_window.stop()
        self.busy_window.accept()
        self.busy_window = None
        Application.restoreOverrideCursor()

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

