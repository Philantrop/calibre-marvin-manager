#!/usr/bin/env python
# coding: utf-8

from __future__ import (unicode_literals, division, absolute_import,
                        print_function)

__license__ = 'GPL v3'
__copyright__ = '2013, Greg Riker <griker@hotmail.com>'
__docformat__ = 'restructuredtext en'

import atexit, cPickle as pickle, glob, hashlib, json, os, re, shutil, sqlite3, sys
import tempfile, threading, time

from datetime import datetime
from dateutil import tz
from functools import partial
from lxml import etree, html

from PyQt4.Qt import (Qt, QApplication, QCursor, QFileDialog, QFont, QIcon,
                      QMenu, QTimer, QUrl,
                      pyqtSignal)

from calibre import browser
from calibre.constants import DEBUG
from calibre.customize.ui import device_plugins, disabled_device_plugins
from calibre.devices.idevice.libimobiledevice import libiMobileDevice
from calibre.devices.usbms.driver import debug_print
from calibre.ebooks.BeautifulSoup import BeautifulSoup, UnicodeDammit
from calibre.gui2 import Application, info_dialog, open_url, question_dialog, warning_dialog
from calibre.gui2.actions import InterfaceAction
from calibre.gui2.device import device_signals
from calibre.gui2.dialogs.message_box import MessageBox
from calibre.library import current_library_name
from calibre.ptempfile import (PersistentTemporaryDirectory, PersistentTemporaryFile,
    TemporaryDirectory, TemporaryFile)
from calibre.utils.config import config_dir, JSONConfig
from calibre.utils.zipfile import ZipFile, ZIP_STORED, is_zipfile

from calibre_plugins.marvin_manager import MarvinManagerPlugin
from calibre_plugins.marvin_manager.annotations_db import AnnotationsDB
from calibre_plugins.marvin_manager.book_status import BookStatusDialog
from calibre_plugins.marvin_manager.common_utils import (AbortRequestException,
    Book, CommandHandler, CompileUI, IndexLibrary, Logger,
    MoveBackup, MyBlockingBusy, PluginMetricsLogger,
    ProgressBar, RestoreBackup, Struct,
    from_json, get_icon, set_plugin_icon_resources, to_json, updateCalibreGUIView)
import calibre_plugins.marvin_manager.config as cfg
#from calibre_plugins.marvin_manager.dropbox import PullDropboxUpdates

# The first icon is the plugin icon, referenced by position.
# The rest of the icons are referenced by name
PLUGIN_ICONS = ['images/connected.png', 'images/disconnected.png']

class MarvinManagerAction(InterfaceAction, Logger):

    INSTALLED_BOOKS_SNAPSHOT = "installed_books.zip"
    REMOTE_CACHE_FOLDER = '/'.join(['/Library', 'calibre.mm'])
    UTF_8_BOM = r'\xef\xbb\xbf'

    icon = PLUGIN_ICONS[0]
    # #mark ~~~ Minimum iOSRA version required ~~~
    minimum_ios_driver_version = (1, 3, 6)
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

    def compare_mainDb_profiles(self, stored_mainDb_profile):
        '''
        '''
        current_mainDb_profile = self.profile_db()
        matched = True
        for key in sorted(current_mainDb_profile.keys()):
            if current_mainDb_profile[key] != stored_mainDb_profile[key]:
                matched = False
                break

        # Display mainDb_profile mismatch
        if not matched:
            self._log_location()
            self._log("   {0:20} {1:37} {2:37}".format('key', 'stored', 'current'))
            self._log("{0:—^23} {1:—^37} {2:—^37}".format('', '', ''))
            self._log("{0}  {1:20} {2:<37} {3:<37}".format(
                'x' if stored_mainDb_profile['device'] != current_mainDb_profile['device'] else ' ',
                'device',
                stored_mainDb_profile['device'],
                current_mainDb_profile['device']))
            keys = current_mainDb_profile.keys()
            keys.pop(keys.index('device'))
            for key in sorted(keys):
                self._log("{0}  {1:20} {2:<37} {3:<37}".format(
                    'x' if stored_mainDb_profile[key] != current_mainDb_profile[key] else ' ',
                    key,
                    repr(stored_mainDb_profile[key]),
                    repr(current_mainDb_profile[key])))
        return matched

    def create_remote_backup(self):
        '''
        iPad1:      500 books in 90 seconds - 5.5 books/second
        iPad Mini:  500 books in 64 seconds - 7.8 books/second
        1) Issue backup command to Marvin
        2) Get destination directory
        3) Move generated backup from /Documents/Backup to local storage
        '''
        IOS_READ_RATE = 7500000  # 11.8 - 16 MB/sec OS X, Windows 6.6MB
        TIMEOUT_PADDING_FACTOR = 1.5
        WORST_CASE_ARCHIVE_RATE = 1800000   # MB/second

        backup_folder = b'/'.join(['/Documents', 'Backup'])
        backup_target = backup_folder + '/marvin.backup'
        last_backup_folder = self.prefs.get('backup_folder', os.path.expanduser("~"))

        def _confirm_overwrite(backup_target):
            '''
            Check for existing backup before overwriting
            stats['st_mtime'], stats['st_size']
            Return True: continue
            Return False: cancel
            '''
            stats = self.ios.exists(backup_target, silent=True)
            if stats:
                d = datetime.fromtimestamp(float(stats['st_mtime']))
                friendly_date = d.strftime("%A, %B %d, %Y")
                friendly_time = d.strftime("%I:%M %p")

                title = "A backup already exists!"
                msg = ('<p>There is an existing backup of your Marvin library '
                        'created {0} at {1}.</p>'
                       '<p>Proceeding with this backup will '
                       'overwrite the existing backup.</p>'
                       '<p>Proceed?</p>'.format(friendly_date, friendly_time))
                dlg = MessageBox(MessageBox.QUESTION, title, msg,
                                 parent=self.gui, show_copy_button=False)
                return dlg.exec_()
            return True

        def _confirm_lengthy_backup(total_books, total_seconds):
            '''
            If this is going to take some time, warn the user
            '''
            estimated_time = self.format_time(total_seconds, show_fractional=False)
            self._log("estimated time to backup {0} books: {1}".format(
                total_books, estimated_time))

            # Confirm that user wants to proceed given estimated time to completion
            book_descriptor = "books" if total_books > 1 else "book"
            title = "Estimated time to create backup"
            msg = ("<p>Creating a backup of " +
                   "{0:,} {1} in your Marvin library ".format(total_books, book_descriptor) +
                   "may take as long as {0}, depending on your iDevice.</p>".format(estimated_time) +
                   "<p>Proceed?</p>")
            dlg = MessageBox(MessageBox.QUESTION, title, msg,
                             parent=self.gui, show_copy_button=False)
            return dlg.exec_()

        def _estimate_size():
            '''
            Estimate uncompressed size of backup
            backup.xml approximately 4kB
            '''
            SMALL_COVER_AVERAGE = 25000
            LARGE_COVER_AVERAGE = 100000

            estimated_size = 0

            books_size = 0
            for book in self.connected_device.cached_books:
                books_size += self.connected_device.cached_books[book]['size']
                books_size += SMALL_COVER_AVERAGE
                books_size += LARGE_COVER_AVERAGE

            estimated_size += books_size

            # Add size of mainDb
            mdbs = os.stat(self.connected_device.local_db_path).st_size
            estimated_size += mdbs

            self._log("estimated size of uncompressed backup: {:,}".format(estimated_size))
            return estimated_size

        # ~~~ Entry point ~~~
        self._log_location()
        analytics = []
        mainDb_profile = self.profile_db()
        estimated_size = _estimate_size()
        total_seconds = int(estimated_size/WORST_CASE_ARCHIVE_RATE)
        timeout = int(total_seconds * TIMEOUT_PADDING_FACTOR)
        estimated_time = self.format_time(total_seconds)

        if timeout > CommandHandler.WATCHDOG_TIMEOUT:
            if not _confirm_lengthy_backup(mainDb_profile['Books'], total_seconds):
                return
        else:
            timeout = CommandHandler.WATCHDOG_TIMEOUT

        if not _confirm_overwrite(backup_target):
            self._log("user declined to overwrite existing backup")
            return

        # Construct the phase 1 ProgressBar
        busy_panel_args = {'book_count': mainDb_profile['Books'],
                           'destination': 'destination folder',
                           'device': self.ios.device_name,
                           'estimated_time': estimated_time}
        BACKUP_MSG_1 = ('<ol style="margin-right:1.5em">'
                        '<li style="margin-bottom:0.5em">Preparing backup of '
                        '{book_count:,} books …</li>'
                        '<li style="color:#bbb;margin-bottom:0.5em">Select destination folder to store backup</li>'
                        '<li style="color:#bbb">Move backup from {device} to {destination}</li>'
                        '</ol>')
        pb = ProgressBar(alignment=Qt.AlignLeft,
                         label=BACKUP_MSG_1.format(**busy_panel_args),
                         parent=self.gui,
                         window_title="Creating backup of {0}".format(self.ios.device_name))

        # Init the command handler
        ch = CommandHandler(self, pb=pb)
        ch.init_pb(total_seconds)
        ch.construct_general_command('backup')

        start_time = time.time()

        # Dispatch the command
        pb.show()
        ch.issue_command(timeout_override=timeout)
        pb.hide()

        actual_time = time.time() - start_time
        args = {'estimated_size': estimated_size,
                'book_count': mainDb_profile['Books'],
                'estimated_time': self.format_time(total_seconds),
                'actual_time': self.format_time(actual_time),
                'pct_complete': pb.get_pct_complete(),
                'archive_rate': estimated_size/actual_time,
                'estimated_archive_rate': WORST_CASE_ARCHIVE_RATE}
        analytics.append((
            '1. Preparing backup\n'
            '         estimated size: {estimated_size:,}\n'
            '             book count: {book_count:,}\n'
            '         estimated time: {estimated_time}\n'
            '            actual time: {actual_time} ({pct_complete}%)\n'
            '      est. archive rate: {estimated_archive_rate:,}\n'
            '    actual archive rate: {archive_rate:,.0f} bytes/second\n'
            ).format(**args))
        del pb

        if ch.results['code']:
            self._log("results: %s" % ch.results)
            title = "Backup unsuccessful"
            msg = ('<p>Unable to create backup of {0}.</p>'
                   '<p>Click <b>Show details</b> for more information.</p>').format(
                   self.ios.device_name)
            det_msg = ch.results['details'] + '\n\n'
            det_msg += self.format_device_profile()

            # Set dialog det_msg to monospace
            dialog = warning_dialog(self.gui, title, msg, det_msg=det_msg)
            font = QFont('monospace')
            font.setFixedPitch(True)
            dialog.det_msg.setFont(font)
            dialog.exec_()
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
                "Select destination folder to store backup",
                last_backup_folder,
                QFileDialog.ShowDirsOnly | QFileDialog.DontResolveSymlinks))

            if destination_folder:
                # Qt apparently sometimes returns a file within the selected directory,
                # rather than the directory itself. Validate destination_folder
                if not os.path.isdir(destination_folder):
                    destination_folder = os.path.dirname(destination_folder)

                # Display status
                busy_panel_args['backup_size'] = int(int(stats['st_size'])/(1024*1024))
                busy_panel_args['destination'] = destination_folder

                BACKUP_MSG_3 = ('<ol style="margin-right:1.5em">'
                                '<li style="color:#bbb;margin-bottom:0.5em">Backup of {book_count:,} '
                                'books prepared</li>'
                                '<li style="color:#bbb;margin-bottom:0.5em">Destination folder selected</li>'
                                '<li>Moving backup ({backup_size:,} MB) '
                                'from {device} to<br/>{destination} …</li>'
                                '</ol>')

                # Create the ProgressBar in the main GUI thread
                pb = ProgressBar(parent=self.gui, window_title="Moving backup",
                                 alignment=Qt.AlignLeft)
                pb.set_label(BACKUP_MSG_3.format(**busy_panel_args))

                # Merge MXD state with backup image
                temp_dir = PersistentTemporaryDirectory()
                zip_dst = os.path.join(destination_folder, storage_name)

                # Init the class
                transfer_estimate = int(stats['st_size']) / IOS_READ_RATE
                if transfer_estimate < 100 * 1024 * 1024:
                    SIDECAR_ESTIMATE = 2.5
                elif transfer_estimate < 1000 * 1024 * 1024:
                    SIDECAR_ESTIMATE = 5.0
                else:
                    SIDECAR_ESTIMATE = 10.0

                kwargs = {
                          'backup_folder': backup_folder,
                          'destination_folder': destination_folder,
                          'ios': self.ios,
                          'parent': self,
                          'pb': pb,
                          'storage_name': storage_name,
                          'src_stats': stats,
                          'total_seconds': transfer_estimate + SIDECAR_ESTIMATE
                         }
                move_operation = MoveBackup(**kwargs)

                # Device cached hashes
                device_cached_hashes = "{0}_cover_hashes.json".format(
                    re.sub('\W', '_', self.ios.device_name))
                dch = os.path.join(self.resources_path, device_cached_hashes)
                if os.path.exists(dch):
                    move_operation.mxd_device_cached_hashes = dch

                # Remote content hashes
                rhc = b'/'.join(['/Library', 'calibre.mm',
                                 BookStatusDialog.HASH_CACHE_FS])
                if self.ios.exists(rhc):
                    base_name = "mxd_{0}".format(BookStatusDialog.HASH_CACHE_FS)
                    thc = os.path.join(temp_dir, base_name)
                    with open(thc, 'wb') as out:
                        self.ios.copy_from_idevice(rhc, out)
                    move_operation.mxd_remote_content_hashes = thc
                    move_operation.mxd_remote_hash_cache_fs = BookStatusDialog.HASH_CACHE_FS

                # mainDb profile
                move_operation.mxd_mainDb_profile = mainDb_profile

                """
                # self.installed_books
                if self.installed_books:
                    move_operation.mxd_installed_books = json.dumps(
                        self.dehydrate_installed_books(self.installed_books),
                        default=to_json,
                        indent=2, sort_keys=True)
                """

                # iOSRA booklist.zip
                archive_path = '/'.join([self.REMOTE_CACHE_FOLDER, 'booklist.zip'])
                if self.ios.exists(archive_path):
                    # Copy the stored booklist to a local temp file
                    with PersistentTemporaryFile(suffix=".zip") as local:
                        with open(local._name, 'w') as f:
                            self.ios.copy_from_idevice(archive_path, f)
                    move_operation.iosra_booklist = local._name

                start_time = time.time()

                pb.show()
                move_operation.start()
                while not move_operation.isFinished():
                    Application.processEvents()

                transfer_size = int(stats['st_size'])
                total_actual = time.time() - start_time

                analytics.append((
                    '2. Destination folder\n'
                    '   {0}\n').format(destination_folder))

                args = {
                        'IOS_READ_RATE': IOS_READ_RATE,
                        'pct_complete': pb.get_pct_complete(),
                        'sidecar_actual': self.format_time(move_operation.sidecar_time),
                        'sidecar_estimate': self.format_time(SIDECAR_ESTIMATE),
                        'total_actual': self.format_time(total_actual),
                        'total_estimate': self.format_time(SIDECAR_ESTIMATE + transfer_estimate),
                        'transfer_actual': self.format_time(move_operation.transfer_time),
                        'transfer_estimate': self.format_time(transfer_estimate),
                        'transfer_rate': transfer_size/move_operation.transfer_time,
                        'transfer_size': transfer_size
                        }
                analytics.append((
                    '3. Transferring backup\n'
                    '      backup image size: {transfer_size:,}\n'
                    '     est. transfer time: {transfer_estimate}\n'
                    '   actual transfer time: {transfer_actual}\n'
                    '      est. sidecar time: {sidecar_estimate}\n'
                    '    actual sidecar time: {sidecar_actual}\n'
                    '        est. total time: {total_estimate}\n'
                    '      actual total time: {total_actual} ({pct_complete}%)\n'
                    '     est. transfer rate: {IOS_READ_RATE:,} bytes/sec\n'
                    '   actual transfer rate: {transfer_rate:,.0f} bytes/sec\n'
                    ).format(**args))

                local.close()

                # Add the environment details to analytics
                analytics.append(self.format_device_profile())

                pb.hide()

                # Inform user backup operation is complete
                title = "Backup operation complete"
                msg = '<p>Marvin library archived to<br/>{0}</p>'.format(zip_dst)
                det_msg = '\n'.join(analytics)

                # Set dialog det_msg to monospace
                dialog = info_dialog(self.gui, title, msg, det_msg=det_msg)
                font = QFont('monospace')
                font.setFixedPitch(True)
                dialog.det_msg.setFont(font)
                dialog.exec_()

                # Save the backup folder
                self.prefs.set('backup_folder', destination_folder)
            else:
                # Inform user backup operation cancelled
                title = "Backup cancelled"
                msg = '<p>Backup of {0} cancelled</p>'.format(self.ios.device_name)
                det_msg = ''
                if analytics:
                    det_msg = '\n'.join(analytics)
                MessageBox(MessageBox.WARNING, title, msg, det_msg=det_msg, parent=self.gui,
                           show_copy_button=False).exec_()
        else:
            self._log("No backup file found at {0}".format(backup_target))

    def create_local_backup(self):
        '''
        Build a backup image locally
        '''
        self._log_location()

        """
        # Get a list of files, covers from mainDb
        con = sqlite3.connect(self.connected_device.local_db_path)
        with con:
            con.row_factory = sqlite3.Row
            cur = con.cursor()
            cur.execute('''SELECT FileName, Hash FROM Books''')
            rows = cur.fetchall()
            mdb_epubs = [row[b'FileName'] for row in rows]
            mdb_covers = ["{0}.jpg".format(row[b'Hash']) for row in rows]
        """

        epubs_path = b'/Documents'
        dir_contents = self.ios.listdir(epubs_path, get_stats=False)
        epubs = []
        for f in dir_contents:
            if f.lower().endswith('.epub'):
                epubs.append(f)
            else:
                self._log("ignoring {0}/{1}".format(epubs_path, f))

        small_covers_path = self.connected_device._cover_subpath(size="small")
        dir_contents = self.ios.listdir(small_covers_path, get_stats=False)
        small_covers = []
        for f in dir_contents:
            if f.lower().endswith('.jpg'):
                small_covers.append(f)
            else:
                self._log("ignoring {0}/{1}".format(small_covers_path, f))

        large_covers_path = self.connected_device._cover_subpath(size="large")
        dir_contents = self.ios.listdir(large_covers_path, get_stats=False)
        large_covers = []
        for f in dir_contents:
            if f.lower().endswith('.jpg'):
                large_covers.append(f)
            else:
                self._log("ignoring {0}/{1}".format(large_covers_path, f))

        total_steps = len(epubs) + len(large_covers)
        total_steps += 5    # MXD components
        total_steps += 2    # backup.xml, mainDb.sqlite
        total_steps += 1    # Move backup to destination folder

        # Set up the progress panel
        busy_panel_args = {'book_count': "{:,}".format(len(epubs)),
                           'destination': 'destination folder',
                           'device': self.ios.device_name,
                           'large_covers': "{:,}".format(len(large_covers)),
                           'small_covers': "{:,}".format(len(small_covers))
                           }
        BACKUP_MSG_1 = (
                        '<ol style="margin-right:1.5em">'
                        '<li style="margin-bottom:0.5em">Preparing backup …</li>'
                        '<li style="color:#bbb;margin-bottom:0.5em">Add {book_count} ePubs to archive</li>'
                        '<li style="color:#bbb;margin-bottom:0.5em">Add {large_covers} covers to archive</li>'
                        '<li style="color:#bbb;margin-bottom:0.5em">Select destination folder</li>'
                        '<li style="color:#bbb">Move backup to destination folder</li>'
                        '</ol>')
        pb = ProgressBar(alignment=Qt.AlignLeft,
                         label=BACKUP_MSG_1.format(**busy_panel_args),
                         parent=self.gui,
                         window_title="Creating backup of {0}".format(self.ios.device_name))
        pb.set_range(0, total_steps)
        pb.set_maximum(total_steps)
        pb.show()
        start_time = time.time()

        with TemporaryFile(suffix=".zip") as local_backup:
            with ZipFile(local_backup, 'w') as zfw:

                # Device cached hashes
                device_cached_hashes = "{0}_cover_hashes.json".format(
                    re.sub('\W', '_', self.ios.device_name))
                dch = os.path.join(self.resources_path, device_cached_hashes)
                if os.path.exists(dch):
                    base_name = "mxd_{0}".format(BookStatusDialog.HASH_CACHE_FS)
                    zfw.write(dch, arcname=base_name)
                pb.increment()

                # Remote content hashes
                rhc = b'/'.join(['/Library', 'calibre.mm',
                                 BookStatusDialog.HASH_CACHE_FS])
                if self.ios.exists(rhc):
                    with TemporaryFile() as lhc:
                        try:
                            with open(lhc, 'wb') as local_copy:
                                self.ios.copy_from_idevice(rhc, local_copy)
                            base_name = "mxd_{0}".format(BookStatusDialog.HASH_CACHE_FS)
                            zfw.write(local_copy.name, arcname=base_name)
                        except:
                            import traceback
                            self._log(traceback.format_exc())
                pb.increment()

                # mainDb profile
                zfw.writestr("mxd_mainDb_profile.json",
                             json.dumps(self.profile_db(), sort_keys=True))
                pb.increment()

                # <current_library>_installed_books.zip
                installed_books_archive = os.path.join(self.resources_path,
                    current_library_name().replace(' ', '_') + '_installed_books.zip')
                if os.path.exists(installed_books_archive):
                    zfw.write(installed_books_archive, arcname="mxd_installed_books.zip")
                pb.increment()

                # iOSRA booklist.db
                lhc = os.path.join(self.connected_device.resources_path, 'booklist.db')
                if os.path.exists(lhc):
                    zfw.write(lhc, arcname="iosra_booklist.db")
                pb.increment()

                # Issue command to request backup.xml
                command_type = "BACKUPMANIFEST"
                ch = CommandHandler(self)
                ch.construct_general_command(command_type)
                ch.issue_command(get_response="backup.xml")
                if ch.results['code']:
                    pb.hide()

                    self._log("results: %s" % ch.results)
                    title = "Unable to get backup.xml"
                    msg = ('<p>Unable to get backup.xml from Marvin.</p>'
                           '<p>Click <b>Show details</b> for more information.</p>').format(
                           self.ios.device_name)
                    det_msg = ch.results['details'] + '\n\n'
                    det_msg += self.format_device_profile()

                    # Set dialog det_msg to monospace
                    dialog = warning_dialog(self.gui, title, msg, det_msg=det_msg)
                    font = QFont('monospace')
                    font.setFixedPitch(True)
                    dialog.det_msg.setFont(font)
                    return dialog.exec_()

                response = ch.results['response']
                if re.match(self.UTF_8_BOM, response):
                    response = UnicodeDammit(response).unicode
                response = "<?xml version='1.0' encoding='utf-8'?>" + response
                zfw.writestr("backup.xml", response)
                pb.increment()

                # mainDb
                zfw.write(self.connected_device.local_db_path, arcname='mainDb.sqlite')
                pb.increment()

                # ePubs
                self._log("archiving {:,} epubs".format(len(epubs)))
                BACKUP_MSG_2 = (
                                '<ol style="margin-right:1.5em">'
                                '<li style="color:#bbb;margin-bottom:0.5em">Backup image initialized</li>'
                                '<li style="margin-bottom:0.5em">Archiving {book_count} ePubs …</li>'
                                '<li style="color:#bbb;margin-bottom:0.5em">Add {large_covers} covers to archive</li>'
                                '<li style="color:#bbb;margin-bottom:0.5em">Select destination folder</li>'
                                '<li style="color:#bbb">Move backup to destination folder</li>'
                                '</ol>')

                pb.set_label(BACKUP_MSG_2.format(**busy_panel_args))

                for path in epubs:
                    # Get a local copy of the book
                    rbp = '/'.join(['/Documents', path])
                    with TemporaryFile() as lbp:
                        try:
                            with open(lbp, 'wb') as local_copy:
                                self.ios.copy_from_idevice(str(rbp), local_copy)
                            zfw.write(local_copy.name, arcname=path)
                        except:
                            import traceback
                            self._log(traceback.format_exc())
                        pb.increment()

                # Large and small covers with one step
                self._log("archiving {:,} covers and thumbs".format(len(large_covers)))
                BACKUP_MSG_3 = (
                                '<ol style="margin-right:1.5em">'
                                '<li style="color:#bbb;margin-bottom:0.5em">Backup image initialized</li>'
                                '<li style="color:#bbb;margin-bottom:0.5em">{book_count} ePubs archived</li>'
                                '<li style="margin-bottom:0.5em">Archiving {large_covers} covers …</li>'
                                '<li style="color:#bbb;margin-bottom:0.5em">Select destination folder</li>'
                                '<li style="color:#bbb">Move backup to destination folder</li>'
                                '</ol>')
                pb.set_label(BACKUP_MSG_3.format(**busy_panel_args))

                # Process large and small covers at the same time
                # Assumes that len(large_covers) == len(small_covers)
                for x in range(len(large_covers)):
                    # Get a local copy of the large cover
                    path = large_covers[x]
                    rcp = b'/'.join([large_covers_path, path])
                    with TemporaryFile() as lcp:
                        try:
                            with open(lcp, 'wb') as local_copy:
                                self.ios.copy_from_idevice(rcp, local_copy)
                            zfw.write(local_copy.name, arcname="L-{0}".format(path))
                        except:
                            import traceback
                            self._log(traceback.format_exc())

                    # Get a local copy of the small cover
                    path = small_covers[x]
                    rcp = b'/'.join([small_covers_path, path])
                    with TemporaryFile() as lcp:
                        try:
                            with open(lcp, 'wb') as local_copy:
                                self.ios.copy_from_idevice(str(rcp), local_copy)
                            zfw.write(local_copy.name, arcname="S-{0}".format(path))
                        except:
                            import traceback
                            self._log(traceback.format_exc())

                    pb.increment()

            pb.hide()
            actual_time = time.time() - start_time
            self._log("archive created in {0} ({1:,.0f} bytes/second)".format(
                self.format_time(actual_time),
                os.stat(local_backup).st_size/actual_time))

            # Get the destination folder
            d = datetime.now()
            storage_name = "{0} {1}.backup".format(
                self.ios.device_name, d.strftime("%Y-%m-%d"))
            destination_folder = str(QFileDialog.getExistingDirectory(
                self.gui,
                "Select destination folder to store backup",
                self.prefs.get('backup_folder', os.path.expanduser("~")),
                QFileDialog.ShowDirsOnly | QFileDialog.DontResolveSymlinks))

            if destination_folder:
                # Qt apparently sometimes returns a file within the selected directory,
                # rather than the directory itself. Validate destination_folder
                if not os.path.isdir(destination_folder):
                    destination_folder = os.path.dirname(destination_folder)

                BACKUP_MSG_4 = (
                                '<ol style="margin-right:1.5em">'
                                '<li style="color:#bbb;margin-bottom:0.5em">Backup image initialized</li>'
                                '<li style="color:#bbb;margin-bottom:0.5em">{book_count} ePubs archived</li>'
                                '<li style="color:#bbb;margin-bottom:0.5em">{large_covers} covers archived</li>'
                                '<li style="color:#bbb;margin-bottom:0.5em">Destination folder selected</li>'
                                '<li style="margin-bottom:0.5em">Moving {backup_size} MB backup to destination folder…</li>'
                                '</ol>')
                busy_panel_args['backup_size'] = "{:,}".format(int(os.stat(local_backup).st_size/(1000*1000)))
                pb.set_label(BACKUP_MSG_4.format(**busy_panel_args))
                pb.show()

                # Copy local_backup to destination folder
                shutil.copy(local_backup, os.path.join(destination_folder, storage_name))

                pb.increment()
                pb.hide()

                # Inform user backup operation is complete
                title = "Backup operation complete"
                msg = '<p>Marvin library backed up to {0}</p>'.format(destination_folder)
                MessageBox(MessageBox.INFO, title, msg, parent=self.gui,
                           show_copy_button=False).exec_()

                # Save the backup folder
                self.prefs.set('backup_folder', destination_folder)
            else:
                # Inform user backup operation cancelled
                title = "Backup cancelled"
                try:
                    msg = '<p>Backup of {0} cancelled</p>'.format(self.ios.device_name)
                except:
                    msg = '<p>Backup cancelled</p>'
                det_msg = ''
                MessageBox(MessageBox.WARNING, title, msg, det_msg=det_msg, parent=self.gui,
                           show_copy_button=False).exec_()

    def create_menu_item(self, m, menu_text, image=None, tooltip=None, shortcut=None):
        ac = self.create_action(spec=(menu_text, None, tooltip, shortcut), attr=menu_text)
        if image:
            ac.setIcon(QIcon(image))
        m.addAction(ac)
        return ac

    def dehydrate_installed_books(self, installed_books):
        '''
        Convert installed_books to JSON-serializable format
        '''
        all_mxd_keys = sorted(Book.mxd_standard_keys + Book.mxd_custom_keys)
        dehydrated = {}
        for key in installed_books:
            dehydrated[key] = {}
            for mxd_attribute in all_mxd_keys:
                dehydrated[key][mxd_attribute] = getattr(
                    installed_books[key], mxd_attribute, None)
        return dehydrated

    def developer_utilities(self, action):
        '''
        '''
        self._log_location(action)
        if action in ['Connected device profile', 'Create remote backup',
                      'Delete calibre hash cache', 'Delete Marvin hash cache',
                      'Delete installed books cache', 'Delete all caches',
                      'Nuke annotations',
                      'Reset column widths']:
            if action == 'Delete Marvin hashes':
                rhc = b'/'.join([self.REMOTE_CACHE_FOLDER, BookStatusDialog.HASH_CACHE_FS])

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

#                 Called by restore_from_backup, so no dialog
#                 info_dialog(self.gui, 'Developer utilities',
#                             'Marvin hashes deleted', show=True,
#                             show_copy_button=False)

            elif action == 'Delete installed books cache':
                # Delete <current_library>_installed_books.zip
                # Reset self.installed_books
                self.installed_books = None
                archive_path = os.path.join(self.resources_path,
                    current_library_name().replace(' ', '_') + '_installed_books.zip')
                if os.path.exists(archive_path):
                    os.remove(archive_path)

                self._log("installed books cache deleted: {}".format(os.path.basename(archive_path)))
                info_dialog(self.gui, 'Developer utilities',
                            'installed books archives deleted: {}'.format(
                            os.path.basename(archive_path)),
                            show=True,
                            show_copy_button=False)

            elif action == 'Connected device profile':
                self.show_connected_device_profile()

            elif action == 'Create remote backup':
                self.create_remote_backup()

            elif action == 'Delete calibre hashes':
                self.gui.current_db.delete_all_custom_book_data('epub_hash')
                self._log("cached epub hashes deleted")

                # Invalidate the library hash map, as library contents may change before reconnection
                if hasattr(self, 'library_scanner'):
                    if hasattr(self.library_scanner, 'hash_map'):
                        self.library_scanner.hash_map = None
                info_dialog(self.gui, 'Developer utilities',
                            'calibre epub hashes deleted', show=True,
                            show_copy_button=False)

            elif action == 'Delete all caches':
                self.reset_caches()

            elif action == 'Nuke annotations':
                # nuke_annotations() has its own dialog informing progress
                self.nuke_annotations()

            elif action == 'Reset column widths':
                self._log("deleting marvin_library_column_widths")
                self.prefs.pop('marvin_library_column_widths')
                self.prefs.commit()
                info_dialog(self.gui, 'Developer utilities',
                            'Column widths reset', show=True,
                            show_copy_button=False)

        else:
            self._log("unsupported action '{0}'".format(action))

    def device_diagnostics(self):
        '''
        Replacement for profile_connected_device()
        '''
        def _add_available_space():
            available = self.connected_device.free_space()[0]
            if available > 1024 * 1024 * 1024:
                available = available / (1024 * 1024 * 1024)
                fmt_str = "{:.2f} GB"
            else:
                available = int(available / (1024 * 1024))
                fmt_str = "{:,} MB"
            device_profile['available_space'] = fmt_str.format(available)

        def _add_cache_files():
            # Cache files
            # Marvin:
            #   Library/mainDb.sqlite
            #   Library/calibre.mm/booklist.db
            #   Library/calibre.mm/content_hashes.db
            # Local:
            #   <calibre resource dir>/iOS_reader_applications_resources/booklist.db
            #   <calibre resource dir>/Marvin_XD_resources/*_cover_hashes.json
            #   <calibre resource dir>/Marvin_XD_resources/installed_books.zip

            from datetime import datetime

            def _get_ios_stats(path):
                mtime = st_size = None
                stats = self.connected_device.ios.stat(path)
                if stats:
                    st_size = int(stats['st_size'])
                    d = datetime.fromtimestamp(int(stats['st_mtime']))
                    mtime = d.strftime('%Y-%m-%d %H:%M:%S')
                return {'mtime': mtime, 'size': st_size}

            def _get_os_stats(path):
                mtime = st_size = None
                if os.path.exists(path):
                    stats = os.stat(path)
                    st_size = stats.st_size
                    d = datetime.fromtimestamp(stats.st_mtime)
                    mtime = d.strftime('%Y-%m-%d %H:%M:%S')
                return {'mtime': mtime, 'size': st_size}

            cache_files = {}

            ''' Marvin-specific cache files '''
            cache_files['mainDb.sqlite (remote)'] = _get_ios_stats('/Library/mainDb.sqlite')
            cache_files['mainDb.sqlite (local)'] = _get_os_stats(self.connected_device.local_db_path)
            cache_files['booklist.db (remote)'] = _get_ios_stats('Library/calibre.mm/booklist.db')
            cache_files['mxd_content_hashes.db (remote)'] = _get_ios_stats('Library/calibre.mm/content_hashes.db')

            # booklist.db from iOSRA resources
            path = os.path.join(self.connected_device.resources_path, 'booklist.db')
            cache_files['booklist.db (local)'] = _get_os_stats(path)

            # Per-library installed_books.zip archives
            mxd_resources_path = self.resources_path
            pattern = os.path.join(mxd_resources_path, "*_installed_books.zip")
            for path in glob.glob(pattern):
                ans = _get_os_stats(path)
                name = path.rsplit(os.path.sep)[-1]
                cache_files['mxd_{}'.format(name)] = ans

            # Per-device cover hashes
            pattern = os.path.join(mxd_resources_path, "/*_cover_hashes.json")
            for path in glob.glob(pattern):
                ans = _get_os_stats(path)
                name = path.rsplit(os.path.sep)[-1]
                cache_files['mxd_{}'.format(name)] = ans

            device_profile['cache_files'] = cache_files

        def _add_caching():
            iosra_prefs = JSONConfig('plugins/iOS reader applications')
            device_caching = {}
            device_caching_enabled = iosra_prefs.get('device_booklist_caching')
            allocation_factor = iosra_prefs.get('device_booklist_cache_limit')
            device_caching['enabled'] = device_caching_enabled
            device_caching['allocation_factor'] = allocation_factor

            allocated_space = int(self.connected_device.free_space()[0] * (allocation_factor / 100))
            if allocated_space > 1024 * 1024 * 1024:
                allocated_space = allocated_space / (1024 * 1024 * 1024)
                fmt_str = "{:.2f} GB"
            else:
                allocated_space = int(allocated_space / (1024 * 1024))
                fmt_str = "{:,} MB"
            device_caching['allocated_space'] = fmt_str.format(allocated_space)
            device_profile['device_caching'] = device_caching

        def _add_installed_plugins():
            # installed plugins
            from calibre.customize.ui import initialized_plugins
            user_installed_plugins = {}
            for plugin in initialized_plugins():
                path = getattr(plugin, 'plugin_path', None)
                if path is not None:
                    name = getattr(plugin, 'name', None)
                    if name == self.name:
                        continue
                    author = getattr(plugin, 'author', None)
                    version = getattr(plugin, 'version', None)
                    user_installed_plugins[name] = {'author': author, 'version': "{0}.{1}.{2}".format(*version)}
            device_profile['user_installed_plugins'] = user_installed_plugins

        def _add_device_book_count():
            # Device book count
            device_profile['device_book_count'] = len(self.connected_device.cached_books)

        def _add_device_info():
            cdp = self.connected_device.device_profile
            device_info = {}
            all_fields = ['DeviceClass', 'DeviceColor', 'DeviceName', 'FSBlockSize',
                          'FSFreeBytes', 'FSTotalBytes', 'FirmwareVersion', 'HardwareModel',
                          'ModelNumber', 'PasswordProtected', 'ProductType', 'ProductVersion',
                          'SerialNumber', 'TimeIntervalSince1970', 'TimeZone',
                          'TimeZoneOffsetFromUTC', 'UniqueDeviceID']
            superfluous = ['DeviceClass', 'DeviceColor', 'FSBlockSize', 'HardwareModel',
                           'SerialNumber', 'TimeIntervalSince1970', 'TimeZoneOffsetFromUTC',
                           'UniqueDeviceID', 'ModelNumber']
            for item in sorted(cdp):
                if item in superfluous:
                    continue
                if item in ['FSTotalBytes', 'FSFreeBytes']:
                    device_info[item] = int(cdp[item])
                else:
                    device_info[item] = cdp[item]
            device_profile['device_info'] = device_info

        def _add_library_profile():
            library_profile = {}
            cdb = self.gui.current_db
            library_profile['epubs'] = len(cdb.search_getting_ids('formats:EPUB', ''))
            #library_profile['pdfs'] = len(cdb.search_getting_ids('formats:PDF', ''))
            #library_profile['mobis'] = len(cdb.search_getting_ids('formats:MOBI', ''))

            device_profile['library_profile'] = library_profile

        def _add_load_time():
            if self.load_time:
                elapsed = _seconds_to_time(self.load_time)
                formatted = "{0:02d}:{1:02d}".format(int(elapsed['mins']), int(elapsed['secs']))
                device_profile['load_time'] = formatted
            else:
                device_profile['load_time'] = "n/a"

        def _add_iOSRA_version():
            device_profile['iOSRA_version'] = "{0}.{1}.{2}".format(*self.connected_device.version)

        def _add_mainDb_profiles():
            archive_path = os.path.join(self.resources_path,
                self.INSTALLED_BOOKS_SNAPSHOT)
            stored_mainDb_profile = {}
            if os.path.exists(archive_path):
                with ZipFile(archive_path, 'r') as zfr:
                    if 'mainDb_profile.json' in zfr.namelist():
                        stored_mainDb_profile = json.loads(zfr.read('mainDb_profile.json'))
            device_profile['stored_mainDb_profile'] = stored_mainDb_profile
            device_profile['current_mainDb_profile'] = self.profile_db()

        def _add_platform_profile():
            # Platform info
            import platform
            from calibre.constants import (__appname__, get_version, isportable, isosx,
                                           isfrozen, is64bit, iswindows)
            calibre_profile = "{0} {1}{2} isfrozen:{3} is64bit:{4}".format(
                __appname__, get_version(),
                ' Portable' if isportable else '', isfrozen, is64bit)
            device_profile['CalibreProfile'] = calibre_profile

            platform_profile = "{0} {1} {2}".format(
                platform.platform(), platform.system(), platform.architecture())
            device_profile['PlatformProfile'] = platform_profile

            try:
                if iswindows:
                    os_profile = "Windows {0}".format(platform.win32_ver())
                    if not is64bit:
                        try:
                            import win32process
                            if win32process.IsWow64Process():
                                os_profile += " 32bit process running on 64bit windows"
                        except:
                            pass
                elif isosx:
                    os_profile = "OS X {0}".format(platform.mac_ver()[0])
                else:
                    os_profile = "Linux {0}".format(platform.linux_distribution())
            except:
                import traceback
                self._log(traceback.format_exc())
                os_profile = "unknown"

            device_profile['OSProfile'] = os_profile

        def _add_prefs():
            if self.prefs.get('debug_plugin', False):
                ignore = ['plugin_version']
            else:
                ignore = ['appearance_css', 'appearance_dialog', 'cc_mappings',
                          'collections_viewer', 'css_editor', 'css_editor_split_points',
                          'html_viewer', 'injected_css', 'marvin_library',
                          'metadata_comparison', 'plugin_version']
            prefs = {'created_under': self.prefs.get('plugin_version')}
            for pref in sorted(self.prefs.keys()):
                if pref in ignore:
                    continue
                prefs[pref] = self.prefs.get(pref, None)
            device_profile['prefs'] = prefs


        def _format_cache_files_info():
            max_fs_width = max([len(v) for v in device_profile['cache_files'].keys()])
            max_size_width = 12
            max_ts_width = 20

            args = {'subtitle': " Caches {} ".format(report_time),
                    'separator_width': separator_width,
                    'report_time_label': "report generated",
                    'fs_width': max_fs_width + 1,
                    'ts_width': max_ts_width + 1}
            for cache_fs in device_profile['cache_files']:
                args[cache_fs] = "{%s}" % cache_fs

            TEMPLATE = ('\n{subtitle:-^{separator_width}}\n')

            # iOSRA cache files
            ans = TEMPLATE.format(**args)
            # iOSRA cache files
            TEMPLATE = 'iOSRA\n'
            for cache_fs, d in sorted(device_profile['cache_files'].iteritems(), key=lambda item: item[0].lower()):
                if cache_fs.startswith('mxd_'):
                    continue
                try:
                    TEMPLATE += " {0:{1}} {2:{3}} {4:{5},}\n".format(
                        cache_fs, max_fs_width + 1,
                        d['mtime'], max_ts_width + 1,
                        d['size'], max_size_width + 1)
                except:
                    TEMPLATE += " {0:{1}} {2:{3}} {4:{5},}\n".format(
                        cache_fs, max_fs_width + 1,
                        d['mtime'], max_ts_width + 1,
                        0, max_size_width + 1)

            ans += TEMPLATE.format(**args)
            ans += '\n'

            # MXD cache files
            TEMPLATE = 'MXD\n'
            for cache_fs, d in sorted(device_profile['cache_files'].iteritems(), key=lambda item: item[0].lower()):
                if not cache_fs.startswith('mxd_'):
                    continue
                try:
                    TEMPLATE += " {0:{1}} {2:{3}} {4:{5},}\n".format(
                        cache_fs[len('mxd_'):], max_fs_width + 1,
                        d['mtime'], max_ts_width + 1,
                        d['size'], max_size_width + 1)
                except:
                    TEMPLATE += " {0:{1}} {2:{3}} {4:{5},}\n".format(
                        cache_fs[len('mxd_'):], max_fs_width + 1,
                        d['mtime'], max_ts_width + 1,
                        0, max_size_width + 1)
            ans += TEMPLATE.format(**args)

            return ans

        def _format_caching_info():
            args = {'subtitle': " Device booklist caching ",
                    'separator_width': separator_width,
                    'enabled': device_profile['device_caching']['enabled'],
                    'available_space': device_profile['available_space'],
                    'allocation_factor': device_profile['device_caching']['allocation_factor'],
                    'allocated_space': device_profile['device_caching']['allocated_space']
                    }
            TEMPLATE = (
                '\n{subtitle:-^{separator_width}}\n'
                ' enabled: {enabled}\n'
                ' available space: {available_space}\n'
                ' allocation factor: {allocation_factor}%\n'
                ' allocated space: {allocated_space}\n'
                )
            return TEMPLATE.format(**args)

        def _format_device_info():
            args = {'subtitle': " iDevice ",
                    'separator_width': separator_width,
                    'iOSRA_version': device_profile['iOSRA_version'],
                    'iOS_version': device_profile['device_info']['ProductVersion'],
                    'ProductType': device_profile['device_info']['ProductType'],
                    'FSTotalBytes': device_profile['device_info']['FSTotalBytes'],
                    'FSFreeBytes': device_profile['device_info']['FSFreeBytes'],
                    'PasswordProtected': device_profile['device_info']['PasswordProtected']
                    }
            TEMPLATE = (
                '\n{subtitle:-^{separator_width}}\n'
                ' iOSRA version: {iOSRA_version}\n'
                ' iOS version: {iOS_version}\n'
                ' model: {ProductType}\n'
                ' FSTotalBytes: {FSTotalBytes:,}\n'
                ' FSFreeBytes: {FSFreeBytes:,}\n'
                ' PasswordProtected: {PasswordProtected}\n'
                )
            return TEMPLATE.format(**args)

        def _format_installed_plugins_info():

            args = {'subtitle': " Installed plugins ",
                    'separator_width': separator_width,
                    }
            ans = '\n{subtitle:-^{separator_width}}\n'.format(**args)

            if device_profile['user_installed_plugins']:
                for plugin in device_profile['user_installed_plugins']:
                    args[plugin] = "{{{0}}}".format(plugin)
                max_name_width = max([len(v) for v in device_profile['user_installed_plugins'].keys()])
                max_author_width = max([len(d['author']) for d in device_profile['user_installed_plugins'].values()])
                max_version_width = max([len(d['version']) for d in device_profile['user_installed_plugins'].values()])
                TEMPLATE = ''
                for plugin, d in sorted(device_profile['user_installed_plugins'].iteritems(), key=lambda item: item[0].lower()):
                    TEMPLATE += " {0:{1}} {2:{3}}\n".format(
                        plugin, max_name_width + 1,
                        d['version'], max_version_width + 1)
                ans += TEMPLATE.format(**args)

            return ans

        def _format_mainDb_profiles():
            CHECKMARK = u"\u2713"
            args = {'subtitle': " mainDb profiles ",
                    'separator_width': separator_width}
            TEMPLATE = '\n{subtitle:-^{separator_width}}\n'
            ans = TEMPLATE.format(**args)

            current = device_profile['current_mainDb_profile']
            stored = device_profile['stored_mainDb_profile']
            ans += "   {0:20} {1:37} {2:37}\n".format('', 'Stored', 'Current')
            #ans += "{0:—^23} {1:—^37} {2:—^37}\n".format('', '', '')
            ans += "{0}  {1:20} {2:<37} {3:<37}\n".format(
                'x' if stored.get('device') != current['device'] else ' ',
                'device',
                stored.get('device'),
                current['device'])
            keys = current.keys()
            keys.pop(keys.index('device'))
            for key in sorted(keys):
                stored_key = stored.get(key, '')
                current_key = current[key]

                if type(current_key) is int:
                    current_key = "{:,}".format(current_key)

                if type(stored_key) is int:
                    stored_key = "{:,}".format(stored_key)

                if key == "max_MetadataUpdated":
                    stored_key = stored.get(key, '')
                    if type(stored_key) is int:
                        stored_ts = datetime.fromtimestamp(stored_key)
                        stored_key = stored_ts.strftime('%Y-%m-%d %H:%M:%S')
                    current_ts = datetime.fromtimestamp(current[key])
                    current_key = current_ts.strftime('%Y-%m-%d %H:%M:%S')

                ans += "{0}  {1:20} {2:<37} {3:<37}\n".format(
                    'x' if stored.get(key) != current[key] else ' ',
                    key,
                    stored_key,
                    current_key)
            return ans

        def _format_prefs_info():
            args = {'subtitle': " Prefs ",
                    'separator_width': separator_width}
            for pref in device_profile['prefs']:
                args[pref] = repr(device_profile['prefs'][pref])

            TEMPLATE = '\n{subtitle:-^{separator_width}}\n'
            for pref in sorted(device_profile['prefs'].keys()):
                TEMPLATE += " {0}: {{{1}}}\n".format(pref, pref)

            return TEMPLATE.format(**args)

        def _format_system_info():
            # System information + MXD details
            args = {'subtitle': " {} ".format(report_time),
                    'separator_width': separator_width,
                    'CalibreProfile': device_profile['CalibreProfile'],
                    'OSProfile': device_profile['OSProfile'],
                    'plugin_version': "{0}.{1}.{2}".format(*self.interface_action_base_plugin.version),
                    'library_books': "{0:,} EPUBs".format(
                        device_profile['library_profile']['epubs']),
                    'is_virtual': bool(self.virtual_library),
                    'load_time': device_profile['load_time'],
                    'device_books': "{0:,} EPUBs".format(device_profile['device_book_count'])
                    }

            TEMPLATE = (
                '{subtitle:-^{separator_width}}\n'
                ' {CalibreProfile}\n'
                ' {OSProfile}\n'
                ' calibre library: {library_books} | virtual: {is_virtual}\n'
                ' Marvin XD version: {plugin_version}\n'
                ' Marvin library: {device_books}\n'
                )

            if self.load_time is not None:
                elapsed = _seconds_to_time(self.load_time)
                formatted = "{0:02d}:{1:02d}".format(int(elapsed['mins']), int(elapsed['secs']))
                args['load_time'] = formatted
                TEMPLATE += ' load time: {load_time}\n'

            return TEMPLATE.format(**args)

        def _seconds_to_time(s):
            years, s = divmod(s, 31556952)
            min, s = divmod(s, 60)
            h, min = divmod(min, 60)
            d, h = divmod(h, 24)
            ans = {'days': d, 'hours': h, 'mins': min, 'secs': s}
            return ans

        # ~~~ Entry point ~~~
        self._log_location()
        report_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        # Collect the diagnostic information
        device_profile = {}
        _add_platform_profile()
        _add_iOSRA_version()
        _add_prefs()
        _add_installed_plugins()
        _add_device_book_count()
        _add_load_time()
        _add_library_profile()
        _add_device_info()
        _add_available_space()
        _add_caching()
        _add_cache_files()
        _add_mainDb_profiles()

        # Format for printing
        det_msg = ''
        separator_width = 96

        det_msg += _format_system_info()
        det_msg += _format_device_info()
        det_msg += _format_installed_plugins_info()
        det_msg += _format_caching_info()
        det_msg += _format_cache_files_info()
        det_msg += _format_mainDb_profiles()
        det_msg += _format_prefs_info()

        # Present the results
        title = "Marvin XD diagnostics"
        msg = (
               '<p>Plugin diagnostics created for {}.</p>'
              ).format(self.connected_device.ios.device_name)

        # Set dialog det_msg to monospace
        dialog = info_dialog(self.gui, title, msg, det_msg=det_msg)
        font = QFont('monospace')
        font.setFixedPitch(True)
        dialog.det_msg.setFont(font)
        dialog.exec_()

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
            MessageBox(MessageBox.WARNING, status, msg, det_msg='',
                       parent=self.gui, show_copy_button=False).exec_()

        return status

    def format_device_profile(self):
        '''
        Return a formatted summary of the connected device
        '''
        device_profile = self.profile_connected_device()

        # Separators
        device_profile['MarvinDetails'] = " Marvin "
        device_profile['SystemDetails'] = " System "
        device_profile['DeviceName'] = " {0} ".format(device_profile['DeviceName'])
        separator_width = len(device_profile['DeviceName']) + 30
        device_profile['SeparatorWidth'] = separator_width

        DEVICE_PROFILE = (
            '{DeviceName:-^{SeparatorWidth}}\n'
            '           Type: {ProductType}\n'
            '          Model: {ModelNumber}\n'
            '            iOS: {ProductVersion}\n'
            '       Password: {PasswordProtected}\n'
            '   FSTotalBytes: {FSTotalBytes:,}\n'
            '    FSFreeBytes: {FSFreeBytes:,}\n'
            '\n{MarvinDetails:-^{SeparatorWidth}}\n'
            '            app: {MarvinAppID}\n'
            '        version: {MarvinVersion}\n'
            '         mainDb: {MarvinMainDbSize:,}\n'
            'installed books: {InstalledBooks}\n'
            '\n{SystemDetails:-^{SeparatorWidth}}\n'
            '{CalibreProfile}\n'
            #'{PlatformProfile}\n'
            '{OSProfile}\n'
            )
        return DEVICE_PROFILE.format(**device_profile) + '\n' + self._get_user_installed_plugins(separator_width)

    def format_time(self, total_seconds, show_fractional=True):
        m, s = divmod(total_seconds, 60)
        h, m = divmod(m, 60)
        if show_fractional:
            if h:
                formatted = "%d:%02d:%05.2f" % (h, m, s)
            else:
                formatted = "%d:%05.2f" % (m, s)
        else:
            if h:
                formatted = "%d:%02d:%02.0f" % (h, m, s)
            else:
                formatted = "%d:%02.0f" % (m, s)
        return formatted

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
        self.load_time = None
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
        '''
        When library changes, reset lib variables and delete installed_books.zip
        '''
        self._log_location(current_library_name())
        self.indexed_library = None
        self.installed_books = None
        self.library_indexed = False
        self.library_scanner = None
        self.library_last_modified = None
        self.load_time = None
        self.virtual_library = None

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
                       parent=self.gui,
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

    def profile_connected_device(self):
        '''
        Return a formatted profile of key device info
        '''
        self._log_location()

        device_profile = self.connected_device.device_profile.copy()

        marvin_version = self.connected_device.marvin_version
        device_profile['MarvinVersion'] = "{0}.{1}.{2}".format(
            marvin_version[0], marvin_version[1], marvin_version[2])

        device_profile['InstalledBooks'] = len(self.connected_device.cached_books)

        device_profile['MarvinAppID'] = self.connected_device.app_id

        for key in ['FSFreeBytes', 'FSTotalBytes']:
            device_profile[key] = int(device_profile[key])

        local_db_path = self.connected_device.local_db_path
        device_profile['MarvinMainDbSize'] = os.stat(local_db_path).st_size

        # Cribbed from calibre.debug:print_basic_debug_info()
        import platform
        from calibre.constants import (__appname__, get_version, isportable, isosx,
                                       isfrozen, is64bit, iswindows)
        calibre_profile = "{0} {1}{2} isfrozen:{3} is64bit:{4}".format(
            __appname__, get_version(),
            ' Portable' if isportable else '', isfrozen, is64bit)
        device_profile['CalibreProfile'] = calibre_profile

        platform_profile = "{0} {1} {2}".format(
            platform.platform(), platform.system(), platform.architecture())
        device_profile['PlatformProfile'] = platform_profile

        try:
            if iswindows:
                os_profile = "Windows {0}".format(platform.win32_ver())
                if not is64bit:
                    try:
                        import win32process
                        if win32process.IsWow64Process():
                            os_profile += " 32bit process running on 64bit windows"
                    except:
                        pass
            elif isosx:
                os_profile = "OS X {0}".format(platform.mac_ver()[0])
            else:
                os_profile = "Linux {0}".format(platform.linux_distribution())
        except:
            import traceback
            self._log(traceback.format_exc())
            os_profile = "unknown"

        device_profile['OSProfile'] = os_profile
        return device_profile

    def profile_db(self):
        '''
        Snapshot key aspects of mainDb
        '''
        profile = {}
        if self.ios.device_name:
            profile = {'device': self.ios.device_name}

            con = sqlite3.connect(self.connected_device.local_db_path)
            with con:
                con.row_factory = sqlite3.Row
                cur = con.cursor()

                # Hash the titles and authors
                m = hashlib.md5()
                cur.execute('''SELECT Title, Author FROM Books''')
                rows = cur.fetchall()
                for row in rows:
                    m.update(row[b'Title'])
                    m.update(row[b'Author'])
                profile['content_hash'] = m.hexdigest()

                # Get the latest MetadataUpdated timestamp
                try:
                    cur.execute('''SELECT max(MetadataUpdated) FROM Books''')
                    row = cur.fetchone()
                    profile['max_MetadataUpdated'] = row[b'max(MetadataUpdated)']
                except:
                    # Outdated version of Marvin
                    profile['max_MetadataUpdated'] = -1

                # Get the table sizes
                for table in ['BookCollections', 'Bookmarks', 'Books', 'Collections',
                              'Highlights', 'PinnedArticles', 'Vocabulary']:
                    cur.execute('''SELECT * FROM '{0}' '''.format(table))
                    profile[table] = len(cur.fetchall())

        return profile

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

            # Backup/restore: Wait for jobs to init and complete before enabling
            #if self.connected_device.marvin_version > (2, 7, 0):

            enabled = (hasattr(self.connected_device, 'local_db_path') and
                       not self.gui.job_manager.unfinished_jobs())
            action = 'Create backup'
            icon = QIcon(os.path.join(self.resources_path, 'icons', 'sync_collections.png'))
            ac = self.create_menu_item(m, action, image=icon)
            ac.triggered.connect(self.create_local_backup)
            ac.setEnabled(enabled)

            action = 'Restore from backup'
            icon = QIcon(os.path.join(self.resources_path, 'icons', 'sync_collections.png'))
            ac = self.create_menu_item(m, action, image=icon)
            ac.triggered.connect(self.restore_from_backup)
            ac.setEnabled(enabled)

            m.addSeparator()
            # Add 'Device diagnostics', enabled when connected device
            ac = self.create_menu_item(m, 'Plugin diagnostics', image=I('dialog_information.png'))
            ac.triggered.connect(self.device_diagnostics)
            ac.setEnabled(enabled)

            """
            # Add 'Reset caches', enabled when connected device
            ac = self.create_menu_item(m, 'Reset Marvin XD caches', image=I('trash.png'))
            ac.triggered.connect(self.reset_caches)
            ac.setEnabled(enabled)
            """

            # Add 'Help'
            ac = self.create_menu_item(m, 'Help', image=I('help.png'))
            ac.triggered.connect(self.show_help)

            # If Alt/Option key pressed, show Developer submenu
            modifiers = Application.keyboardModifiers()
            if bool(modifiers & Qt.AltModifier):
                m.addSeparator()
                self.developer_menu = m.addMenu(QIcon(I('config.png')),
                                                "Developer…")
                action = 'Delete calibre hash cache'
                ac = self.create_menu_item(self.developer_menu, action, image=I('trash.png'))
                ac.triggered.connect(partial(self.developer_utilities, action))

                action = 'Delete Marvin hash cache'
                ac = self.create_menu_item(self.developer_menu, action, image=I('trash.png'))
                ac.triggered.connect(partial(self.developer_utilities, action))
                ac.setEnabled(marvin_connected)

                action = 'Delete installed books cache'
                ac = self.create_menu_item(self.developer_menu, action, image=I('trash.png'))
                ac.triggered.connect(partial(self.developer_utilities, action))
                ac.setEnabled(marvin_connected)

                action = 'Delete all caches'
                ac = self.create_menu_item(self.developer_menu, action, image=I('trash.png'))
                ac.triggered.connect(partial(self.developer_utilities, action))
                ac.setEnabled(marvin_connected)

                action = 'Nuke annotations'
                ac = self.create_menu_item(self.developer_menu, action, image=I('trash.png'))
                ac.triggered.connect(partial(self.developer_utilities, action))

                action = 'Reset column widths'
                ac = self.create_menu_item(self.developer_menu, action, image=I('trash.png'))
                ac.triggered.connect(partial(self.developer_utilities, action))

                self.developer_menu.addSeparator()
                action = 'Create remote backup'
                icon = QIcon(os.path.join(self.resources_path, 'icons', 'sync_collections.png'))
                ac = self.create_menu_item(self.developer_menu, action, image=icon)
                ac.triggered.connect(partial(self.developer_utilities, action))
                ac.setEnabled(enabled)

            # Process Dropbox sync records automatically once only.
            if process_dropbox:
                self.process_dropbox_sync_records()
                self.dropbox_processed = True

    def rehydrate_installed_books(self, stored, silent=False):
        '''
        Reconstruct self.installed_books from stored image
        '''
        if not silent:
            self._log_location()

        rehydrated = None
        try:
            all_mxd_keys = sorted(Book.mxd_standard_keys + Book.mxd_custom_keys)
            for key in ['title', 'authors']:
                all_mxd_keys.remove(key)

            rehydrated = {}
            for cid in stored:
                rehydrated[int(cid)] = Book(stored[cid]['title'], stored[cid]['authors'])
                for prop in all_mxd_keys:
                    setattr(rehydrated[int(cid)], prop, stored[cid].get(prop))
        except:
            import traceback
            self._log_location()
            self._log(traceback.format_exc())
            rehydrated = None
        return rehydrated

    def remove_paths_from_hash_cache(self, paths):
        '''
        Remove cached hashes when iOSRA deletes books
        '''
        self._log_location()
        rhc = '/'.join([self.REMOTE_CACHE_FOLDER, BookStatusDialog.HASH_CACHE_FS])

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

    def reset_caches(self):
        '''
        Delete all cache content locally and remotely
        Mirrors developer_utilities()
        '''
        self._log_location()

        # Confirm
        if question_dialog(self.gui, 'Marvin XD', 'Reset all Marvin XD caches?'):

            det_msg = ''

            # Delete Marvin hashes
            rhc = b'/'.join([self.REMOTE_CACHE_FOLDER, BookStatusDialog.HASH_CACHE_FS])

            if self.ios.exists(rhc):
                self.ios.remove(rhc)
                self._log("remote hash cache deleted: {}".format(rhc))
                det_msg += "remote hash cache deleted:\n {}\n".format(rhc)
            else:
                self._log("no remote hash cache found")
                det_msg += "no remote hash cache found\n"

            # Delete cover hashes for all devices
            pattern = os.path.join(self.resources_path, "*_cover_hashes.json")
            if glob.glob(pattern):
                det_msg += "deleting device cover hashes:\n"
                for path in glob.glob(pattern):
                    self._log("deleting device cover hashes: {}".format(path))
                    det_msg += " {}\n".format(path)
                    os.remove(path)
            else:
                self._log("no device cover hashes found")
                det_msg += "no device cover hashes found\n"

            # Delete _installed_books.zip archives
            pattern = os.path.join(self.resources_path, "*_installed_books.zip")
            if glob.glob(pattern):
                for path in glob.glob(pattern):
                    self._log("deleting {}".format(path))
                    det_msg += "deleted:\n {}\n".format(path)
                    os.remove(path)
            else:
                self._log("no _installed_books.zip archives found")
                det_msg += "no _installed_books.zip archives found\n"

            # Delete calibre hashes from db
            self.gui.current_db.delete_all_custom_book_data('epub_hash')
            self._log("calibre epub hashes deleted from db")
            det_msg += "calibre epub hashes deleted from db\n"

            # Invalidate the library hash map, as library contents may change before reconnection
            if hasattr(self, 'library_scanner'):
                if hasattr(self.library_scanner, 'hash_map'):
                    self.library_scanner.hash_map = None

            dialog = info_dialog(self.gui, "Marvin XD", "All Marvin XD caches reset",
                show_copy_button=False, det_msg=det_msg)
            font = QFont('monospace')
            font.setFixedPitch(True)
            dialog.det_msg.setFont(font)
            dialog.exec_()

    def reset_marvin_library(self):
        self._log_location("not implemented")

    def restore_from_backup(self):
        '''
        Invoke RestoreBackup() in a separate thread
        Display a dialog telling user how to complete restore
        '''
        RESTORE_MSG_1 = ('<ol style="margin-right:1.5em">'
                         '<li style="margin-bottom:0.5em">Transferring backup image of '
                         '{book_count:,} books …</li>'
                         '<li style="color:#bbb">Complete restore process in Marvin</li>'
                         '</ol>')

        self._log_location()

        # Get the backup file
        backup_image = unicode(QFileDialog.getOpenFileName(
            self.gui,
            "Select Marvin backup file to restore",
            self.prefs.get('backup_folder', os.path.expanduser("~")),
            "*.backup").toUtf8())
        if backup_image:
            # Analyze the candidate file
            if not is_zipfile(backup_image):
                title = "Invalid backup file"
                msg = "{0} is not a valid Marvin backup image.".format(backup_image)
                return MessageBox(MessageBox.WARNING, title, msg,
                                  parent=self.gui, show_copy_button=False).exec_()

            archive = ZipFile(backup_image, 'r')
            archive_list = archive.infolist()
            epub_count = 0
            for f in archive_list:
                if f.filename.lower().endswith('.epub'):
                    epub_count += 1

            # Confirm space available on device
            src_size = os.stat(backup_image).st_size
            space_required = src_size * 2
            fs_free_bytes = int(self.connected_device.device_profile['FSFreeBytes'])
            if fs_free_bytes < space_required:
                title = "Insufficient space available"
                msg = ("<p>Not enough space available on {0} to restore backup.</p>"
                       "<p>{1:,.0f} MB required<br/>{2:,.0f} MB available.</p>".format(
                        self.ios.device_name,
                        space_required / (1024*1024),
                        fs_free_bytes / (1024*1024)))
                return MessageBox(MessageBox.WARNING, title, msg,
                                  parent=self.gui, show_copy_button=False).exec_()

            backup_source = ''
            components = re.match(r'(.*?) (\d{4}-\d{2}-\d{2})', os.path.basename(backup_image))
            if components:
                backup_source = components.group(1)

            ts = os.path.getmtime(backup_image)
            dt = datetime.fromtimestamp(ts)
            #backup_date = dt.strftime("%A, %b %e %Y")  # Windows chokes on this
            try:
                backup_date = components.group(2)
            except:
                backup_date = 'unknown'

            # Estimate transfer time @ IOS_WRITE_RATE
            IOS_WRITE_RATE = 5800000

            total_seconds = int(src_size / IOS_WRITE_RATE) + 1
            estimated_time = self.format_time(total_seconds)
            self._log("estimated_time: {0}".format(estimated_time))

            # Confirm
            title = "Confirm restore operation"
            msg = "<p>This will restore a backup of {0:,} books".format(epub_count)
            if backup_source:
                msg += ", from {0},".format(backup_source)
            msg += (' created {0}, to {1}.</p>'
                    '<p>It should take about {2} to restore the {3:,} MB '
                    'backup image.</p>'
                    '<p>Proceed?</p>').format(
                        backup_date,
                        self.ios.device_name,
                        self.format_time(total_seconds, show_fractional=False),
                        int(src_size/(1024*1024)))
            if not MessageBox(MessageBox.QUESTION, title, msg, parent=self.gui,
                              show_copy_button=False).exec_():
                return

            # Save the selected backup folder
            self.prefs.set('backup_folder', os.path.dirname(backup_image))

            # Create the ProgressBar in the main GUI thread
            busy_panel_args = {'book_count': epub_count,
                               'device': self.ios.device_name,
                               'estimated_time': estimated_time}
            pb = ProgressBar(parent=self.gui,
                             alignment=Qt.AlignLeft,
                             window_title="Restoring backup to {0}".format(self.ios.device_name))
            pb.set_label(RESTORE_MSG_1.format(**busy_panel_args))

            kwargs = {
                      'backup_image': backup_image,
                      'ios': self.ios,
                      'msg': RESTORE_MSG_1.format(**busy_panel_args),
                      'parent': self,
                      'pb': pb,
                      'total_seconds': total_seconds
                     }
            copy_operation = RestoreBackup(**kwargs)

            start_time = time.time()

            pb.show()

            # Start the copy operation
            copy_operation.start()
            while not copy_operation.isFinished():
                Application.processEvents()

            analytics = []
            actual_size = src_size
            actual_time = time.time() - start_time
            args = {'actual_size': actual_size,
                    'actual_time': self.format_time(actual_time),
                    'epub_count': epub_count,
                    'estimated_time': estimated_time,
                    'IOS_WRITE_RATE': IOS_WRITE_RATE,
                    'pct_complete': pb.get_pct_complete(),
                    'transfer_rate': actual_size/actual_time}
            analytics.append((
                '1. Restore from backup\n'
                '    backup image size: {actual_size:,}\n'
                '     est. backup time: {estimated_time}\n'
                '   actual backup time: {actual_time} ({pct_complete}%)\n'
                '   est. transfer rate: {IOS_WRITE_RATE:,} bytes/sec\n'
                ' actual transfer rate: {transfer_rate:,.0f} bytes/sec\n'
                ).format(**args))

            # Verify transferred size
            if copy_operation.success:
                start_time = time.time()

                # Delete cached Marvin data
                self.developer_utilities('Delete Marvin hashes')

                # Recover cached Marvin data if available
                if 'mxd_cover_hashes.json' in archive.namelist():
                    cover_hash_data = archive.read('mxd_cover_hashes.json')
                    device_cached_hashes = "{0}_cover_hashes.json".format(
                        re.sub('\W', '_', self.ios.device_name))
                    ch_dst = os.path.join(self.resources_path, device_cached_hashes)
                    with open(ch_dst, 'w') as out:
                        out.write(cover_hash_data)
                    self._log("cover hashes restored")
                else:
                    self._log('MXD cover hashes not found in archive')

                if 'mxd_{0}'.format(BookStatusDialog.HASH_CACHE_FS) in archive.namelist():
                    content_hash_data = archive.read("mxd_{0}".format(BookStatusDialog.HASH_CACHE_FS))
                    temp_dir = PersistentTemporaryDirectory()
                    temp_ch = os.path.join(temp_dir, BookStatusDialog.HASH_CACHE_FS)
                    with open(temp_ch, 'w') as lhc:
                        lhc.write(content_hash_data)
                    # Copy to iDevice
                    rhc = b'/'.join([self.REMOTE_CACHE_FOLDER, BookStatusDialog.HASH_CACHE_FS])
                    self.ios.remove(rhc)
                    self.ios.copy_to_idevice(temp_ch, rhc)
                    self._log("content hashes restored")
                else:
                    self._log('MXD content hashes not found in archive')

                # Recover _installed_books.zip
                previous_library_name = None
                if 'mxd_installed_books.zip' in archive.namelist():
                    basename = 'mxd_installed_books.zip'
                    with TemporaryDirectory() as tdir:
                        archive.extract(basename, tdir)
                        archive_path = os.path.join(tdir, basename)
                        with ZipFile(archive_path, 'r') as zfr:
                            if 'active_library.txt' in zfr.namelist():
                                previous_library_name = zfr.read('active_library.txt')
                        if previous_library_name:
                            dst = os.path.join(self.resources_path,
                                previous_library_name.replace(' ', '_') + "_installed_books.zip")
                            shutil.copy(archive_path, dst)
                            self._log("'{}' restored".format(os.path.basename(dst)))
                else:
                    self._log("mxd_installed_books.zip not found in archive")

                # Recover iosra_booklist.db, restore to connected device and iOSRA resource folder
                local_booklist_path = os.path.join(self.connected_device.resources_path, 'booklist.db')
                if os.path.exists(local_booklist_path):
                    os.remove(local_booklist_path)
                if 'iosra_booklist.db' in archive.namelist():
                    with TemporaryDirectory() as tdir:
                        basename = 'iosra_booklist.db'
                        archive.extract(basename, tdir)
                        source = str(os.path.join(tdir, basename))

                        # Copy to iOSRA resource folder
                        shutil.copy(source, local_booklist_path)

                        # If iOSRA device_booklist_caching enabled and within allocated size, copy to device
                        # Modeled from iOSRA:Marvin_overlays:_restore_from_snapshot(#2294)
                        from calibre.utils.config import JSONConfig
                        iOSRA_prefs = JSONConfig('plugins/iOS reader applications')
                        if iOSRA_prefs.get('device_booklist_caching'):
                            db_size = os.stat(source).st_size
                            #self._log("*** size of iosra_booklist.db: {:,}".format(db_size))
                            available = int(self.connected_device.device_profile['FSFreeBytes'])
                            #self._log("*** available: {:,}".format(available))
                            allocated = float(iOSRA_prefs.get('device_booklist_cache_limit', 10.0) / 100)
                            max_allowed = int(available * allocated)
                            #self._log("*** allocated: {0}% of free space ({1:,})".format(allocated * 100, max_allowed))
                            if db_size < max_allowed:
                                self._log("restoring iosra_booklist snapshot to device")
                                destination = '/'.join([self.REMOTE_CACHE_FOLDER, 'booklist.db'])
                                self.ios.copy_to_idevice(source, destination)
                            else:
                                self._log("iosra_booklist snapshot ({0:,}) exceeds allocated cache size ({1:,})".format(
                                    db_size, max_allowed))
                        else:
                            self._log("device_booklist_caching disabled, booklist.db not restored to device")

                else:
                    self._log("iosra_booklist snapshot not found in archive")

                actual_time = time.time() - start_time
                analytics.append((
                    '2. Recover iOSRA and MXD caches\n'
                    '        recovery time: {0}\n'
                    ).format(self.format_time(actual_time)))

                analytics.append(self.format_device_profile())

                pb.hide()

                self._log("Backup verifies: {0:,} bytes".format(copy_operation.src_size))
                # Display dialog detailing how to complete restore
                title = "Complete restore process"
                msg = ('<p>To complete the restore process of your Marvin library:</p>'
                       '<ol>'
                       '<li>Touch <b>Disconnect</b> on the calibre connector</li>'
                       '<li>Press the Home button to return to the Home screen</li>'
                       '<li>Double-click the Home button to display running apps</li>'
                       '<li>Force-quit and restart Marvin</li>'
                       '<li>Wait until restore finishes in Marvin</li>'
                       '<li>Restart the calibre connector</li>'
                       '</ol>'
                       )
                det_msg = '\n'.join(analytics)

                # Set dialog det_msg to monospace
                dialog = info_dialog(self.gui, title, msg, det_msg=det_msg)
                font = QFont('monospace')
                font.setFixedPitch(True)
                dialog.det_msg.setFont(font)
                dialog.exec_()

            else:
                pb.hide()

                self._log("Backup does not verify: source {0:,} dest {1:,}".format(
                    copy_operation.src_size, copy_operation.dst_size))
                title = "Restore unsuccessful"
                msg = ('<p>Backup does not verify.</p>'
                       '<p>'
                       '<tt>Src: {0:,}</tt><br/>'
                       '<tt>Dst: {1:,}</tt></p>'
                       '<p>Restore cancelled.</p>'
                      ).format(copy_operation.src_size, copy_operation.dst_size)
                d = MessageBox(MessageBox.WARNING, title, msg, det_msg='',
                               parent=self.gui, show_copy_button=False).exec_()
                self._log("Restore cancelled")

    def restore_installed_books(self):
        '''
        Use existing self.installed_books, or try to restore self.installed_books
        Evaluations:
            Do we have a warm self.installed_books?
            Is there an _installed_books.zip for the current library?
            Do the calibre metadata last_modified timestamps match?
            Do the mainDb profiles match?
        book_status:_get_installed_books() uses self.installed_books if available
        '''
        self._log_location()
        # Do we already have a populated self.installed_books?
        if self.installed_books:
            self._log("self.installed_books already populated")
            return

        # Do we have a library-specific stored image?
        archive_path = os.path.join(self.resources_path,
            current_library_name().replace(' ', '_') + '_installed_books.zip')
        if not os.path.exists(archive_path):
            return

        # Fetch the archived _installed_books contents
        previous_library_name = None
        stored_mainDb_profile = None
        dehydrated = {}
        with ZipFile(archive_path, 'r') as zfr:
            if 'mainDb_profile.json' in zfr.namelist():
                stored_mainDb_profile = json.loads(zfr.read('mainDb_profile.json'))
            if 'installed_books.json' in zfr.namelist():
                dehydrated = json.loads(zfr.read('installed_books.json'), object_hook=from_json)
            if 'active_library.txt' in zfr.namelist():
                previous_library_name = zfr.read('active_library.txt')

        # Make sure we're dealing with the same library used when archive was created
        if previous_library_name != current_library_name():
            self._log("previous_library: {}".format(previous_library_name))
            self._log("current_library: {}".format(current_library_name()))
            self._log("unable to restore from installed_books archive, different libraries")
            return

        # Test calibre metadata last_updated timestamps against archived
        self.installed_books_metadata_changes = []
        db = self.opts.gui.current_db
        for mid, d in dehydrated.iteritems():
            if d['cid']:
                mi = db.get_metadata(d['cid'], index_is_id=True)

                # Confirm CID is still active
                if mi.uuid == 'dummy':
                    self.installed_books_metadata_changes.append(d['mid'])
                    continue

                # Validate last_modified
                calibre_last_modified = time.mktime(mi.last_modified.astimezone(tz.tzlocal()).timetuple())
                archive_last_modified = d['last_updated']
                if calibre_last_modified != archive_last_modified:
                    self.installed_books_metadata_changes.append(d['mid'])

        if not self.installed_books_metadata_changes:
            self._log("individual book metadata is up to date")

        if self.compare_mainDb_profiles(stored_mainDb_profile):
            self._log("restoring self.installed_books from {}".format(os.path.basename(archive_path)))
            self.installed_books = self.rehydrate_installed_books(dehydrated)

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
        d = MessageBox(MessageBox.INFO, title, msg, det_msg=text,
                       parent=self.gui, show_copy_button=False)
        d.exec_()

    def show_connected_device_profile(self):
        '''
        '''
        device_profile = self.profile_connected_device()
        # Display connected device profile
        title = "Connected device profile"
        msg = (
               '<p>{0}<br/>Marvin v{1}</p>'
               '<p>Click <b>Show details</b> for summary.</p>'
              ).format(self.ios.device_name, device_profile['MarvinVersion'])
        det_msg = self.format_device_profile()

        # Set dialog det_msg to monospace
        dialog = info_dialog(self.gui, title, msg, det_msg=det_msg)
        font = QFont('monospace')
        font.setFixedPitch(True)
        dialog.det_msg.setFont(font)
        dialog.exec_()

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
            MessageBox(MessageBox.INFO, title, msg, det_msg='',
                       parent=self.gui, show_copy_button=False).exec_()
        else:
            start_time = time.time()
            self.launch_library_scanner()

            # Assure that Library is active view. Avoids problems with _delete_books
            restore_to = None
            if self.current_location != 'library':
                restore_to = self.current_location
                self.gui.location_selected('library')

            # Try to restore previous snapshot of self.installed_books
            self.restore_installed_books()

            # Open MXD dialog
            self.book_status_dialog = BookStatusDialog(self, 'marvin_library')
            self.book_status_dialog.initialize(self)
            self._log_location("{0} books".format(len(self.book_status_dialog.installed_books)))
            self.load_time = time.time() - start_time
            args = {'device_book_count': len(self.book_status_dialog.installed_books),
                    'library_book_count': len(self.library_scanner.uuid_map),
                    'is_virtual_library': bool(self.virtual_library),
                    'load_time': int(self.load_time)}
            self._log_metrics(args)
            self.book_status_dialog.exec_()

            # MXD dialog closed

            # Keep an in-memory snapshot of installed_books in case user reopens w/o disconnect
            self.installed_books = self.book_status_dialog.installed_books

            self.snapshot_installed_books(self.profile_db())

            # Restore the Device view if active before MXD window launched
            if restore_to:
                self.gui.location_selected(restore_to)

            self.book_status_dialog = None

    # subclass override
    def shutting_down(self):
        self._log_location()

    def snapshot_installed_books(self, profile):
        '''
        Store a snapshot of the connected Marvin library, dehydrated installed_books
        Enables optimized reload after disconnect
        '''
        self._log_location()

        if False:
            db = self.opts.gui.current_db

            for book in self.installed_books:
                #self._log(dir(self.installed_books[book]))
                title = getattr(self.installed_books[book], 'title')
                last_updated = getattr(self.installed_books[book], 'last_updated')
                self._log("{}: {}".format(title, last_updated))
                '''
                cid = getattr(self.installed_books[book], 'cid')
                if cid:
                    mi = db.get_metadata(cid, index_is_id=True)
                    last_modified = mi.last_modified.astimezone(tz.tzlocal())
                    ts = time.mktime(mi.last_modified.astimezone(tz.tzlocal()).timetuple())
                    self._log("{0} last_modified: {1} {2}".format(cid, last_modified, ts))
                '''

            self._log("*** snapshotting disabled ***")
            return

        dehydrated = self.dehydrate_installed_books(self.installed_books)
        if self.validate_dehydrated_books(dehydrated):
            archive_path = os.path.join(self.resources_path,
                current_library_name().replace(' ', '_') + '_installed_books.zip')
            with ZipFile(archive_path, 'w', compression=ZIP_STORED) as zfw:
                zfw.writestr("active_library.txt", current_library_name())
                zfw.writestr("mainDb_profile.json", json.dumps(profile, indent=2, sort_keys=True))
                zfw.writestr("installed_books.json", json.dumps(dehydrated, default=to_json, indent=2, sort_keys=True))

    def start_library_indexing(self):
        self._log_location()
        self._busy_panel_setup("Indexing calibre library…")
        self.library_scanner.start()

    def validate_dehydrated_books(self, dehydrated):
        '''
        A sanity test to confirm stored version of self.installed_books is legit
        '''
        self._log_location()
        rehydrated = self.rehydrate_installed_books(dehydrated, silent=True)
        ans = None
        if rehydrated == self.installed_books:
            self._log("validated: dehydrated matches self.installed")
            ans = True
        else:
            self._log("not validated: dehydrated does not match self.installed")
            """
            for cid in sorted(rehydrated.keys()):
                self._log("{0:>3} {1:30} {2}  |  {3}".format(cid, rehydrated[cid].title,
                    rehydrated[cid].pubdate, self.installed_books[cid].pubdate))
            """
            ans = False
        return ans

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

    def _delete_installed_books_archive(self):
        '''
        Delete all _installed_book.zip archives
        '''
        self._log_location()
        pattern = os.path.join(self.resources_path, "*_installed_books.zip")
        if glob.glob(pattern):
            for path in glob.glob(pattern):
                self._log("deleting {}".format(os.path.basename(path)))
                os.remove(path)
        else:
            self._log("no _installed_books.zip archives found")

    def _get_user_installed_plugins(self, width):
        '''
        Return a formatted list of user-installed plugins
        '''
        from calibre.customize.ui import initialized_plugins
        user_installed_plugins = {}
        for plugin in initialized_plugins():
            path = getattr(plugin, 'plugin_path', None)
            if path is not None:
                name = getattr(plugin, 'name', None)
                if name == self.name:
                    continue
                author = getattr(plugin, 'author', None)
                version = getattr(plugin, 'version', None)
                user_installed_plugins[name] = {'author': author,
                                                'version': "{0}.{1}.{2}".format(*version)}
        max_name_width = max([len(v) for v in user_installed_plugins.keys()])
        max_author_width = max([len(d['author']) for d in user_installed_plugins.values()])
        max_version_width = max([len(d['version']) for d in user_installed_plugins.values()])
        ans = "{:-^{}}\n".format(" User installed plugins ", width)
        for plugin, d in sorted(user_installed_plugins.iteritems(), key=lambda item: item[0].lower()):
            ans += "{0:{1}} {2:{3}}\n".format(plugin, max_name_width + 1,
                                              d['version'], max_version_width + 1)
        return ans

    def _log_metrics(self, kwargs):
        '''
        Post logging event
        No identifying information or library metadata is included in the logging event.
        device_udid is securely encrypted before logging so we have an anonymous (but unique) device id,
        used to determine number of unique devices using plugin.
        '''
        if self.prefs.get('plugin_diagnostics', True):
            self._log_location()
            br = browser()
            try:
                br.open(PluginMetricsLogger.URL)
                args = {'plugin': "Marvin XD",
                        'version': "%d.%d.%d" % MarvinManagerPlugin.version}
                post = PluginMetricsLogger(**args)
                post.req.add_header("DEVICE_MODEL", self.connected_device.device_profile['ProductType'])
                post.req.add_header("DEVICE_OS", self.connected_device.device_profile['ProductVersion'])
                # Encrypt the udid using MD5 encryption before logging
                m = hashlib.md5()
                m.update(self.connected_device.device_profile['UniqueDeviceID'])
                post.req.add_header("DEVICE_UDID", m.hexdigest())
                post.req.add_header("PLUGIN_BOOK_COUNT", kwargs.get('device_book_count', -1))
                post.req.add_header("PLUGIN_LIBRARY_BOOK_COUNT", kwargs.get('library_book_count', -1))
                post.req.add_header("PLUGIN_LIBRARY_IS_VIRTUAL", int(kwargs.get('is_virtual_library', False)))
                post.req.add_header("PLUGIN_LOAD_TIME", kwargs.get('load_time', 0))
                #self._log(post.req.header_items())
                post.start()
            except Exception as e:
                self._log("Plugin logger unreachable: {0}".format(e))

