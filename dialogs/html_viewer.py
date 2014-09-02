#!/usr/bin/env python
# coding: utf-8

from __future__ import (unicode_literals, division, absolute_import,
                        print_function)

__license__ = 'GPL v3'
__copyright__ = '2010, Gregory Riker'
__docformat__ = 'restructuredtext en'

import os, sys

from calibre.devices.usbms.driver import debug_print
from calibre.ebooks.BeautifulSoup import BeautifulSoup
from calibre.gui2 import Application, open_url

from calibre_plugins.marvin_manager.book_status import dialog_resources_path
from calibre_plugins.marvin_manager.common_utils import Logger, SizePersistedDialog

try:
    from PyQt5.Qt import (Qt, QAction, QApplication, QDialogButtonBox, QIcon, QKeySequence,
                          QPalette, QSize, QSizePolicy,
                          pyqtSignal)
    from PyQt5.QtWebKitWidgets import QWebPage, QWebView
except ImportError:
    from PyQt4.Qt import (Qt, QAction, QApplication, QDialogButtonBox, QIcon, QKeySequence,
                          QPalette, QSize, QSizePolicy,
                          pyqtSignal)
    from PyQt4.QtWebKit import QWebPage, QWebView

# Import Ui_Form from form generated dynamically during initialization
if True:
    sys.path.insert(0, dialog_resources_path)
    from html_viewer_ui import Ui_Dialog
    sys.path.remove(dialog_resources_path)


class HTMLViewerDialog(SizePersistedDialog, Ui_Dialog, Logger):
    marvin_device_status_changed = pyqtSignal(dict)

    def accept(self):
        self._log_location()
        super(HTMLViewerDialog, self).accept()

    def close(self):
        self._log_location()
        super(HTMLViewerDialog, self).close()

    def copy_to_clipboard(self, *args):
        '''
        Store window contents to system clipboard
        '''
        modifiers = Application.keyboardModifiers()
        if bool(modifiers & Qt.AltModifier):
            contents = self.html_wv.page().currentFrame().toHtml()
            #contents = BeautifulSoup(str(contents)).prettify()
            QApplication.clipboard().setText(contents)
        else:
            contents = self.html_wv.page().currentFrame().toPlainText()
            QApplication.clipboard().setText(unicode(contents))

        if hasattr(self, 'ctc_button'):
            self.ctc_button.setText('Copied')
            self.ctc_button.setIcon(QIcon(I('ok.png')))

    def dispatch_button_click(self, button):
        '''
        BUTTON_ROLES = ['AcceptRole', 'RejectRole', 'DestructiveRole', 'ActionRole',
                        'HelpRole', 'YesRole', 'NoRole', 'ApplyRole', 'ResetRole']
        '''
        self._log_location()
        if self.bb.buttonRole(button) == QDialogButtonBox.AcceptRole:
            # Save content
            self.accept()

        elif self.bb.buttonRole(button) == QDialogButtonBox.ActionRole:
            if button.objectName() == 'refresh_button':
                self.refresh_custom_column()
            elif button.objectName() == 'copy_to_clipboard_button':
                self.copy_to_clipboard()

        elif self.bb.buttonRole(button) == QDialogButtonBox.RejectRole:
            # Cancelled
            self.close()

    def esc(self, *args):
        self.close()

    def initialize(self, parent, content, book_id, installed_book, marvin_db_path, use_qwv=True):
        '''
        __init__ is called on SizePersistedDialog()
        '''
        self.setupUi(self)
        self.book_id = book_id
        self.connected_device = parent.opts.gui.device_manager.device
        self.installed_book = installed_book
        self.marvin_db_path = marvin_db_path
        self.opts = parent.opts
        self.parent = parent
        self.stored_command = None
        self.verbose = parent.verbose
        self._log_location(installed_book.title)

        # Subscribe to Marvin driver change events
        self.connected_device.marvin_device_signals.reader_app_status_changed.connect(
            self.marvin_status_changed)

        # Set the icon
        self.setWindowIcon(self.parent.icon)

        # Set or hide the header
        if content['header']:
            self.header.setText(content['header'])
        else:
            self.header.setVisible(False)

        # Set the titles
        self.setWindowTitle(content['title'])
        self.html_gb.setTitle(content['group_box_title'])
        if content['toolTip']:
            self.html_gb.setToolTip(content['toolTip'])

        # Set the bg color of the content to the dialog bg color
        bgcolor = self.palette().color(QPalette.Background)
        palette = QPalette()
        palette.setColor(QPalette.Base, bgcolor)

        #self._log(repr(content['html_content']))

        # Initialize the window content
        if use_qwv:
            # Add a QWebView to layout
            self.html_wv = QWebView()
            self.html_wv.setHtml(content['html_content'])
            self.html_wv.sizeHint = self.wv_sizeHint
            self.html_wv.setSizePolicy(QSizePolicy.MinimumExpanding, QSizePolicy.MinimumExpanding)
            self.html_wv.page().setLinkDelegationPolicy(QWebPage.DelegateAllLinks)
            self.html_wv.linkClicked.connect(self.link_clicked)

            self.html_gb_vl.addWidget(self.html_wv)
            self.html_tb.setVisible(False)
        else:
            # Initialize the contents of the TextBrowser
            self.html_tb.setText(content['html_content'])
            #self.html_tb.setPalette(palette)

        # Set or hide the footer
        if content['footer']:
            self.footer.setText(content['footer'])
        else:
            self.footer.setVisible(False)

        # Add Copy to Clipboard button
        self.ctc_button = self.bb.addButton('&Copy to clipboard',
                                            self.bb.ActionRole)
        self.ctc_button.clicked.connect(self.copy_to_clipboard)
        self.ctc_button.setIcon(QIcon(I('edit-copy.png')))
        self.ctc_button.setObjectName('copy_to_clipboard_button')
        self.ctc_button.setToolTip('<p>Copy plain text to clipboard, <b>Alt/Option-click</b> for HTML</p>')

        self.copy_action = QAction(self)
        self.addAction(self.copy_action)
        self.copy_action.setShortcuts(QKeySequence.Copy)
        self.copy_action.triggered.connect(self.copy_to_clipboard)

        # Add Refresh button if enabled
        if content['refresh']:
            self.refresh_method = content['refresh']['method']
            self.refresh_button = self.bb.addButton("Refresh '%s'" % content['refresh']['name'],
                                                    self.bb.ActionRole)
            self.refresh_button.setIcon(QIcon(os.path.join(self.parent.opts.resources_path,
                                              'icons',
                                              'from_marvin.png')))
            self.refresh_button.setObjectName('refresh_button')
            self.refresh_button.setEnabled(bool(self.installed_book.cid))

        # Hook the button events
        self.bb.clicked.connect(self.dispatch_button_click)

        # Restore position
        self.resize_dialog()

    def link_clicked(self, url):
        '''
        Open clicked link in regular browser
        '''
        open_url(url)

    def marvin_status_changed(self, cmd_dict):
        '''

        '''
        self.marvin_device_status_changed.emit(cmd_dict)
        command = cmd_dict['cmd']

        self._log_location(command)

        if command in ['disconnected', 'yanked']:
            self._log("closing dialog: %s" % command)
            self.close()

    def refresh_custom_column(self):
        '''
        If enabled, pass window content to custom column
        '''
        refresh = getattr(self.parent, self.refresh_method, None)
        if refresh is not None:
            refresh()
            self.refresh_button.setText('Refreshed')
            self.refresh_button.setIcon(QIcon(I('ok.png')))
        else:
            self._log_location("ERROR: Can't execute '%s'" % self.refresh_method)

    def store_command(self, command):
        '''
        '''
        self._log_location(command)
        self.stored_command = command
        self.close()

    def wv_sizeHint(self):
        '''
        QWebVew apparently has a default size of 800, 600
        '''
        return QSize(400,200)
