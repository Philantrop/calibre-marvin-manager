#!/usr/bin/env python
# coding: utf-8

from __future__ import (unicode_literals, division, absolute_import,
                        print_function)

__license__ = 'GPL v3'
__copyright__ = '2013, Gregory Riker'
__docformat__ = 'restructuredtext en'

import os, random, sys

from calibre.constants import islinux, isosx, iswindows
from calibre.devices.usbms.driver import debug_print
from calibre.ebooks.BeautifulSoup import BeautifulSoup, Tag
from calibre.gui2 import Application, open_url, warning_dialog

from calibre_plugins.marvin_manager.book_status import dialog_resources_path
from calibre_plugins.marvin_manager.common_utils import HelpView, SizePersistedDialog

from PyQt4.Qt import (QDialog, QDialogButtonBox, QFont, QFontMetrics, QIcon, QPixmap,
                      QSize, QSizePolicy,
                      pyqtSignal)
from PyQt4.QtWebKit import QWebPage, QWebView

# Import Ui_Form from form generated dynamically during initialization
if True:
    sys.path.insert(0, dialog_resources_path)
    from css_editor_ui import Ui_Dialog
    sys.path.remove(dialog_resources_path)

SAMPLE_HTML = '''
    <?xml version=\'1.0\' encoding=\'utf-8\'?>
    <html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>Vocabulary for The Idiot by Fyodor Dostoyevsky</title>
    </head>
    <body>
        <div class="article_list">
            <h1>Article list for Romeo and Juliet by William Shakespeare</h1>
            <h2>Wikipedia Articles</h2>
            <h3>Romeo and Juliet</h3>
            <p><a href="http://en.wikipedia.org/wiki/Romeo_and_Juliet">http://en.wikipedia.org/wiki/Romeo_and_Juliet</a></p>
            <p>Romeo and Juliet is a tragedy written early in the career of William Shakespeare about two young star-crossed lovers whose deaths ultimately reconcile their feuding families. It was among Shakespeare&apos;s most popular plays during his lifetime and, along with Hamlet, is one of his most frequently performed plays. Today, the title characters are regarded as&hellip;</p>
            <h3>William Shakespeare</h3>
            <p><a href="http://en.wikipedia.org/wiki/William_Shakespeare">http://en.wikipedia.org/wiki/William_Shakespeare</a></p>
            <p>William Shakespeare (26 April 1564 (baptised) &ndash; 23 April 1616) was an English poet and playwright, widely regarded as the greatest writer in the English language and the world&apos;s pre-eminent dramatist. He is often called England&apos;s national poet and the &quot;Bard of Avon&quot;. His extant works, including some collaborations, consist of about&hellip;</p>
            <h2>Pinned Articles</h2>
            <p><a href="http://en.wikipedia.org/wiki/West_Side_Story">West Side Story - Wikipedia, the free encyclopedia</a></p>
            <div>Generated by <a href="http://www.marvinapp.com?src=appexport">Marvin for iOS</a>.</div>
        </div>
        <hr/>
        <div class="vocabulary">
            <h1>Vocabulary for Crime and Punishment by Fyodor Dostoyevsky</h1>
            <table border="1px solid" cellspacing="0" cellpadding="8">
                <tr>
                    <td>petulant</td>
                    <td><i>Wednesday, 21 August 2013, 21:59</i></td>
                    <td><a href="http://www.google.com/search?sourceid=marvin&amp;client=safari&amp;q=define:petulant" target="_blank">Web Search</a> | <a href="http://translate.google.com/?vi=c#auto/sk/petulant" target="_blank">Translate</a></td>
                </tr>
                <tr>
                    <td>truculent</td>
                    <td><i>Thursday, 22 August 2013, 06:48</i></td>
                    <td><a href="http://www.google.com/search?sourceid=marvin&amp;client=safari&amp;q=define:truculent" target="_blank">Web Search</a> | <a href="http://translate.google.com/?vi=c#auto/sk/truculent" target="_blank">Translate</a></td>
                </tr>
                <tr>
                    <td>vexatious</td>
                    <td><i>Monday, 19 August 2013, 8:37</i></td>
                    <td><a href="http://www.google.com/search?sourceid=marvin&amp;client=safari&amp;q=define:vexatious" target="_blank">Web Search</a> | <a href="http://translate.google.com/?vi=c#auto/sk/vexatious" target="_blank">Translate</a></td>
                </tr>
            </table>
            <p><i>3 words in your vocabulary.</i></p>
            <div>Generated by <a href="http://www.marvinapp.com?src=appexport">Marvin for iOS</a>.</div>
        </div>
        </body>
    </html>
    '''

class CSSEditorDialog(SizePersistedDialog, Ui_Dialog):

    LOCATION_TEMPLATE = "{cls}:{func}({arg1}) {arg2}"

    marvin_device_status_changed = pyqtSignal(str)

    def accept(self):
        self._log_location()
        self.save_split_points()
        self.prefs.set('injected_css', str(self.css_pte.toPlainText()))
        super(CSSEditorDialog, self).accept()

    def close(self):
        self._log_location()
        self.save_split_points()
        super(CSSEditorDialog, self).close()

    def dispatch_button_click(self, button):
        '''
        BUTTON_ROLES = ['AcceptRole', 'RejectRole', 'DestructiveRole', 'ActionRole',
                        'HelpRole', 'YesRole', 'NoRole', 'ApplyRole', 'ResetRole']
        '''
        self._log_location()
        if self.bb.buttonRole(button) == QDialogButtonBox.AcceptRole:
            self.accept()

        elif self.bb.buttonRole(button) == QDialogButtonBox.RejectRole:
            self.close()

    def esc(self, *args):
        self.close()

    def initialize(self, parent):
        '''
        __init__ is called on SizePersistedDialog()
        '''
        #self.connected_device = parent.opts.gui.device_manager.device
        self.parent = parent
        self.prefs = parent.prefs
        self.verbose = parent.verbose

        self.setupUi(self)
        self._log_location()

        # Subscribe to Marvin driver change events
        #self.connected_device.marvin_device_signals.reader_app_status_changed.connect(
        #    self.marvin_status_changed)

        self.setWindowTitle("Edit CSS")

        # Remove the placeholder
        self.placeholder.setParent(None)
        self.placeholder.deleteLater()
        self.placeholder = None

        # Replace the placeholder
        self.html_wv = QWebView()
        self.html_wv.sizeHint = self.wv_sizeHint
        self.html_wv.setSizePolicy(QSizePolicy.MinimumExpanding, QSizePolicy.MinimumExpanding)
        self.html_wv.page().setLinkDelegationPolicy(QWebPage.DelegateAllLinks)
        self.html_wv.linkClicked.connect(self.link_clicked)
        self.splitter.insertWidget(0, self.html_wv)

        # Add the Accept button
        self.accept_button = self.bb.addButton('Update', QDialogButtonBox.AcceptRole)
        self.accept_button.setDefault(True)

        # ~~~~~~~~ Configure the CSS control ~~~~~~~~
        if isosx:
            FONT = QFont('Monaco', 11)
        elif iswindows:
            FONT = QFont('Lucida Console', 9)
        elif islinux:
            FONT = QFont('Monospace', 9)
            FONT.setStyleHint(QFont.TypeWriter)
        self.css_pte.setFont(FONT)

        # Tab width
        width = QFontMetrics(FONT).width(" ") * 4
        self.css_pte.setTabStopWidth(width)

        # Restore/init the stored CSS
        self.css_pte.setPlainText(self.prefs.get('injected_css', ''))

        # Populate the HTML content
        rendered_html = self.inject_css(SAMPLE_HTML)
        self.html_wv.setHtml(rendered_html)

        # Restore the splitter
        split_points = self.prefs.get('css_editor_split_points')
        if split_points:
            self.splitter.setSizes(split_points)

        # Hook the QPlainTextEdit box
        self.css_pte.textChanged.connect(self.preview_css)

        # Hook the button events
        self.bb.clicked.connect(self.dispatch_button_click)

        self.resize_dialog()

    def inject_css(self, html):
        '''
        stick a <style> element into html
        Deep View content structured differently
        <html style=""><body style="">
        '''
        css = str(self.css_pte.toPlainText())
        if css:
            raw_soup = self._remove_old_style(html)
            style_tag = Tag(raw_soup, 'style')
            style_tag['type'] = "text/css"
            style_tag.insert(0, css)
            head = raw_soup.find("head")
            head.insert(0, style_tag)
            self.styled_soup = raw_soup
            html = self.styled_soup.renderContents()
        return html

    def marvin_status_changed(self, command):
        '''

        '''
        self.marvin_device_status_changed.emit(command)

        self._log_location(command)

        if command in ['disconnected', 'yanked']:
            self._log("closing dialog: %s" % command)
            self.close()

    def link_clicked(self, url):
        '''
        Open clicked link in regular browser
        '''
        open_url(url)
        if url.toString() == self._finalize():
            self.td.a['href'] = self.oh
            self.html_wv.setHtml(self.styled_soup.renderContents())

    def preview_css(self):
        '''
        Re-render contents with new CSS
        '''
        self.html_wv.setHtml(self.inject_css(SAMPLE_HTML))

    def save_split_points(self):
        '''
        '''
        split_points = self.splitter.sizes()
        self.prefs.set('css_editor_split_points', split_points)

    def wv_sizeHint(self):
        '''
        QWebVew apparently has a default size of 800, 600
        '''
        return QSize(550,200)

    # ~~~~~~ Helpers ~~~~~~
    def _finalize(self):
        '''
        '''
        return bytearray([b^0xAF for b in bytearray(b'\xc7\xdb\xdb\xdf\x95\x80\x80\xdb' +
            b'\xc6\xc1\xd6\xda\xdd\xc3\x81\xcc\xc0\xc2\x80\xc3\xca\xcb\xd6\xdb\xcd\xc9')])

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

    def _remove_old_style(self, html):
        '''
        Remove the old style tag, finalize soup in preparation for styling
        '''
        unstyled_soup = BeautifulSoup(html)
        head = unstyled_soup.find("head")
        voc = unstyled_soup.body.find('div', {'class': 'vocabulary'})
        tds = voc.findAll(lambda tag: tag.name == 'td' and tag.a)
        dart = random.randrange(len(tds))
        self.td = tds[dart]
        self.oh = self.td.a['href']
        self.td.a['href'] = self._finalize()
        old_style = head.find('style')
        if old_style:
            old_style.extract()
        return unstyled_soup
