#!/usr/bin/env python
# coding: utf-8
from __future__ import (unicode_literals, division, absolute_import,
                        print_function)

__license__ = 'GPL v3'
__copyright__ = '2013, Greg Riker <griker@hotmail.com>'
__docformat__ = 'restructuredtext en'

import glob, os, sys, time

from lxml import etree

from calibre.devices.usbms.driver import debug_print
from calibre.gui2.dialogs.message_box import MessageBox

import calibre_plugins.marvin_manager.config as cfg
from calibre_plugins.marvin_manager.common_utils import ProgressBar

class PullDropboxUpdates():
    # Location reporting template
    LOCATION_TEMPLATE = "{cls}:{func}({arg1}) {arg2}"
    prefs = cfg.plugin_prefs

    def __init__(self, parent):
        self.opts = parent.opts
        self.parent = parent
        self.verbose = parent.verbose

        self.process_updates()

    def process_updates(self):
        '''
        '''
        self._log_location()

        db_folder = self._get_folder_location()
        if db_folder:
            sub_folder = os.path.join(db_folder, 'Metadata')
            updates = glob.glob(os.path.join(sub_folder, '*.xml'))
            if os.path.exists(sub_folder) and updates:
                self._log("processing Marvin Dropbox folder at '{0}'".format(db_folder))

                pb = ProgressBar(parent=self.opts.gui,
                    window_title="Processing Dropbox updates",
                    on_top=True)
                total_books = len(updates)
                pb.set_maximum(total_books)
                pb.set_value(0)
                pb.show()

                for update in updates:
                    with file(update) as f:
                        doc = etree.fromstring(f.read())

                    self._log("tag: %s" % doc.tag)
                    self._log("timestamp: %s" % doc.attrib['timestamp'])
                    self._log("creator: %s" % doc.attrib['creator'])

                    book = doc.xpath('//book')[0]
                    pb.set_label('{:^100}'.format("Merging '{0}' updates".format(book.attrib['title'])))
                    pb.increment()
                    cid = self._find_in_calibre_db(book)
                    if cid:
                        #self._log("attributes: %s" % book.attrib)
                        author = book.attrib['author']
                        title = book.attrib['title']
                        uuid = book.attrib['uuid']
                        self._log("'{0}' by {1} {2}".format(title, author, uuid))

                        for el in book.getchildren():
                            self._log("{0}: {1}".format(el.tag, el.text))

                    time.sleep(1.0)

                pb.hide()
                del pb

            else:
                self._log("No MAX updates found")

    def _find_in_calibre_db(self, book):
        '''
        Try to find book in calibre, prefer UUID match
        Return cid
        '''
        self._log_location(book.attrib['title'])

        cid = None
        uuid = book.attrib['uuid']
        title = book.attrib['title']
        authors = book.attrib['author'].split(', ')
        if uuid in self.parent.library_scanner.uuid_map:
            cid = self.parent.library_scanner.uuid_map[uuid]['id']
            self._log("UUID match: %d" % cid)
        elif title in self.parent.library_scanner.title_map and \
            self.parent.library_scanner.title_map[title]['authors'] == authors:
            cid = self.parent.library_scanner.title_map[title]['id']
            self._log("Title/Author match: %d" % cid)
        else:
            self._log("No match")
        return cid


    def _get_folder_location(self):
        '''
        Confirm specified folder location contains Marvin subfolder
        '''
        dfl = self.prefs.get('dropbox_folder', None)
        msg = None
        title = 'Invalid Dropbox folder'
        folder_location = None
        if not dfl:
            msg = '<p>No Dropbox folder location specified in Configuration dialog.</p>'
        else:
            # Confirm presence of Marvin subfolder
            if not os.path.exists(dfl):
                msg = "<p>Specified Dropbox folder <tt>{0}</tt> not found.".format(dfl)
            else:
                path = os.path.join(dfl, 'Apps', 'com.marvinapp')
                if os.path.exists(path):
                    folder_location = path
                else:
                    msg = '<p>com.marvinapp not found in Apps folder.</p>'
        if msg:
            self._log_location("{0}: {1}".format(title, msg))
            MessageBox(MessageBox.WARNING, title, msg, det_msg='',
                show_copy_button=False).exec_()
        return folder_location

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

