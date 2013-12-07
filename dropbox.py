#!/usr/bin/env python
# coding: utf-8
from __future__ import (unicode_literals, division, absolute_import,
                        print_function)

__license__ = 'GPL v3'
__copyright__ = '2013, Greg Riker <griker@hotmail.com>'
__docformat__ = 'restructuredtext en'

import glob, os, re, sys, time

from datetime import datetime, timedelta
from dateutil import tz
from lxml import etree

from calibre.devices.usbms.driver import debug_print
from calibre.ebooks.BeautifulSoup import BeautifulSoup, Tag, UnicodeDammit
from calibre.gui2 import Application
from calibre.gui2.dialogs.message_box import MessageBox
from calibre.utils.date import strptime

import calibre_plugins.marvin_manager.config as cfg
from calibre_plugins.marvin_manager.common_utils import (Logger, ProgressBar,
    updateCalibreGUIView)

class PullDropboxUpdates(Logger):
    # Location reporting template
    UTF_8_BOM = r'\xef\xbb\xbf'

    prefs = cfg.plugin_prefs

    def __init__(self, parent):
        self.db = parent.opts.gui.current_db
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

                    if doc.tag == "metadatasnapshot" and doc.attrib['creator'] == "Marvin":
                        book = doc.xpath('//book')[0]
                        title = book.attrib['title']
                        pb.set_label('{:^100}'.format("{0}".format(title)))
                        pb.increment()
                        cid = self._find_in_calibre_db(book)
                        if cid:
                            # *** Are we going to care about who's most recent?
                            c_last_modified = self._get_calibre_last_modified(cid)
                            m_last_modified = self._get_marvin_last_modified(book)

                            if c_last_modified > m_last_modified:
                                self._log("calibre metadata is newer")
                            elif m_last_modified > c_last_modified:
                                self._log("Marvin metadata is newer")
                            self._update_calibre_metadata(book, cid)

                            self._log("*** Deleting metadata update record: NOT IMPLEMENTED ***")
                            #os.remove(update)

                    Application.processEvents()
                    time.sleep(1.0)

                pb.hide()
                del pb

                updateCalibreGUIView()

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

    def _inject_css(self, html):
        '''
        stick a <style> element into html
        '''
        css = self.prefs.get('injected_css', None)
        if css:
            try:
                styled_soup = BeautifulSoup(html)
                head = styled_soup.find("head")
                style_tag = Tag(styled_soup, 'style')
                style_tag['type'] = "text/css"
                style_tag.insert(0, css)
                head.insert(0, style_tag)
                html = styled_soup.renderContents()
            except:
                return html
        return(html)

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

    def _get_calibre_last_modified(self, cid):
        '''
        '''
        mi = self.db.get_metadata(cid, index_is_id=True)
        c_last_modified = mi.last_modified.astimezone(tz.tzlocal())
        self._log_location(c_last_modified)
        return c_last_modified

    def _get_marvin_last_modified(self, book):
        '''
        Return a datetime object in local tz
        '''
        timestamp = float(book.attrib['lastmodified'])
        m_last_modified = datetime.utcfromtimestamp(timestamp).replace(tzinfo=tz.tzutc())
        m_last_modified = m_last_modified.astimezone(tz.tzlocal())
        self._log_location(m_last_modified)
        return m_last_modified

    def _update_calibre_metadata(self, book, cid):
        '''
        Update cid mapped custom columns from book metadata
        Annotations     comments    html                annotations
        Collections     text                            collections
        Last read       datetime                        date_read
        *Notes          comments    html
        Progress        float       50.0                progress
        *Rating
        Read            bool        True|False|None     read
        Reading list    bool        True|False|None     reading_list
        Word count      int         12345               word_count
        '''
        CUSTOM_COLUMN_MAPPINGS = {
            'Annotations': {
                'attribute': './annotations',
                'datatype': 'comments',
                'lookup': get_cc_mapping('annotations', 'field', None)
            },
            'Collections': {
                'attribute': './collections/collection',
                'datatype': 'text',
                'lookup': get_cc_mapping('collections', 'field', None)
            },
            'Last read': {
                'attribute': 'dateopened',
                'datatype': 'datetime',
                'lookup': get_cc_mapping('date_read', 'field', None)
            },
            'Progress': {
                'attribute': 'progress',
                'datatype': 'float',
                'lookup': get_cc_mapping('progress', 'field', None)
            },
            'Read': {
                'attribute': 'isread',
                'datatype': 'bool',
                'lookup': get_cc_mapping('read', 'field', None)
            },
            'Reading list': {
                'attribute': 'readinglist',
                'datatype': 'bool',
                'lookup': get_cc_mapping('reading_list', 'field', None)
            },
            'Word count': {
                'attribute': 'wordcount',
                'datatype': 'int',
                'lookup': get_cc_mapping('word_count', 'field', None)
            }
        }

        # Don't show floats less than FLOAT_THRESHOLD
        FLOAT_THRESHOLD = 1.0

        self._log_location(book.attrib['title'])
        mi = self.db.get_metadata(cid, index_is_id=True)
        mi_updated = False

        for ccm in CUSTOM_COLUMN_MAPPINGS:
            mapping = CUSTOM_COLUMN_MAPPINGS[ccm]
            lookup = self.prefs.get(mapping['lookup'], None)
            if lookup:
                um = mi.metadata_for_field(lookup)
                datatype = mapping['datatype']

                if ccm in ['Annotations']:
                    '''
                    # Not sure if we'll need this, depends on how Kris returns annotations
                    if re.match(self.UTF_8_BOM, anns):
                        anns = UnicodeDammit(anns).unicode
                    '''
                    ann_el = book.find(mapping['attribute'])
                    if ann_el is not None:
                        anns = ''
                        for sub_element in ['head', 'body']:
                            el = ann_el.find(sub_element)
                            anns += etree.tostring(el)
                        anns = self._inject_css(anns).encode('utf-8')
                        anns = "<?xml version='1.0' encoding='utf-8'?>" + anns
                        um['#value#'] = anns
                        mi.set_user_metadata(lookup, um)
                        mi_updated = True

                elif ccm in ['Collections']:
                    cels = book.findall(mapping['attribute'])
                    um['#value#'] = [unicode(cel.text) for cel in cels]
                    mi.set_user_metadata(lookup, um)
                    mi_updated = True

                else:
                    if datatype == 'bool':
                        um['#value#'] = None
                        if book.attrib[mapping['attribute']] == '1':
                            um['#value#'] = True

                    elif datatype == 'datetime':
                        ts = time.strftime("%Y-%m-%d %H:%M",
                            time.localtime(float(book.attrib[mapping['attribute']])))
                        um['#value#'] = ts

                    elif datatype == 'float':
                        val = float(book.attrib[mapping['attribute']]) * 100
                        if val < FLOAT_THRESHOLD:
                            val = None
                        um['#value#'] = val

                    elif datatype == 'int':
                        val = book.attrib[mapping['attribute']]
                        if val > '':
                            um['#value#'] = int(val)

                    else:
                        self._log("*** datatype '{0}' not handled ***".format(datatype))

                    mi.set_user_metadata(lookup, um)
                    mi_updated = True

            if mi_updated:
                self.db.set_metadata(cid, mi, set_title=False, set_authors=False,
                    commit=True, force_changes=True)
