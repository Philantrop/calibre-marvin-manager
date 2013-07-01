#!/usr/bin/env python
from __future__ import (unicode_literals, division, absolute_import,
                        print_function)

__license__ = 'GPL v3'
__copyright__ = '2013, Greg Riker <griker@hotmail.com>'
__docformat__ = 'restructuredtext en'

from calibre.customize import InterfaceActionBase
from calibre.utils.config import JSONConfig

class MarvinManagerPlugin(InterfaceActionBase):
    name                = 'Marvin Mangler'
    description         = 'Support functions for Marvin'
    supported_platforms = ['linux', 'osx', 'windows']
    author              = 'Greg Riker'
    version             = (0, 0, 1)
    minimum_calibre_version = (0, 9, 34)

    actual_plugin       = 'calibre_plugins.marvin_manager.action:MarvinManagerAction'
    prefs = JSONConfig('plugins/Marvin Mangler')

    def is_customizable(self):
        return True

    def config_widget(self):
        '''
        See devices.usbms.deviceconfig:DeviceConfig()
        '''
        from calibre_plugins.marvin_manager.config import ConfigWidget
        self.icon = getattr(self.actual_plugin, 'icon', None)
        self.resources_path = getattr(self.actual_plugin, 'resources_path', None)
        self.verbose = getattr(self.actual_plugin, 'verbose', False)
        self.cw = ConfigWidget(self)
        return self.cw

    def save_settings(self, config_widget):
        config_widget.save_settings()
        if self.actual_plugin_:
            self.actual_plugin_.rebuild_menus()

# For testing ConfigWidget, run from command line:
# cd ~/Documents/calibredev/Marvin_Manager
# calibre-debug __init__.py
if __name__ == '__main__':
    from PyQt4.Qt import QApplication
    from calibre.gui2.preferences import test_widget
    app = QApplication([])
    test_widget('Advanced', 'Plugins')
