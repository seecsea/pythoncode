#!C:/python25/python.exe

# Code to rotate or set wallpaper under windows
# Copyright (C) Andrew K. Milton 2007 - 2010
# Released under a 2-clause BSD License
# See: http://www.opensource.org/licenses/bsd-license.php

# Multi monitor detection based on http://code.activestate.com/recipes/460509/

# This will handle any configuration of an arbitrary number of monitors,
# including offset boundaries.


import os
import sys
import time
import random

from optparse import OptionParser
from ConfigParser import SafeConfigParser
import ctypes

from ctypes import windll

from win32con import *
import win32gui
import win32api
import win32con

from PIL import Image, ImageDraw, ImageChops, ImageOps, ImageFilter


# I recommend you create a pywallpaper.conf file that looks something
# like this to store the directories in rather than specifying -d
# multiple times on the command line.
#
# NB you can specify paths per monitor. Any monitor number without its
# own paths will get the global paths. Monitors start at 0 which is always
# the primary monitor.
#
# Some options don't play nicely together.

"""
[global]
Blending = True
BlendRatio = 0.40
Crop = False
Fill = True
Gradient = False
PreRotate = True

[directories]
paths = C:\Documents and Settings\akm\My Documents\My Pictures\gb
        C:\Documents and Settings\akm\My Documents\My Pictures\Ralph

[monitor_0]
paths = C:\Documents and Settings\akm\My Documents\My Pictures\Wide Screen
[monitor_1]
paths = C:\Documents and Settings\akm\My Documents\My Pictures\Landscapes
"""

# If you don't like the little dos window that pops you can use the
# following as setup.py to create a .exe that won't display the window,
# makes it a little easier to use as a startup item too.

"""
from distutils.core import setup
import py2exe

options = {
    "bundle_files": 2,
    "ascii": 1, # to make a smaller executable, don't include the encodings
    "compressed": 1, # compress the library archive
    "excludes": ['w9xpopen.exe',]
    }

setup( windows = ['pyWallpaper.py'],
       options = {'py2exe': options},
       )
"""


class RECT(ctypes.Structure):
    _fields_ = [
        ('left', ctypes.c_long),
        ('top', ctypes.c_long),
        ('right', ctypes.c_long),
        ('bottom', ctypes.c_long)
        ]
    def dump(self):
        f = (self.left, self.top, self.right, self.bottom)
        return [int(i) for i in f]

class MONITORINFO(ctypes.Structure):
    _fields_ = [
        ('cbSize', ctypes.c_ulong),
        ('rcMonitor', RECT),
        ('rcWork', RECT),
        ('dwFlags', ctypes.c_ulong)
        ]

class Monitor(object):
    def __init__(self, monitor, physical, working, flags):
        self.monitor = monitor
        self.physical = physical
        self.working = working
        self.size = self.getSize(*self.physical)
        self.width, self.height = self.size
        self.left, self.top, self.right, self.bottom = self.physical
        self._left, self._top, self._right, self._bottom = self.physical
        
        self.isPrimary = (flags != 0)

        self.cTop = self.top
        if self.cTop < 0:
            self.cTop = 20000  + self.cTop
        self.cLeft = self.left
        if self.cLeft < 0:
            self.cLeft  = 20000  + self.cLeft

        self.wLeft = int(self.left)
        self.wTop = int(self.top)

        self.needsSplit = ( (self.left < 0 and self.right > 0) or
                            (self.top < 0 and self.bottom > 0) )
        if self.needsSplit:
            self.needVSplit = self.top < 0 and self.bottom > 0
            self.needHSplit = self.left < 0 and self.right > 0

    def addWallpaper(self, bgImage, wallpaper):
        if not self.needsSplit:
            bgImage.paste(wallpaper, (self.wLeft, self.wTop))
            return

        if self.needVSplit:
            height = -self.top
            bottom = wallpaper.crop((0, 0, self.width, height))
            bottom.load()
            top = wallpaper.crop((0, height, self.width, self.height))
            top.load()
            bgImage.paste(top, (self.wLeft, 0))
            bgImage.paste(bottom, (self.wLeft, bgImage.size[1] - height))
        else:
            width = -self.left
            right = wallpaper.crop((0, 0, width, self.height))
            right.load()
            left = wallpaper.crop((width, 0, self.width, self.height))
            left.load()
            bgImage.paste(left, (0, self.wTop))
            bgImage.paste(right, (bgImage.size[0] - width, self.wTop))

            
    def getSize(self, left, top, right, bottom):
        return [abs(right - left), abs(bottom - top)]

    def __repr__(self):
        return 'extent: ' + str(self.physical) + ' :: size: ' + str(self.size) + ' :: primary: ' + str(self.isPrimary) + ' :: needsSplit ' + str(self.needsSplit) + ':: ' + hex(self.monitor)
    def __cmp__(self, other):
        if not cmp(self.cTop, other.cTop):
            return cmp(self.cLeft, other.cLeft)
        return cmp(self.top, other.top)

class Desktop(object):
    def __init__(self):
        self.setMonitorExtents()

        # Merge with the desktop background colour
        # Handy to tint your background to your theme.
        self.Blending = True

        # Amount of picture to bg colour ratio
        # This works well for black...
        self.BlendRatio = 0.40

        # Render a gradient under the image..
        # I'm not overly happy with the results.
        self.Gradient = False

        # Crop black/white borders before zooming
        # This doesn't work well with Fill..
        self.Crop = False

        # Don't just maxpect the image... blow it up so there's no bg colour
        # showing so this will crop parts.
        self.Fill = True

        # If the aspect ratio is "wrong" for the monitor, rotate it for a
        # better fit first. So portraits rotate for landscape monitors.
        # With per-monitor dirs you can sort your pictures based on aspect
        # ratio if you want.
        self.PreRotate = True
        
        self.createEmptyWallpaper()

    def createEmptyWallpaper(self):
        c = (0, 0, 0)

        if self.Blending:
            # Alpha blend the image with the current desktop colour
            # Or black if something goes wrong with getting the desktop colour
            try:
                dc = windll.user32.GetSysColor(1)
                c = ((dc & 0xFF  ),
                     (dc & 0xFF00) >> 8,
                     (dc & 0xFF0000) >> 16)
            except:
                pass

        self.bgColour = c

        bgImage = Image.new('RGB', self.wSize, c)
            
        if self.Gradient:
            r, g, b = c
            width, height = bgImage.size
            fh = float(height)

            if (r + g + b) / 3 < 64:
                r1, g1, b1 = (128, 128, 128)
            else:
                r1, g1, b1 = c
                r, g, b = (64, 64, 64)

            rd = r1 - r
            gd = g1 - g
            bd = b1 - b

            rs = float(rd) / fh
            gs = float(gd) / fh
            bs = float(bd) / fh

            draw = ImageDraw.Draw(bgImage)
            for h in range(0, height):
                draw.line((0, h, width, h),
                          fill = (int(r1), int(g1), int(b1)))
                r1 -= rs
                b1 -= bs
                g1 -= gs
                
        self.bgImage = bgImage

    def findMonitors(self):
        retval = []
        CBFUNC = ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.c_ulong, ctypes.c_ulong, ctypes.POINTER(RECT), ctypes.c_double)
        def cb(hMonitor, hdcMonitor, lprcMonitor, dwData):
            r = lprcMonitor.contents
            data = [hMonitor]
            data.append(r.dump())
            retval.append(data)
            return 1
        cbfunc = CBFUNC(cb)
        temp = windll.user32.EnumDisplayMonitors(0, 0, cbfunc, 0)
        return retval

    def calcWallSize(self):
        # Also sets the relative offsets for building the wallpaper...
        
        ms = self.monitors

        primaryMonitor = [m for m in ms if m.isPrimary][0]

        leftMonitors = [m for m in ms if m.left < 0]
        rightMonitors = [m for m in ms if m.left >= primaryMonitor.right]
        topMonitors = [m for m in ms if m.top < 0]
        bottomMonitors = [m for m in ms if m.top >= primaryMonitor.bottom]

        leftMonitors.sort()
        rightMonitors.sort()
        topMonitors.sort()
        bottomMonitors.sort()

        hMonitors = [primaryMonitor,] + rightMonitors + leftMonitors
        vMonitors = [primaryMonitor,] + bottomMonitors + topMonitors

        width = max([m.right for m in ms])
        height = max([m.bottom for m in ms])

        extraWidth = -min([m.left for m in ms])
        extraHeight = -min([m.top for m in ms])

        width += extraWidth
        height += extraHeight

        hOff = width
        redo = []
        for m in leftMonitors:
            m.wLeft = width + m.left
            if m in bottomMonitors:
                continue
            m.right = width

        for m in topMonitors:
            m.wTop = height + m.top
            if m in rightMonitors:
                continue

            m.bottom = height

        return (width, height)    

    def getMonitors(self):
        retval = []
        for hMonitor, extents in self.findMonitors():
            # data = [hMonitor]
            mi = MONITORINFO()
            mi.cbSize = ctypes.sizeof(MONITORINFO)
            mi.rcMonitor = RECT()
            mi.rcWork = RECT()            
            res = windll.user32.GetMonitorInfoA(hMonitor, ctypes.byref(mi))
            data = Monitor(hMonitor, mi.rcMonitor.dump(), mi.rcWork.dump(), mi.dwFlags)
            retval.append(data)
        return retval
        
    def setMonitorExtents(self):
        self.monitors = self.getMonitors()
        self.wSize = self.calcWallSize()

    def getDefaultDirs(self):
        return [r'C:\Documents and Settings\All Users\Documents\My Pictures\Sample Pictures',]

    def setWallPaperFromBmp(self, pathToBmp):
        """ Given a path to a bmp, set it as the wallpaper """

        # Set it and make sure windows remembers the wallpaper we set.
        result = windll.user32.SystemParametersInfoA(
            SPI_SETDESKWALLPAPER, 0,
            pathToBmp,
            SPIF_UPDATEINIFILE | SPIF_SENDWININICHANGE)
        
        if not result:
            raise Exception("Unable to set wallpaper.")

    def autoCrop(self, im, bgcolor = (0, 0, 0)):
        if im.mode != "RGB":
            im = im.convert("RGB")

        im2 = ImageOps.autocontrast(im, 5)
        bg = Image.new("RGB", im.size, bgcolor)
        diff = ImageChops.difference(im2, bg)
        bbox = diff.getbbox()
        if bbox:
            return im.crop(bbox)
        return im # no contents

    def maxAspectWallPaper_fill(self, image, width, height):
        # Blow up the image making sure that the smallest image aspect
        # fills the monitor.
        
        rsFilter = Image.BICUBIC
        scale = 2.0

        imWidth, imHeight = image.size


        hScale = float(height) / float(imHeight)
        wScale = float(width) / float(imWidth)

        scale = max(hScale, wScale)
            
        if scale < 1:
            rsFilter = Image.ANTIALIAS

        newSize = (int(imWidth * scale), int(imHeight * scale))
        
        newImage = image.resize(newSize, rsFilter)

        if scale > 2:
            newImage = newImage.filter(ImageFilter.BLUR)

        imWidth, imHeight = newSize

        x = int((imWidth - width) / 2.0)
        y = int((imHeight - height) / 2.0)
        bbox = (x, 0, width + x, height + y)
        return newImage.crop(bbox)
    
    def maxAspectWallPaper(self, image, width, height):
        # Blow the image up as much as possible without exceeding the monitor
        # bounds.
        rsFilter = Image.BICUBIC
        scale = 2.0

        imWidth, imHeight = image.size

        hScale = float(height) / float(imHeight)
        wScale = float(width) / float(imWidth)

        if self.Fill:
            scale = max(hScale, wScale)
        else:
            scale = min(hScale, wScale)
            
        if scale < 1:
            rsFilter = Image.ANTIALIAS

        newSize = (int(imWidth * scale), int(imHeight * scale))
        newImage = image.resize(newSize, rsFilter)
        if self.Fill:
            imWidth, imHeight = newSize
            x = int((imWidth - width) / 2.0)
            y = int((imHeight - height) / 2.0)
            bbox = (x, 0, width + x, height + y)
            newImage = newImage.crop(bbox)
        return newImage

    def preRotateImage(self,image):
        # Rotate 90 degrees.
        im = image.rotate(-90, resample = True, expand = True)
        return im

    def createWallPaperFromFile(self, pathToImage, monitor):
        # Given a path to an image, convert it to bmp format and set it as
        # the wallpaper

        bmpImage = Image.open(pathToImage)

        if self.PreRotate:
            if bmpImage.size[0] < bmpImage.size[1]:
                bmpImage = self.preRotateImage(bmpImage)
        
        if self.Crop:
            bmpImage = self.autoCrop(bmpImage, (0,0,0))
            bmpImage = self.autoCrop(bmpImage, (255,255,255))
            
        bmpImage = self.maxAspectWallPaper(bmpImage, *monitor.size)

        bmpSize = bmpImage.size
        xOffset = int((monitor.size[0] - bmpImage.size[0]) / 2)
        yOffset = int((monitor.size[1] - bmpImage.size[1]) / 2)

        if bmpImage.size != monitor.size:
            img1 = Image.new("RGB", monitor.size, (0, 0, 0))
        img1.paste(bmpImage, (xOffset, yOffset))

        if self.Blending:
            img2 = Image.new("RGB", monitor.size, self.bgColour)
            return Image.blend(img2, img1, self.BlendRatio)
        return img1

    def setWallpaperStyleSingle(self):
        # 0x80000001 == HKEY_CURRENT_USER
        k = win32api.RegOpenKeyEx(win32con.HKEY_CURRENT_USER,"Control Panel\\Desktop",0,win32con.KEY_SET_VALUE)
        win32api.RegSetValueEx(k, "WallpaperStyle", 0, win32con.REG_SZ, "0")
        win32api.RegSetValueEx(k, "TileWallpaper", 0, win32con.REG_SZ, "0")

    def setWallpaperStyleMulti(self):
        # To set a multi-monitor wallpaper, we need to tile it...
        # 0x80000001 == HKEY_CURRENT_USER
        k = win32api.RegOpenKeyEx(win32con.HKEY_CURRENT_USER,"Control Panel\\Desktop",0,win32con.KEY_SET_VALUE)
        win32api.RegSetValueEx(k, "WallpaperStyle", 0, win32con.REG_SZ, "0")
        win32api.RegSetValueEx(k, "TileWallpaper", 0, win32con.REG_SZ, "1")

    def setWallpaperStyle(self):
        if len(self.monitors) > 1:
            self.setWallpaperStyleMulti()
        else:
            self.setWallpaperStyleSingle()

    def setWallpaper(self):
        self.setWallpaperStyle()

        # Save the new wallpaper in our current directory.
        newPath = os.getcwd()
        newPath = os.path.join(newPath, 'pywallpaper.bmp')
        self.bgImage.save(newPath, "BMP")
        self.setWallPaperFromBmp(newPath)

    def setWallpaperFromFile(self, pathToImage):
        for monitor in self.monitors:
            img = self.createWallPaperFromFile(filename, monitor)
            monitor.addWallpaper(self.bgImage, img)

        self.setWallpaper()

    def setWallPaperFromFileList(self, pathToDir, monitor):
        """ Given a directory choose an image from it and set it as a wallpaper """
        # Image directories often contain Thumbs.db or other non-image
        # files or directories, try a few times to set a wallpaper, and then just give up.

        tries = 0
        done = False
        filenames = []
        files = os.listdir(pathToDir)

        try:
            # Priorwalls.txt is used so that we don't repeat an image
            # until every other image in that directory has been seen
            # The file is rewritten when necessary.
            prevList = open('priorWalls.txt', 'rb')
            filenames = [l.strip() for l in prevList.readlines()]
            prevList.close()
        except:
            pass

        files = [os.path.join(pathToDir, f) for f in files if not f.endswith('.db')]
        availChoices = [f for f in files if not f in filenames and not f.endswith('.db')]

        if not availChoices:
            # This entire directory has been "done" remove them
            # from the previously seen wallpapers and rewrite
            # the cache file without any of these entries.
            filenames = [f for f in filenames if not f in files]
            availChoices = files
            p = open('priorWalls.txt', 'wb')
            for f in filenames:
                p.write("%s\n"%(f))
            p.close()

        files = availChoices

        while (not done) and tries < 3:
            # Thumbs.db and other stuff can live in the same Folder
            # So try three times to set a wallpaper before giving up.
            try:
                image = random.choice(files)
                filename = image
                img = self.createWallPaperFromFile(filename, monitor)
                monitor.addWallpaper(self.bgImage, img)
                prevList = open('priorWalls.txt', 'ab')
                prevList.write("%s\n"%(filename))
                done = True
            except:
                import traceback; traceback.print_exc()
                print >> sys.stderr, filename, "failed"
                tries += 1
        return done

    def getMonitorDirs(self, monIndex):
        section = 'monitor_%d'%(monIndex)
        # Check for [monitor_0]
        if self.config.has_section(section):
            # check for its own paths section
            if self.config.has_option(section, 'paths'):
                dirs = self.config.get(section, 'paths').split('\n')
        else:
            dirs = self.dirs

        return dirs

    def setWallPaperFromDirList(self):
        """ Given a list of directories choose a directory """
        for monNum, monitor in enumerate(self.monitors):
            imageDir = random.choice(self.getMonitorDirs(monNum))
            status = self.setWallPaperFromFileList(imageDir, monitor)

        self.setWallpaper()

    def getImageDirectories(self):
        # Set global image directories.
        dirs = []
        try:
            dirs = self.config.get('directories', 'paths').split('\n')
        except:
            pass

        if not dirs:
            dirs = self.getDefaultDirs()
        return dirs

    def setWallPaperFromConfigDirs(self):
        self.setWallPaperFromDirList()

    def getConfigFileOptions(self, options):
        configFile = 'pywallpaper.conf'

        if options.configFile:
            configFile = options.configFile

        dirs = options.directories

        self.config = SafeConfigParser()
        self.config.readfp(open(configFile))

        if self.config.has_section('global'):
            self.Blending = self.config.getboolean('global', 'Blending')
            self.BlendRatio = self.config.getfloat('global', 'BlendRatio')
            self.Crop = self.config.getboolean('global', 'Crop')
            self.Fill = self.config.getboolean('global', 'Fill')
            self.Gradient = self.config.getboolean('global', 'Gradient')
            self.PreRotate = self.config.getboolean('global', 'PreRotate')

    def getCommandLineOptions(self):
        parser = OptionParser()
        parser.add_option("-t", "--time", dest="change_time",
                          help = "Change wallpaper time in minutes (0 = change once and exit [default])",
                          default = 0, type = "int")
        parser.add_option("-i", "--image", dest="singleImage", default = None,
                          help = "Set wallpaper to this image and exit (overrides -d)")

        parser.add_option("-d", "--directory", dest="directories", default = [], 
                          action="append", type="string",
                          help = "Add an image directory")

        parser.add_option("-c", "--config", dest="configFile", default = None,
                          help = "path to alternate config file (default <working dir>/pywallpaper.conf)")

        parser.add_option("-w", "--workingdir", dest="cwd", default=".",
                          help = "Working Directory (default .)")

        (options, args) = parser.parse_args()
        return (options, args)

    def go(self):
        options, args = self.getCommandLineOptions()
        self.getConfigFileOptions(options)

        self.dirs = self.getImageDirectories()
        
        if options.cwd and options.cwd != '.':
            os.cwd(cwd)
        if options.singleImage:
            self.setWallPaper(options.singleImage)
            
        elif not options.change_time:
            self.setWallPaperFromConfigDirs()
        else:
            sleepTime = options.change_time * 60.0
            while True:
                self.setWallPaperFromDirList()
                time.sleep(sleepTime)

if __name__ == '__main__':
    d = Desktop()
    d.go()
