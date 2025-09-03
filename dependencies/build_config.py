#!/usr/bin/env python3
"""
PyInstaller build configuration for ArtWork.py
Handles both Windows and Mac builds
"""

import sys
import os
import platform
import PyInstaller.__main__

def build_app():
    """Build the executable using PyInstaller"""
    
    # Determine platform-specific settings
    is_windows = platform.system() == 'Windows'
    is_mac = platform.system() == 'Darwin'
    
    # Base PyInstaller arguments
    args = [
        'ArtWork.py',  # Your main script
        '--name=DataProcessor',  # Name of the executable
        '--onefile',  # Single executable file
        '--windowed',  # No console window (GUI app)
        '--clean',  # Clean PyInstaller cache
        '--noconfirm',  # Overwrite output without confirmation
        
        # Hidden imports for pandas, kivy, and other dependencies
        '--hidden-import=pandas._libs.tslibs.timedeltas',
        '--hidden-import=pandas._libs.tslibs.np_datetime',
        '--hidden-import=pandas._libs.tslibs.nattype',
        '--hidden-import=pandas._libs.skiplist',
        '--hidden-import=pandas._libs.reshape',
        '--hidden-import=pandas._libs.tslibs.offsets',
        '--hidden-import=pandas._libs.tslibs.parsing',
        '--hidden-import=pandas._libs.tslibs.conversion',
        '--hidden-import=pandas._libs.tslibs.period',
        '--hidden-import=pandas._libs.tslibs.timestamps',
        '--hidden-import=pandas._libs.tslibs.timezones',
        '--hidden-import=pandas._libs.tslibs.fields',
        '--hidden-import=pandas._libs.tslibs.dtypes',
        '--hidden-import=pandas._libs.window.aggregations',
        '--hidden-import=pandas._libs.window.indexers',
        '--hidden-import=pandas._libs.indexing',
        '--hidden-import=pandas._libs.index',
        '--hidden-import=pandas._libs.algos',
        '--hidden-import=pandas._libs.join',
        '--hidden-import=pandas._libs.sparse',
        '--hidden-import=pandas._libs.reduction',
        '--hidden-import=pandas._libs.parsers',
        '--hidden-import=pandas._libs.groupby',
        '--hidden-import=pandas._libs.properties',
        '--hidden-import=pandas._libs.writers',
        '--hidden-import=pandas._libs.ops_dispatch',
        '--hidden-import=pandas._libs.missing',
        '--hidden-import=pandas._libs.hashtable',
        '--hidden-import=pandas._libs.lib',
        '--hidden-import=numpy.random.common',
        '--hidden-import=numpy.random.bounded_integers',
        '--hidden-import=numpy.random.entropy',
        '--hidden-import=openpyxl.cell._writer',
        '--hidden-import=xlsxwriter',
        '--hidden-import=xlrd',
        '--hidden-import=kivy.core.window',
        '--hidden-import=kivy.core.text',
        '--hidden-import=kivy.core.image',
        '--hidden-import=kivy.core.gl',
        '--hidden-import=kivy.graphics.instructions',
        '--hidden-import=kivy.graphics.context_instructions',
        '--hidden-import=kivy.graphics.vertex_instructions',
        '--hidden-import=kivy.graphics.canvas',
        '--hidden-import=kivy.graphics.texture',
        '--hidden-import=kivy.graphics.transformation',
        '--hidden-import=kivy.uix.widget',
        '--hidden-import=kivy.uix.label',
        '--hidden-import=kivy.uix.button',
        '--hidden-import=kivy.uix.textinput',
        '--hidden-import=kivy.uix.boxlayout',
        '--hidden-import=kivy.uix.anchorlayout',
        '--hidden-import=kivy.uix.progressbar',
        '--hidden-import=kivy.uix.popup',
        
        # Collect all data from these packages
        '--collect-all=kivy',
        '--collect-all=pandas',
        '--collect-all=openpyxl',
        '--collect-all=xlsxwriter',
        
        # Add binary files
        '--collect-binaries=pandas',
        '--collect-binaries=numpy',
    ]
    
    # Windows-specific settings
    if is_windows:
        args.extend([
            '--icon=NONE',  # You can add a .ico file if you want
            '--hidden-import=win32com.client',
            '--hidden-import=pythoncom',
        ])
    
    # Mac-specific settings
    elif is_mac:
        args.extend([
            '--icon=NONE',  # You can add a .icns file if you want
            '--hidden-import=tkinter',
            '--hidden-import=tkinter.filedialog',
            '--osx-bundle-identifier=com.dataprocessor.app',
        ])
        
        # For Mac .app bundle
        args[2] = '--onedir'  # Change to onedir for Mac .app
    
    # Add data files for tkinter on Mac
    if is_mac:
        # Try to locate tkinter data files
        import tkinter
        tk_path = tkinter.__file__
        tk_dir = os.path.dirname(tk_path)
        if os.path.exists(tk_dir):
            args.append(f'--add-data={tk_dir}:tkinter')
    
    print(f"Building for {platform.system()}...")
    print("Arguments:", args)
    
    # Run PyInstaller
    PyInstaller.__main__.run(args)
    
    print("Build complete!")
    
    # Platform-specific output location
    if is_windows:
        print(f"Executable location: dist/DataProcessor.exe")
    elif is_mac:
        print(f"App bundle location: dist/DataProcessor.app")

if __name__ == "__main__":
    build_app()
