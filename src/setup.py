'''
Created on 13.09.2016

@author: ChrisCuts
'''

from distutils.core import setup
import py2exe
import sys

sys.argv.append('py2exe')
 
options = {'py2exe':{'compressed':1,  
                    'bundle_files': 2, 
                    'dist_dir': "../dist/"}} 

setup(name='TravelExp',
      version='1.1',
      description='Travel Expense Creator',
      author='ChrisCuts',
      url='https://github.com/ChrisCuts/TravelExp',
      console=['travelexp.py'],
      options=options,
      zipfile = None)