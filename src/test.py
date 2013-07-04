'''
Created on 2013-07-03

@author: glifchits
'''
import unittest
import powerpoint

import re

x = ["slide1.xml", "slide25.xml", "slides10.xml", "slide", "", "_rels"]

def matcher( string ):
    x = re.match( "(^slide)(\d+)(\.xml)", string )
    if x is not None:
        return x.groups()

for f in x:
    print matcher( f ) != None

class Test( unittest.TestCase ):

    def setUp( self ):
        pass

    def tearDown( self ):
        pass

    def testName( self ):
        pass


'''
if __name__ == "__main__":
    # import sys;sys.argv = ['', 'Test.testName']
    unittest.main()
'''
