'''
Created on 2013-07-05

@author: glifchits
'''
import unittest
from traverse_xml import Slide

class SampleTest( unittest.TestCase ):

    def setUp( self ):
        self.xml_file = open( 'test_files/sample.xml' )
        self.slide = Slide( self.xml_file.read() )

    def tearDown( self ):
        self.xml_file.close()

    def test_getTitle( self ):
        self.slide

    def test_str_MarkDown( self ):
        md = open( 'test_files/sample_xml.md' )
        markdown = md.read()
        self.assertEqual( str( self.slide ).strip(), markdown.strip() )
        md.close()


if __name__ == "__main__":
    # import sys;sys.argv = ['', 'Test.testName']
    unittest.main()
