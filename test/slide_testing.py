'''
Created on 2013-07-05

@author: glifchits
'''
import unittest
from slide import Slide

class SampleTest( unittest.TestCase ):

    def setUp( self ):
        self.xml_file = open( 'files/sample.xml' )
        self.slide = Slide( self.xml_file.read() )

    def tearDown( self ):
        self.xml_file.close()

    def test_str_MarkDown( self ):
        md = open( 'files/sample_xml.md' )
        markdown = md.read().splitlines()
        md.close()
        slide = str( self.slide ).splitlines()
        for real_line, test_line in zip( slide, markdown ):
            self.assertEqual( real_line.strip(), test_line.strip() )


class BasicTest( unittest.TestCase ):

    def setUp( self ):
        self.xml_file = open( 'files/basic_slides/slide1.xml' )
        self.slide = Slide( self.xml_file.read() )

    def tearDown( self ):
        self.xml_file.close()

    def test_str_MarkDown( self ):
        md = open( 'files/basic_slides/slide1.md' )
        markdown = md.read().splitlines()
        md.close()
        slide = str( self.slide ).splitlines()
        for real_line, test_line in zip( slide, markdown ):
            self.assertEqual( real_line.strip(), test_line.strip() )


if __name__ == "__main__":
    # import sys;sys.argv = ['', 'Test.testName']
    unittest.main()
