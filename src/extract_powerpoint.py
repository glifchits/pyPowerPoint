'''
extract_powerpoint
==================
The intention of this program is to open a PPTX file and extract the text from it.

Created on 2013-07-03
@author: glifchits
'''
import os
import sys
import re
from zipfile import ZipFile
from xml.etree import ElementTree

def slide_matcher( fname ):
    r = re.match( "(^slide)(\d+)(\.xml)", fname )
    return r.groups() if r != None else None

def traverse_xml( slide ):
    xml = ElementTree.parse( slide )
    root = xml.getroot()
    traverse_xml_aux( root, 0 )

def traverse_xml_aux( root, level ):
    for child in root:
        space = "".join( [str( x + 1 ) for x in range( level )] )
        print space, "<" + child.tag[child.tag.find( '}' ) + 1:] + ">",
        print child.attrib,
        print child.text if child.text != None else ""
        # print space, child.tag, child.text
        traverse_xml_aux( child, level + 1 )

def get_slide_text( slide ):
    xml = ElementTree.parse( slide )
    root = xml.getroot()
    tag = '{http://schemas.openxmlformats.org/drawingml/2006/main}t'
    for child in root.iter():
        print child.attrib, child.text

def get_slides( fname ):
    try:
        os.mkdir( 'temp' )
    except OSError:
        pass
    ZipFile( fname ).extractall( "temp" )
    os.chdir( "temp/ppt/slides" )
    # get all files in dir
    files = os.listdir( os.getcwd() )
    # filter out what isn't a valid slide
    files = filter( lambda x: slide_matcher( x ) != None, files )
    # sort by slide index in a four-digit leading 0 number (eg. 0003)
    files.sort( key = lambda x: "{0:>04}".format( slide_matcher( x )[1] ) )
    for slide in [files[4]]:  # will need to parse all files
        print "\n" + slide
        traverse_xml( slide )
        # get_slide_text( slide )

if __name__ == '__main__':
    # file = sys.argv[1]
    file = "../test/07 Impeachment of K.pptx"
    # get_slides( file )
    traverse_xml( "sample.xml" )
