'''
powerpoint
==========

This contains a class which can extract the textual information from a .pptx file.
The slides themselves are parsed by a class in the `slide` module.

@author: George Lifchits
Created on 2013-07-05
'''
import os
import shutil
import sys
import re
from zipfile import ZipFile

from slide import Slide

class PowerPoint:

    def __init__( self, file ):
        self.pptx = file
        slide_files = self._get_slides()
        slides_json = []
        for slide_file_name in slide_files:
            with open( slide_file_name ) as slide_file:
                slides_json.append ( Slide( slide_file.read() ) )
        self.slides = slides_json

    def __del__( self ):
        ''' /temp directory will delete when this class is destroyed '''
        if os.path.isdir( 'temp' ):
            shutil.rmtree( 'temp' )

    def __str__( self ):
        s = ''
        for slide in self.slides:
            s += str( slide )
        return s

    def get_slides( self ):
        return self.slides

    def _get_slides( self ):
        '''
        Gets the slideX.xml files from a .pptx file
        '''

        def slide_matcher( fname ):
            ''' Regex to match files that are valid slides '''
            r = re.match( "(^slide)(\d+)(\.xml)", fname )
            return r.groups() if r != None else None

        original_cwd = os.getcwd()
        files = []
        try:
            fname = self.pptx
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
            # append the full path to file names
            files = [os.getcwd() + '/' + file for file in files]
        finally:
            os.chdir( original_cwd )

        return files


if __name__ == "__main__":
    if len( sys.argv ) != 2:
        sys.stderr.write( "Usage: python %s <some_file>[.pptx]\n" % sys.argv[0] )
        sys.exit()
    else:
        filename = sys.argv[1]
        p = PowerPoint( filename )
        print p
