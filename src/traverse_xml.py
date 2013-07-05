'''
Created on 2013-07-03

@author: glifchits
'''
from bs4 import BeautifulSoup
import pprint

class Slide:
    '''
    A slide. Given the .pptx XML for a specific PowerPoint slide, this class creates a nice nested
    dictionary representation of the bare minimum content.
    '''

    def __init__( self, xml_string ):
        self.xml = BeautifulSoup( xml_string )

    def get_attrs( self, parent, tag ):
        attrs = []
        a = parent.find( tag )
        if a is None:
            return []

        if tag == 'p:ph':  # paragraph attributes
            if 'type' in a.attrs:
                x = ( 'type', a['type'] )
                attrs.append( x )

        elif tag == 'a:ppr':
            if 'lvl' in a.attrs:
                x = ( 'indent', int( a['lvl'] ) )
                attrs.append( x )

        elif tag == 'a:rpr':
            if 'i' in a.attrs:
                x = ( 'italic', int( a['i'] ) )
                attrs.append( x )

            if 'baseline' in a.attrs:
                x = ( 'baseline', int( a['baseline'] ) )
                attrs.append( x )

        return attrs

    def get_slide( self ):
        '''
        <p:sp>      a part, a logically contiguous piece. eg, title, content, slide number
            <p:ph>    attribute tag

        <p:txbody>  the tag that contains the text of a part

        <a:p>       a paragraph (subdivided part). eg, each point in the bullet list
            <a:ppr>   attribute tag

        <a:r>       "sentences". from what I see, its just for different formatting styles used
                    within a paragraph
            <a:rpr>   attribute tag
            <a:t>     text tag
        '''
        parts = []
        for sp_id, sp in enumerate( self.xml.find_all( 'p:sp' ) ):
            part = {'id':sp_id, 'paragraphs':[]}

            for key, value in self.get_attrs( sp, 'p:ph' ):
                part[key] = value

            txt = sp.find( 'p:txbody' )
            for p_id, p in enumerate( txt.find_all( 'a:p' ) ):
                if p.find( 'a:t' ) is None:  # skip any paragraph with no text
                    continue

                par = {'id': p_id + 1, 'words':[]}

                for key, value in self.get_attrs( p, "a:ppr" ):
                    par[key] = value

                for r_id, r in enumerate( p.find_all( 'a:r' ) ):
                    word = {'id':r_id,
                            'text': r.find( 'a:t' ).text.strip()}  # this is the text portion!

                    for key, value in self.get_attrs( r, "a:rpr" ):
                        word[key] = value

                    par['words'].append( word )

                part['paragraphs'].append( par )

            parts.append( part )

        pprint.pprint( parts )

if __name__ == "__main__":
    with open( "sample.xml" ) as xml_file:
        s = Slide( xml_file.read() )
        s.get_slide()
