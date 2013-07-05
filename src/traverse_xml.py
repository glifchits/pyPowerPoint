'''
Created on 2013-07-03

@author: glifchits
'''
from bs4 import BeautifulSoup
import pprint

class Slide:

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

        return attrs

    def get_slide( self ):
        parts = []
        for sp in self.xml.find_all( 'p:sp' ):
            part = {'paragraphs':[]}

            for key, value in self.get_attrs( sp, 'p:ph' ):
                part[key] = value

            txt = sp.find( 'p:txbody' )
            for id, p in enumerate( txt.find_all( 'a:p' ) ):
                if p.find( 'a:t' ) is None:
                    continue
                par = {'id': id + 1}

                for key, value in self.get_attrs( p, "a:ppr" ):
                    par[key] = value

                paragraph = ""
                for r in p.find_all( 'a:r' ):
                    text = r.find( 'a:t' ).text  # this is the text portion!
                    formatting = r.find( 'a:rpr' ).attrs
                    # print formatting
                    if 'i' in formatting and formatting['i'] == u'1':
                        text = "<i>" + text + "</i>"
                    if 'baseline' in formatting and formatting['baseline'] == u'30000':
                        text = "<sup>" + text + "</sup>"
                    paragraph += text.strip() + ' '
                paragraph = paragraph.replace( " ,", "," )
                par['text'] = paragraph
                part['paragraphs'].append( par )
            parts.append( part )
        pprint.pprint( parts )

if __name__ == "__main__":
    with open( "sample.xml" ) as xml_file:
        s = Slide( xml_file.read() )
        s.get_slide()
