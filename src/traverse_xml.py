'''
Created on 2013-07-03

@author: glifchits
'''
from bs4 import BeautifulSoup
import pprint

def traverse2( xml ):
    for sp in xml.find_all( 'p:sp' ):
        part = {'paragraphs':[]}

        ph = sp.find( 'p:ph' )
        if 'type' in ph.attrs:
            part['type'] = ph['type']

        txt = sp.find( 'p:txbody' )
        for id, p in enumerate( txt.find_all( 'a:p' ) ):
            if p.find( 'a:t' ) is None:
                continue
            par = {'id': id + 1}
            formatting = p.find( 'a:ppr' )
            if formatting is not None:
                formatting = formatting.attrs
                if 'lvl' in formatting:
                    par['indent'] = int( formatting['lvl'] )

            paragraph = ""
            for r in p.find_all( 'a:r' ):
                text = r.find( 'a:t' ).text
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
        pprint.pprint( part )

if __name__ == "__main__":
    with open( "sample.xml" ) as xml_file:
        xml = BeautifulSoup( xml_file.read() )
        traverse2( xml )
