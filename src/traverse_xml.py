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

    self.slide will be parsed and the structure looks like this:

    [{ id: int
       paragraphs: [{ id: int
                      words: [{ id: int
                                text: ""
                                italic: int
                                baseline: int }]
                    }]
    }]



    '''
    INDENT = '  '

    def __init__( self, xml_string ):
        self.xml = BeautifulSoup( xml_string )
        self.slide = self._parse_slide()

    def __str__( self ):
        def wrap( string, tag ):
            if tag.startswith( '<' ) and tag.endswith( '>' ):
                open = tag
                close = tag[:-1] + '/>'
            else:
                open, close = tag, tag

            return open + string + close

        def delimiters( string ):
            pass

        s = ""
        by_id = lambda x: x['id']
        for part in sorted( self.slide, key = by_id ):
            para = ""
            if 'type' in part and part['type'] == 'title':
                para += "# "

            for paragraph in sorted( part['paragraphs'], key = by_id ):
                sentence = ""
                if 'indent' in paragraph:
                    sentence += self.INDENT * paragraph['indent'] + ' '

                for word in sorted( paragraph['words'], key = by_id ):
                    this_word = word['text']
                    if 'italic' in word:
                        this_word = wrap( this_word, "<i>" )
                    if 'baseline' in word:
                        this_word = wrap( this_word, "<sup>" )
                    delimiters( this_word )  # TODO: Markdown characters that need to be delimited
                    sentence = sentence.strip() + this_word + ' '
                para += sentence + '\n'
            s += para + '\n'
        return s


    def _get_attrs( self, parent, tag ):
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

    def _parse_slide( self ):
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

            for key, value in self._get_attrs( sp, 'p:ph' ):
                part[key] = value

            txt = sp.find( 'p:txbody' )
            for p_id, p in enumerate( txt.find_all( 'a:p' ) ):
                if p.find( 'a:t' ) is None:  # skip any paragraph with no text
                    continue

                par = {'id': p_id + 1, 'words':[]}

                for key, value in self._get_attrs( p, "a:ppr" ):
                    par[key] = value

                for r_id, r in enumerate( p.find_all( 'a:r' ) ):
                    word = {'id':r_id,
                            'text': r.find( 'a:t' ).text.strip()}  # this is the text portion!

                    for key, value in self._get_attrs( r, "a:rpr" ):
                        word[key] = value

                    par['words'].append( word )

                part['paragraphs'].append( par )

            parts.append( part )

        return parts

if __name__ == "__main__":
    with open( "sample.xml" ) as xml_file:
        s = Slide( xml_file.read() )
        print s
