'''
slide
=====

This contains a class which parses the XML data of a .pptx slide and returns a JSON-like dictionary
structure containing the barebones text and metadata of that slide.

The string encoding of this class returns a Markdown friendly formatting of the slide.

@author: George Lifchits
Created on 2013-07-03
'''
from bs4 import BeautifulSoup

class Slide:
    '''
    A slide. Given the .pptx XML for a specific PowerPoint slide, this class creates a nice nested
    dictionary representation of the bare minimum content.

    The slide will be parsed and the resulting (JSON-like) structure looks like this:

    [{ id: int
       type: 'ctrTitle', 'subTitle', 'title', 'sldNum'
       paragraphs: [{ id: int
                      words: [{ id: int
                                text: ""
                                italic: int
                                baseline: int }]
                    }]
    }]
    '''

    MAIN_TITLE = '# '
    SUBTITLE = '## '
    SLIDE_TITLE = '## '
    INDENT = '  '
    BULLET = '* '

    def __init__( self, xml_string ):
        self.xml = BeautifulSoup( xml_string )
        self.slide = self._parse_slide()

    def __str__( self ):
        '''
        Markdown representation of the slide's text
        '''
        def wrap( string, tag ):
            '''
            Helper function to easily wrap a string with tags. Auto closes HTML tags
            '''
            if tag.startswith( '<' ) and tag.endswith( '>' ):
                open = tag
                close = '</' + tag[1:]
            else:
                open, close = tag, tag
            return open + string + close

        def delimiters( string ):
            '''
            Helper function to clean up characters that should be delimited in the Markdown syntax
            '''
            chars = '*`'
            for char in chars:
                i = string.find( char )
                while i >= 0:
                    string = string[:i] + '\\' + string[i:]
                    i = string.find( char )

            return string

        s = ""
        by_id = lambda x: x['id']

        for part in sorted( self.slide, key = by_id ):
            str_part = ""

            #### Part formatting <p:ph>
            type = part.get( 'type', 0 )
            if type == 'ctrTitle':
                str_part += self.MAIN_TITLE
            if type == 'subTitle':
                str_part += self.SUBTITLE
            if type == 'title':
                str_part += self.SLIDE_TITLE
            #### END

            for paragraph in sorted( part['paragraphs'], key = by_id ):
                str_paragraph = ""

                # ### Paragraph formatting <p:ppr>
                indent = paragraph.get( 'indent', 0 )
                list = paragraph.get( 'list', 0 )

                str_paragraph += self.INDENT * indent
                if list == 'bullet':
                    str_paragraph += self.BULLET
                elif list == 'arabicPeriod':
                    str_paragraph += "1."
                # ### END

                for word in sorted( paragraph['words'], key = by_id ):
                    str_word = word['text']

                    # ### Word formatting <a:rpr>
                    baseline = word.get( 'baseline', 0 )
                    italic = word.get( 'italic', 0 )

                    if italic == 1:
                        str_word = wrap( str_word, "<i>" )

                    if baseline == 30000:
                        str_word = wrap( str_word, "<sup>" )
                    # ### END

                    str_paragraph += str_word + ' '

                str_part += str_paragraph + '\n'

            s += str_part + '\n'

        return s.strip().encode( 'ascii', 'ignore' )

    def _parse_slide( self ):
        '''
        Parses the XML of a slide and produces the JSON-like structure of text content of the slide

        <p:sp>      a part, a logically contiguous piece. eg, title, content, slide number
            <p:ph>     attribute tag
            <p:txbody> the tag that contains the text of a part

        <a:p>       a paragraph (subdivided part). eg, each point in the bullet list
            <a:ppr>    attribute tag

        <a:r>       words. from what I see, its just for different formatting styles used within a
                    paragraph
            <a:rpr>    attribute tag
            <a:t>      text tag
        '''
        def _get_attrs( parent, tag ):
            a = parent.find( tag )
            if a is None:
                return []

            attrs = []
            if tag == 'p:ph':  # part attributes
                if 'type' in a.attrs:
                    attrs.append( ( 'type', a['type'] ) )

            elif tag == 'a:ppr':  # paragraph attributes
                if 'lvl' in a.attrs:
                    attrs.append( ( 'indent', int( a['lvl'] ) ) )

                x = a.find( 'a:buautonum' )
                if x is not None:
                    if 'arabicPeriod' in x.attrs:
                        attrs.append( ( 'list', 'numbered' ) )
                elif a.find( 'a:bunone' ) is None:
                    attrs.append( ( 'list', 'bullet' ) )

            elif tag == 'a:rpr':  # sentence attributes
                if 'i' in a.attrs:
                    attrs.append( ( 'italic', int( a['i'] ) ) )

                if 'baseline' in a.attrs:
                    attrs.append( ( 'baseline', int( a['baseline'] ) ) )

            return attrs

        parts = []
        for sp_id, sp in enumerate( self.xml.find_all( 'p:sp' ) ):
            part = {'id':sp_id, 'paragraphs':[]}

            for key, value in _get_attrs( sp, 'p:ph' ):
                part[key] = value

            txt = sp.find( 'p:txbody' )
            for p_id, p in enumerate( txt.find_all( 'a:p' ) ):
                if p.find( 'a:t' ) is None:  # skip any paragraph with no text
                    continue

                par = {'id': p_id + 1, 'words':[]}

                for key, value in _get_attrs( p, "a:ppr" ):
                    par[key] = value

                for r_id, r in enumerate( p.find_all( 'a:r' ) ):
                    if r.find( 'a:t' ) is None:
                        continue

                    word = {'id':r_id,
                            'text': r.find( 'a:t' ).text.strip()}  # this is the text portion!

                    for key, value in _get_attrs( r, "a:rpr" ):
                        word[key] = value

                    par['words'].append( word )

                part['paragraphs'].append( par )

            parts.append( part )

        return parts

if __name__ == "__main__":
    with open( "../test/test_files/sample.xml" ) as xml_file:
        s = Slide( xml_file.read() )
        print s
