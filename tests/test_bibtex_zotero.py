#!/usr/bin/env python3
"""
Unit Test Suite for BibTeX to Zotero Word Converter
"""

import os
import sys
import unittest

# Add scripts directory to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'scripts'))

from bibtex_to_zotero_word import (
    clean_latex_accents,
    clean_bib_value,
    parse_authors,
    parse_bibtex,
    parse_bibtex_fields,
    format_citation_label,
    generate_random_id,
    build_custom_xml,
    update_rels,
    update_content_types
)

class TestBibtexZotero(unittest.TestCase):

    def test_clean_latex_accents(self):
        """Test conversion of LaTeX accent codes to UTF-8 Unicode."""
        self.assertEqual(clean_latex_accents(r"{\L}ukasz"), "Łukasz")
        self.assertEqual(clean_latex_accents(r"Timoth{\'e}e"), "Timothée")
        self.assertEqual(clean_latex_accents(r"Rozi{\`e}re"), "Rozière")
        self.assertEqual(clean_latex_accents(r"M\"uller"), "Müller")
        self.assertEqual(clean_latex_accents(r"Karpukhin, Vladimir and O{\u{g}}uz, Barlas"), "Karpukhin, Vladimir and Oğuz, Barlas")
        self.assertEqual(clean_latex_accents(r"LaTeX \& Symbol \% test \$"), "LaTeX & Symbol % test $")

    def test_parse_authors(self):
        """Test author parsing and filtering of 'others'."""
        author_str = r"Vaswani, Ashish and Shazeer, Noam and Parmar, Niki and others"
        authors = parse_authors(author_str)
        self.assertEqual(len(authors), 3)
        self.assertEqual(authors[0], {'family': 'Vaswani', 'given': 'Ashish'})
        self.assertEqual(authors[1], {'family': 'Shazeer', 'given': 'Noam'})
        self.assertEqual(authors[2], {'family': 'Parmar', 'given': 'Niki'})

    def test_parse_bibtex_depth_balanced(self):
        """Test depth-balanced brace parsing for nested LaTeX values."""
        bib_body = r"""
            title={Dense passage retrieval for open-domain question answering},
            author={Karpukhin, Vladimir and O{\u{g}}uz, Barlas and Sewon, Min},
            journal={arXiv preprint arXiv:2004.04906},
            year={2020}
        """
        fields = parse_bibtex_fields(bib_body)
        self.assertEqual(fields['title'], "Dense passage retrieval for open-domain question answering")
        self.assertEqual(fields['author'], "Karpukhin, Vladimir and Oğuz, Barlas and Sewon, Min")
        self.assertEqual(fields['year'], "2020")

    def test_format_citation_label(self):
        """Test formatted label generation for IEEE vs APA styles."""
        entry = {
            'fields': {
                'author': 'Vaswani, Ashish and Shazeer, Noam',
                'year': '2017'
            }
        }
        self.assertEqual(format_citation_label(entry, 1, style='ieee'), "[1]")
        self.assertEqual(format_citation_label(entry, 1, style='apa'), "(Vaswani & Shazeer, 2017)")
        
        entry_et_al = {
            'fields': {
                'author': 'Devlin, Jacob and Chang, Ming-Wei and Lee, Kenton',
                'year': '2018'
            }
        }
        self.assertEqual(format_citation_label(entry_et_al, 2, style='apa'), "(Devlin et al., 2018)")

    def test_custom_xml_property_chunking(self):
        """Test 255-character property chunking in docProps/custom.xml."""
        custom_bytes = build_custom_xml(style='ieee')
        custom_str = custom_bytes.decode('utf-8')
        
        # Check property names
        self.assertIn('name="ZOTERO_PREF_1"', custom_str)
        self.assertIn('name="ZOTERO_PREF_2"', custom_str)
        
        # Check pid allocation
        self.assertIn('pid="2"', custom_str)
        self.assertIn('pid="3"', custom_str)
        
        # Check double quote escaping (must be unescaped inside vt:lpwstr)
        self.assertIn('data-version="3"', custom_str)
        self.assertNotIn('data-version=&quot;3&quot;', custom_str)

    def test_update_rels_unique_id(self):
        """Test unique rId allocation in _rels/.rels to prevent collisions."""
        existing_rels = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Target="word/document.xml"/>'
            '<Relationship Id="rId2" Target="docProps/core.xml"/>'
            '<Relationship Id="rId3" Target="docProps/app.xml"/>'
            '<Relationship Id="rId4" Target="docProps/thumbnail.jpeg"/>'
            '</Relationships>'
        ).encode('utf-8')
        
        updated_bytes = update_rels(existing_rels)
        updated_str = updated_bytes.decode('utf-8')
        
        # Should allocate rId5 cleanly
        self.assertIn('Id="rId5"', updated_str)
        self.assertIn('Target="docProps/custom.xml"', updated_str)

if __name__ == '__main__':
    unittest.main()
