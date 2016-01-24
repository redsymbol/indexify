import unittest

import compactify

class TestCompactify(unittest.TestCase):
    def test_pagespec2rawlocator_pagesonly(self):
        pagespecs = [
            # pagespec, set of page #'s, set of "see also"
            ("137",     {137}, set()),
            ("137-8",   {137, 138}, set()),
            ("137-38",  {137, 138}, set()),
            ("137-138", {137, 138}, set()),
            ("137-42",  {137, 138, 139, 140, 141, 142}, set()),
            ("137-142", {137, 138, 139, 140, 141, 142}, set()),
            ("7, 3",    {7, 3}, set()),
            ("7-9, 3",  {7, 8, 9, 3}, set()),
            ]
        for pagespec, pages, see_also in pagespecs:
            with self.subTest(pagespec=pagespec, pages=pages, see_also=see_also):
                rawlocator = compactify.pagespec2rawlocator(pagespec)
                self.assertEqual(pages, rawlocator.pages)
                self.assertEqual(see_also, rawlocator.see_also)

    def test_render_locator__render_pages(self):
        page_sets = [
            # page_set, render string
            ({3}, "3"),
            ({3, 7}, "3, 7"),
            ({3, 4}, "3-4"),
            ({3, 4, 7}, "3-4, 7"),
            ({1, 2, 3, 4, 5}, "1-5"),
            ({1, 2, 3, 20, 21, 22}, "1-3, 20-22"),
            ({28, 29, 30, 31, 32}, "28-32"),
        ]
        for page_set, rendered in page_sets:
            with self.subTest(page_set=page_set, rendered=rendered):
                self.assertEqual(rendered, compactify._render_pages(page_set))
