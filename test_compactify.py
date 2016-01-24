import unittest

from compactify import pagespec2rawlocator

class TestCompactify(unittest.TestCase):
    def test_pagespec2rawlocator_pagesonly(self):
        pagespecs = [
            # pagespec, set of page #'s, set of "see also"
            ("137",     {137}, {}),
            ("137-8",   {137, 138}, {}),
            ("137-38",  {137, 138}, {}),
            ("137-138", {137, 138}, {}),
            ("137-42",  {137, 138, 139, 140, 141, 142}, {}),
            ("137-142", {137, 138, 139, 140, 141, 142}, {}),
            ("7, 3",    {7, 3}, {}),
            ("7-9, 3",  {7, 8, 9, 3}, {}),
            ]
        for pagespec, pages, see_also in pagespecs:
            with self.subTest(pagespec=pagespec, pages=pages, see_also=see_also):
                rawlocator = pagespec2rawlocator(pagespec)
                self.assertEqual(pages, rawlocator.pages)
                self.assertEqual(see_also, rawlocator.see_also)
