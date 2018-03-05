import unittest

from readData import Section

class SectionMethods(unittest.TestCase):
    def test_colon_stripped_from_section_name(self):
        # Act
        section = Section("Systems:")
        # Assert
        self.assertEqual(section.name, 'Systems', msg='trailing colons are expected to be removed from section name')

    def test_leading_whitespaces_are_removed_from_section_name(self):
        # Act
        section = Section(" Systems")
        # Assert
        self.assertEqual(section.name, 'Systems', msg='leading whitespaces are expected to be removed from section name')

    def test_trailing_whitespaces_are_removed_from_section_name(self):
        # Act
        section = Section("Systems ")
        # Assert
        self.assertEqual(section.name, 'Systems', msg='trailing whitespaces are expected to be removed from section name')

if __name__ == '__main__':
    unittest.main()
