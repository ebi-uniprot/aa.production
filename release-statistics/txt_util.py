
def has_text(s):
    """checks if passed-in string is not None and has non-whitespace text"""
    return s and not s.isspace();


def skip_first_section(in_file):
    """skips the 'first*' text section of the input file, till the 1st empty line

    * 'first' relative to current cursor position in in_file.
    The function will skip initial empty lines (if any) before the section
    actually containing text"""
    text_found = False
    for line in in_file:
        line_has_txt = has_text(line)
        if not line_has_txt and text_found:
            # we've done the job: it's empty line, and we've already seen text
            return
        text_found = line_has_txt
