import sys
import os
# Add the root directory to sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from main import has_edits

def test_has_edits_with_insertion():
    # Test case with insertion (**text**)
    content = "This is a **test**."
    assert has_edits(content) is True

def test_has_edits_with_deletion():
    # Test case with deletion (--text--)
    content = "This is a --test--."
    assert has_edits(content) is True

def test_has_edits_with_no_edits():
    # Test case with no edits
    content = "This is a test."
    assert has_edits(content) is False

def test_has_edits_with_both_insertion_and_deletion():
    # Test case with both insertion and deletion
    content = "This is a **test** and a --sample--."
    assert has_edits(content) is True