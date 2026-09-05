#!/usr/bin/env python3
"""
Test Runner for BibTeX to Zotero Skill & Engine
"""

import os
import sys
import unittest

def main():
    print("===================================================")
    print("  Running BibTeX to Zotero Unit Test Suite")
    print("===================================================")
    
    loader = unittest.TestLoader()
    suite = loader.discover(start_dir=os.path.join(os.path.dirname(__file__), 'tests'), pattern='test_*.py')
    
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    
    if result.wasSuccessful():
        print("\n[SUCCESS] All unit tests passed cleanly!")
        sys.exit(0)
    else:
        print("\n[FAILURE] Some unit tests failed. Please review errors above.")
        sys.exit(1)

if __name__ == '__main__':
    main()
