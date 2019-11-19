#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "jmsMaupin1"

import sys
import os
import argparse
from zipfile import ZipFile


def get_files_at_dir(directory):
    """Returns a list of all files in directory dir"""
    files = filter(lambda f: f[-4:] == "dotm", os.listdir(directory))
    path = directory + "/" if directory else "./"
    return map(lambda f: path + f, files)

def search_for_text(search, dotm, f):
    """
    Prints a list of all files that contain search in its text
    as well as the 40 characters before and after the index of the search term
    """
    line = dotm.read()
    if search in line:
        idx = line.index(search)
        print("Match found in file %s" % f)
        print("...%s..." % line[idx - 40: idx + 40])
        return 1
    
    return 0

def main(args):
    file_paths = get_files_at_dir(args.dir)
    searched_count = 0
    found_count = 0

    for f in file_paths:
        searched_count += 1
        with ZipFile(f) as zip:
            with zip.open("word/document.xml", "r") as dotm:
                found_count += search_for_text(args.search, dotm, f)

    print("Total dotm files searched: %s" % len(file_paths))
    print("Total dotm files matched: %s" % found_count)

    return 0


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument("--dir", help="Directory to search")
    parser.add_argument("search", help="Text to search documents for")
    args = parser.parse_args()
    status = main(args)
    sys.exit(status)