#!/usr/bin/env python
# -*- coding: utf-8 -*-
import argparse
from zipfile import ZipFile
import os
import re
"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "Nathaniel Lyttle"


def main():
    #initialize command line parser
    parser = argparse.ArgumentParser()
    parser.add_argument('--dir', type=str, default=os.getcwd(), help="path to directory to be serched. (Default = current working directory)")
    parser.add_argument('--search_str', type=str, default="$", help="substring to search for. (Default = '$')")

    #initalize environment
    os.chdir(parser.parse_args().dir)
    matches = {}
    file_count = 0
    r_search = "((.{1,20}|)\\"+parser.parse_args().search_str+"(.{1,20}|))"
    matched_string = 0 #index of matched string in matches touple

    #Main search loop
    for filename in os.listdir(os.getcwd()):
        file_count += 1
        if filename.endswith(".dotm"):
            with ZipFile(filename).open("word/document.xml") as search: #dotm files are zipped files
                for i,word in enumerate(search):
                    r = re.search(r_search, word) #finds search_str surrounded by 20 characters on each side
                    try:
                        matches[filename] = (r.group(), i)
                        continue
                    except AttributeError:
                        pass

    for key in list(matches.keys()):
        print(key + " "*(15-len(key)) +": " + matches[key][matched_string])
    print("Total number of files searched :" + str(file_count))
    print("Total number of matches        :" + str(len(matches.keys())))

if __name__ == '__main__':
    main()
