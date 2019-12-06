#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "Bryan" + "revision"

import os
import sys
import zipfile
import argparse


def create_parser():  # creating arguments
    parser = argparse.ArgumentParser()  #an object for arguments
    parser.add_argument("--dir", default=".", help="Directory of dotm files to search") # -- means optional argument
    #default is current directory
    parser.add_argument("text", help="Text to search for within dotm files")
    return parser

def main(args):
    parser = create_parser()
    args = parser.parse_args(args)  #arguments created associated with command-line arguments

    print("Searching directory {} for dotm files with text '{}' ...".format(args.dir, args.text))


    file_list = os.listdir(args.dir)    #making a list of everything in args.dir
    # print(file_list)
    match_count = 0
    search_count = 0

    for file in file_list:

        if not file.endswith(".dotm"):
            print("Disregarding file {}".format(file))
            continue

        search_count += 1
        full_path = os.path.join(args.dir, file)   #combine directory path with file to form full_path

        with zipfile.ZipFile(full_path) as z_file:


            with z_file.open("word/document.xml") as doc:
                for line in doc:
                    print(line)
                    text_position = line.find(args.text)
                    if text_position >= 0:
                        match_count += 1
                        print("Match found in file {}".format(file))
                        print("... {} ...".format(line[text_position-40: text_position+40]))


    print("Files searched: {}".format(search_count))
    print("Times matched: {}".format(match_count))

if __name__ == '__main__':
    main(sys.argv[1:])
