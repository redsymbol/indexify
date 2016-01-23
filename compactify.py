#!/usr/bin/env python3

import argparse
import openpyxl

def get_args():
    parser = argparse.ArgumentParser()
    parser.add_argument('input-xls')
    parser.add_argument('output-xls')
    return parser.parse_args()

if __name__ == '__main__':
    args = get_args()
