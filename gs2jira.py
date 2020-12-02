#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""Template for python3 terminal scripts.
This gist allows you to quickly create a functioning
python3 terminal script using argparse and subprocess.
"""

import gspread

__author__ = "bursno22"
__license__ = "MIT"
__version__ = "0.0.1"


def main():
    gc = gspread.oauth()
    sh = gc.open("ticket_management")
    print(sh.sheet1.get('G8'))

if __name__ == '__main__':
    main()