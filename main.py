import argparse

from datetime import datetime
import math

from utils import parse_xls, create_csv


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Find distance")
    parser.add_argument("-i", "--input", help="Input file (default: ./files/test.xlsx)", default="./files/test.xlsx")
    parser.add_argument("-s", "--slice", help="Slice number of final list (default: None)", default=None)
    args = parser.parse_args()

    places = parse_xls(args.input, int(args.slice))
    create_csv(places)