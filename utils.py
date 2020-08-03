import csv
import xlrd

from math import cos, sin, radians, acos, ceil, floor


def parse_xls(filename, slice_num=None):
    """
    :param filename: name of xls file (str)
    :param slice_num: slice of the final list, default None (int)
    :return: list with all places (list)
    """
    
    result = []

    rb = xlrd.open_workbook(filename)
    sheet = rb.sheet_by_index(0)
    for rownum in range(1, sheet.nrows):
        row = sheet.row_values(rownum)
        _temp = {
            "name": row[0],
            "lon": row[-2],
            "lat": row[-1],
        }
        result.append(_temp)
    if slice_num:
        return result[:slice_num]
    else:
        return result


def find_distance(lat_1, lat_2, lon_1, lon_2):
    """
    :param lat_1: latitude of fisrt point
    :param lat_2: latitude of second point
    :param lon_1: longtitude of first point
    :param lon_2: longtitude of second point
    :return: distance between two points in kilometers
    """
    
    rlat_1 = radians(lat_1) 
    rlat_2 = radians(lat_2) 
    rlon_1 = radians(lon_1) 
    rlon_2 = radians(lon_2)
    cos_distance = sin(rlat_1) * sin(rlat_2) + cos(rlat_1) * cos(rlat_2) * cos(rlon_1 - rlon_2)
    try:
        return acos(cos_distance) * 6371
    except:
        return acos(round(cos_distance, 3)) * 6371


def get_places_names(places):
    """
    :param places: list of all places
    :return: list with places' names
    """

    names = [""]
    for place in places:
        names.append(place["name"])
    return names


def create_csv(places):
    f = csv.writer(open("distances.csv", "w+", encoding="cp1251", newline='', errors='replace'), delimiter=';')
    places_names = get_places_names(places)
    f.writerow(places_names)

    for place_1 in places:
        _temp = [find_distance(place_1["lat"], places[j]["lat"], place_1["lon"], places[j]["lon"]) for j in range(0, len(places))]
        _temp.insert(0, place_1["name"])
        f.writerow(_temp)