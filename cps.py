#!/bin/python

from enum import Enum
import re
import pandas as pd
import numpy as np
import math


#
# class
#

class Device_t(Enum):
    NULL    = 0
    MINI    = 1
    LAPTOP  = 2


class Brand_t(Enum):
    NULL    = 0
    HP      = 1
    SURFACE = 2
    Lenovo  = 3


# this represents one row of PileDevice class
class DeviceInfo:
    #
    # var
    #

    def __init__(self, _location, _serial_num, _id_tag):
        self.location   = str(_location)
        self.serial_num = str(_serial_num)

        try:
            self.id_tag     = str(int(_id_tag))
        except:
            self.id_tag     = str(_id_tag)



# the class is used for all type/model of devices  -mini, laptop, hp...
# this basically represents the 3-column sets on sheets
class PileDevice:
    #
    # var
    #

    def __init__(self, _dev_type, _brand, _model):
        self.dev_type   = _dev_type
        self.brand      = _brand
        self.model      = _model
        self.lst_dev   = []
        self.len_lst_dev = 0

    def add(self, location, serial_num, id_tag):
        self.lst_dev.append(DeviceInfo(location, serial_num, id_tag))
        self.len_lst_dev += 1

    def pop(self, index):
        if index >= self.len_lst_dev:
            raise IndexError(
            f"Index of <{index}> is out of len <{self.len_lst_dev}>")
        else:
            self.len_lst_dev -= 1
            return self.lst_dev.pop(index)
        


#
# func
#

def sheet_to_listPile(df_sheet, dev_type, brand):

    # get all model names
    # {
    lst_keys = df_sheet.keys()

    lst_mdl = [_key for _key in lst_keys if not
                re.match(r"Unnamed: \d+", _key)]

    lst_mdl = [_key for _key in lst_mdl if not any(
        re.match(rf"{_mdl}\.\w+", _key) for _mdl in lst_mdl)]
    # }

    lst_len = [max(
        [df_sheet[_key].last_valid_index() + 1 if 
        df_sheet[_key].last_valid_index() is not None else 0
                    for _key in [f"{_mdl}",
                                 f"{_mdl}.srl",
                                 f"{_mdl}.tag"]]
         )
        for _mdl in lst_mdl]

    lst_pile = []
    for _i, _mdl in enumerate(lst_mdl):
        lst_pile.append(PileDevice(dev_type, brand, _mdl))

        for _j in range(lst_len[_i]):
            lst_pile[_i].add(
                    df_sheet[f"{_mdl}"][_j],
                    df_sheet[f"{_mdl}.srl"][_j],
                    df_sheet[f"{_mdl}.tag"][_j]
                    )


    return lst_pile


#
# main
#

def main():
    file_xlsx = "cps.xlsx"

    df_laptop_hp        = pd.read_excel(file_xlsx, sheet_name='laptop_hp')
    df_laptop_surface   = pd.read_excel(file_xlsx, sheet_name='laptop_surface')
    df_mini_hp          = pd.read_excel(file_xlsx, sheet_name='mini_hp')

    lst_pile_1 = sheet_to_listPile(
            df_laptop_hp, Device_t.LAPTOP, Brand_t.HP)

    lst_pile_2 = sheet_to_listPile(
            df_laptop_surface, Device_t.LAPTOP, Brand_t.SURFACE)

    lst_pile_3 = sheet_to_listPile(
            df_mini_hp, Device_t.MINI, Brand_t.HP)


if __name__ == '__main__':
    main()
