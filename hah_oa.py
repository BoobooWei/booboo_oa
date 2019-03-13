# -*- coding:utf8 -*-
"""
Created on:
@author: BoobooWei
Email: rgweiyaping@hotmail.com
Version: V.18.07.09.0
Description:
Help:
"""
from app import oa_booboo
import sys

reload(sys)
sys.setdefaultencoding('utf8')


if __name__ == "__main__":
    items = [
        {"input": "20190313091400007.xlsx",
         "output": "20190313091400007_end.xlsx"
         },
        {"input": "20190313091600016.xlsx",
         "output": "20190313091600016_ing.xlsx"
         },
    ]
    for params in items:
        oa_booboo.starup(**params)