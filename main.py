#!/usr/bin/python

import xlwings as xw
import openpyxl
import datetime
import re
import sys
import time
from people import People 
from parser_file import load_main_file
from parser_file import update_people
from parser_file import find_missing_contractor 
from parser_file import find_missing_racker
from parser_file import compare_with_main_and_Contractor_not_regitst_list 
from parser_file import generate_report 
from parser_file import generate_name_list 
from parser_file import system_call 
from parser_file import create_dic_for_name_mapping


if __name__ == '__main__':

    main_file = sys.argv[1]
    #racker_file = sys.argv[2]
    #contractor_file = sys.argv[3]
    #extra_file = sys.argv[4]

    main_dic = {}

    #load data from Introduction_to_PCI_DSS_3.1_for_Developers
    load_main_file(main_file, main_dic)


    #generate racker and contractor name list
    generate_name_list("rackspace", main_dic)
    generate_name_list("contractor", main_dic)

    #call henry code generate list with leader
    system_call("rackspace.xlsx")
    system_call("contractor.xlsx")

    name_mapping_dic = create_dic_for_name_mapping("name_mapping.xlsx")

    #update for racker 
    name_not_in_list = update_people("rackspace_with_leaders.xlsx", main_dic, name_mapping_dic)
    #update for contractor
    name_not_in_list = update_people("contractor_with_leaders.xlsx", main_dic, name_mapping_dic)

    #generate report named "result.xlsx"
    generate_report(main_dic)

    #find_missing_contractor(extra_file, main_dic)

