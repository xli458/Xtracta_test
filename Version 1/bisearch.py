# -*- coding: utf-8 -*-
"""
Created on Wed Nov 16 12:55:06 2016

This program defines a function for binary search

@author: xli458
"""

import math

def binary_search(word_tmp, Sup_split_name):
    # if the first word is found, return the index of the row number, or else return -1
    indx = -1

    start_pt = 0
    end_pt = len(Sup_split_name)
    
    while(end_pt > start_pt+1):
        mid = math.ceil((end_pt + start_pt) / 2)
        if Sup_split_name[mid][0] == word_tmp:
            indx = mid
            break
        elif Sup_split_name[mid][0] > word_tmp:
            end_pt = mid
        elif Sup_split_name[mid][0] < word_tmp: 
            start_pt = mid
    
    if end_pt == start_pt+1:
        if Sup_split_name[end_pt][0] == word_tmp:
            indx = end_pt
        if Sup_split_name[start_pt][0] == word_tmp:    
            indx = start_pt
    return indx