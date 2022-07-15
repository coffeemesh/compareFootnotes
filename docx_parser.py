#! /usr/bin/env python
"""
A python script to extract footers from .docx documents and compare them in a .csv.
Uses the docx2python library. You must define your own paths.
"""
from __future__ import print_function

import os

import numpy as np
import pandas as pd
from docx2python import docx2python

# Change this with a fitting path
DOCX_FILES_DIRECTORY = "./path/"

# Lists all the files in for the given path
files = os.listdir(DOCX_FILES_DIRECTORY)
# Filters the previous result by keeping only .docx files
docx_files = [docxFile for docxFile in files if docxFile.endswith(".docx")]


def get_base_footnotes_as_list():
    """ Returns all the footnotes of a specific .docx document.

    Returns:
        List of footnotes, each footnote being on a single row.
    """
    base_footnotes_array = []
    base_document = docx2python("./path/to/BaseDocument.docx") # Change this with a fitting path
    for x in np.array(base_document.footnotes, dtype=object):
        for y in x:
            for z in y:
                base_footnotes_array.append(z)

    base_df = pd.DataFrame(base_footnotes_array, columns=[
        'Footnote1', 'Footnote2'])
    base_df['Footnote'] = base_df['Footnote1'].map(
        str) + base_df['Footnote2'].map(str)
    base_footnotes_list = base_df['Footnote'].tolist()
    base_footnotes_list.sort()
    return base_footnotes_list


def get_variant_footnotes_as_list(docx_file):
    """ Returns all the Footnotes for the given Document

    Args:
        docx_file (File): A Word Document ending with .docx

    Returns:
        List of footnotes, each footnote being on a single row.
    """
    variant_footnotes_array = []
    file_path = os.path.join(DOCX_FILES_DIRECTORY, docx_file)
    variant_document = docx2python(file_path)
    for x in np.array(variant_document.footnotes, dtype=object):
        for y in x:
            for z in y:
                variant_footnotes_array.append(z)
        variant_df = pd.DataFrame(variant_footnotes_array, columns=[
            'Footnote1', 'Footnote2'])
        variant_df['Footnote'] = variant_df['Footnote1'].map(
            str) + variant_df['Footnote2'].map(str)
        variant_footnotes_list = variant_df['Footnote'].tolist()
        variant_footnotes_list.sort()
        return variant_footnotes_list


def find_missing_footnotes(variant_footnotes_list, original_footnotes_list):
    missing_footnotes = set(variant_footnotes_list).difference(
        original_footnotes_list)
    return missing_footnotes


def find_additional_footnotes(original_footnotes_list, variant_footnotes_list):
    additional_footnotes = set(
        original_footnotes_list).difference(variant_footnotes_list)
    return additional_footnotes


def create_csv_for_missing_footnotes():
    """ Creates a .csv file listing the missing footnotes on the base document

    Returns:
        A message informing about the operation.
    """
    for docx_file in docx_files:
        variant_footnotes_list = get_variant_footnotes_as_list(docx_file)
        original_footnotes_list = get_base_footnotes_as_list()

        missing_footnotes = find_missing_footnotes(
            variant_footnotes_list, original_footnotes_list)

        missing_footnotes_df = pd.DataFrame(missing_footnotes)
        # You might need to create those folders
        missing_footnotes_df.to_csv("./CSV/MissingFootnotes/Missing_Footnotes_" +
                                    docx_file + ".csv", sep=",")
    return print("CSVs finished")


def create_csv_for_additional_footnotes():
    """ Creates a .csv file listing the additional footnotes on the base document

    Returns:
        A message informing about the operation.
    """
    for docx_file in docx_files:
        variant_footnotes_list = get_variant_footnotes_as_list(docx_file)
        original_footnotes_list = get_base_footnotes_as_list()

        additional_footnotes = find_additional_footnotes(
            original_footnotes_list, variant_footnotes_list)

        additional_footnotes_df = pd.DataFrame(additional_footnotes)
        # You might need to create those folders
        additional_footnotes_df.to_csv(
            "./CSV/AdditionalFootnotes/Additional_Footnotes_" + docx_file + ".csv",
            sep=",")
    return print("CSVs finished")


def create_footnotes_comparison_csv(docx_file):
    """ Creates a .csv file where footnotes from both files are displayed and the missing ones are marked

    Returns:
        A message informing about the operation.
    """
    variant_footnotes_list = get_variant_footnotes_as_list(docx_file)
    original_footnotes_list = get_base_footnotes_as_list()

    missing_footnotes_list = []
    for x in original_footnotes_list:
        for y in variant_footnotes_list:
            if x == y:
                missing_footnotes_list.append('MATCH')
            else:
                missing_footnotes_list.append('MISSING')

    comparison_df = pd.DataFrame({'Original Footnotes': pd.Series(original_footnotes_list),
                                  'Variant Footnotes': pd.Series(variant_footnotes_list),
                                  'Missing Footer?': pd.Series(missing_footnotes_list)})
    # You might need to create those folders
    comparison_csv = comparison_df.to_csv(
        "./CSV/ComparisonCSV/Comparison_CSV" + docx_file + ".csv", sep=",")
    return comparison_csv


def compare_base_document_with_variant():
    for docx_file in docx_files:
        create_footnotes_comparison_csv(docx_file)
    return print("CSVs finished")
