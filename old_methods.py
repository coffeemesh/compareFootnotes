def extract_footnotes(docx_file):
    """Extracts all footnotes from a given .docx file.

    Args:
        docx_file (str): Path to the .docx file.

    Returns:
        List of footnotes, each footnote being on a single row.
    """
    footnotes_array = []
    document = docx2python(docx_file)
    for x in np.array(document.footnotes, dtype=object):
        for y in x:
            for z in y:
                footnotes_array.append(z)
    return footnotes_array


def get_footnotes_as_list(docx_file):
    """Returns all the footnotes of a specific .docx document as a sorted list.

    Args:
        docx_file (str): Path to the .docx file.

    Returns:
        List of footnotes, each footnote being on a single row.
    """
    footnotes_array = extract_footnotes(docx_file)
    footnotes_df = pd.DataFrame(footnotes_array, columns=['Footnote'])
    footnotes_list = footnotes_df['Footnote'].tolist()
    footnotes_list.sort()
    return footnotes_list


def get_base_footnotes_as_list():
    """Returns all the footnotes of the base .docx document as a sorted list.

    Returns:
        List of footnotes, each footnote being on a single row.
    """
    return get_footnotes_as_list(BASE_DOCUMENT)


def get_variant_footnotes_as_list(docx_file):
    """Returns all the footnotes of a variant .docx document as a sorted list.

    Args:
        docx_file (str): Name of the .docx file.

    Returns:
        List of footnotes, each footnote being on a single row.
    """
    variant_docx_file = os.path.join(DOCX_FILES_DIRECTORY, docx_file)
    return get_footnotes_as_list(variant_docx_file)


def find_missing_footnotes(variant_footnotes_list, original_footnotes_list):
    """Finds footnotes that are in the variant list but not in the original list.

    Args:
        variant_footnotes_list (list): List of footnotes from the variant document.
        original_footnotes_list (list): List of footnotes from the original document.

    Returns:
        Set of missing footnotes.
    """
    return set(variant_footnotes_list).difference(original_footnotes_list)


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

