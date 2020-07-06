import const, os
import re
from tika import parser


def extract_text_from_pdfs_recursively():
    """
    This function run over all files recursively in directory and extract it's test to text file.
    """
    if not os.path.isdir(const.TEXT_FILES_PATH):
        os.mkdir(const.TEXT_FILES_PATH)
        print('text file directory created in {0}'.format(const.TEXT_FILES_PATH))
    for root, dirs, files in os.walk(const.PDF_FILES_PATH):
        for file in files:
            path_to_pdf = os.path.join(root, file)
            if path_to_pdf.endswith('.pdf'):
                print("Processing " + path_to_pdf)
                pdf_contents = parser.from_file(path_to_pdf)
                path_to_txt = const.TEXT_FILES_PATH + '/' + file.replace('.pdf', '.txt')
                pdf_contents = pdf_contents['content']
                with open(path_to_txt, 'w+', encoding="utf-8") as txt_file:
                    print("Writing contents o " + path_to_txt)
                    txt_file.write(pdf_contents)
                with open(path_to_txt, 'r+', encoding='utf-8') as rewrite_file:
                    for line in rewrite_file.readlines():
                        if line != '\n':
                            rewrite_file.seek(0)
                            rewrite_file.write(line)


def regex_find_keys(string):
    """
    This function extract specific data from a string using regex.
    1. transaction id
    2. counterparty details
    3. credit amount
    :param string: given data
    :return: dictionary with wanted values
    """
    s = string
    pattern = "name and address: (.+?) Title"
    counterparty_details = re.search(pattern, s).group(1)
    pattern = "Transaction identifier: (.*?)  Credit"
    transaction_id = re.search(pattern, s).group(1)
    try:
        pattern = "Credit to account (.*?(PLN|EUR|GBP))"
        credit_amount = re.search(pattern, s).group(1)
    except AttributeError:
        return 'failed to parse the data from the text file - unknown type of details'
    return {'transaction_id': [transaction_id], 'counterparty_details': [counterparty_details],
            'credit_amount': [credit_amount]}


def extract_fields_from_textfile():
    """
    This function catch the bulk of data that relevance to work with and find or values with regex find keys function
    :return: list that includes dictionary with keys and values that we wanted to catch.
    """
    gold_keys_list = []
    for root, dirs, files in os.walk(const.TEXT_FILES_PATH):
        for file in files:
            path_to_file = os.path.join(root, file)
            [name, ext] = os.path.splitext(path_to_file)
            if ext == '.txt':
                file_path = name + '.txt'
                with open(file_path, 'r+', encoding='utf-8') as wf:
                    wf_one_line = (wf.read().replace('\n', ' '))
                    if 'CURRENT ACCOUNT' in wf_one_line:
                        account_id = re.search('CURRENT ACCOUNT (.*?) Amount', wf_one_line).group(1)
                        customer_details_option1 = re.findall('Counterparty name and address:.+?PLN', wf_one_line)
                        customer_details_option2 = re.findall('Counterparty name and address:.+?EUR', wf_one_line)
                        customer_details_option3 = re.findall('Counterparty name and address:.+?GBP', wf_one_line)
                        customer_details = customer_details_option1 + customer_details_option2 + customer_details_option3
                        for i in customer_details:
                            if len(i) > 450:
                                customer_details_option1.remove(i)
                        matches = []
                        for i in customer_details:
                            if i is not None and 'Credit to account' in i:
                                matches.append(i)
                for i in matches:
                    gold_keys = regex_find_keys(string=i)
                    gold_keys['account_id'] = [account_id]
                    gold_keys_list.append(gold_keys)
    return gold_keys_list


def data_extractor():
    extract_text_from_pdfs_recursively()
    statements = extract_fields_from_textfile()
    return statements
