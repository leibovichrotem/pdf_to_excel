import const, shutil
import data_extractor, excel_writer


def main():
    """
    This function start the project using the project functions.
    """
    statements = data_extractor.data_extractor()
    for statement in statements:
        sheet_name = statement['account_id'][0][-10:]
        excel_writer.manager(sheet_name,statement)
    shutil.rmtree(const.TEXT_FILES_PATH)


if __name__ == '__main__':
    main()
