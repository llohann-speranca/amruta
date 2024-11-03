import pandas as pd

from talk_reader import TalkReader, get_metadata_file


def test_metadata():
    df = get_metadata_file()
    print(df.columns.to_list())


def test_talk_reader():
    test_path = "/Users/llohann/Python/amruta/data/test/P04 - Prt8 - 1982-04-29 till 1982-05-14 _ WORKFILE.docx"
    metadata_file_path = '/Users/llohann/Python/amruta/data/test/metadata.csv'
    talk_reader = TalkReader(test_path, metadata_file_path=metadata_file_path)
    print(talk_reader.talks)
    talk_reader.dump_metadata()
    metadata = pd.read_csv(metadata_file_path)
    print(metadata)

# def test_open_word_document():
#     test_path = "/Users/llohann/Python/amruta/data/test/P04 - Prt8 - 1982-04-29 till 1982-05-14 _ WORKFILE.docx"
#     doc = open_word_document(test_path)
#     print(doc)


if __name__ == '__main__':
    # test_metadata()
    # test_open_word_document()
    test_talk_reader()

