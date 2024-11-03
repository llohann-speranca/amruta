from talk_reader import TalkReader
from glob import glob

for path in glob('data/*.docx'):
    talk_reader = TalkReader(path, '/Users/llohann/Python/amruta/data/metadata/metadata.csv')
    talk_reader.dump_metadata()
