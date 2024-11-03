import pandas as pd
from docx import Document
import os
from csv import DictWriter
from datetime import datetime
import re
from tqdm import tqdm
import traceback


def get_metadata_file(path="./data/metadata/2024-02-xx Maha talks list .xlsx"):
    metadata_df = pd.read_excel(path, index_col=0)
    return metadata_df


def parse_metadata_from_talk(talk: list[str]):
    # [Title, Type of talk, Location, Talks status, Number of words]
    # Example talk:
    # 04 May 1982
    # Shri Mataji Teaches Yogis to Sing Bhajans in Ashram in Le Raincy
    # Talk to Sahaja Yogis on Eve of Sahasrara Puja
    # Le Raincy, France
    # Talk Language: English | Transcript: Draft
    while len(talk[0]) < 2:
        talk.pop(0)
    metadata_list = talk[:5]
    number_of_words = sum([len(paragraph.split(" ")) for paragraph in talk])
    metadata = {
        "Date": re.sub(
            "[^A-Za-z0-9 ]",
            "",
            metadata_list[0]
            .strip()
            .replace("th", "")
            .replace("  ", " ")
            .replace("1001", ""),
        ),
        "Title": metadata_list[1],
        "Type of talk": metadata_list[2],
        "Location": metadata_list[3],
        "Talk Language": metadata_list[-1]
        .split(" | ")[0]
        .removeprefix("Talk Language: "),
        "Talks status": metadata_list[-1].split(" | ")[-1].removeprefix("Transcript: "),
        "Number of words": number_of_words,
    }
    return metadata


class TalkReader:
    def __init__(
        self,
        list_of_talks_path: str,
        metadata_file_path: str,
        save_folder: str = "data/output",
    ):
        # print("Initiated")
        self.document = Document(list_of_talks_path)
        self.metadata = []
        self.metadata_file_path = metadata_file_path
        self.talks = []
        self.save_folder = save_folder
        self.metadata_fields = [
            "Date",
            "Title",
            "Type of talk",
            "Location",
            "Talks status",
            "Number of words",
            "Talk Language",
        ]
        self.parsing_that_went_wrong = []
        self.paragraphs_do_not_starting_with_hashtag = []
        if not os.path.exists(metadata_file_path):
            with open(self.metadata_file_path, "w") as metadata_file:
                csv_writer = DictWriter(metadata_file, fieldnames=self.metadata_fields)
                csv_writer.writeheader()

        # self.talks = self.process_document()

    def get_all_text(self):
        document_text = [p.text for p in self.document.paragraphs]
        document_text = "\n\n".join(document_text)
        return document_text
        

    def process_document(self):
        # document = [p.text for p in document.paragraphs]
        talk = []
        if os.getenv("DEBUG") == "true":
            paragraphs = self.document.paragraphs[:1000]
        else:
            paragraphs = self.document.paragraphs

        for paragraph in tqdm(paragraphs):
            text: str = paragraph.text
            if text.startswith("##"):
                self.add_talk_to_database(talk)
                text = text.removeprefix("##").removeprefix("\u2003").strip()
                if text:
                    talk = [text]
            # elif '##' in text:
            #     self.paragraphs_do_not_starting_with_hashtag.append(text)
            else:
                talk.append(text)
        self.add_talk_to_database(talk)
        return self.metadata

    def add_talk_to_database(self, talk):
        try:
            if len(talk) < 5:
                self.parsing_that_went_wrong.append(
                    {
                        "talk": talk[:10],
                        "error": "len(talk) < 5",
                        "traceback": "",
                    }
                )

            else:
                metadata = self.parse_metadata_from_talk(talk)
                self.metadata.append(metadata)
                self.dump_talk(talk=talk, metadata=metadata)
                # self.talks.append(talk)
        except Exception as e:
            self.parsing_that_went_wrong.append(
                {
                    "talk": talk[:10],
                    "error": str(e),
                    "traceback": traceback.format_exc(),
                }
            )

    def dump_talk(self, talk: list[str], metadata: dict, save_folder=None):
        if save_folder is None:
            save_folder = self.save_folder
        document = Document()
        # style = document.styles['Normal']
        # font = style.font
        # font.name = 'Calibri'
        # font.size = Pt(11)

        for paragraph in talk:
            _ = document.add_paragraph(paragraph)
            # para.style = document.styles['Normal']
        try:
            date_iso = datetime.strptime(metadata["Date"], "%d %B %Y").strftime(
                "%Y-%m-%d"
            )
        except ValueError:
            date_iso = datetime.strptime(metadata["Date"], "%B %d %Y").strftime(
                "%Y-%m-%d"
            )
        save_file_name = f"{date_iso} {metadata['Title']}, {metadata['Location']}.docx"
        document.save(os.path.join(save_folder, save_file_name))

    @staticmethod
    def parse_metadata_from_talk(talk: list[str]):
        # [Title, Type of talk, Location, Talks status, Number of words]
        # Example talk:
        # 04 May 1982
        # Shri Mataji Teaches Yogis to Sing Bhajans in Ashram in Le Raincy
        # Talk to Sahaja Yogis on Eve of Sahasrara Puja
        # Le Raincy, France
        # Talk Language: English | Transcript: Draft
        while len(talk[0]) < 2:
            talk.pop(0)
        metadata_list = talk[:5]
        number_of_words = sum([len(paragraph.split(" ")) for paragraph in talk])
        metadata = {
            "Date": re.sub(
                "[^A-Za-z0-9 ]",
                "",
                metadata_list[0]
                .strip()
                .replace("th", "")
                .replace("  ", " ")
                .replace("1001", ""),
            ),
            "Title": metadata_list[1],
            "Type of talk": metadata_list[2],
            "Location": metadata_list[3],
            "Talk Language": metadata_list[-1]
            .split(" | ")[0]
            .removeprefix("Talk Language: "),
            "Talks status": metadata_list[-1]
            .split(" | ")[-1]
            .removeprefix("Transcript: "),
            "Number of words": number_of_words,
        }
        return metadata

    def dump_metadata(self):
        with open(self.metadata_file_path, "a") as metadata_file:
            csv_writer = DictWriter(metadata_file, fieldnames=self.metadata_fields)
            csv_writer.writerows(self.metadata)


# def break_document
