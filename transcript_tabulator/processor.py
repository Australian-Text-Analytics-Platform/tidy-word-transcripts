"""
Processes transcripts in Word docs to a tidy form of transcripts in xlsx.

"""

from collections import Counter
import dataclasses as dc
import re
import typing

from openpyxl import Workbook, load_workbook

from docx import Document


def extract_turns(transcript_doc: Document, split_speaker_on: re.Pattern):
    """
    Extract the turns from the provided docx Document.

    split_speaker_on is a regular expression Pattern to split the turn into speaker and
    remaining components. Only the first match is split.

    If no match is found to split on, the speaker will be marked as absent and the
    full text of the turn used as the spoken text.

    """

    segment_no = 1
    last_para_has_content = False

    # We're going to treat each paragraph in the word file as a turn
    for turn_no, paragraph in enumerate(transcript_doc.paragraphs):

        # Note this is where we discard styling information
        para_text = paragraph.text

        # Check for a segment break in the transcript.
        # Make sure to merge multiple empty paragraphs into one segment break.
        if not para_text and last_para_has_content:
            segment_no += 1
            continue

        if para_text:
            last_para_has_content = True

        # Use the provided splitter and the first place it matches to separate the
        # text into speaker and turn.
        potential_split = split_speaker_on.split(para_text, 1)

        if len(potential_split) == 2:
            speaker_code, text = potential_split
        else:
            speaker_code, text = "", para_text

        # turn_no + 1 so we count from 1 like most people expect.
        yield [segment_no, turn_no + 1, speaker_code, text]


@dc.dataclass
class Turn:
    source_file: str
    segment_no: int
    turn_no: int
    speaker_code: str
    transcription: str


@dc.dataclass
class Segment:
    source_file: str
    segment_no: int
    name: typing.Optional[str] = ""
    turn_count: typing.Optional[int] = 0


@dc.dataclass
class SpeakerCode:
    source_file: str
    speaker_code: str
    turn_count: typing.Optional[int] = 0


@dc.dataclass
class TidyTranscripts:
    """

    Limitations:

    - Spreadsheet styling is not retained.
    - Word doc styling is not retained.

    """

    transcripts: dict[str, Document]
    split_speaker_on: re.Pattern | str = ":\t"
    spreadsheet: typing.Optional[Workbook] = None

    turns: list[Turn] = dc.field(init=False, default_factory=list)
    segments: dict[tuple[str, int], Segment] = dc.field(
        init=False, default_factory=Counter
    )
    speaker_codes: dict[tuple[str, str], SpeakerCode] = dc.field(
        init=False, default_factory=Counter
    )
    transcript_turns: dict[str, int] = dc.field(init=False, default_factory=Counter)

    @classmethod
    def from_filepaths(
        cls, transcript_paths, spreadsheet_path=None, split_speaker_on=":\t"
    ):
        """Load transcripts from files on disk given by paths."""

        transcripts = {}

        for path in transcript_paths:
            with open(path, "rb") as transcript:
                transcripts[path] = Document(transcript)

        if spreadsheet_path:
            spreadsheet = load_workbook(spreadsheet_path)
        else:
            spreadsheet = Workbook()

        return cls(
            transcripts=transcripts,
            spreadsheet=spreadsheet,
            split_speaker_on=split_speaker_on,
        )

    @classmethod
    def from_ipywidgets(cls, doc_widget, spreadsheet_widget, split_speaker_on=":\t"):
        """Tidy transcripts from the upload widgets"""

        transcripts = {}

        for uploaded in doc_widget.value:
            transcripts[uploaded.name] = Document(uploaded.content)

        if spreadsheet_widget.value:
            spreadsheet = load_workbook(spreadsheet_widget.value[0].content)
        else:
            spreadsheet = Workbook()

        return cls(
            transcripts=transcripts,
            spreadsheet=spreadsheet,
            split_speaker_on=split_speaker_on,
        )

    def __post_init__(self):

        if isinstance(self.split_speaker_on, str):
            self.split_speaker_on = re.compile(self.split_speaker_on)

        self.turns = []

        self.speaker_codes = Counter()
        self.segments = Counter()
        self.transcript_turns = Counter()

        # Extract and populate all the data we need.
        for source_file, doc in self.transcripts.items():
            for turn_details in extract_turns(doc, self.split_speaker_on):

                turn = Turn(source_file, *turn_details)
                self.turns.append(turn)

                # Extract keys for segments and speaker_codes as we go.
                self.speaker_codes[(turn.source_file, turn.speaker_code)] += 1
                self.segments[(turn.source_file, turn.speaker_code)] += 1
                self.transcript_turns[turn.source_file] += 1

    def process_files(self):
        pass

    def as_xlsx(self):
        pass

    def tidy(self, include_files, include_segments=None, align_overlap=True):
        pass


if __name__ == "__main__":

    tidied = TidyTranscripts.from_filepaths(
        ["../example_transcript/transcript_format_example.docx"],
    )

    print(tidied.turns)
