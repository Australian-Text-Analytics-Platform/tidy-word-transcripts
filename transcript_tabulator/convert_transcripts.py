# ---
# jupyter:
#   jupytext:
#     cell_metadata_filter: -all
#     formats: py:percent,ipynb
#     text_representation:
#       extension: .py
#       format_name: percent
#       format_version: '1.3'
#       jupytext_version: 1.19.1
#   kernelspec:
#     display_name: Python 3 (ipykernel)
#     language: python
#     name: python3
# ---

# %% [markdown]
# # Transcript Tabulator
#
# ## To support computational analysis of qualitative interview transcripts.
#


# %% [markdown]
# # Getting Started
#
# Download the <a href="../example_transcript/transcript_format_example.docx"
# download="">annotated example transcript</a> to understand the expected format.
#

# %% [markdown]
#
# # Step 1 - Upload your Transcripts
#
# Upload your transcripts to the uploaded_transcripts folder - you can drag and drop.
#


# %% [markdown]
#
# # Step 2 - Run This Cell
#
#

# %%
import glob
from pathlib import Path

from docx import Document

# Find all the uploaded .docx files and process them one by one

from docx import Document

transcripts_folder = Path("uploaded_transcripts")

transcript_rows = []

# This looks for colon-tab (the \t is a tab character) as the speaker-text separator.
# If you use a different convention you can try changing this.
speaker_text_separator = ":\t"

for transcript_filepath in glob.glob(str(transcripts_folder / "**.docx")):

    filename = Path(transcript_filepath).relative_to(transcripts_folder)

    print("Processing", filename)

    with open(transcript_filepath, mode="rb") as transcript_file:

        # Load the word document
        doc = Document(transcript_file)

        para_no = 1
        segment_no = 1

        # We're going to treat each paragraph in the file as a single speaker's turn.
        # This is why headers need to be wrapped in a table or similar, otherwise we
        # end up the header included as part of the transcript.
        for paragraph in doc.paragraphs:

            # The text of this paragraph
            para_text = paragraph.text

            # Handle blank lines.
            # If they're at the start of the file just ignore them, otherwise use them
            # as segment boundaries within the transcript.
            if not para_text.strip():
                if para_no == 1:
                    continue

                segment_no += 1
                continue

            # Identify speakers by looking for the first colon character then a tab.
            # If colon-tab is not matched, no speaker will be assigned to this turn.
            speaker_code, sep, text = para_text.partition(speaker_text_separator)

            # If nothing in the line matches the speaker/text separator, keep the whole
            # line, and don't record a speaker
            if sep == "":
                text = para_text
                speaker_code = None

            transcript_rows.append(
                (str(filename), para_no, speaker_code, text, segment_no)
            )

            para_no += 1

print(transcript_rows)

# %% [markdown]
# # Finding and fixing issues
#
# 1. Identify and fix speaker code issues first.
#   - One off mispellings: fix in the source file and re-upload that file.
#   - Systematic inconsistencies: one transcript may have Interviewer, the other
#   - may have interviewer (lowercase).
# 2. Then add speaker information as relevant. This will vary depending on your

# %%
from collections import Counter
from itertools import groupby


def extract_transcript_info(transcript_rows):
    """
    Extract a summary of the extracted transcript rows for each transcript.

    """

    by_transcript = groupby(transcript_rows, key=lambda x: x[0])

    for transcript, rows in by_transcript:
        all_rows = list(rows)
        print(transcript, all_rows)


def extract_segment_info(transcript_rows):
    """
    Extract a summary of the extracted segments for each transcript.

    """
    by_segment = groupby(transcript_rows, key=lambda x: (x[0], x[-1]))

    for segment, rows in by_segment:
        all_rows = list(rows)
        print(segment, all_rows)


def extract_speaker_code_info(transcript_rows):
    """
    Extract a summary of the extracted speaker codes for each transcript.

    """
    key = lambda x: (x[0], x[2])
    by_speaker = groupby(sorted(transcript_rows, key=key), key=key)

    for speaker, rows in by_speaker:
        all_rows = list(rows)
        print(speaker, all_rows)


extract_transcript_info(transcript_rows)
extract_segment_info(transcript_rows)
extract_speaker_code_info(transcript_rows)


for filename, para_no, speaker_code, text, segment_no in transcript_rows:
    pass


# %%
# Write this all out to the XLSX file.


# %% [markdown]
# # Working with conversation data


# %% [markdown]
# # Requirements and assumptions


# %%
# Create a folder to hold the conversations
from pathlib import Path

conversations_path = Path("conversations")
conversations_path.mkdir(exist_ok=True)

# %% [markdown]
# # Upload your transcripts
#
# You can either click the button below, or upload your files directly in the
# conversations folder from the lefthand panel.

# %%
import ipywidgets

uploaded_via_jupyter = False

uploader = ipywidgets.FileUpload(accept=".docx", multiple=True)
display(uploader)


# %% [markdown]
# Process the uploaded documents.
#
# Uploaded files are saved in the conversations folder. All .docx files in that folder
# will be included in the processing below.

# %%
from pathlib import Path

for uploaded_file in uploader.value:
    with open(conversations_path / uploaded_file.name, "wb") as f:
        f.write(uploaded_file.content)


# %% [markdown]
# # Pre-process conversations
#
# This step will aim to extract each turn of each uploaded conversation, and separate
# the speaker from the turn content. This step will only be as consistent as your
# transcripts are.
#
# This will completely delete and recreate the state of processed conversations - if
# you want to upload new or edited files, make the contents of the conversations folder
# match what you want, and re-run this cell.

# %%
# Start by setting up a small database to hold the processed information.

import sqlite3

convo_db_path = conversations_path / "conversations.db"

convo_db = sqlite3.connect(convo_db_path, isolation_level=None)

convo_db.executescript("""
    DROP table if exists turn;
    CREATE table turn (
        turn_id integer primary key,
        source_file,
        turn_no,
        speaker,
        turn,
        unique(source_file, turn_no)
    )

    """)

# Then we'll identify and load conversation turns from each transcript.
import glob

from docx import Document

convo_db.execute("begin")

for transcript_filepath in glob.glob(str(conversations_path / "**.docx")):
    filename = Path(transcript_filepath).relative_to(conversations_path)
    print("Processing", transcript_filepath)

    with open(transcript_filepath, "rb") as transcript_file:

        # Load the word document
        doc = Document(transcript_file)

        # We're going to treat each paragraph in the word file as a turn
        for turn_no, paragraph in enumerate(doc.paragraphs):
            para_text = paragraph.text

            # Identify speakers by looking for the colon character then a tab.
            # If colon-tab is not matched, no speaker will be assigned to this turn.
            speaker_split = para_text.split(":\t")

            if len(speaker_split) == 2:
                speaker, text = speaker_split
            else:
                speaker, text = None, para_text

            convo_db.execute(
                "INSERT into turn values (?, ?, ?, ?, ?)",
                (None, transcript_filepath, turn_no, speaker, text),
            )

convo_db.execute("commit")

# %% [markdown]
#
# The previous step created a little database holding the extracted turns from all
# uploaded files. Now we're going to index this dataset in a way that is aware of the
# within and across-turn structure of conversations.


# %% [markdown]
#
# First lets define two functions for breaking down turns into word-like units
# (tokenisation). We'll create two tokenisers - the first one normalises case by
# lowercasing all of the text, then finds and splits the turn up at characters
# indicating word boundaries('\b'), or whitespace characters like space, newlines, and
# tabs.
#
# The second, display_tokenise, breaks on the same places, but does not lowercase, or
# remove spaces, so we can recreate and highlight search result matches/concordances
# of search terms.
