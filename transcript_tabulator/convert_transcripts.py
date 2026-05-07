# ---
# jupyter:
#   jupytext:
#     cell_metadata_filter: title,-all
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
# # The Transcript Compiler
#
# This notebook will assist you in creating a computationally consistent view of
# transcripts of spoken materials that are recorded in Word documents (.docx). It will
# help you do things like:
#
# 1. Extract speaker codes.
# 2. Extract transcribed text.
# 3. Add further information about who's speaking
# 4. Add further information about the context of your transcripts (eg. if they're from
#    different parts of a multi phase study).
# 4. Identify and correct inconsistencies in formatting and speaker information.
# 6. Identify possible quality issues, such as missing speaker codes.
# 5. Connect your transcripts to their associated audio recordings.
# 6. Segment your transcripts into different components with labels - for example to
#    break up a semi-structured interview into topical segments for comparison.
#
# Apart from providing some structure for entering consistent information in and about
# your transcripts, this is primarily intended to help you get the most out of the
# transcripts you already have by allowing you to use them with computational tools for
# searching and filtering.
#
# Note that this tool assumes you already have transcripts in Word format (.docx) - if
# you are just starting out transcribing audio you may want to consider other tools and
# formats for transcribing that solve many of the issues we attempt to address here. We
# also assume that these Word transcripts you have will continue to be the version of
# record for your transcripts: the spreadsheet we compile here does not replace these
# but instead complements them.
#
# TODO: Working with Word files or PDF's that aren't transcripts - checkout our [document
# text extractor tool]().
#
# TODO: link to guidance on transcribing speech.
#
# TODO: this is way too text heavy and needs some explanation. Maybe I should start with
# the quick and dirty example first and see what happens? Or link to the explanation at
# the end?


# %% [markdown]
# # Word Transcript Conventions and How We Compile Them
#
# This notebook relies on a common set of conventions that we have observed in many Word
# based transcripts: because Word is a document preparation tool and not a data tool
# you may need to make adjustments to your Word files or change the configuration of
# this notebook. We aim to have clear and transparent failure modes, so even if things
# aren't quite right it can still be useful.
#
# We principally rely on two common conventions:
#
# 1. A new paragraph in Word is a new turn in the transcript.
# 2. The speaker code is separated from the text of what they said by colon-tab `:  `.
#
# Download the <a href="../example_transcript/transcript_format_example.docx"
# download="">annotated example transcript</a> to understand the expected format.
#
# What doesn't work:
#
# - If line-breaks are manually inserted to wrap text.
# - If you include non transcript material in your transcripts such as headers.
# - If your speaker codes aren't consistently marked with the same punctuation.
# - Information only present in styles like *bold* or _italics_ are ignored.
# - If your transcript lines are in tables.

# %% [markdown]
# # Recommended Workflow
#
# TODO: Make a diagram for this not just another list.
#
# 1. Start with whatever Word documents you have and upload them.
# 2. Download the output spreadsheet and examine the different sheets.
# 3. If major consistency problems are evident, identify and fix worst parts.
# 4. When major problems with transcripts are fixed, move onto entering metadata.


# %% [markdown]
#
# # 1. Upload your Transcripts
#
# Upload your Word documents (.docx format) with the button on the left.
#
# *Optionally*, upload the spreadsheet you created from an earlier run of the tool with the button on the right.
# The contents of that spreadsheet will be merged into the output of rerunning this tool so you don't have to enter any information again.
#

# %%
from uploader import upload_widget

layout, doc_widget, xl_widget = upload_widget()
display(layout)


# %% [markdown]
# # 2. Process Uploaded Transcripts
#
# This next step will process your uploaded transcripts to extract:
#
# - each turn of transcription
# - split the speaker code from the text of the transcript
# - 
#

# %%
from IPython.display import HTML
from processor import TidyTranscripts

transcripts = TidyTranscripts.from_ipywidgets(doc_widget, xl_widget)
output = transcripts.as_xlsx()
output.save('combined_transcripts.xlsx')

display(HTML("<a href=combined_transcripts.xlsx>Download Combined Transcripts</a>"))


# %% [markdown]
# # 3. Download and Review the Created File

# %%
