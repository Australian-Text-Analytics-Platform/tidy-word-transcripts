"""
A widget to allow users to upload their files.

Takes both word files and an optional metadata spreadsheet created by an earlier run.

Display the list of uploaded files.

"""

import ipywidgets


def upload_widget():
    button_layout = ipywidgets.Layout(width="40%", height="3lh")

    def list_uploaded_files(change):

        out.clear_output()
        with out:
            display(f"{len(change['new'])} files uploaded:")
            for file in change["new"]:
                display(file["name"])

    out = ipywidgets.Output()

    doc_uploader = ipywidgets.FileUpload(
        accept=".docx",
        multiple=True,
        description="Upload your Word docs:",
        layout=button_layout,
    )
    xl_uploader = ipywidgets.FileUpload(
        accept=".xlsx",
        multiple=False,
        description="(Optional) Upload your matching spreadsheet:",
        layout=button_layout,
    )

    doc_uploader.observe(list_uploaded_files, names="value")
    xl_uploader.observe(list_uploaded_files, names="value")

    layout = ipywidgets.VBox([ipywidgets.HBox([doc_uploader, xl_uploader]), out])

    return layout


# def handle_upload(change):
#     out.clear_output()
#     with out:
#         if upload.value:
#             uploaded_file = upload.value[0]
#             content = uploaded_file["content"]
#             df = pd.read_csv(pd.io.common.BytesIO(content))
#             display(df)


# display(widgets.VBox([upload, out]))
