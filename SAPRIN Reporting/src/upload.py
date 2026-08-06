"""
upload.py

Uploads consolidated reports back to SharePoint.

Author: Freedom Mukomana
"""

import logging
from pathlib import Path
from datetime import datetime

from office365.sharepoint.client_context import ClientContext

logger = logging.getLogger(__name__)


class SharePointUploader:

    def __init__(self, context: ClientContext):

        self.ctx = context

    ###########################################################################
    # Create Folder
    ###########################################################################

    def ensure_folder(self, folder_url):

        try:

            folder = self.ctx.web.get_folder_by_server_relative_url(
                folder_url
            )

            self.ctx.load(folder)

            self.ctx.execute_query()

            return folder

        except Exception:

            logger.info(f"Creating folder: {folder_url}")

            parent = "/".join(folder_url.split("/")[:-1])

            name = folder_url.split("/")[-1]

            parent_folder = self.ctx.web.get_folder_by_server_relative_url(
                parent
            )

            new_folder = parent_folder.folders.add(name)

            self.ctx.execute_query()

            return new_folder

    ###########################################################################
    # Upload One File
    ###########################################################################

    def upload_file(
        self,
        local_file,
        destination_folder
    ):

        local_file = Path(local_file)

        if not local_file.exists():

            raise FileNotFoundError(local_file)

        folder = self.ensure_folder(destination_folder)

        with open(local_file, "rb") as file:

            folder.upload_file(
                local_file.name,
                file.read()
            ).execute_query()

        logger.info(

            f"Uploaded {local_file.name}"

        )

    ###########################################################################
    # Upload Folder
    ###########################################################################

    def upload_folder(
        self,
        local_folder,
        destination_folder
    ):

        local_folder = Path(local_folder)

        uploaded = []

        for file in local_folder.iterdir():

            if not file.is_file():

                continue

            self.upload_file(

                file,

                destination_folder

            )

            uploaded.append(file.name)

        return uploaded

    ###########################################################################
    # Archive
    ###########################################################################

    def archive_report(
        self,
        file,
        archive_root,
        report_date
    ):

        year = report_date.strftime("%Y")

        week = report_date.strftime("Week_%U")

        folder = f"{archive_root}/{year}/{week}"

        self.ensure_folder(folder)

        self.upload_file(file, folder)

    ###########################################################################
    # Upload Weekly Outputs
    ###########################################################################

    def upload_outputs(
        self,
        output_folder,
        sharepoint_folder
    ):

        output_folder = Path(output_folder)

        uploaded = []

        for filename in [

            "Weekly_Consolidated_Report.xlsx",

            "Weekly_Consolidated_Report.csv",

            "Validation_Report.xlsx",

            "Historical_Weekly_Data.xlsx",

            "Failed_Reports.xlsx",

            "Warnings.xlsx"

        ]:

            file = output_folder / filename

            if not file.exists():

                continue

            self.upload_file(

                file,

                sharepoint_folder

            )

            uploaded.append(filename)

        return uploaded


###########################################################################
# Convenience Function
###########################################################################

def upload_results(
    context,
    output_folder,
    sharepoint_folder
):

    uploader = SharePointUploader(context)

    return uploader.upload_outputs(

        output_folder,

        sharepoint_folder

    )


###########################################################################
# Test
###########################################################################

if __name__ == "__main__":

    from auth import get_context

    ctx = get_context(

        "https://tenant.sharepoint.com/sites/Management"

    )

    uploaded = upload_results(

        ctx,

        "output",

        "Shared Documents/Consolidated Reports"

    )

    print(uploaded)