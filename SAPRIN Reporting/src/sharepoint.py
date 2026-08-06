"""
sharepoint.py

SharePoint helper functions for downloading and uploading files.

Author: Freedom Mukomana
"""

import logging
from pathlib import Path
from typing import List

from office365.sharepoint.client_context import ClientContext

logger = logging.getLogger(__name__)


class SharePointManager:
    """
    Handles SharePoint file operations.
    """

    def __init__(self, context: ClientContext):
        self.ctx = context

    ###########################################################################
    # LIST FILES
    ###########################################################################

    def list_files(self, folder_url: str):
        """
        List all files in a SharePoint folder.

        Example:
            Shared Documents/Weekly Reports
        """

        folder = self.ctx.web.get_folder_by_server_relative_url(folder_url)

        files = folder.files

        self.ctx.load(files)

        self.ctx.execute_query()

        return list(files)

    ###########################################################################
    # DOWNLOAD FILE
    ###########################################################################

    def download_file(self, server_relative_url: str, destination: Path):

        destination.parent.mkdir(
            parents=True,
            exist_ok=True
        )

        sp_file = self.ctx.web.get_file_by_server_relative_url(
            server_relative_url
        )

        with open(destination, "wb") as local_file:

            sp_file.download(local_file).execute_query()

        logger.info(f"Downloaded {destination.name}")

        return destination

    ###########################################################################
    # DOWNLOAD FOLDER
    ###########################################################################

    def download_folder(
        self,
        folder_url: str,
        destination_folder: Path
    ) -> List[Path]:

        downloaded = []

        files = self.list_files(folder_url)

        destination_folder.mkdir(
            parents=True,
            exist_ok=True
        )

        for file in files:

            filename = file.properties["Name"]

            if filename.startswith("~$"):
                continue

            destination = destination_folder / filename

            self.download_file(
                file.serverRelativeUrl,
                destination
            )

            downloaded.append(destination)

        logger.info(f"{len(downloaded)} files downloaded.")

        return downloaded

    ###########################################################################
    # UPLOAD FILE
    ###########################################################################

    def upload_file(
        self,
        local_file: Path,
        folder_url: str
    ):

        folder = self.ctx.web.get_folder_by_server_relative_url(
            folder_url
        )

        with open(local_file, "rb") as f:

            folder.upload_file(
                local_file.name,
                f.read()
            ).execute_query()

        logger.info(f"Uploaded {local_file.name}")

    ###########################################################################
    # UPLOAD MULTIPLE FILES
    ###########################################################################

    def upload_folder(
        self,
        local_folder: Path,
        folder_url: str
    ):

        count = 0

        for file in local_folder.iterdir():

            if not file.is_file():
                continue

            self.upload_file(
                file,
                folder_url
            )

            count += 1

        logger.info(f"{count} files uploaded.")

    ###########################################################################
    # DELETE FILE
    ###########################################################################

    def delete_file(
        self,
        server_relative_url: str
    ):

        file = self.ctx.web.get_file_by_server_relative_url(
            server_relative_url
        )

        file.delete_object()

        self.ctx.execute_query()

        logger.info(f"Deleted {server_relative_url}")

    ###########################################################################
    # CREATE FOLDER IF REQUIRED
    ###########################################################################

    def ensure_folder(self, folder_url: str):

        try:

            folder = self.ctx.web.get_folder_by_server_relative_url(
                folder_url
            )

            self.ctx.load(folder)

            self.ctx.execute_query()

            return folder

        except Exception:

            parent = "/".join(folder_url.split("/")[:-1])

            name = folder_url.split("/")[-1]

            parent_folder = self.ctx.web.get_folder_by_server_relative_url(
                parent
            )

            new_folder = parent_folder.folders.add(name)

            self.ctx.execute_query()

            logger.info(f"Created folder {folder_url}")

            return new_folder

    ###########################################################################
    # DOWNLOAD ALL WEEKLY REPORTS
    ###########################################################################

    def download_weekly_reports(
        self,
        reports_folder: str,
        local_folder: Path
    ) -> List[Path]:
        """
        Downloads every workbook from the Weekly Reports folder.
        The report_selector.py module decides which one to use.
        """

        return self.download_folder(
            reports_folder,
            local_folder
        )


###############################################################################
# TEST
###############################################################################

if __name__ == "__main__":

    from auth import get_context

    SITE = "https://YOURTENANT.sharepoint.com/sites/Test"

    ctx = get_context(SITE)

    sp = SharePointManager(ctx)

    files = sp.list_files(
        "Shared Documents/Weekly Reports"
    )

    for file in files:
        print(file.properties["Name"])