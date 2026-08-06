"""
auth.py

Authentication helper for SharePoint Online using an Azure App Registration.

Author: Freedom Mukomana
"""

import os
from functools import lru_cache

from dotenv import load_dotenv
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.client_context import ClientContext

###############################################################################
# LOAD ENVIRONMENT VARIABLES
###############################################################################

load_dotenv()

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")

###############################################################################
# VALIDATE CONFIGURATION
###############################################################################

def validate_configuration():
    """
    Ensure all required environment variables are available.
    """

    missing = []

    if not TENANT_ID:
        missing.append("TENANT_ID")

    if not CLIENT_ID:
        missing.append("CLIENT_ID")

    if not CLIENT_SECRET:
        missing.append("CLIENT_SECRET")

    if missing:
        raise RuntimeError(
            "Missing environment variables:\n"
            + "\n".join(missing)
        )

###############################################################################
# CLIENT CREDENTIAL
###############################################################################

@lru_cache(maxsize=1)
def get_credentials():
    """
    Returns a cached ClientCredential object.
    """

    validate_configuration()

    return ClientCredential(
        CLIENT_ID,
        CLIENT_SECRET
    )

###############################################################################
# SHAREPOINT CONTEXT
###############################################################################

def get_context(site_url):
    """
    Authenticate against a SharePoint site.

    Parameters
    ----------
    site_url : str
        Example:
        https://contoso.sharepoint.com/sites/Johannesburg

    Returns
    -------
    ClientContext
    """

    credentials = get_credentials()

    context = ClientContext(site_url).with_credentials(
        credentials
    )

    # Test the connection
    context.load(context.web)
    context.execute_query()

    return context

###############################################################################
# CONNECTION TEST
###############################################################################

def test_connection(site_url):
    """
    Tests whether authentication succeeds.
    """

    try:

        ctx = get_context(site_url)

        print("=" * 60)
        print("Connection Successful")
        print("=" * 60)
        print("Site:", ctx.web.properties["Title"])
        print("URL :", site_url)
        print("=" * 60)

        return True

    except Exception as ex:

        print("=" * 60)
        print("Connection Failed")
        print("=" * 60)
        print(ex)

        return False


###############################################################################
# TEST
###############################################################################

if __name__ == "__main__":

    SITE = "https://yourtenant.sharepoint.com/sites/Management"

    test_connection(SITE)