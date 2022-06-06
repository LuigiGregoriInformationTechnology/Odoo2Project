import fnmatch, os, pythoncom, sys, win32com.client
import os
import sys
import argparse
import logging
import psycopg2
import odoorpc

try:
    from dev_settings import (
        ODOO_HOST,
        ODOO_PORT,
        ODOO_DB,
        ODOO_PROTOCOL,
        ODOO_USER,
        ODOO_PASSWORD,
        CONN_STRING,
    )
except ImportError:
    print("Missing dev_settings.py")
    sys.exit(-1)


try:
    filename = os.path.abspath(__file__)[:-3]
    with open(f"{filename}.mpp") as f:
        sql = f.read()
except FileNotFoundError:
    logging.error("Missing %s.sql", filename)
    sys.exit(-1)


pjApp= win32com.client.gencache.EnsureDispatch("MSProject.Application")

pjApp.Visible = True 
pjApp.FileNew 
pjApp.ActiveProject.Tasks.Add ("Hang clocks") 
pjApp.FileSaveAs (filename )
pjApp.FileClose 
pjApp.Quit 